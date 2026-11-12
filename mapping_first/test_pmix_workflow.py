import datetime as dt
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
import openpyxl.styles
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pmix_workflow as workflow

# Tests build small workbooks from real item names; the live reviewer decisions must not apply.
workflow.REVIEW_FILE = Path(__file__).resolve().parent / "no-review-decisions-in-tests.json"
# Likewise the names reviewers added to the real Menu Items List.
workflow.MENU_ADDITIONS_FILE = Path(__file__).resolve().parent / "no-menu-additions-in-tests.json"


class MappingWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.source = pd.DataFrame([
            {"POS Sheet": "TOAST", "Source Group": "Pancakes", "Source Item": "Signature Pancakes",
             "Location": "Store A", "Date": pd.Timestamp("2026-08-03"), "Qty": 2, "Source Row": 2},
            {"POS Sheet": "TOAST", "Source Group": "Pancakes", "Source Item": "Signature Pancakes",
             "Location": "Store B", "Date": pd.Timestamp("2026-10-25"), "Qty": -1, "Source Row": 3},
            {"POS Sheet": "SQUARE", "Source Group": "", "Source Item": "D",
             "Location": "Store C", "Date": pd.Timestamp("2026-10-26"), "Qty": 0, "Source Row": 2},
        ])
        self.source["Mapping ID"] = [
            workflow.mapping_id(*(row[c] for c in workflow.KEY_COLUMNS))
            for row in self.source.to_dict("records")
        ]
        self.reference = pd.DataFrame([
            {"Reference Row": 5, "Reference Top": "All Pancakes", "Category": "Pancakes",
             "Menu": "Pancakes", "Menu group": "Pancakes", "Item": "Signature Pancakes"},
        ])
        self.mapping = workflow.propose_mapping(self.source, self.reference)

    def test_map_covers_all_rows_and_blank_groups(self):
        self.assertEqual(len(self.mapping), 2)
        self.assertEqual(self.mapping["Rows"].sum(), 3)
        self.assertEqual(self.mapping["Net Qty"].sum(), 1)
        self.assertEqual(self.mapping["Review Status"].unique().tolist(), ["Pending"])
        self.assertEqual(self.mapping.iloc[1]["Category"], "")

    def test_report_uses_edited_saved_mapping_not_classifier(self):
        changed = self.mapping.copy()
        changed.loc[0, ["Category", "Menu", "Menu group", "Item"]] = [
            "Promotions", "Reviewed menu", "Reviewed group", "Reviewed product",
        ]
        result = workflow.apply_mapping(self.source, changed, draft=True)
        self.assertEqual(result["Category"].tolist(), ["Promotions", "Promotions", ""])
        self.assertEqual(result.iloc[0]["Item"], "Reviewed product")
        self.assertEqual(result["Qty"].tolist(), [2, -1, 0])
        self.assertEqual(len(result), len(self.source))

    def test_final_blocks_pending(self):
        with self.assertRaisesRegex(ValueError, "Pending"):
            workflow.apply_mapping(self.source, self.mapping, draft=False)

    def test_approved_plus_explicit_exclusion_preserves_all_rows(self):
        changed = self.mapping.copy()
        changed.loc[0, "Review Status"] = "Approved"
        changed.loc[1, ["Review Status", "Review Notes"]] = ["Excluded", "Confirmed register test"]
        result = workflow.apply_mapping(self.source, changed, draft=False)
        self.assertEqual(result.iloc[2]["Exception Reason"], "Confirmed register test")
        self.assertEqual(len(result), 3)

    def test_bad_mapping_rejected(self):
        for column, value, message in [
            ("Category", "Invented", "Invalid Category"),
            ("Source Item", "Changed identity", "identity"),
            ("Review Status", "Maybe", "Invalid Review Status"),
            ("Menu", "", "require Menu"),
        ]:
            with self.subTest(column=column):
                bad = self.mapping.copy()
                bad.loc[0, column] = value
                with self.assertRaisesRegex(ValueError, message):
                    workflow.apply_mapping(self.source, bad, draft=True)

    def test_allowed_categories_include_the_business_additions(self):
        self.assertEqual(len(workflow.CATEGORIES), 12)
        for category in ("Sandwiches", "Appetizers", "Miscellaneous", "Combos"):
            self.assertIn(category, workflow.CATEGORIES)

    def test_business_rules(self):
        for group, item, draft, expected in [
            ("Sandwiches", "Ham and Cheese Sandwich", "Fluffy All Day", "Sandwiches"),
            ("Savory Croffles", "BELT Croffle", "Fluffy All Day", "Sandwiches"),
            ("Savory Croffles", "BLT", "Fluffy All Day", "Sandwiches"),
            ("Crunchy Croissant", "Ham and Cheese Croissant", "Fluffy All Day", "Sandwiches"),
            ("Appetizers", "Crispy Calamari", "", "Appetizers"),
            ("Appetizer", "Fries", "Add-Ons", "Appetizers"),
            ("Ice Cream", "Vanilla Ice Cream", "", "Add-Ons"),
            ("Ice Cream", "Single Scoop", "", "Add-Ons"),
            ("Soft Serve", "Vanilla Soft Serve", "", "Add-Ons"),
            ("Sweet Pancakes", "Ice Cream (1)", "Pancakes", "Add-Ons"),
            ("Mango Pudding", "Mango Pudding", "Promotions", "Pastries"),
            ("Sweets", "Tiramisu Cake", "", "Pastries"),
            ("None", "Chocolate Cake (Slice)", "Drinks", "Pastries"),
            ("Small Bites", "Spicy Yuzu Chicken Bites (8 pcs)", "Japanese Eats", "Add-Ons"),
            ("Special Snacks", "Crispy Chicken Wings (6 pc)", "", "Appetizers"),
            ("Sides", "Chicken Wings (10)", "Add-Ons", "Appetizers"),
            ("Sides", "Deep Fried Chicken Gyoza", "Japanese Eats", "Appetizers"),
            ("Sides", "Deep Fried Edamame Gyoza", "Japanese Eats", "Appetizers"),
            ("Special Snacks", "Karaage Chicken", "Japanese Eats", "Appetizers"),
            ("Top 10 Picks", "Karaage Chicken", "Japanese Eats", "Appetizers"),
            ("", "Karrage Chicken", "Japanese Eats", "Appetizers"),
            ("Sides", "Takoyaki", "Japanese Eats", "Appetizers"),
            ("", "3 Chicken Strips", "Add-Ons", "Add-Ons"),
            ("Extra Sweet", "Affogato", "", "Pastries"),
            ("", "Single Pancake", "Pancakes", "Add-Ons"),
            ("", "1 Plain Pancake", "Pancakes", "Add-Ons"),
            ("", "Gift Card", "", "Miscellaneous"),
            ("?? MUST SELECT FOR UTENSILS ??", "Need Utensils", "", "Miscellaneous"),
            ("FREE ICE CREAM", "Signature Pancakes", "Pancakes", "Promotions"),
            ("Buy Any Chicken Bowl Get Refresher $1.99", "Teriyaki Chicken Plate",
             "Japanese Eats", "Promotions"),
        ]:
            with self.subTest(item=item):
                self.assertEqual(workflow.business_rule(group, item, draft)[0], expected)

    def test_business_rules_leave_other_products_alone(self):
        for group, item, draft in [
            ("Fluffy Chicken Bowl", "Teriyaki Chicken Bowl", "Japanese Eats"),
            ("Chicken Feast", "Shareable chicken 8 Pieces", "Japanese Eats"),
            ("Chicken Feast", "8 Pieces Karaage", "Japanese Eats"),
            ("Fluffy Chicken Bowl", "Teriyaki Karaage Don", "Japanese Eats"),
            ("Drinks", "Affogato Float", "Drinks"),
            ("Basque Cheesecake", "Original Basque Cheesecake - Slice", "Cheesecakes"),
            ("Basque Cheesecake", "Warm Brownie with Ice Cream Scoop", "Pastries"),
            ("Fluffy Pancakes", "Signature Pancakes", "Pancakes"),
            ("Pancakes", "Special Pancake Combo (come with a free ice cream)", "Promotions"),
            ("Sweet Croffles", "Oreo Croffles", "Pastries"),
            ("Espresso Latte", "Fluffy Pumpkin Spice Tiramisu Latte", "Drinks"),
        ]:
            with self.subTest(item=item):
                self.assertEqual(workflow.business_rule(group, item, draft), (draft, ""))

    def test_custom_amount_is_an_intentional_exclusion(self):
        source = self.source.iloc[[0]].copy()
        source["Source Item"], source["Source Group"] = "Custom Amount", "None"
        source["POS Sheet"] = "SQUARE"
        source["Mapping ID"] = workflow.mapping_id("SQUARE", "None", "Custom Amount")
        row = workflow.propose_mapping(source, self.reference).iloc[0]
        self.assertEqual((row["Category"], row["Review Status"]), ("", "Excluded"))
        result = workflow.apply_mapping(source, pd.DataFrame([row]), draft=True)
        self.assertIn("intentionally", result.iloc[0]["Exception Reason"])

    def test_cheena_answers_are_dated_in_the_evidence(self):
        source = self.source.iloc[[0]].copy()
        source["Source Item"], source["Source Group"] = "Affogato", "Extra Sweet"
        source["POS Sheet"] = "TOAST"
        source["Mapping ID"] = workflow.mapping_id("TOAST", "Extra Sweet", "Affogato")
        row = workflow.propose_mapping(source, self.reference).iloc[0]
        self.assertEqual((row["Category"], row["Review Status"]), ("Pastries", "Pending"))
        self.assertEqual(row["Evidence"], "Business decision 2026-11-05")

    def test_changed_category_drops_the_old_reference_path(self):
        reference = pd.DataFrame([
            {"Reference Row": 1, "Reference Top": "All Fluffy All Day", "Category": "Fluffy All Day",
             "Menu": "Lunch", "Menu group": "Croissant Sandwiches", "Item": "BLT Sandwich"},
            {"Reference Row": 2, "Reference Top": "All Promotions", "Category": "Promotions",
             "Menu": "Special Promotions", "Menu group": "Mango Pudding", "Item": "Mango Pudding"},
        ])
        source = pd.concat([self.source.iloc[[0]]] * 2, ignore_index=True)
        source["Source Item"] = ["BLT Sandwich", "Mango Pudding"]
        source["Source Group"] = ["Lunch Sandwich", "Pastries"]
        source["Mapping ID"] = [workflow.mapping_id("TOAST", g, i) for g, i in
                                zip(source["Source Group"], source["Source Item"])]
        mapping = workflow.propose_mapping(source, reference).set_index("Source Item")
        self.assertEqual(mapping.loc["BLT Sandwich", ["Category", "Menu", "Menu group"]].tolist(),
                         ["Sandwiches", "Sandwiches", "Croissant Sandwiches"])
        self.assertEqual(mapping.loc["Mango Pudding", ["Category", "Menu", "Menu group"]].tolist(),
                         ["Pastries", "Pastries", "Pastries"])

    def small_workbook(self, directory):
        book = openpyxl.Workbook()
        toast = book.active
        toast.title = "TOAST"
        toast.append(["Menu Group", "Menu Item", "Sent Date", "Qty", "Location"])
        toast.append(["Sandwiches", "BLT Sandwich", dt.datetime(2026, 8, 3), 1, "Fluffy Fluffy - Tucker"])
        square = book.create_sheet("SQUARE")
        square.append(["Category", "Item", "Date", "Qty", "Location"])
        square.append(["Coffee", "Latte", dt.datetime(2026, 8, 4), 2, "Fluffy Fluffy Tucker GA "])
        menu = book.create_sheet("Menu Category")
        for level, name in ((0, "All Pancakes"), (1, "Pancakes"), (3, "Pancakes"), (4, "Signature")):
            menu.append([None, None, None, None, name])
            menu.cell(menu.max_row, 5).alignment = openpyxl.styles.Alignment(indent=level)
        path = Path(directory) / "source.xlsx"
        book.save(path)
        return path

    def test_tucker_rename_keeps_the_source_label(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            _, source = workflow.load_sources(self.small_workbook(directory))
        self.assertEqual(source["Location"].unique().tolist(), ["Fluffy Fluffy - Tucker"])
        self.assertEqual(source["Source Location"].nunique(), 2)

    def small_menu_list(self, directory):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = workflow.MENU_LIST_SHEET
        for value, bold in (
            ("MENU", True), ("SANDWICHES", True), ("BLT Sandwich", False), ("Ham and Cheese Sandwich", False),
            ("DRINKS - COFFEE", True), ("Latte ", False), ("Crème Brulee", False),
            ("UNCATEGORIZED", True), ("LATTE", False), ("Cafe Mocha", False), ("Ham & Cheese Sandwich", False),
        ):
            sheet.append([value])
            sheet.cell(sheet.max_row, 1).font = openpyxl.styles.Font(bold=bold)
        path = Path(directory) / "menu list.xlsx"
        book.save(path)
        return path

    def menu_list(self):
        return workflow.MenuList([
            ("SANDWICHES", "BLT Sandwich"), ("SANDWICHES", "Ham and Cheese Sandwich"),
            ("DRINKS", "Latte"), ("PASTRIES", "Crème Brulee"), ("APPETIZER", "Tater Tots"),
            ("ADD-ONS", "Tater Tots"),
            (workflow.UNSORTED_SECTION, "LATTE"), (workflow.UNSORTED_SECTION, "Cafe Mocha"),
        ])

    def test_menu_list_reads_sections_from_bold_cells_and_tidies_names(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            entries = workflow.read_menu_list(self.small_menu_list(directory))
        self.assertEqual(entries[:3], [("SANDWICHES", "BLT Sandwich"),
                                       ("SANDWICHES", "Ham and Cheese Sandwich"),
                                       ("DRINKS - COFFEE", "Latte")])
        self.assertIn((workflow.UNSORTED_SECTION, "Cafe Mocha"), entries)
        self.assertNotIn("MENU", [name for _, name in entries])

    def test_listed_names_take_the_list_spelling(self):
        menu = self.menu_list()
        for source, expected, status in [
            ("BLT Sandwich", "BLT Sandwich", workflow.LISTED),
            ("Ham & Cheese Sandwich", "Ham and Cheese Sandwich", workflow.LISTED_CLEANED),
            ("  blt   sandwich ", "BLT Sandwich", workflow.LISTED_CLEANED),
            ("BLT  Sandwich", "BLT Sandwich", workflow.LISTED_CLEANED),  # extra space is a cleaning
            ("Creme Brulee", "Crème Brulee", workflow.LISTED_CLEANED),
            ("#Latte", "Latte", workflow.LISTED_CLEANED),
            ("LATTE", "Latte", workflow.LISTED_CLEANED),  # curated spelling beats UNCATEGORIZED
        ]:
            with self.subTest(source=source):
                name, got, match = workflow.check_menu_name(menu, {}, "id", source, source)
                self.assertEqual((name, got), (expected, status))
                self.assertIn("[", match)

    def test_a_name_on_two_sections_reports_both(self):
        _, _, match = workflow.check_menu_name(self.menu_list(), {}, "id", "Tater Tots", "Tater Tots")
        self.assertEqual(match, "Tater Tots [APPETIZER / ADD-ONS]")

    def test_uncategorized_only_is_noted_but_not_respelled(self):
        name, status, match = workflow.check_menu_name(
            self.menu_list(), {}, "id", "cafe  mocha", "cafe  mocha")
        self.assertEqual((name, status, match), ("cafe mocha", workflow.UNSORTED_ONLY, "Cafe Mocha"))

    def test_unlisted_names_keep_their_own_spelling_and_get_a_note(self):
        name, status, match = workflow.check_menu_name(
            self.menu_list(), {}, "id", "Bleh ", "Bleh ")
        self.assertEqual((name, status, match), ("Bleh", workflow.NOT_LISTED, ""))

    def test_judged_close_match_renames_and_a_suggestion_does_not(self):
        menu = self.menu_list()
        overrides = {
            "typo": {"action": "rename", "list_name": "Latte", "reason": "typo"},
            "maybe": {"action": "suggest", "list_name": "BLT Sandwich", "reason": "unsure"},
        }
        name, status, match = workflow.check_menu_name(menu, overrides, "typo", "Lattte", "Lattte")
        self.assertEqual((name, status), ("Latte", workflow.CLOSE_MATCH))
        self.assertIn("Latte [DRINKS]", match)
        name, status, match = workflow.check_menu_name(menu, overrides, "maybe", "BLT Wrap", "BLT Wrap")
        self.assertEqual((name, status), ("BLT Wrap", workflow.NOT_LISTED))
        self.assertTrue(match.startswith("Possible match, not applied: BLT Sandwich"))

    def test_override_for_a_name_no_longer_listed_fails_loudly(self):
        overrides = {"id": {"action": "rename", "list_name": "Gone", "reason": "x"}}
        with self.assertRaisesRegex(ValueError, "no longer in the Menu Items List"):
            workflow.check_menu_name(self.menu_list(), overrides, "id", "Gonee", "Gonee")

    def test_without_a_menu_list_names_and_columns_are_untouched(self):
        self.assertEqual(self.mapping["Menu List Status"].unique().tolist(), [""])
        self.assertEqual(self.mapping.iloc[0]["Item"], "Signature Pancakes")

    def test_mapping_applies_the_menu_list_to_item_names(self):
        source = self.source.iloc[[0]].copy()
        source["Source Item"] = "signature  pancakes"
        source["Mapping ID"] = workflow.mapping_id("TOAST", "Pancakes", "signature  pancakes")
        menu = workflow.MenuList([("PANCAKES", "Signature Pancakes")])
        row = workflow.propose_mapping(source, self.reference, menu, {}).iloc[0]
        self.assertEqual((row["Item"], row["Menu List Status"]), ("Signature Pancakes", workflow.LISTED_CLEANED))
        self.assertEqual(row["Source Item"], "signature  pancakes")  # identity is never edited

    def test_name_overrides_file_is_well_formed(self):
        overrides = workflow.load_name_overrides()
        for ident, pick in overrides.items():
            with self.subTest(ident=ident):
                self.assertRegex(ident, r"^[0-9a-f]{64}$")
                self.assertIn(pick["action"], ("rename", "suggest"))
                self.assertTrue(pick["list_name"].strip() and pick["reason"].strip())
        if workflow.MENU_LIST_FILE.exists():
            menu = workflow.MenuList(workflow.read_menu_list(workflow.MENU_LIST_FILE))
            missing = {p["list_name"] for p in overrides.values()
                       if workflow.normalized(p["list_name"]) not in menu.names}
            self.assertEqual(missing, set())

    def test_prepare_builds_and_a_failed_build_leaves_nothing_behind(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            source = self.small_workbook(directory)
            mapping = Path(directory) / "mapping.xlsx"
            workflow.make_mapping(source, mapping, self.small_menu_list(directory))
            output = Path(directory) / "report.xlsx"
            audit = workflow.make_report(source, mapping, output, None, None,
                                         draft=True, all_dates=True, native=False)
            self.assertTrue(output.exists())
            self.assertEqual(audit["Selected net Qty"], audit["Pivot net Qty"] + audit["Exception net Qty"])
            failing = Path(directory) / "failing.xlsx"
            real = workflow.write_table
            workflow.write_table = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom"))
            try:
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    workflow.make_report(source, mapping, failing, None, None,
                                         draft=True, all_dates=True, native=False)
            finally:
                workflow.write_table = real
            self.assertFalse(failing.exists())
            self.assertFalse(failing.with_name("failing.building.xlsx").exists())

    def test_list_section_follows_the_list_and_the_entry_category(self):
        menu = self.menu_list()
        for item, category, expected in [
            ("BLT Sandwich", "Sandwiches", "SANDWICHES"),
            ("latte", "Drinks", "DRINKS"),                  # curated section beats UNCATEGORIZED
            ("Tater Tots", "Appetizers", "APPETIZER"),      # listed twice: follow the category
            ("Tater Tots", "Add-Ons", "ADD-ONS"),
            ("Tater Tots", "Drinks", "APPETIZER"),          # no match: first section on the list
            ("Cafe Mocha", "Drinks", workflow.UNSORTED_SECTION),
            ("Bleh", "", workflow.NOT_LISTED),
        ]:
            with self.subTest(item=item, category=category):
                self.assertEqual(workflow.list_section(menu, item, category), expected)

    def test_menu_list_layout_adds_sections_and_keeps_totals(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            source = self.small_workbook(directory)
            mapping = Path(directory) / "mapping.xlsx"
            workflow.make_mapping(source, mapping, self.small_menu_list(directory))
            output = Path(directory) / "list report.xlsx"
            audit = workflow.make_report(source, mapping, output, None, None, draft=True,
                                         all_dates=True, native=False, layout="menu-list")
            data = pd.read_excel(output, sheet_name="PMIXData", dtype=object)
        self.assertEqual(audit["Row layout"], "Menu List Section > Item")
        self.assertEqual(dict(zip(data["Item"], data[workflow.LIST_SECTION])),
                         {"BLT Sandwich": "SANDWICHES", "Latte": "DRINKS - COFFEE"})
        self.assertEqual(audit["Selected net Qty"], audit["Pivot net Qty"] + audit["Exception net Qty"])

    def test_mapping_dropdown_covers_every_allowed_category(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            output = Path(directory) / "mapping.xlsx"
            result = workflow.make_mapping(self.small_workbook(directory), output,
                                           self.small_menu_list(directory))
            self.assertEqual(result["mapping_rows"], 2)
            self.assertEqual(result["menu_list_status"], {workflow.LISTED: 2})
            book = openpyxl.load_workbook(output)
            try:
                listed = [row[0].value for row in book["Allowed Categories"].iter_rows(min_row=2)]
                rules = book["Menu Category"].data_validations.dataValidation
                self.assertEqual(listed, workflow.CATEGORIES)
                self.assertIn(f"$A${len(listed) + 1}", rules[0].formula1)
            finally:
                book.close()

    def test_place_sets_caption_geometry_and_free_floating(self):
        class FakeShape:
            Placement = None

        class FakeSlicer:
            Shape = FakeShape()

        slicer = FakeSlicer()
        workflow.place(slicer, "Sale date window", 1, 2, 3, 4)
        self.assertEqual((slicer.Caption, slicer.Top, slicer.Left, slicer.Width, slicer.Height),
                         ("Sale date window", 1, 2, 3, 4))
        self.assertEqual(slicer.Shape.Placement, 3)

    def test_missing_and_duplicate_mapping_rejected(self):
        with self.assertRaisesRegex(ValueError, "missing"):
            workflow.apply_mapping(self.source, self.mapping.iloc[:1], draft=True)
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            workflow.apply_mapping(self.source, pd.concat([self.mapping, self.mapping]), draft=True)

    def test_window_is_inclusive_and_rejects_reversed_or_empty(self):
        selected = workflow.select_window(self.source, "2026-08-03", "2026-10-25")
        self.assertEqual(selected["Source Row"].tolist(), [2, 3])
        with self.assertRaisesRegex(ValueError, "later"):
            workflow.select_window(self.source, "2026-10-25", "2026-08-03")
        with self.assertRaisesRegex(ValueError, "No source"):
            workflow.select_window(self.source, "2027-01-01", "2027-01-02")

    def test_draft_hints_do_not_repeat_known_bad_rule_precedence(self):
        self.assertEqual(workflow.item_hint("Brunch", "Jalapenos Maple Chicken and Pancakes"),
                         "Fluffy All Day")
        self.assertEqual(workflow.item_hint("Drinks", "Cookies and Cream Frappe"), "Drinks")
        self.assertEqual(workflow.item_hint("Pastry", "#Chocolate Roll Cake"), "Pastries")
        self.assertEqual(workflow.item_hint("", "D"), "")

    def test_savory_filled_croissants_are_not_pastries(self):
        for item in (
            "Egg and Cheese Croissant",
            "Turkey and Cheese Croissant",
            "Double Egg and Cheese Croissant",
            "Sausage and Cheese Croissant",
            "Ham and Cheese Croissant",
        ):
            with self.subTest(item=item):
                self.assertEqual(workflow.item_hint("Crunchy Croissant", item), "Fluffy All Day")
        self.assertEqual(workflow.item_hint("Pastries", "Toasted Croissant"), "Pastries")

    def test_cheesecake_identity_beats_tea_flavor(self):
        self.assertEqual(
            workflow.item_hint("Basque Cheesecake", "Ruby Black Tea Basque Cheesecake - Slice"),
            "Cheesecakes",
        )

    def test_explicit_pancake_and_drink_bundles_are_not_drinks(self):
        for group, item in (
            ("Ube Pancake + $3 Ube Latte", "Ube Pancakes + $3 Ube Latte"),
            ("Signature Pancakes + Frappe", "Signature Pancakes + Frappe"),
        ):
            with self.subTest(item=item):
                self.assertEqual(workflow.item_hint(group, item), "Promotions")

    def test_explicit_free_cheesecake_offer_is_not_a_cheesecake_product(self):
        self.assertEqual(
            workflow.item_hint("FREE CHEESECAKE",
                               "Buy 2 Pancakes and Get A Free Slice of Cheesecake"),
            "Promotions",
        )

    def test_excel_preserves_time_datetime_boolean_and_literal_none(self):
        frame = pd.DataFrame({
            "Mixed Time": [dt.time(12, 34, 56), dt.datetime(2026, 8, 24, 9, 12)],
            "Void?": [False, True], "Group": ["None", ""], "Qty": [-1, 0],
        })
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            path = Path(directory) / "typed.xlsx"
            with pd.ExcelWriter(path, engine="xlsxwriter",
                                datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
                workflow.write_table(writer, "Source", frame)
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                rows = list(book["Source"].iter_rows(values_only=True))
                self.assertEqual(rows[1], (dt.time(12, 34, 56), False, "None", -1))
                self.assertEqual(rows[2], (dt.datetime(2026, 8, 24, 9, 12), True, "", 0))
            finally:
                book.close()


class NameAndAmountTests(unittest.TestCase):
    """Cheena's test run of 2026-11-12: '#' in item names, blank Menu, and the amount columns."""

    def test_hash_marker_is_dropped_from_item_names(self):
        for raw in ("#Double Shot Espresso", "# Nutella Hazelnut Latte", "##Mocha ", "Item #2"):
            with self.subTest(raw=raw):
                cleaned = workflow.clean_item_name(raw)
                self.assertFalse(cleaned.startswith("#") or cleaned.startswith(" "))
        self.assertEqual(workflow.clean_item_name("#Double  Shot Espresso"), "Double Shot Espresso")
        self.assertEqual(workflow.clean_item_name("Item #2"), "Item #2")  # only a leading marker

    def test_unlisted_items_lose_the_hash_too(self):
        menu_list = workflow.MenuList([("DRINKS", "Latte")])
        name, status, _ = workflow.check_menu_name(menu_list, {}, "x", "#Matcha Espresso Latte",
                                                   "#Matcha Espresso Latte")
        self.assertEqual((name, status), ("Matcha Espresso Latte", workflow.NOT_LISTED))

    def test_a_category_alone_is_enough_the_rest_is_filled_in(self):
        mapping = pd.DataFrame([{
            "Source Group": "Specialty Latte", "Source Item": "#Raspberry Hot Cocoa",
            "Category": "Drinks", "Menu": "", "Menu group": "", "Item": "#Raspberry Hot Cocoa"},
            {"Source Group": "", "Source Item": "Pepsi", "Category": "Drinks", "Menu": "",
             "Menu group": "", "Item": ""},
            {"Source Group": "G", "Source Item": "Custom", "Category": "", "Menu": "",
             "Menu group": "", "Item": ""}])
        done = workflow.complete_names(mapping)
        self.assertEqual(done.loc[0, ["Menu", "Menu group", "Item"]].tolist(),
                         ["Drinks", "Specialty Latte", "Raspberry Hot Cocoa"])
        self.assertEqual(done.loc[1, ["Menu", "Menu group", "Item"]].tolist(),
                         ["Drinks", "Drinks", "Pepsi"])
        self.assertEqual(done.loc[2, ["Menu", "Menu group", "Item"]].tolist(), ["", "", ""])

    def test_unusable_entries_are_all_reported_together(self):
        tests = MappingWorkflowTests("test_map_covers_all_rows_and_blank_groups")
        tests.setUp()
        bad = tests.mapping.copy()
        bad.loc[0, "Category"] = "Invented"
        bad.loc[1, ["Review Status"]] = ["Approved"]
        with self.assertRaises(workflow.MappingError) as caught:
            workflow.apply_mapping(tests.source, bad, draft=True)
        message = str(caught.exception)
        self.assertIn("Invalid Category", message)
        self.assertIn("require one of the allowed categories", message)

    def book_with_amounts(self, directory, gross="", tax=None):
        book = openpyxl.Workbook()
        toast = book.active
        toast.title = "TOAST"
        toast.append(["Menu Group", "Menu Item", "Sent Date", "Qty", *workflow.MONEY_COLUMNS, "Location"])
        toast.append(["Sandwiches", "BLT Sandwich", dt.datetime(2026, 8, 3), 1, gross, None, 10.5, tax,
                      "Fluffy Fluffy - Tucker"])
        square = book.create_sheet("SQUARE")
        square.append(["Category", "Item", "Date", "Qty", "Location"])  # no amount columns at all
        square.append(["Coffee", "Latte", dt.datetime(2026, 8, 4), 2, "Fluffy Fluffy Miami"])
        path = Path(directory) / "amounts.xlsx"
        book.save(path)
        return path

    def test_amount_columns_are_read_and_may_be_blank_or_missing(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            sheets, source = workflow.load_sources(self.book_with_amounts(directory, gross=12.25))
        toast = source[source["POS Sheet"] == "TOAST"].iloc[0]
        self.assertEqual((toast["Gross Sales"], toast["Net Sales"]), (12.25, 10.5))
        self.assertTrue(pd.isna(toast["Discount"]) and pd.isna(toast["Tax"]))
        self.assertTrue(source[source["POS Sheet"] == "SQUARE"][workflow.MONEY_COLUMNS].isna().all().all())
        self.assertEqual(sheets["TOAST"]["Qty"].tolist(), [1])  # the source sheet copy is untouched

    def test_text_in_an_amount_column_is_refused_in_plain_words(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            with self.assertRaisesRegex(ValueError, "'Gross Sales' must contain numbers only"):
                workflow.load_sources(self.book_with_amounts(directory, gross="n/a"))

    def test_report_data_sheet_carries_the_amount_columns(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            source = self.book_with_amounts(directory, gross=12.25)
            mapping = Path(directory) / "mapping.xlsx"
            checks = MappingWorkflowTests("test_map_covers_all_rows_and_blank_groups")
            workflow.make_mapping(source, mapping, checks.small_menu_list(directory),
                                  reference_path=checks.small_workbook(directory))
            frame = pd.read_excel(mapping, sheet_name="Menu Category", dtype=object, keep_default_na=False)
            frame.loc[:, "Review Status"] = "Pending"
            output = Path(directory) / "report.xlsx"
            workflow.make_report(source, mapping, output, None, None, draft=True, all_dates=True,
                                 native=False, layout="menu-list")
            data = pd.read_excel(output, sheet_name="PMIXData")
        columns = list(data.columns)
        self.assertEqual(columns[columns.index("Qty") + 1:columns.index("Qty") + 5],
                         workflow.MONEY_COLUMNS)
        self.assertEqual(data["Gross Sales"].dropna().tolist(), [12.25])


class NameSuggestionTests(unittest.TestCase):
    """Cheena, 2026-11-12: rename items to the Menu Items List automatically where it is safe."""

    LIST = [("PANCAKES", "Apple Caramel Pancakes"), ("FLUFFY ALL DAY - ENTREES", "Eggs Benedict (Bacon)"),
            ("DRINKS - COFFEE", "Cappuccino"), ("DRINKS - COFFEE", "Hot Chocolate"),
            ("SIDES", "2 Eggs"), ("SIDES", "Side Stack (1 piece)"), ("PASTRIES", "Strawberry Cake"),
            ("DRINKS - TEA", "Blueberry Matcha Latte"), ("DRINKS - TEA", "Matcha Blueberry Latte")]

    def menu_list(self):
        return workflow.MenuList(self.LIST)

    def test_same_words_is_high_and_a_typo_is_medium(self):
        menu = self.menu_list()
        self.assertEqual(menu.suggest("Apple Caramel Pancake"), ("Apple Caramel Pancakes", workflow.HIGH_MATCH))
        self.assertEqual(menu.suggest("Bacon Eggs Benedict"), ("Eggs Benedict (Bacon)", workflow.HIGH_MATCH))
        self.assertEqual(menu.suggest("Cappucino"), ("Cappuccino", workflow.MEDIUM_MATCH))

    def test_different_numbers_never_match_and_unknown_items_get_nothing(self):
        menu = self.menu_list()
        self.assertIsNone(menu.suggest("Side Stack 2 Pieces"))
        self.assertIsNone(menu.suggest("Biscoff Coffee"))
        self.assertIsNone(menu.suggest("Blueberry Latte Matcha"))  # two list names share these words

    def mapping(self, status, choice="", item="Apple Caramel Pancake", listed=workflow.NOT_LISTED):
        row = {c: "" for c in workflow.MAPPING_COLUMNS}
        row.update({"Mapping ID": "a", "POS Sheet": "TOAST", "Source Group": "Pancakes",
                    "Source Item": item, "Category": "Pancakes", "Menu": "Pancakes",
                    "Menu group": "Pancakes", "Item": item, "Review Status": status,
                    "Menu List Status": listed, "Menu List Section": workflow.NOT_LISTED})
        frame = pd.DataFrame([row])
        previous = frame.copy()
        for column in workflow.SUGGEST_COLUMNS:
            previous[column] = ""
        previous["Accept Suggestion"] = choice
        return frame, previous

    def test_a_pending_high_match_is_renamed_automatically(self):
        frame, previous = self.mapping("Pending")
        done, changes = workflow.apply_suggestions(frame, self.menu_list(), previous)
        row = done.iloc[0]
        self.assertEqual((row["Item"], row["Accept Suggestion"], row["Menu List Status"], row["Menu List Section"]),
                         ("Apple Caramel Pancakes", "Auto", workflow.CLOSE_MATCH, "PANCAKES"))
        self.assertEqual(changes["Change"].tolist(), [workflow.SUGGESTION_APPLIED])

    def test_a_reviewed_entry_is_only_suggested_until_the_reviewer_says_yes(self):
        frame, previous = self.mapping("Approved")
        done, changes = workflow.apply_suggestions(frame, self.menu_list(), previous)
        self.assertEqual((done.iloc[0]["Item"], done.iloc[0]["Suggested Name"]),
                         ("Apple Caramel Pancake", "Apple Caramel Pancakes"))
        self.assertTrue(changes.empty)
        frame, previous = self.mapping("Approved", choice="Yes")
        done, changes = workflow.apply_suggestions(frame, self.menu_list(), previous)
        self.assertEqual((done.iloc[0]["Item"], done.iloc[0]["Menu List Status"], done.iloc[0]["Menu List Section"]),
                         ("Apple Caramel Pancakes", workflow.RENAMED, "PANCAKES"))

    def test_a_section_the_reviewer_chose_is_kept_when_the_name_is_accepted(self):
        frame, previous = self.mapping("Approved", choice="Yes")
        frame["Menu List Section"] = "ADD-ONS"
        done, _ = workflow.apply_suggestions(frame, self.menu_list(), previous)
        self.assertEqual((done.iloc[0]["Item"], done.iloc[0]["Menu List Section"]),
                         ("Apple Caramel Pancakes", "ADD-ONS"))

    def test_a_count_and_its_unit_agree_with_or_without_a_space(self):
        menu = workflow.MenuList([("SMALL BITES", "Spicy Yuzu Chicken Bites (8pcs)")])
        self.assertEqual(menu.suggest("Spicy Yuzu Chicken Bites (8 pcs)"),
                         ("Spicy Yuzu Chicken Bites (8pcs)", workflow.HIGH_MATCH))
        self.assertIsNone(menu.suggest("Spicy Yuzu Chicken Bites (12 pcs)"))

    def test_no_keeps_the_original_name_even_for_a_pending_high_match(self):
        frame, previous = self.mapping("Pending", choice="No")
        done, changes = workflow.apply_suggestions(frame, self.menu_list(), previous)
        self.assertEqual((done.iloc[0]["Item"], done.iloc[0]["Accept Suggestion"]),
                         ("Apple Caramel Pancake", "No"))
        self.assertTrue(changes.empty)

    def test_a_medium_match_is_never_applied_without_yes(self):
        frame, previous = self.mapping("Pending", item="Cappucino")
        done, _ = workflow.apply_suggestions(frame, self.menu_list(), previous)
        self.assertEqual((done.iloc[0]["Item"], done.iloc[0]["Suggested Name"]), ("Cappucino", "Cappuccino"))

    def test_an_accepted_rename_is_remembered_on_the_next_run(self):
        frame, previous = self.mapping("Approved", choice="Yes")
        first, _ = workflow.apply_suggestions(frame, self.menu_list(), previous)
        again, changes = workflow.apply_suggestions(first, self.menu_list(), first)
        self.assertEqual((again.iloc[0]["Item"], again.iloc[0]["Accept Suggestion"],
                          again.iloc[0]["Suggested Name"]),
                         ("Apple Caramel Pancakes", "Yes", "Apple Caramel Pancakes"))
        self.assertTrue(changes.empty)


class CarryOverTests(unittest.TestCase):
    """A new map run on new data keeps what a reviewer decided in the previous master."""

    def setUp(self):
        tests = MappingWorkflowTests("test_map_covers_all_rows_and_blank_groups")
        tests.setUp()
        self.mapping = tests.mapping  # Signature Pancakes (Pancakes) and D (no category)
        self.pancake = self.mapping.iloc[0]["Mapping ID"]
        self.blank = self.mapping.iloc[1]["Mapping ID"]

    def previous(self, **changes):
        frame = self.mapping.copy()
        for column, value in changes.items():
            frame.loc[0, column] = value
        return frame

    def test_approved_decision_and_its_item_name_are_kept(self):
        previous = self.previous(**{"Category": "Fluffy All Day", "Item": "Signature Pancake",
                                    "Review Status": "Approved", "Evidence": "Cheena review",
                                    "Menu List Status": "Listed", "Menu List Match": "x [PANCAKES]"})
        mapping, earlier, changes = workflow.carry_over(self.mapping, previous)
        row = mapping.set_index("Mapping ID").loc[self.pancake]
        self.assertEqual((row["Category"], row["Item"], row["Review Status"], row["Evidence"]),
                         ("Fluffy All Day", "Signature Pancake", "Approved", "Cheena review"))
        self.assertEqual(row["Menu List Status"], "Listed")
        self.assertEqual(row["Rows"], 2)  # counts come from the new data
        self.assertEqual(changes["Change"].tolist(), [workflow.KEPT])
        self.assertIn("rules now give 'Pancakes'", changes.iloc[0]["Details"])
        self.assertTrue(earlier.empty)

    def test_new_items_are_listed_and_left_pending(self):
        mapping, _, changes = workflow.carry_over(self.mapping, self.mapping.iloc[[0]])
        self.assertEqual(changes.set_index("Mapping ID").loc[self.blank, "Change"], workflow.NEW_ITEM)
        self.assertEqual(mapping.set_index("Mapping ID").loc[self.blank, "Review Status"], "Pending")
        self.assertEqual(len(changes), 1)

    def test_pending_hand_edit_follows_the_rules_and_is_reported(self):
        mapping, _, changes = workflow.carry_over(self.mapping, self.previous(Category="Drinks"))
        self.assertEqual(mapping.set_index("Mapping ID").loc[self.pancake, "Category"], "Pancakes")
        self.assertEqual(changes["Change"].tolist(), [workflow.REBUILT])
        self.assertIn("Category: was 'Drinks', now 'Pancakes'", changes.iloc[0]["Details"])

    def test_unchanged_previous_mapping_reports_nothing(self):
        mapping, earlier, changes = workflow.carry_over(self.mapping, self.mapping)
        pd.testing.assert_frame_equal(mapping, self.mapping)
        self.assertTrue(changes.empty and earlier.empty)

    def test_reviewed_items_not_in_this_data_are_kept_for_later(self):
        gone = self.previous(**{"Mapping ID": "old", "Source Item": "Old Item",
                                "Review Status": "Approved"})
        pending_gone = self.previous(**{"Mapping ID": "old2", "Source Item": "Old Pending"})
        previous = pd.concat([self.mapping, gone.iloc[[0]], pending_gone.iloc[[0]]])
        _, earlier, _ = workflow.carry_over(self.mapping, previous)
        self.assertEqual(earlier["Mapping ID"].tolist(), ["old"])

    def test_invalid_approved_category_is_refused(self):
        previous = self.previous(**{"Category": "Desserts", "Review Status": "Approved"})
        with self.assertRaisesRegex(ValueError, "no valid Category"):
            workflow.carry_over(self.mapping, previous)

    def test_every_approved_entry_without_a_category_is_listed_at_once(self):
        previous = self.mapping.copy()
        previous["Review Status"] = "Approved"
        previous["Category"] = ""
        with self.assertRaises(workflow.MappingError) as caught:
            workflow.carry_over(self.mapping, previous)
        message = str(caught.exception)
        self.assertIn("'Signature Pancakes'", message)
        self.assertIn("'D'", message)  # both entries, not just the first
        self.assertIn("Excluded", message)  # and what to do about it

    def test_read_previous_adds_earlier_decisions_and_current_rows_win(self):
        earlier = self.previous(**{"Category": "Drinks", "Review Status": "Approved"}).iloc[[0]]
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            path = Path(directory) / "previous.xlsx"
            with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
                workflow.write_table(writer, "Menu Category", self.mapping)
                workflow.write_table(writer, workflow.EARLIER_SHEET, earlier)
            previous = workflow.read_previous(path)
        self.assertEqual(len(previous), 2)
        self.assertEqual(previous.set_index("Mapping ID").loc[self.pancake, "Category"], "Pancakes")


class ReviewDecisionTests(unittest.TestCase):
    """Reviewer decisions (review_decisions.json) and the recategorize.py helpers."""

    def setUp(self):
        tests = MappingWorkflowTests("test_map_covers_all_rows_and_blank_groups")
        tests.setUp()
        self.source, self.reference = tests.source, tests.reference
        self.menu = workflow.MenuList([("PANCAKES", "Signature Pancakes"), ("ADD-ONS", "Avocado"),
                                       ("DRINKS - JUICE", "Apple Juice")])
        self.pancake = workflow.mapping_id("TOAST", "Pancakes", "Signature Pancakes")
        self.blank = workflow.mapping_id("SQUARE", "", "D")

    def propose(self, decisions):
        return workflow.propose_mapping(self.source, self.reference, self.menu, {},
                                        decisions).set_index("Mapping ID")

    def test_sections_are_stored_in_the_mapping(self):
        mapping = self.propose({})
        self.assertEqual(mapping.loc[self.pancake, workflow.LIST_SECTION], "PANCAKES")
        self.assertEqual(mapping.loc[self.blank, workflow.LIST_SECTION], workflow.NOT_LISTED)

    def test_decision_moves_renames_and_approves_last(self):
        mapping = self.propose({self.blank: {"status": "Approved", "section": "DRINKS - JUICE",
                                             "item": "Apple Juice", "by": "Cheena 2026-11-06"}})
        row = mapping.loc[self.blank]
        self.assertEqual((row["Category"], row[workflow.LIST_SECTION], row["Item"], row["Review Status"]),
                         ("Drinks", "DRINKS - JUICE", "Apple Juice", "Approved"))
        self.assertEqual(row["Evidence"], "Reviewed by Cheena 2026-11-06")
        self.assertEqual(row["Menu List Status"], workflow.RENAMED)
        self.assertEqual((row["Menu"], row["Menu group"]), ("Drinks", "Drinks"))

    def test_extra_section_and_its_category(self):
        mapping = self.propose({self.pancake: {"status": "Approved", "section": "PROMOTIONS", "by": "x"}})
        self.assertEqual(mapping.loc[self.pancake, "Category"], "Promotions")
        self.assertEqual(mapping.loc[self.pancake, workflow.LIST_SECTION], "PROMOTIONS")

    def test_section_category_keeps_a_fitting_category(self):
        self.assertEqual(workflow.section_category("FLUFFY ALL DAY - SANDWICHES", "Fluffy All Day"),
                         "Fluffy All Day")
        self.assertEqual(workflow.section_category("COMBOS", "Promotions"), "Combos")
        self.assertEqual(workflow.section_category("DRINKS - TEA", "Pancakes"), "Drinks")
        self.assertEqual(workflow.section_category("UNCATEGORIZED", "Pancakes"), "Pancakes")

    def test_canonical_section_reads_reviewer_spellings(self):
        sections = ["FLUFFY ALL DAY - ENTREES", "FLUFFY ALL DAY - SANDWICHES", "SWEET CROFFLES",
                    "APPETIZER", "ADD-ONS", *workflow.EXTRA_SECTIONS]
        for typed, expected in (("Fluffy All Day - Entrée", "FLUFFY ALL DAY - ENTREES"),
                                ("Fluffy All Day -Sandwiches", "FLUFFY ALL DAY - SANDWICHES"),
                                ("Croffles", "SWEET CROFFLES"), ("Appetizer", "APPETIZER"),
                                ("Add-Ons", "ADD-ONS"), ("Promotions", "PROMOTIONS"),
                                ("Desserts", None)):
            with self.subTest(typed=typed):
                self.assertEqual(workflow.canonical_section(typed, sections), expected)

    def test_carry_over_does_not_undo_a_newer_decision(self):
        decisions = {self.pancake: {"status": "Approved", "section": "PROMOTIONS", "by": "x"}}
        mapping = workflow.propose_mapping(self.source, self.reference, self.menu, {}, decisions)
        previous = mapping.copy()
        previous.loc[0, ["Category", "Review Status"]] = ["Pancakes", "Approved"]
        carried, _, changes = workflow.carry_over(mapping, previous, set(decisions))
        self.assertEqual(carried.iloc[0]["Category"], "Promotions")
        self.assertEqual(changes["Change"].tolist(), [workflow.REVIEW_APPLIED])

    def test_recategorize_reads_remarks(self):
        import recategorize as rc
        self.assertEqual(rc.instruction("Rename to Apple Juice"), ("rename", "Apple Juice", ""))
        self.assertEqual(rc.instruction("Raname to Roll Cake"), ("rename", "Roll Cake", ""))
        self.assertEqual(rc.instruction("Rename to Mango Pancakes, if item is under Pancakes Category"),
                         ("rename", "Mango Pancakes", "Pancakes"))
        self.assertEqual(rc.instruction('Remove " (Cold or Hot)" in item name'),
                         ("remove", " (Cold or Hot)", ""))
        self.assertEqual(rc.instruction("Remove # character in the item name"), ("remove", "#", ""))
        self.assertEqual(rc.instruction("Ube Pancakes"), ("rename", "Ube Pancakes", ""))
        self.assertEqual(rc.apply_instruction("Matcha Latte - Cold", "remove", " -Cold"), "Matcha Latte")
        self.assertEqual(rc.apply_instruction("#Add On Ice Cream", "remove", "#"), "Add On Ice Cream")
        self.assertIsNone(rc.apply_instruction("Latte", "remove", " (Hot)"))
        self.assertEqual(rc.apply_instruction("Thai Tea Iced Tea", "remove", " Iced"), "Thai Tea")
        self.assertEqual(rc.apply_instruction("Fluffy Fluffy Latte - Hot", "remove", " -Hot"),
                         "Fluffy Fluffy Latte")

    def test_recategorize_reads_threaded_cell_comments(self):
        import recategorize as rc
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            path = Path(directory) / "commented.xlsx"
            book = openpyxl.Workbook()
            book.active.title = "PMIXData"
            book.active["P3"] = "Pending"
            book.active["P3"].comment = openpyxl.comments.Comment("Comment:\nRename to Strawberry Pancakes", "x")
            book.save(path)
            self.assertEqual(rc.read_cell_comments(path, "PMIXData"), [(3, "Rename to Strawberry Pancakes")])

    def test_recategorize_snaps_typos_but_never_drops_words(self):
        import recategorize as rc
        menu = workflow.MenuList([("X", "Mocha Espresso"), ("X", "Basque Cheesecake - Slice")])
        self.assertEqual(rc.snap("Mocha Expresso", menu), "Mocha Espresso")
        self.assertEqual(rc.snap("Oreo Basque Cheesecake - Slice", menu), "Oreo Basque Cheesecake - Slice")

    def test_menu_additions_go_at_the_end_of_their_section(self):
        entries = [("PANCAKES", "Signature Pancakes"), ("PANCAKES", "Ube Pancakes"),
                   ("DRINKS - TEA", "Matcha")]
        added = workflow.add_menu_names(entries, {"PANCAKES": ["Nutella Pancakes", "ube  pancakes"]})
        self.assertEqual(added, [("PANCAKES", "Signature Pancakes"), ("PANCAKES", "Ube Pancakes"),
                                 ("PANCAKES", "Nutella Pancakes"), ("DRINKS - TEA", "Matcha")])
        self.assertEqual(entries[0], ("PANCAKES", "Signature Pancakes"))  # input is not changed
        with self.assertRaises(ValueError):
            workflow.add_menu_names(entries, {"SALADS": ["Caesar"]})

    def test_added_name_is_listed_and_keeps_its_section(self):
        entries = workflow.add_menu_names([("PANCAKES", "Signature Pancakes")],
                                          {"PANCAKES": ["Nutella Pancakes"]})
        menu = workflow.MenuList(entries)
        name, status, _ = workflow.check_menu_name(menu, {}, "id", "nutella pancakes", "Nutella Pancakes")
        self.assertEqual((name, status), ("Nutella Pancakes", workflow.LISTED_CLEANED))
        self.assertEqual(workflow.list_section(menu, "Nutella Pancakes", "Pancakes"), "PANCAKES")

    def test_list_copy_adds_names_and_leaves_the_original_alone(self):
        import recategorize as rc
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            source = MappingWorkflowTests("test_listed_names_take_the_list_spelling").small_menu_list(directory)
            additions = Path(directory) / "additions.json"
            additions.write_text(json.dumps({"sections": {"SANDWICHES": ["Tuna Sandwich", "BLT sandwich"],
                                                          "DRINKS - COFFEE": ["Mocha"]}}), encoding="utf-8")
            before = source.read_bytes()
            output = Path(directory) / "copy.xlsx"
            with mock.patch.object(workflow, "MENU_ADDITIONS_FILE", additions):
                rc.list_copy(source, output)
            self.assertEqual(source.read_bytes(), before)
            self.assertEqual([e for e in workflow.read_menu_list(output) if e[0] != workflow.UNSORTED_SECTION],
                             [("SANDWICHES", "BLT Sandwich"), ("SANDWICHES", "Ham and Cheese Sandwich"),
                              ("SANDWICHES", "Tuna Sandwich"), ("DRINKS - COFFEE", "Latte"),
                              ("DRINKS - COFFEE", "Crème Brulee"),
                              ("DRINKS - COFFEE", "Mocha")])  # "BLT sandwich" was already listed
            self.assertFalse(openpyxl.load_workbook(output)[workflow.MENU_LIST_SHEET]["A5"].font.b)

    def test_exception_statuses_are_read_by_mapping_id(self):
        import recategorize as rc
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            path = Path(directory) / "report.xlsx"
            pd.DataFrame({"Mapping ID": ["a", "a", "b", "c"],
                          "Review Status": ["Add-Ons", " Add-Ons", "Uncategorized", "Add-Ons"]}
                         ).to_excel(path, sheet_name="Exceptions", index=False)
            self.assertEqual(rc.read_exception_statuses(path),
                             {"a": ["Add-Ons"], "b": ["Uncategorized"], "c": ["Add-Ons"]})
            pd.DataFrame({"x": [1]}).to_excel(path, sheet_name="Other", index=False)
            self.assertEqual(rc.read_exception_statuses(path), {})  # no Exceptions sheet

    def test_unplaced_exception_stays_without_a_category(self):
        import recategorize as rc
        entry = {"POS Sheet": "TOAST", "Source Group": "", "Source Item": "Oat", "Category": ""}
        made = rc.decision(entry, "x", workflow.UNSORTED_SECTION, "Oat", "Pending", "note")
        self.assertEqual((made["category"], made["status"]), ("", "Pending"))
        placed = rc.decision(entry, "x", "ADD-ONS", "Oat", "Approved", "note")
        self.assertEqual(placed["category"], "Add-Ons")

    def test_decisions_already_on_file_are_not_recorded_again(self):
        import recategorize as rc
        old = {"source": "s", "status": "Approved", "category": "Pancakes", "section": "PANCAKES",
               "item": "Ube Pancakes", "by": "Cheena 2026-11-06", "note": "n", "renamed_from": "Ube"}
        same = {**old, "by": "Cheena 2026-11-09", "note": "again"}
        moved = {**old, "category": "Promotions", "section": "PROMOTIONS", "by": "Cheena 2026-11-09"}
        del moved["renamed_from"]
        with mock.patch.object(workflow, "load_review_decisions", return_value={"a": old, "b": old}):
            check = {}
            kept = rc.drop_unchanged({"a": same, "b": moved, "c": {**old, "source": "t"}}, check)
        self.assertEqual(sorted(kept), ["b", "c"])
        self.assertEqual(kept["b"]["renamed_from"], "Ube")  # an earlier rename is not forgotten
        self.assertEqual(len(check["Already on file"]), 1)


if __name__ == "__main__":
    unittest.main()
