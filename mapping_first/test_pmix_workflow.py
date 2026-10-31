import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
import openpyxl.styles
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pmix_workflow as workflow


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
        self.assertEqual(len(workflow.CATEGORIES), 11)
        for category in ("Sandwiches", "Appetizers", "Miscellaneous"):
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
            ("Special Snacks", "Crispy Chicken Wings (6 pc)", "", "Add-Ons"),
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
            ("Special Snacks", "Karaage Chicken", "Japanese Eats"),
            ("Sides", "Deep Fried Chicken Gyoza", "Japanese Eats"),
            ("Basque Cheesecake", "Original Basque Cheesecake - Slice", "Cheesecakes"),
            ("Basque Cheesecake", "Warm Brownie with Ice Cream Scoop", "Pastries"),
            ("Fluffy Pancakes", "Signature Pancakes", "Pancakes"),
            ("Pancakes", "Special Pancake Combo (come with a free ice cream)", "Promotions"),
            ("Sweet Croffles", "Oreo Croffles", "Pastries"),
            ("Extra Sweet", "Affogato", ""),
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

    def test_mapping_dropdown_covers_every_allowed_category(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent) as directory:
            output = Path(directory) / "mapping.xlsx"
            result = workflow.make_mapping(self.small_workbook(directory), output)
            self.assertEqual(result["mapping_rows"], 2)
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


if __name__ == "__main__":
    unittest.main()
