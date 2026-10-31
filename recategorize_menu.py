"""Recategorize USA PMix into the 8 Menu Categories and produce Quarterly PMIX Reports for all locations.

Target 8 Menu Categories (from Menu Category Col A and user specification):
1. Pancakes
2. Fluffy All Day
3. Drinks
4. Japanese Eats
5. Cheesecakes
6. Pastries
7. Promotions
8. Add-Ons
(Non-Menu items like Custom Amount, Gift Cards, and Utensils are tracked separately)

Locations:
- TOAST (17 locations): Aurora, Boston, Buford, Cranford, Davis, Doral, Doraville, Duluth,
                        East Cobb, Edison, Elk Grove, Huntsville, NYC, Orlando, Philly,
                        Schaumburg, Tucker
- SQUARE (3 locations): Carrollton, Miami, Tucker GA
Total: 20 locations.
"""

import sys
import unicodedata
import openpyxl
import pandas as pd

SRC = "USA PMix.xlsx"
OUT = "USA PMix - Recategorized.xlsx"

CATEGORIES = [
    "Pancakes",
    "Fluffy All Day",
    "Drinks",
    "Japanese Eats",
    "Cheesecakes",
    "Pastries",
    "Promotions",
    "Add-Ons",
]


def normalize_text(text):
    if pd.isna(text):
        return ""
    s = str(text).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower()


def classify_item(group, item):
    g = normalize_text(group)
    i = normalize_text(item)

    # 0. Non-Menu (POS registers, cash codes, merchandise, utensils, testing)
    if any(k in i for k in [
        "custom amount", "gift card", "utensil", "need ketchup", "t shirt", "condiment",
        "donation", "delivery fee", "testing printer", "private event", "empty plate",
        "difference", "gkuten", "gluten", "hats"
    ]):
        return "Non-Menu"
    if i in [
        "d", "a", "m", "1", "g", "w", "o", "r", "s", "t", "for", "diff", "al",
        "bleh", "bcbnjvx", "byrij", "candy", "coco", "claudia", "open", "pomp",
        "pump", "test", "tip", "oreo"
    ]:
        return "Non-Menu"

    # 1. Promotions
    if any(k in i or k in g for k in [
        "promo", "promotion", "combo", "deal", "follow ig", "google review", "buy any",
        "can dine", "tasty tuesday", "chicken and croffle thursday", "mango pudding",
        "review sftraww"
    ]):
        if not ("anniversary" in g or "celebration" in g or "free ice cream" in g):
            return "Promotions"

    # 2. Add-Ons
    if any(k in i for k in [
        "add bacon", "add tater", "add an egg", "add salmon", "add extra", "add on ice cream",
        "add on slice", "add value", "side bacon", "side eggs", "side of", "extra protein",
        "extra stack", "side stack", "mini hash brown", "hash brown", "hashbrown", "tater tot",
        "tots", "set of eggs", "2 eggs", "1 egg", "sunny-side", "sauce", "suace", "maple syrup",
        "hot sauce", "dressing", "furikake tots", "truffle fries", "fries", "salted butter",
        "guacamole", "whipped cream", "fresh cream", "tiramisu cream", "strawberry sauce",
        "nutella sauce", "nutella suace", "nutellasuace", "strzwberry suace", "strawberry syrup",
        "white rice", "bowl of rice", "extra topping", "extra strawberry", "extra nutella",
        "extra chicken", "extra caramel", "extra fc", "xtra stack", "whip", "side fruit",
        "fruit cup", "fruit side", "scambles eggs", "hollandaise", "berries and fruit",
        "tropical fruit bowl", "melon fruit bowl", "strawberries 4x", "cutted strawberries",
        "banana side", "add avacado", "cheese slice", "2 pieces of salmon", "1 sausage",
        "5 slices of turkey meat", "taro"
    ]):
        if not any(k in i for k in ["sandwich", "benedict", "croffle sandwich", "croissant sandwich", "bowl", "pancake sandwich"]):
            return "Add-Ons"
    if g in ["side orders", "side order", "sides", "side orders ", "a la cart", "add ons", "all add ons", "brunchy brunchy add-ons"]:
        if not any(k in i for k in ["pancake", "waffle", "benedict", "sandwich", "burger", "bowl"]):
            return "Add-Ons"
    if i in [
        "bacon (2 slices)", "bacon (2 pieces)", "egg", "egg (1 pc)", "sausage (2 pieces)",
        "sausage (2 pcs)", "salmon (2 slices)", "avocado (1/2)", "side salad", "salad",
        "mini hash browns", "golden croffle (1 piece)", "croffle (1 pc w/syrup)", "rice",
        "strawberries", "cup of strawberries", "fresh straw", "honey", "jalapeno",
        "jalapenos", "pork bacon (2 pcs)", "turkey bacon (2 pcs)", "side of bacon",
        "strawberry", "strabwerry", "nutella", "fruit", "mango", "banana", "apple", "grape",
        "melon", "maple", "cheese", "strawberrynana", "stawrberrnana", "strawberry + nana",
        "strawberry+nana", "mangonana", "raspberry", "rasberry", "raspberry blackberries",
        "blueberry toppings", "fruit forest", "apple caramel", "bluerry", "tropical fruit bowl",
        "bowl of rice", "cream and strawberries on side", "fruit. cup", "mango banana",
        "melon fruit bowl"
    ]:
        return "Add-Ons"

    # 3. Japanese Eats
    if any(k in i or k in g for k in [
        "chicken bowl", "donburi", "chicken feast", "karaage", "yuzu fire", "sweet heat",
        "teriyaki chicken", "chicken bites", "chicken tender", "tenders", "chicken wings",
        "calamari", "gyoza", "takoyaki", "taiyaki", "mozzarella", "japanese eats",
        "japanese eat", "shareable chicken", "katsu don", "spring roll", "tebders",
        "chicken 8peices", "karrage chicken"
    ]):
        if not any(k in i for k in ["pancake", "croffle", "sandwich", "benedict"]):
            return "Japanese Eats"
    if i in [
        "chicken", "chiken x3", "chickeb x5", "chicken x 6", "chicken x2", "tenders x6",
        "5chicken tenders+ tots", "8 pieces chicken", "12 pieces chicken",
        "spicy yuzu chicken bites (8 pcs)", "spicy yuzu chicken bites (12 pcs)",
        "spicy yuzu chicken bites (30 pcs)", "2 chicken strips", "3 strips chicken", "3 chicken strips"
    ]:
        return "Japanese Eats"

    # 4. Pancakes
    if any(k in i for k in [
        "pancake", "pancakes", "french toast", "souffle pancake", "creme brulee (fluffy all day long)",
        "blueberry cheese (fluffy all day long)", "plain fluffy pancakes", "minis (8 per order)",
        "creme brule", "creme bruelee", "creme brullee", "plain souffle", "sigle stack",
        "pancaake", "panckaes", "chicken pa", "nutella banana (2 layers"
    ]):
        if not any(k in i for k in ["sandwich", "benedict", "croffle", "chicken and pancake", "chicken & pancake"]):
            return "Pancakes"
    if g in ["fluffy pancakes", "pancakes", "sweet pancakes", "souffle pancakes", "souffle pancakes*", "all pancakes", "1 year anniversary celebration", "free ice cream", "fluffy pancakes + $1.99 refresher"]:
        if not any(k in i for k in ["sandwich", "benedict", "croffle", "side", "egg", "bacon", "sausage"]):
            return "Pancakes"

    # 5. Fluffy All Day
    if any(k in i or k in g for k in [
        "benedict", "sandwich", "croissant sandwich", "all day breakfast", "breakfast sandwich",
        "fluffy breakfast", "croffle sandwich", "savory croffle", "savory pancake",
        "fluffy all day", "brunch", "entrees", "savory", "salad", "omelette", "samosa",
        "masala aloo", "avocado toast", "avacado toast", "pizza croffle", "ham & cheese",
        "ham and cheese", "turkey and cheese", "smoked salmon", "blt", "belt",
        "cheese lava nachos", "chicken and croffle", "chicken & croffle", "chicken & croffles",
        "jalapenos maple chicken", "jalapeños maple chicken", "croissant", "avocado & bacon waffle",
        "toasre", "acacado toaste"
    ]):
        if not any(k in i for k in ["sweet croffle", "fruit salad", "cheesecake", "latte", "tea", "frappe", "toasted croissant"]):
            return "Fluffy All Day"

    # 6. Cheesecakes
    if any(k in i or k in g for k in [
        "cheesecake", "cheese cake", "chessecake", "cheescske", "chese cake", "basque",
        "burnt top", "ice cream", "soft serve", "gelato", "icre cream", "ice crea", "icecream"
    ]):
        if not any(k in i for k in ["pancake", "croffle", "latte", "frappe"]):
            return "Cheesecakes"
    if i in [
        "cake", "cake original", "cake choco", "slice of chocolate", "slice of cheescake",
        "slice cheesecake", "vanilla flavor", "single scoop", "double scoop",
        "two scoop vanilla ice cream", "vanilla ice cream", "ice cream scoop (single)",
        "ice cream scoop (double)", "chocolate cake (slice)", "cakee", "original burnt top cheesecake (slice)",
        "slice earl gray"
    ]:
        return "Cheesecakes"

    # 7. Pastries
    if any(k in i or k in g for k in [
        "croffle", "croffles", "macaron", "macarons", "macron", "roll cake", "rollcake",
        "roll cakes", "pastry", "pastries", "mousse", "flan", "cookie", "cookies",
        "mascarpone", "creme brulee dessert", "fried oreo", "fried banana", "brownie",
        "muffin", "fluffy cloud", "sweet waffles", "waffle", "toasted croissant",
        "cske pop", "caie pop", "cqkr oioo", "peccan brownnie", "roll czke", "swiss roll"
    ]):
        if not any(k in i for k in ["savory", "sandwich", "benedict", "egg", "bacon", "salmon", "cheese and", "ham and", "turkey and", "pizza", "avocado and", "pancake", "avocado & bacon waffle"]):
            return "Pastries"

    # 8. Drinks
    if any(k in i or k in g for k in [
        "latte", "coffee", "espresso", "americano", "cappuccino", "refresher", "tea",
        "frappe", "frappachino", "water", "soda", "coke", "sprite", "milk", "juice",
        "lemonade", "smoothie", "boba", "drink", "drinks", "beverage", "beverages",
        "slush", "slushie", "shake", "mocktail", "cocktail", "mimosa", "beer", "sapporo",
        "sapporro", "ramune", "calpico", "mogu", "izze", "celsius", "pellegrino",
        "perrier", "alani", "evian", "fiji", "dasani", "yakult", "matcha", "hot chocolate",
        "affogato", "cold brew", "sparkling", "coffe", "pepsi", "aquafina", "oolong",
        "genmaicha", "cold", "hot", "bliss", "ginger ale", "espolon", "lemon ice",
        "cup of ice", "ice cup", "smoo", "smoothje", "sprakling"
    ]):
        if not any(k in i for k in ["pancake", "cheesecake", "croffle", "sandwich", "chicken", "waffle"]):
            return "Drinks"
    if i in [
        "water", "milk", "oj", "orange juice", "apple juice", "lemonade", "hotwater",
        "drink", "wter", "sprie", "binggrae", "bingrae", "monday oj", "hot milk",
        "whole milk", "chocolate milk", "cup of milk", "cream", "oat milk", "chai",
        "match clud", "pepsi", "aquafina", "coca-cola", "coca cola", "orange", "orsnge",
        "ornge juicd", "apple j", "apple juic", "applejuic", "oat", "alanj", "mili",
        "mlk", "warm wilk", "watee", "watr", "monday lemoand", "monday lemonad",
        "refreser strawwberry", "vanilla late", "strawberry re", "iced vanilla late",
        "mango smothie almond", "mango smotthie", "match witybmango and dragon fruit",
        "strawberry and cream frappachini"
    ]:
        return "Drinks"

    # Fallback checks
    if "pancake" in i or "creme brulee" in i:
        return "Pancakes"
    if "croffle" in i:
        return "Pastries"
    if "cake" in i:
        return "Cheesecakes"
    if "tea" in i or "coffee" in i or "juice" in i or "latte" in i:
        return "Drinks"

    return "Unmapped"


def get_quarter(ts):
    month = ts.month
    q = (month - 1) // 3 + 1
    suffixes = {1: "st", 2: "nd", 3: "rd", 4: "th"}
    return f"{q}{suffixes.get(q, 'th')} Quarter"


def clean_location(loc):
    s = str(loc).strip()
    return s


def format_location_name(loc):
    s = clean_location(loc)
    s = s.replace("Fluffy Fluffy - ", "").replace("Fluffy Fluffy ", "")
    return s


def build_pmix_table(df, title):
    quarters = ["3rd Quarter", "4th Quarter"]
    loc_total_qty = df["Qty"].sum()
    cats = CATEGORIES + ["Non-Menu"] if (df["Menu Category"] == "Non-Menu").any() else CATEGORIES

    pv = df.pivot_table(index="Menu Category", columns="Quarter", values="Qty", aggfunc="sum", fill_value=0)
    for q in quarters:
        if q not in pv.columns:
            pv[q] = 0
    pv = pv.reindex(index=cats, fill_value=0)

    rows = []
    for cat in cats:
        q3_qty = pv.loc[cat, "3rd Quarter"]
        q4_qty = pv.loc[cat, "4th Quarter"]
        tot_qty = q3_qty + q4_qty
        rows.append({
            "Menu Item": cat,
            "3rd Quarter Qty (#)": int(q3_qty),
            "3rd Quarter Qty (%)": q3_qty / loc_total_qty if loc_total_qty else 0,
            "4th Quarter Qty (#)": int(q4_qty),
            "4th Quarter Qty (%)": q4_qty / loc_total_qty if loc_total_qty else 0,
            "Total Qty (#)": int(tot_qty),
            "Total Qty (%)": tot_qty / loc_total_qty if loc_total_qty else 0,
        })

    # Grand Total row
    q3_tot = pv["3rd Quarter"].sum()
    q4_tot = pv["4th Quarter"].sum()
    grand_tot = q3_tot + q4_tot
    rows.append({
        "Menu Item": "Grand Total",
        "3rd Quarter Qty (#)": int(q3_tot),
        "3rd Quarter Qty (%)": q3_tot / loc_total_qty if loc_total_qty else 0,
        "4th Quarter Qty (#)": int(q4_tot),
        "4th Quarter Qty (%)": q4_tot / loc_total_qty if loc_total_qty else 0,
        "Total Qty (#)": int(grand_tot),
        "Total Qty (%)": 1.0 if loc_total_qty else 0,
    })

    return pd.DataFrame(rows)


def main():
    print("Reading USA PMix.xlsx...")
    xls = pd.ExcelFile(SRC)
    toast = pd.read_excel(xls, "TOAST")
    square = pd.read_excel(xls, "SQUARE")

    print(f"TOAST rows: {len(toast)}, SQUARE rows: {len(square)}")

    # Clean dates and locations
    toast["Location_Clean"] = toast["Location"].apply(clean_location)
    square["Location_Clean"] = square["Location"].apply(clean_location)

    toast["Sent Date Clean"] = pd.to_datetime(toast["Sent Date"])
    square["Date Clean"] = pd.to_datetime(square["Date"])

    toast["Year"] = toast["Sent Date Clean"].dt.year
    toast["Quarter"] = toast["Sent Date Clean"].apply(get_quarter)

    square["Year"] = square["Date Clean"].dt.year
    square["Quarter"] = square["Date Clean"].apply(get_quarter)

    # Apply categorization
    toast["Menu Category"] = [classify_item(g, i) for g, i in zip(toast["Menu Group"], toast["Menu Item"])]
    square["Menu Category"] = [classify_item(c, i) for c, i in zip(square["Category"], square["Item"])]

    print("TOAST category counts:")
    print(toast["Menu Category"].value_counts())
    print("\nSQUARE category counts:")
    print(square["Menu Category"].value_counts())

    # Combine for reporting
    t_sub = toast[["Location_Clean", "Year", "Quarter", "Menu Group", "Menu Item", "Qty", "Menu Category"]].copy()
    t_sub.rename(columns={"Location_Clean": "Location", "Menu Group": "Source Group", "Menu Item": "Item"}, inplace=True)
    t_sub["Sheet"] = "TOAST"

    s_sub = square[["Location_Clean", "Year", "Quarter", "Category", "Item", "Qty", "Menu Category"]].copy()
    s_sub.rename(columns={"Location_Clean": "Location", "Category": "Source Group"}, inplace=True)
    s_sub["Sheet"] = "SQUARE"

    all_orders = pd.concat([t_sub, s_sub], ignore_index=True)
    all_orders["Qty"] = pd.to_numeric(all_orders["Qty"], errors="coerce").fillna(0)

    # Unique locations
    locations = sorted(all_orders["Location"].unique())
    print(f"\nTotal locations: {len(locations)}")
    print(locations)

    # 1. Summary table across all locations
    summary_all = build_pmix_table(all_orders, "Quarterly PMIX Report - All Locations")

    # 2. Location Summary (Side by side comparison)
    loc_summary_rows = []
    quarters = ["3rd Quarter", "4th Quarter"]
    for loc in locations:
        ldf = all_orders[all_orders["Location"] == loc]
        ltot = ldf["Qty"].sum()
        pv = ldf.pivot_table(index="Menu Category", columns="Quarter", values="Qty", aggfunc="sum", fill_value=0)
        for q in quarters:
            if q not in pv.columns:
                pv[q] = 0
        for cat in CATEGORIES:
            q3 = pv.loc[cat, "3rd Quarter"] if cat in pv.index else 0
            q4 = pv.loc[cat, "4th Quarter"] if cat in pv.index else 0
            tot = q3 + q4
            loc_summary_rows.append({
                "Location": loc,
                "Menu Category": cat,
                "3rd Quarter Qty (#)": int(q3),
                "3rd Quarter Qty (%)": q3 / ltot if ltot else 0,
                "4th Quarter Qty (#)": int(q4),
                "4th Quarter Qty (%)": q4 / ltot if ltot else 0,
                "Total Qty (#)": int(tot),
                "Total Qty (%)": tot / ltot if ltot else 0,
            })
    location_summary_df = pd.DataFrame(loc_summary_rows)

    # 3. Report Data (Detailed rollup)
    report_data = (
        all_orders.groupby(["Location", "Year", "Quarter", "Menu Category", "Source Group", "Item"], dropna=False)["Qty"]
        .sum()
        .reset_index()
        .rename(columns={"Source Group": "Menu group", "Qty": "Qty sold"})
    )
    report_data = report_data.sort_values(
        by=["Location", "Quarter", "Menu Category", "Qty sold"], ascending=[True, True, True, False]
    )

    # 4. Item Mapping reference
    item_mapping = (
        all_orders.groupby(["Sheet", "Source Group", "Item", "Menu Category"], dropna=False)
        .agg(Rows=("Item", "count"), Total_Qty=("Qty", "sum"))
        .reset_index()
        .sort_values(by=["Sheet", "Total_Qty"], ascending=[True, False])
    )

    # Write out to Excel with professional formatting matching Quarterly PMIX Report
    print(f"\nWriting output workbook to {OUT}...")
    with pd.ExcelWriter(OUT, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_format = workbook.add_format({"bold": True, "font_size": 12})
        header_format = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#D9E1F2", "border": 1
        })
        cat_format = workbook.add_format({"bold": True, "border": 1})
        grand_total_format = workbook.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        num_format = workbook.add_format({"num_format": "#,##0", "align": "right", "border": 1})
        pct_format = workbook.add_format({"num_format": "0.0%", "align": "right", "border": 1})
        grand_num_format = workbook.add_format({"bold": True, "num_format": "#,##0", "align": "right", "bg_color": "#F2F2F2", "border": 1})
        grand_pct_format = workbook.add_format({"bold": True, "num_format": "0.0%", "align": "right", "bg_color": "#F2F2F2", "border": 1})

        def write_report_sheet(ws, title, df_table):
            ws.write(0, 0, title, title_format)
            ws.write(1, 0, "", header_format)
            ws.merge_range(1, 1, 1, 2, "3rd Quarter", header_format)
            ws.merge_range(1, 3, 1, 4, "4th Quarter", header_format)
            ws.merge_range(1, 5, 1, 6, "Total", header_format)

            headers = ["Menu Item", "Qty (#)", "Qty (%)", "Qty (#)", "Qty (%)", "Total Qty (#)", "Total Qty (%)"]
            for col_idx, h in enumerate(headers):
                ws.write(2, col_idx, h, header_format)

            for row_idx, r in df_table.iterrows():
                excel_row = 3 + row_idx
                is_grand = (r["Menu Item"] == "Grand Total")
                c_fmt = grand_total_format if is_grand else cat_format
                n_fmt = grand_num_format if is_grand else num_format
                p_fmt = grand_pct_format if is_grand else pct_format

                ws.write(excel_row, 0, r["Menu Item"], c_fmt)
                ws.write(excel_row, 1, r["3rd Quarter Qty (#)"], n_fmt)
                ws.write(excel_row, 2, r["3rd Quarter Qty (%)"], p_fmt)
                ws.write(excel_row, 3, r["4th Quarter Qty (#)"], n_fmt)
                ws.write(excel_row, 4, r["4th Quarter Qty (%)"], p_fmt)
                ws.write(excel_row, 5, r["Total Qty (#)"], n_fmt)
                ws.write(excel_row, 6, r["Total Qty (%)"], p_fmt)

            ws.set_column(0, 0, 24)
            ws.set_column(1, 6, 14)

        # 1. Summary - All Locations
        ws_all = workbook.add_worksheet("Summary - All Locations")
        write_report_sheet(ws_all, "Quarterly PMIX Report - All Locations (8 Menu Categories)", summary_all)

        # 2. Location Summary comparison table
        location_summary_df.to_excel(writer, sheet_name="Location Summary", index=False)
        ws_ls = writer.sheets["Location Summary"]
        ws_ls.set_column(0, 0, 26)
        ws_ls.set_column(1, 1, 20)
        ws_ls.set_column(2, 7, 14)

        # 3. Individual Location Sheets (All 20 Locations!)
        for loc in locations:
            short_name = format_location_name(loc)[:31]
            ws_loc = workbook.add_worksheet(short_name)
            ldf = all_orders[all_orders["Location"] == loc]
            loc_table = build_pmix_table(ldf, f"Quarterly PMIX Report - {loc}")
            write_report_sheet(ws_loc, f"Quarterly PMIX Report - {loc}", loc_table)

        # 4. Report Data
        report_data.to_excel(writer, sheet_name="Report Data", index=False)
        ws_rd = writer.sheets["Report Data"]
        ws_rd.freeze_panes(1, 0)
        ws_rd.set_column(0, 0, 26)
        ws_rd.set_column(1, 2, 14)
        ws_rd.set_column(3, 5, 26)
        ws_rd.set_column(6, 6, 12)

        # 5. TOAST sheet with Menu Category column appended
        toast_out = toast.drop(columns=["Location_Clean", "Sent Date Clean", "Year", "Quarter"])
        toast_out.to_excel(writer, sheet_name="TOAST", index=False)
        ws_t = writer.sheets["TOAST"]
        ws_t.freeze_panes(1, 0)
        ws_t.set_column(0, len(toast_out.columns) - 1, 22)

        # 6. SQUARE sheet with Menu Category column appended
        square_out = square.drop(columns=["Location_Clean", "Date Clean", "Year", "Quarter"])
        square_out.to_excel(writer, sheet_name="SQUARE", index=False)
        ws_s = writer.sheets["SQUARE"]
        ws_s.freeze_panes(1, 0)
        ws_s.set_column(0, len(square_out.columns) - 1, 22)

        # 7. Item Mapping
        item_mapping.to_excel(writer, sheet_name="Item Mapping", index=False)
        ws_im = writer.sheets["Item Mapping"]
        ws_im.freeze_panes(1, 0)
        ws_im.set_column(0, 4, 26)

        # 8. Menu Category sheet preserved from original
        mc_ws = workbook.add_worksheet("Menu Category")
        src_mc = openpyxl.load_workbook(SRC)["Menu Category"]
        mc_ws.set_column(0, 0, 14)
        mc_ws.set_column(4, 4, 54)
        for r in range(1, src_mc.max_row + 1):
            for col in (1, 5):
                c = src_mc.cell(r, col)
                if c.value is not None:
                    fmt = workbook.add_format({
                        "bold": bool(c.font.b),
                        "indent": int(c.alignment.indent or 0),
                    })
                    mc_ws.write(r - 1, col - 1, c.value, fmt)

    print("Completed successfully! File saved to:", OUT)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
