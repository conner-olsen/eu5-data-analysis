"""Analyze EU5 army unit data and export to Excel spreadsheet."""

import json
import os
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
OUTPUT_DIR = DATA_DIR

AGE_ORDER = [
    "age_1_traditions",
    "age_2_renaissance",
    "age_3_discovery",
    "age_4_reformation",
    "age_5_absolutism",
    "age_6_revolutions",
]
AGE_LABELS = {
    "age_1_traditions": "1 - Traditions",
    "age_2_renaissance": "2 - Renaissance",
    "age_3_discovery": "3 - Discovery",
    "age_4_reformation": "4 - Reformation",
    "age_5_absolutism": "5 - Absolutism",
    "age_6_revolutions": "6 - Revolutions",
}
CAT_LABELS = {
    "army_infantry": "Infantry",
    "army_cavalry": "Cavalry",
    "army_artillery": "Artillery",
    "army_auxiliary": "Auxiliary",
}
LAND_CATS = ["army_infantry", "army_cavalry", "army_artillery", "army_auxiliary"]

# Styling
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT_WHITE = Font(bold=True, size=11, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
SUBTITLE_FONT = Font(bold=True, size=11, italic=True)
NUM_FMT_2 = "0.00"
NUM_FMT_3 = "0.000"
NUM_FMT_PCT = "0%"
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Category highlight colors
CAT_FILLS = {
    "Infantry": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),
    "Cavalry": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "Artillery": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "Auxiliary": PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid"),
}


def load_data():
    with open(DATA_DIR / "land_units.json") as f:
        land_units = json.load(f)
    with open(DATA_DIR / "unit_categories.json") as f:
        categories = json.load(f)
    with open(DATA_DIR / "age_progression.json") as f:
        age_progression = json.load(f)
    with open(DATA_DIR / "unit_prices.json") as f:
        prices = json.load(f)
    with open(DATA_DIR / "combined_arms.json") as f:
        combined_arms = json.load(f)
    with open(DATA_DIR / "goods_demands.json") as f:
        goods_demands = json.load(f)
    with open(DATA_DIR / "production_recipes.json") as f:
        production_recipes = json.load(f)
    with open(DATA_DIR / "localizations.json") as f:
        localizations = json.load(f)
    return land_units, categories, age_progression, prices, combined_arms, goods_demands, production_recipes, localizations


LOC = {}  # populated in main(), used by loc() helper


def loc(internal_name: str) -> str:
    """Get display name for a unit, falling back to internal name."""
    return LOC.get(internal_name, internal_name)


def style_header_row(ws, row, num_cols):
    """Style a header row with blue background and white bold text."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def safe_num(v, default=0):
    """Safely extract a numeric value, handling lists from duplicate keys."""
    if isinstance(v, list):
        # Sum duplicate modifier values (e.g., strength_damage_taken appearing twice)
        return sum(x for x in v if isinstance(x, (int, float)))
    if isinstance(v, (int, float)):
        return v
    return default


def auto_width(ws, min_width=8, max_width=30):
    """Auto-fit column widths based on content."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = min_width
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)) + 2, max_width))
        ws.column_dimensions[col_letter].width = max_len


def calc_flank_power(damage, sfd, dmg_taken, str_dmg_done=0, str_dmg_taken=0):
    """FlankPower = EffDamage * 10 / ((1 - SFD) * DamageTaken * (1 + str_dmg_taken))

    str_dmg_done/taken: unit-specific strength damage modifiers (negative = better).
    """
    eff_damage = damage * (1 + str_dmg_done)
    denom = (1 - sfd) * dmg_taken * (1 + str_dmg_taken)
    if denom == 0:
        return 0
    return eff_damage * 10 / denom


def calc_center_power(flank_power, sfd):
    """CenterPower = FlankPower * (1 + 2 * SFD)"""
    return flank_power * (1 + 2 * sfd)


def calc_cost(strength, category, prices):
    """BuildCost = max_strength * base_build_gold * 10 (costs scale by max_strength)."""
    base = prices.get(category, {}).get("build_gold", 100)
    return strength * base * 10


def calc_maintenance(strength, category, prices):
    """Maintenance = max_strength * base_maintenance_gold * 10."""
    base = prices.get(category, {}).get("maintenance_gold", 1)
    return strength * base * 10


BEST_FONT = Font(bold=True)
BEST_OVERALL_FONT = Font(bold=True, color="FF0000")
BEST_CAT_FONT = Font(color="0070C0")  # blue


def highlight_best_in_age(ws, data_rows, header_row, col_indices):
    """Bold the best value per age for specified columns.

    data_rows: list of (excel_row_number, age_key, {col_index: value}) dicts
    col_indices: list of column indices (1-based) to check for best values
    """
    by_age = {}
    for row_num, age, vals in data_rows:
        by_age.setdefault(age, []).append((row_num, vals))

    for age, rows in by_age.items():
        for col in col_indices:
            best_val = None
            best_row = None
            for row_num, vals in rows:
                v = vals.get(col, 0)
                if v is not None and (best_val is None or v > best_val):
                    best_val = v
                    best_row = row_num
            if best_row is not None and best_val is not None and best_val > 0:
                ws.cell(row=best_row, column=col).font = BEST_FONT


def highlight_best_in_age_by_cat(ws, data_rows, col_indices):
    """Highlight best overall (green) and best per category (blue) per age.

    data_rows: list of (excel_row_number, age_key, category, {col_index: value})
    """
    by_age = {}
    for row_num, age, cat, vals in data_rows:
        by_age.setdefault(age, []).append((row_num, cat, vals))

    for age, rows in by_age.items():
        for col in col_indices:
            # Find overall best
            best_val = None
            best_row = None
            # Find best per category
            cat_best = {}  # cat -> (row, val)
            for row_num, cat, vals in rows:
                v = vals.get(col, 0)
                if v is not None and v > 0:
                    if best_val is None or v > best_val:
                        best_val = v
                        best_row = row_num
                    cb = cat_best.get(cat)
                    if cb is None or v > cb[1]:
                        cat_best[cat] = (row_num, v)

            # Apply overall best = green
            if best_row is not None:
                ws.cell(row=best_row, column=col).font = BEST_OVERALL_FONT

            # Apply best per category = blue (skip the overall winner's category
            # and categories with only one unit)
            if best_row is not None:
                winner_cat = None
                for row_num, cat, vals in rows:
                    if row_num == best_row:
                        winner_cat = cat
                        break
                # Count units per category
                cat_counts = {}
                for row_num, cat, vals in rows:
                    cat_counts[cat] = cat_counts.get(cat, 0) + 1
                for cat, (cr, cv) in cat_best.items():
                    if cat != winner_cat and cv > 0 and cat_counts.get(cat, 0) > 1:
                        ws.cell(row=cr, column=col).font = BEST_CAT_FONT


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def build_army_meta(wb, age_data, categories, prices):
    """Sheet 1: Army Composition Meta - per-age flank/center analysis."""
    ws = wb.active
    ws.title = "Army Meta"

    # Title
    ws.cell(row=1, column=1, value="EU5 Army Composition Meta").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Per-age template stats with flank/center power analysis").font = SUBTITLE_FONT

    # Formula explanation
    ws.cell(row=3, column=1, value="Formulas:").font = Font(bold=True)
    ws.cell(row=4, column=1, value="  Damage = Strength x CombatPower")
    ws.cell(row=5, column=1, value="  Flank Power = Damage x 10 / ((1 - SecureFlanks) x DamageTaken)")
    ws.cell(row=6, column=1, value="  Center Power = Flank Power x (1 + 2 x SecureFlanks)")
    ws.cell(row=7, column=1, value="  Power Per Gold = Power / BuildCost x 100")
    ws.cell(row=8, column=1, value="  Gold cells = best in age for that column").font = Font(italic=True)

    # Headers
    headers = [
        "Age", "Category", "Strength", "Combat Power", "Combat Speed",
        "Flanking Ability", "Secure Flanks", "Damage Taken",
        "Build Cost", "Maintenance",
        "Damage",
        "Flank Power", "Center Power",
        "Flank Power/Gold", "Center Power/Gold",
    ]
    header_row = 10
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))

    # Column indices for highlighting (1-based): FP=12, CP=13, FP/G=14, CP/G=15
    HIGHLIGHT_COLS = [12, 13, 14, 15]

    tracked_rows = []  # (row_num, age_key, {col: value})
    row = header_row + 1
    for age in AGE_ORDER:
        age_rows = [r for r in age_data if r["age"] == age]
        for ar in age_rows:
            cat = ar["category"]
            if cat not in LAND_CATS:
                continue

            cat_data = categories.get(cat, {})
            cat_label = CAT_LABELS.get(cat, cat)

            strength = ar["max_strength"]
            cp = ar["combat_power"]
            cs = cat_data.get("combat_speed", 1)
            fa = cat_data.get("flanking_ability", 1.0)
            sfd = cat_data.get("secure_flanks_defense", 0)
            dmg_taken = cat_data.get("damage_taken", 1.0)

            damage = strength * cp
            build_cost = calc_cost(strength, cat, prices)
            maint = calc_maintenance(strength, cat, prices)

            fp = calc_flank_power(damage, sfd, dmg_taken)
            cp_val = calc_center_power(fp, sfd)
            fp_gold = fp / build_cost * 100 if build_cost > 0 else 0
            cp_gold = cp_val / build_cost * 100 if build_cost > 0 else 0

            values = [
                AGE_LABELS[age], cat_label, strength, cp, cs,
                fa, sfd, dmg_taken,
                build_cost, maint,
                damage,
                fp, cp_val,
                fp_gold, cp_gold,
            ]
            for i, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.border = THIN_BORDER
                if isinstance(v, float):
                    cell.number_format = NUM_FMT_2
                cat_fill = CAT_FILLS.get(cat_label)
                if cat_fill:
                    cell.fill = cat_fill

            tracked_rows.append((row, age, {c: values[c - 1] for c in HIGHLIGHT_COLS}))
            row += 1

        row += 1  # blank row between ages

    highlight_best_in_age(ws, tracked_rows, header_row, HIGHLIGHT_COLS)
    auto_width(ws)
    ws.freeze_panes = f"A{header_row + 1}"


def build_buildable_units(wb, land_units, categories, prices):
    """Sheet 2: All buildable units with detailed stats and power calculations."""
    ws = wb.create_sheet("Unique Units")

    ws.cell(row=1, column=1, value="Unique Units by Age").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Green = best overall in age, Blue = best in category").font = Font(italic=True)

    headers = [
        "Age", "Category", "Unit",
        "Flank Power", "Center Power", "Flank P/Gold", "Center P/Gold",
        "Light", "Special",
        "Strength", "Combat Power", "Combat Speed", "Initiative",
        "Flanking", "Secure Flanks", "Damage Taken",
        "Str Dmg Taken", "Morale Dmg Taken",
        "Str Dmg Done", "Morale Dmg Done",
        "Damage", "Build Cost",
        "Upgrades To",
    ]
    header_row = 4
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))

    # Highlight columns: FP=4, CP=5, FP/G=6, CP/G=7
    HIGHLIGHT_COLS = [4, 5, 6, 7]

    units = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
    ]

    tracked_rows = []
    row = header_row + 1
    for age in AGE_ORDER:
        age_units = [u for u in units if u.get("age") == age]
        if not age_units:
            continue

        for u in sorted(age_units, key=lambda x: (x["category"], x["name"])):
            cat = u["category"]
            cat_data = categories.get(cat, {})
            cat_label = CAT_LABELS.get(cat, cat)

            strength = u.get("max_strength", 0)
            cp = u.get("combat_power", 0)
            cs = u.get("combat_speed", cat_data.get("combat_speed", 1))
            fa = u.get("flanking_ability", cat_data.get("flanking_ability", 1.0))
            sfd = u.get("secure_flanks_defense", cat_data.get("secure_flanks_defense", 0))
            dmg_taken = u.get("damage_taken", cat_data.get("damage_taken", 1.0))

            sdd = safe_num(u.get("strength_damage_done", 0))
            sdt = safe_num(u.get("strength_damage_taken", 0))

            damage = strength * cp
            build_cost = calc_cost(strength, cat, prices)

            fp = calc_flank_power(damage, sfd, dmg_taken, sdd, sdt)
            center = calc_center_power(fp, sfd)
            fp_gold = fp / build_cost * 100 if build_cost > 0 else 0
            cp_gold = center / build_cost * 100 if build_cost > 0 else 0

            values = [
                AGE_LABELS.get(u.get("age", ""), "?"),
                cat_label,
                loc(u["name"]),
                fp, center, fp_gold, cp_gold,
                "Yes" if u.get("light") else "",
                "Yes" if u.get("is_special") else "",
                strength, cp, cs,
                safe_num(u.get("initiative", cat_data.get("initiative", 1))),
                fa, sfd, dmg_taken,
                safe_num(u.get("strength_damage_taken", 0)),
                safe_num(u.get("morale_damage_taken", 0)),
                safe_num(u.get("strength_damage_done", 0)),
                safe_num(u.get("morale_damage_done", 0)),
                damage, build_cost,
                u.get("upgrades_to", ""),
            ]
            for i, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=i, value=v)
                cell.border = THIN_BORDER
                if isinstance(v, float):
                    cell.number_format = NUM_FMT_2
                cat_fill = CAT_FILLS.get(cat_label)
                if cat_fill:
                    cell.fill = cat_fill

            tracked_rows.append((row, age, cat_label, {c: values[c - 1] for c in HIGHLIGHT_COLS}))
            row += 1

        row += 1  # gap between ages

    highlight_best_in_age_by_cat(ws, tracked_rows, HIGHLIGHT_COLS)
    auto_width(ws)
    ws.freeze_panes = f"A{header_row + 1}"


def build_upgrade_chains(wb, land_units):
    """Sheet 3: Upgrade chains laid out horizontally."""
    ws = wb.create_sheet("Upgrade Chains")

    ws.cell(row=1, column=1, value="Unit Upgrade Chains").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Each row shows a full upgrade path from earliest to latest age").font = SUBTITLE_FONT

    headers = ["Category", "Type"] + [AGE_LABELS[a] for a in AGE_ORDER]
    header_row = 4
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))

    units_by_name = {u["name"]: u for u in land_units}
    buildable = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
    ]

    upgraded_to = {u.get("upgrades_to") for u in buildable if u.get("upgrades_to")}
    starters = [u for u in buildable if u["name"] not in upgraded_to]

    row = header_row + 1
    for start in sorted(starters, key=lambda x: (x["category"], x.get("is_special", False), x["name"])):
        cat_label = CAT_LABELS.get(start["category"], start["category"])
        is_light = start.get("light", False)
        is_special = start.get("is_special", False)
        type_label = ("Special" if is_special else "Light" if is_light else "Heavy")

        ws.cell(row=row, column=1, value=cat_label).border = THIN_BORDER
        ws.cell(row=row, column=2, value=type_label).border = THIN_BORDER

        # Walk the chain and place units in their age column
        current = start
        visited = set()
        while current and current["name"] not in visited:
            visited.add(current["name"])
            age = current.get("age", "")
            if age in AGE_ORDER:
                col = AGE_ORDER.index(age) + 3  # offset for Category, Type columns
                ep = round(current.get("max_strength", 0) * current.get("combat_power", 0), 2)
                cell = ws.cell(row=row, column=col, value=f"{loc(current['name'])} (EP:{ep})")
                cell.border = THIN_BORDER
                cat_fill = CAT_FILLS.get(cat_label)
                if cat_fill:
                    cell.fill = cat_fill

            next_name = current.get("upgrades_to", "")
            current = units_by_name.get(next_name) if next_name else None

        row += 1

    auto_width(ws, min_width=10, max_width=45)
    ws.freeze_panes = f"A{header_row + 1}"


def build_special_units(wb, land_units, categories):
    """Sheet 4: Special/unique units with terrain modifiers."""
    ws = wb.create_sheet("Special Units")

    ws.cell(row=1, column=1, value="Special / Unique Units").font = TITLE_FONT

    specials = [
        u for u in land_units
        if u.get("is_special", False)
        and not u.get("levy", False)
        and u.get("buildable", True)
    ]

    headers = [
        "Age", "Unit", "Category", "Light",
        "Strength", "Combat Power", "Effective Power",
        "Combat Speed", "Initiative",
        "Str Dmg Taken", "Morale Dmg Taken",
        "Str Dmg Done", "Morale Dmg Done",
        "Flank Power", "Center Power",
        "Terrain Combat", "Terrain Impact",
    ]
    header_row = 3
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))

    row = header_row + 1
    for u in sorted(specials, key=lambda x: (x.get("age", ""), x["category"], x["name"])):
        cat = u["category"]
        cat_data = categories.get(cat, {})
        cat_label = CAT_LABELS.get(cat, cat)

        strength = u.get("max_strength", 0)
        cp = u.get("combat_power", 0)
        sfd = u.get("secure_flanks_defense", cat_data.get("secure_flanks_defense", 0))
        dmg_taken = u.get("damage_taken", cat_data.get("damage_taken", 1.0))
        sdd = safe_num(u.get("strength_damage_done", 0))
        sdt = safe_num(u.get("strength_damage_taken", 0))
        damage = strength * cp

        fp = calc_flank_power(damage, sfd, dmg_taken, sdd, sdt)
        center = calc_center_power(fp, sfd)

        tc = u.get("terrain_combat", {})
        ti = u.get("terrain_impact", {})
        tc_str = ", ".join(f"{k}: {v:+.2f}" for k, v in tc.items()) if tc else ""
        ti_str = ", ".join(f"{k}: {v:+.2f}" for k, v in ti.items()) if ti else ""

        values = [
            AGE_LABELS.get(u.get("age", ""), "?"),
            loc(u["name"]), cat_label,
            "Yes" if u.get("light") else "",
            strength, cp, round(damage, 2),
            safe_num(u.get("combat_speed", cat_data.get("combat_speed", 1))),
            safe_num(u.get("initiative", cat_data.get("initiative", 1))),
            safe_num(u.get("strength_damage_taken", 0)),
            safe_num(u.get("morale_damage_taken", 0)),
            safe_num(u.get("strength_damage_done", 0)),
            safe_num(u.get("morale_damage_done", 0)),
            fp, center,
            tc_str, ti_str,
        ]
        for i, v in enumerate(values, 1):
            cell = ws.cell(row=row, column=i, value=v)
            cell.border = THIN_BORDER
            if isinstance(v, float):
                cell.number_format = NUM_FMT_2
            cat_fill = CAT_FILLS.get(cat_label)
            if cat_fill:
                cell.fill = cat_fill

        row += 1

    auto_width(ws, max_width=50)
    ws.freeze_panes = f"A{header_row + 1}"


def build_category_reference(wb, categories, prices):
    """Sheet 5: Base category stats reference."""
    ws = wb.create_sheet("Category Stats")

    ws.cell(row=1, column=1, value="Unit Category Base Stats").font = TITLE_FONT
    ws.cell(row=2, column=1, value="These are inherent stats from the category, before any unit-type modifiers").font = SUBTITLE_FONT

    stats = [
        ("damage_taken", "Damage Taken Multiplier"),
        ("combat_speed", "Combat Speed"),
        ("initiative", "Initiative"),
        ("frontage", "Frontage"),
        ("flanking_ability", "Flanking Ability"),
        ("secure_flanks_defense", "Secure Flanks Defense"),
        ("supply_weight", "Supply Weight"),
        ("attrition_loss", "Extra Attrition"),
        ("food_storage_per_strength", "Food Storage / Strength"),
        ("food_consumption_per_strength", "Food Consumption / Strength"),
        ("startup_amount", "Starting Army Amount"),
    ]

    header_row = 4
    ws.cell(row=header_row, column=1, value="Stat")
    for i, cat in enumerate(LAND_CATS, 2):
        ws.cell(row=header_row, column=i, value=CAT_LABELS[cat])
    style_header_row(ws, header_row, len(LAND_CATS) + 1)

    for r, (key, label) in enumerate(stats, header_row + 1):
        ws.cell(row=r, column=1, value=label).border = THIN_BORDER
        ws.cell(row=r, column=1).font = Font(bold=True)
        for i, cat in enumerate(LAND_CATS, 2):
            val = categories.get(cat, {}).get(key, 0)
            cell = ws.cell(row=r, column=i, value=val)
            cell.border = THIN_BORDER
            if isinstance(val, float):
                cell.number_format = NUM_FMT_2

    # Cost reference (scraped from prices/02_units.txt)
    cost_row = header_row + len(stats) + 2
    ws.cell(row=cost_row, column=1, value="Cost Reference (scraped from prices/02_units.txt)").font = Font(bold=True, size=12)
    cost_headers = ["Category", "Build Cost (gold)", "Reinforce Cost", "Maintenance (gold)"]
    for i, h in enumerate(cost_headers, 1):
        ws.cell(row=cost_row + 1, column=i, value=h)
    style_header_row(ws, cost_row + 1, len(cost_headers))

    for r, cat in enumerate(LAND_CATS, cost_row + 2):
        p = prices.get(cat, {})
        values = [
            CAT_LABELS[cat],
            p.get("build_gold", 0),
            p.get("reinforce_gold", 0),
            p.get("maintenance_gold", 0),
        ]
        for i, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.border = THIN_BORDER

    auto_width(ws)


def build_light_vs_heavy(wb, land_units, categories):
    """Sheet 6: Light vs Heavy comparison within same age/category."""
    ws = wb.create_sheet("Light vs Heavy")

    ws.cell(row=1, column=1, value="Light vs Heavy Unit Comparison").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Generic (non-special) units only - same age, same category").font = SUBTITLE_FONT

    headers = [
        "Age", "Category", "Unit", "Type",
        "Strength", "Combat Power", "Effective Power",
        "Combat Speed", "Initiative",
        "Flanking", "Secure Flanks",
        "Str Dmg Taken", "Morale Dmg Taken",
        "Str Dmg Done", "Morale Dmg Done",
        "Flank Power", "Center Power",
    ]
    header_row = 4
    for i, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=i, value=h)
    style_header_row(ws, header_row, len(headers))

    buildable = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
        and not u.get("is_special", False)
    ]

    row = header_row + 1
    for age in AGE_ORDER:
        age_units = [u for u in buildable if u.get("age") == age]
        for cat in ["army_infantry", "army_cavalry"]:
            cat_units = [u for u in age_units if u["category"] == cat]
            lights = [u for u in cat_units if u.get("light")]
            heavies = [u for u in cat_units if not u.get("light")]

            if not lights or not heavies:
                continue

            for u in heavies + lights:
                cat_data = categories.get(cat, {})
                cat_label = CAT_LABELS.get(cat, cat)

                strength = u.get("max_strength", 0)
                cp = u.get("combat_power", 0)
                sfd = u.get("secure_flanks_defense", cat_data.get("secure_flanks_defense", 0))
                dmg_taken = u.get("damage_taken", cat_data.get("damage_taken", 1.0))
                sdd = safe_num(u.get("strength_damage_done", 0))
                sdt = safe_num(u.get("strength_damage_taken", 0))
                damage = strength * cp

                fp = calc_flank_power(damage, sfd, dmg_taken, sdd, sdt)
                center = calc_center_power(fp, sfd)

                values = [
                    AGE_LABELS[age], cat_label,
                    loc(u["name"]),
                    "Light" if u.get("light") else "Heavy",
                    strength, cp, round(damage, 2),
                    safe_num(u.get("combat_speed", cat_data.get("combat_speed", 1))),
                    safe_num(u.get("initiative", cat_data.get("initiative", 1))),
                    safe_num(u.get("flanking_ability", cat_data.get("flanking_ability", 1.0))),
                    safe_num(u.get("secure_flanks_defense", cat_data.get("secure_flanks_defense", 0))),
                    safe_num(u.get("strength_damage_taken", 0)),
                    safe_num(u.get("morale_damage_taken", 0)),
                    safe_num(u.get("strength_damage_done", 0)),
                    safe_num(u.get("morale_damage_done", 0)),
                    fp, center,
                ]
                for i, v in enumerate(values, 1):
                    cell = ws.cell(row=row, column=i, value=v)
                    cell.border = THIN_BORDER
                    if isinstance(v, float):
                        cell.number_format = NUM_FMT_2
                    cat_fill = CAT_FILLS.get(cat_label)
                    if cat_fill:
                        cell.fill = cat_fill

                row += 1
            row += 1  # gap between comparisons

    auto_width(ws)
    ws.freeze_panes = f"A{header_row + 1}"


# ---------------------------------------------------------------------------
# Combined Arms Optimization
# ---------------------------------------------------------------------------

# The 6 combined-arms types (light/heavy count as separate)
CA_TYPES = [
    ("Heavy Infantry", "army_infantry", False),
    ("Light Infantry", "army_infantry", True),
    ("Heavy Cavalry", "army_cavalry", False),
    ("Light Cavalry", "army_cavalry", True),
    ("Artillery", "army_artillery", None),  # None = no light/heavy split
    ("Auxiliary", "army_auxiliary", None),
]


def get_best_generic_units(land_units, categories):
    """For each age and each of the 6 CA types, find the best generic unit by flank power.

    Returns: { age: [ { "type_label", "unit_name", "flank_power", "center_power", ... }, ... ] }
    """
    result = {}
    for age in AGE_ORDER:
        age_types = []
        for type_label, cat, is_light in CA_TYPES:
            cat_data = categories.get(cat, {})
            sfd = cat_data.get("secure_flanks_defense", 0)
            dmg_taken = cat_data.get("damage_taken", 1.0)

            candidates = [
                u for u in land_units
                if u.get("age") == age
                and u["category"] == cat
                and u.get("buildable", True)
                and not u.get("levy", False)
                and not u["name"].startswith("a_age_")
                and not u.get("is_special", False)
                and (is_light is None or u.get("light", False) == is_light)
            ]

            if not candidates:
                age_types.append({
                    "type_label": type_label, "unit_name": "-",
                    "flank_power": 0, "center_power": 0,
                    "strength": 0, "combat_power": 0,
                    "str_dmg_done": 0, "str_dmg_taken": 0,
                    "initiative": 0,
                })
                continue

            best = None
            best_fp = -1
            for u in candidates:
                strength = u.get("max_strength", 0)
                cp = u.get("combat_power", 0)
                sdd = safe_num(u.get("strength_damage_done", 0))
                sdt = safe_num(u.get("strength_damage_taken", 0))
                damage = strength * cp
                fp = calc_flank_power(damage, sfd, dmg_taken, sdd, sdt)
                if fp > best_fp:
                    best_fp = fp
                    best = {
                        "type_label": type_label,
                        "unit_name": u["name"],
                        "flank_power": fp,
                        "center_power": calc_center_power(fp, sfd),
                        "strength": strength,
                        "combat_power": cp,
                        "str_dmg_done": sdd,
                        "str_dmg_taken": sdt,
                        "initiative": safe_num(u.get("initiative", cat_data.get("initiative", 0))),
                    }
            age_types.append(best)

        result[age] = age_types
    return result


CENTER_RATIO = 1 / 3  # ~33% of army is center, ~67% flanks
FLANK_RATIO = 1 - CENTER_RATIO


def calc_positional_power(pcts, flank_powers, center_powers):
    """Calculate total army power with optimal positional placement.

    Units are assigned to center (33%) or flanks (67%) to maximize power.
    Each unit type uses its center_power when in center, flank_power on flanks.
    The center/flank benefit ratio determines which types go where.
    """
    n = len(pcts)
    if sum(pcts) < 1e-9:
        return 0

    # For each type, compute the center advantage ratio:
    # how much MORE power it gets from center vs flank (relative)
    # Types with highest center advantage should go to center first.
    type_info = []
    for i in range(n):
        if pcts[i] < 1e-9:
            continue
        fp = flank_powers[i]
        cp = center_powers[i]
        # center_advantage = extra power gained per unit in center vs flank
        advantage = cp - fp  # positive means center is better
        type_info.append((i, pcts[i], fp, cp, advantage))

    # Sort by center advantage descending — best center units first
    type_info.sort(key=lambda x: x[4], reverse=True)

    # Fill center slots (33% of army), then flanks (67%)
    center_remaining = CENTER_RATIO
    total_power = 0

    for idx, pct, fp, cp, adv in type_info:
        # How much of this type goes to center?
        in_center = min(pct, center_remaining)
        in_flank = pct - in_center
        center_remaining -= in_center

        total_power += in_center * cp + in_flank * fp

    return total_power


def optimize_composition(flank_powers, center_powers, combined_arms):
    """Find optimal percentage allocation to maximize total positional power.

    Uses positional placement: 33% center, 67% flanks.
    Each type placed optimally based on its center vs flank advantage.

    Returns: (best_pcts, best_total, bonus_used, n_qualifying)
    """
    bonus_per_type = combined_arms["bonus_per_type"]
    min_pct = combined_arms["min_percent"]
    max_pct = combined_arms["max_threshold"]
    n = len(flank_powers)

    best_total = -1
    best_pcts = [0.0] * n
    best_bonus = 0.0
    best_n_qual = 0

    # Enumerate all 2^n subsets of qualifying types
    for mask in range(1 << n):
        qualifying = [i for i in range(n) if mask & (1 << i)]
        k = len(qualifying)

        if k * min_pct > 1.0:
            continue

        if any(flank_powers[i] <= 0 for i in qualifying):
            continue

        # Allocate minimums
        pcts = [0.0] * n
        for i in qualifying:
            pcts[i] = min_pct
        remaining = 1.0 - k * min_pct

        # Greedily allocate remaining to types that contribute most power.
        # Use a blended power estimate for greedy ordering:
        # assume the marginal unit goes ~67% flank, ~33% center
        blended = [FLANK_RATIO * flank_powers[i] + CENTER_RATIO * center_powers[i]
                    for i in range(n)]
        sorted_indices = sorted(range(n), key=lambda i: blended[i], reverse=True)
        for i in sorted_indices:
            if blended[i] <= 0:
                continue
            room = max_pct - pcts[i]
            add = min(remaining, room)
            if add > 0:
                pcts[i] += add
                remaining -= add
            if remaining <= 1e-9:
                break

        bonus = bonus_per_type * k
        if any(p > max_pct + 1e-9 for p in pcts):
            bonus = 0

        weighted = calc_positional_power(pcts, flank_powers, center_powers)
        total = weighted * (1 + bonus)

        if total > best_total:
            best_total = total
            best_pcts = pcts[:]
            best_bonus = bonus
            best_n_qual = k

    # Also try no-bonus: 100% in best type
    for i in range(n):
        # 100% of one type — it fills both center and flank
        power = CENTER_RATIO * center_powers[i] + FLANK_RATIO * flank_powers[i]
        if power > best_total:
            best_total = power
            best_pcts = [0.0] * n
            best_pcts[i] = 1.0
            best_bonus = 0.0
            best_n_qual = 0

    return best_pcts, best_total, best_bonus, best_n_qual


def build_optimal_composition(wb, land_units, categories, combined_arms):
    """Sheet: Optimal army composition per age using combined arms."""
    ws = wb.create_sheet("Optimal Composition")

    ws.cell(row=1, column=1, value="Optimal Army Composition (Combined Arms)").font = TITLE_FONT

    # Show scraped combined arms values
    ca = combined_arms
    ws.cell(row=2, column=1,
            value=f"Scraped: bonus/type={ca['bonus_per_type']:.1%}, "
                  f"min threshold={ca['min_percent']:.0%}, "
                  f"max cap={ca['max_threshold']:.0%}").font = SUBTITLE_FONT
    ws.cell(row=3, column=1,
            value="Positional placement: 33% center / 67% flanks. "
                  "Units assigned to center/flank to maximize power.").font = Font(italic=True)
    ws.cell(row=4, column=1,
            value="Formulas: EffDmg = Str*CP*(1+StrDmgDone), "
                  "Power = EffDmg*10/((1-SFD)*DmgTaken*(1+StrDmgTaken)), "
                  "Total = sum(positional_power) * (1+bonus)").font = Font(italic=True)

    best_units = get_best_generic_units(land_units, categories)

    row = 6
    for age in AGE_ORDER:
        types = best_units[age]

        # Age header
        ws.cell(row=row, column=1, value=AGE_LABELS[age]).font = Font(bold=True, size=13)
        row += 1

        # Column headers
        headers = [
            "Type", "Unit", "Optimal %", "Positional Power",
            "Flank Power", "Center Power",
            "Strength", "CombatPower",
            "Str Dmg Done", "Str Dmg Taken", "Initiative",
        ]
        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        row += 1

        flank_powers = [t["flank_power"] for t in types]
        center_powers = [t["center_power"] for t in types]

        pcts, total, bonus, nq = optimize_composition(
            flank_powers, center_powers, combined_arms
        )

        # Compute per-type positional contribution
        center_remaining = CENTER_RATIO
        type_positional = [0.0] * len(types)
        for idx, pct, fp, cp, adv in sorted(
            [(i, pcts[i], flank_powers[i], center_powers[i], center_powers[i] - flank_powers[i])
             for i in range(len(types)) if pcts[i] > 1e-9],
            key=lambda x: x[4], reverse=True,
        ):
            in_center = min(pct, center_remaining)
            in_flank = pct - in_center
            center_remaining -= in_center
            type_positional[idx] = in_center * cp + in_flank * fp

        for i, t in enumerate(types):
            pct = pcts[i]
            values = [
                t["type_label"], loc(t["unit_name"]),
                pct, type_positional[i],
                t["flank_power"], t["center_power"],
                t["strength"], t["combat_power"],
                t["str_dmg_done"], t["str_dmg_taken"], t["initiative"],
            ]
            for j, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=j, value=v)
                cell.border = THIN_BORDER
                if isinstance(v, float):
                    cell.number_format = NUM_FMT_2
                if j == 3:  # Optimal %
                    cell.number_format = "0%"
            if pct >= combined_arms["min_percent"] - 1e-9:
                ws.cell(row=row, column=1).font = BEST_FONT
                ws.cell(row=row, column=3).font = BEST_FONT
            row += 1

        # Summary rows
        base_power = calc_positional_power(pcts, flank_powers, center_powers)

        ws.cell(row=row, column=1, value="Base Power (no bonus)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=base_power).number_format = NUM_FMT_2
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value=f"Combined Arms Bonus ({nq} types)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=bonus).number_format = "0.0%"
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value="Total Power (with bonus)").font = Font(bold=True, size=12)
        ws.cell(row=row, column=4, value=total).number_format = NUM_FMT_2
        ws.cell(row=row, column=4).font = Font(bold=True, size=12)
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 2

    auto_width(ws, max_width=35)
    ws.freeze_panes = "A6"


def build_optimal_composition_gold(wb, land_units, categories, combined_arms, prices):
    """Sheet: Optimal army composition per age by power-per-gold."""
    ws = wb.create_sheet("Optimal Comp (Gold)")

    ws.cell(row=1, column=1, value="Optimal Army Composition (Power per Gold)").font = TITLE_FONT

    ca = combined_arms
    ws.cell(row=2, column=1,
            value=f"Scraped: bonus/type={ca['bonus_per_type']:.1%}, "
                  f"min threshold={ca['min_percent']:.0%}, "
                  f"max cap={ca['max_threshold']:.0%}").font = SUBTITLE_FONT
    ws.cell(row=3, column=1,
            value="Positional placement: 33% center / 67% flanks. "
                  "Optimizes power/gold instead of raw power.").font = Font(italic=True)
    ws.cell(row=4, column=1,
            value="Power/Gold = Power / (Strength * BaseBuildCost * 10) * 100").font = Font(italic=True)

    best_units = get_best_generic_units(land_units, categories)

    row = 6
    for age in AGE_ORDER:
        types = best_units[age]

        ws.cell(row=row, column=1, value=AGE_LABELS[age]).font = Font(bold=True, size=13)
        row += 1

        headers = [
            "Type", "Unit", "Optimal %", "Positional P/Gold",
            "Flank P/Gold", "Center P/Gold",
            "Strength", "CombatPower",
            "Str Dmg Done", "Str Dmg Taken", "Initiative",
        ]
        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        row += 1

        # Compute power-per-gold for each type
        flank_pg = []
        center_pg = []
        for t in types:
            cat_key = [c for label, c, _ in CA_TYPES if label == t["type_label"]][0]
            cost = calc_cost(t["strength"], cat_key, prices)
            if cost > 0:
                flank_pg.append(t["flank_power"] / cost * 100)
                center_pg.append(t["center_power"] / cost * 100)
            else:
                flank_pg.append(0)
                center_pg.append(0)

        pcts, total, bonus, nq = optimize_composition(
            flank_pg, center_pg, combined_arms
        )

        # Per-type positional contribution
        center_remaining = CENTER_RATIO
        type_positional = [0.0] * len(types)
        for idx, pct, fpg, cpg, adv in sorted(
            [(i, pcts[i], flank_pg[i], center_pg[i], center_pg[i] - flank_pg[i])
             for i in range(len(types)) if pcts[i] > 1e-9],
            key=lambda x: x[4], reverse=True,
        ):
            in_center = min(pct, center_remaining)
            in_flank = pct - in_center
            center_remaining -= in_center
            type_positional[idx] = in_center * cpg + in_flank * fpg

        for i, t in enumerate(types):
            pct = pcts[i]
            values = [
                t["type_label"], loc(t["unit_name"]),
                pct, type_positional[i],
                flank_pg[i], center_pg[i],
                t["strength"], t["combat_power"],
                t["str_dmg_done"], t["str_dmg_taken"], t["initiative"],
            ]
            for j, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=j, value=v)
                cell.border = THIN_BORDER
                if isinstance(v, float):
                    cell.number_format = NUM_FMT_2
                if j == 3:
                    cell.number_format = "0%"
            if pct >= combined_arms["min_percent"] - 1e-9:
                ws.cell(row=row, column=1).font = BEST_FONT
                ws.cell(row=row, column=3).font = BEST_FONT
            row += 1

        base_power = calc_positional_power(pcts, flank_pg, center_pg)

        ws.cell(row=row, column=1, value="Base P/Gold (no bonus)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=base_power).number_format = NUM_FMT_2
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value=f"Combined Arms Bonus ({nq} types)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=bonus).number_format = "0.0%"
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value="Total P/Gold (with bonus)").font = Font(bold=True, size=12)
        ws.cell(row=row, column=4, value=total).number_format = NUM_FMT_2
        ws.cell(row=row, column=4).font = Font(bold=True, size=12)
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 2

    auto_width(ws, max_width=35)
    ws.freeze_panes = "A6"


def calc_unit_iron(unit, categories, goods_demands, production_recipes):
    """Get total iron cost for a unit by resolving its construction demand."""
    ref = (unit.get("construction_demand", "")
           or categories.get(unit["category"], {}).get("construction_demand", ""))
    goods = goods_demands.get(ref, {})
    if not goods:
        return 0
    raw = resolve_raw_materials(goods, production_recipes)
    return raw.get("iron", 0)


def build_optimal_composition_iron(wb, land_units, categories, combined_arms,
                                    goods_demands, production_recipes):
    """Sheet: Optimal army composition per age by power-per-iron."""
    ws = wb.create_sheet("Optimal Comp (Iron)")

    ws.cell(row=1, column=1, value="Optimal Army Composition (Power per Iron)").font = TITLE_FONT

    ca = combined_arms
    ws.cell(row=2, column=1,
            value=f"Scraped: bonus/type={ca['bonus_per_type']:.1%}, "
                  f"min threshold={ca['min_percent']:.0%}, "
                  f"max cap={ca['max_threshold']:.0%}").font = SUBTITLE_FONT
    ws.cell(row=3, column=1,
            value="Positional placement: 33% center / 67% flanks. "
                  "Optimizes power/iron (resolved via workshop recipes).").font = Font(italic=True)

    best_units = get_best_generic_units(land_units, categories)

    row = 5
    for age in AGE_ORDER:
        types = best_units[age]

        ws.cell(row=row, column=1, value=AGE_LABELS[age]).font = Font(bold=True, size=13)
        row += 1

        headers = [
            "Type", "Unit", "Optimal %", "Positional P/Iron",
            "Iron Cost", "Flank P/Iron", "Center P/Iron",
            "Flank Power", "Center Power",
            "Strength", "CombatPower",
            "Str Dmg Done", "Str Dmg Taken", "Initiative",
        ]
        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        row += 1

        # Compute power-per-iron for each type
        flank_pi = []
        center_pi = []
        iron_costs = []
        for t in types:
            if t["unit_name"] == "-":
                flank_pi.append(0)
                center_pi.append(0)
                iron_costs.append(0)
                continue
            u = next((u for u in land_units if u["name"] == t["unit_name"]), None)
            iron = calc_unit_iron(u, categories, goods_demands, production_recipes) if u else 0
            iron_costs.append(iron)
            if iron > 0:
                flank_pi.append(t["flank_power"] / iron)
                center_pi.append(t["center_power"] / iron)
            else:
                flank_pi.append(t["flank_power"] * 1000 if t["flank_power"] > 0 else 0)
                center_pi.append(t["center_power"] * 1000 if t["center_power"] > 0 else 0)

        pcts, total, bonus, nq = optimize_composition(
            flank_pi, center_pi, combined_arms
        )

        # Per-type positional contribution
        center_remaining = CENTER_RATIO
        type_positional = [0.0] * len(types)
        for idx, pct, fpi, cpi, adv in sorted(
            [(i, pcts[i], flank_pi[i], center_pi[i], center_pi[i] - flank_pi[i])
             for i in range(len(types)) if pcts[i] > 1e-9],
            key=lambda x: x[4], reverse=True,
        ):
            in_center = min(pct, center_remaining)
            in_flank = pct - in_center
            center_remaining -= in_center
            type_positional[idx] = in_center * cpi + in_flank * fpi

        for i, t in enumerate(types):
            pct = pcts[i]
            values = [
                t["type_label"], loc(t["unit_name"]),
                pct, type_positional[i],
                iron_costs[i],
                flank_pi[i] if iron_costs[i] > 0 else "",
                center_pi[i] if iron_costs[i] > 0 else "",
                t["flank_power"], t["center_power"],
                t["strength"], t["combat_power"],
                t["str_dmg_done"], t["str_dmg_taken"], t["initiative"],
            ]
            for j, v in enumerate(values, 1):
                cell = ws.cell(row=row, column=j, value=v)
                cell.border = THIN_BORDER
                if isinstance(v, float):
                    cell.number_format = NUM_FMT_2
                if j == 3:  # optimal %
                    cell.number_format = "0%"
                if j == 5:  # iron cost
                    cell.number_format = NUM_FMT_3
            if pct >= combined_arms["min_percent"] - 1e-9:
                ws.cell(row=row, column=1).font = BEST_FONT
                ws.cell(row=row, column=3).font = BEST_FONT
            row += 1

        base_power = calc_positional_power(pcts, flank_pi, center_pi)

        ws.cell(row=row, column=1, value="Base P/Iron (no bonus)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=base_power).number_format = NUM_FMT_2
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value=f"Combined Arms Bonus ({nq} types)").font = Font(bold=True)
        ws.cell(row=row, column=4, value=bonus).number_format = "0.0%"
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 1

        ws.cell(row=row, column=1, value="Total P/Iron (with bonus)").font = Font(bold=True, size=12)
        ws.cell(row=row, column=4, value=total).number_format = NUM_FMT_2
        ws.cell(row=row, column=4).font = Font(bold=True, size=12)
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        row += 2

    auto_width(ws, max_width=35)
    ws.freeze_panes = "A5"


def build_goods_demands(wb, land_units, categories, goods_demands):
    """Sheet: Resource requirements per buildable unit per age."""
    ws = wb.create_sheet("Goods Demands")

    ws.cell(row=1, column=1, value="Unit Goods Demands (Construction & Maintenance)").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Scraped from goods_demand/army_demands.txt. "
                  "Quantities are per unit at max_strength.").font = SUBTITLE_FONT

    # Collect all goods that appear across all demands
    all_goods = set()
    for goods in goods_demands.values():
        all_goods.update(goods.keys())
    all_goods = sorted(all_goods)

    units = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
    ]

    headers = ["Age", "Unit", "Category", "Light", "Demand Ref"] + all_goods

    row = 4
    for table_type, demand_field in [
        ("Construction", "construction_demand"),
        ("Maintenance", "maintenance_demand"),
    ]:
        ws.cell(row=row, column=1, value=f"{table_type} Demands").font = Font(bold=True, size=13)
        row += 1

        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        header_row = row
        row += 1

        for age in AGE_ORDER:
            age_units = [u for u in units if u.get("age") == age]
            if not age_units:
                continue

            for u in sorted(age_units, key=lambda x: (x["category"], x["name"])):
                cat_label = CAT_LABELS.get(u["category"], u["category"])
                demand_ref = u.get(demand_field, "") or categories.get(u["category"], {}).get(demand_field, "")
                goods = goods_demands.get(demand_ref, {})

                values = [
                    AGE_LABELS.get(u.get("age", ""), "?"),
                    loc(u["name"]),
                    cat_label,
                    "Yes" if u.get("light") else "",
                    demand_ref,
                ] + [goods.get(g, "") for g in all_goods]

                for j, v in enumerate(values, 1):
                    cell = ws.cell(row=row, column=j, value=v)
                    cell.border = THIN_BORDER
                    if isinstance(v, float):
                        cell.number_format = NUM_FMT_3
                    cat_fill = CAT_FILLS.get(cat_label)
                    if cat_fill:
                        cell.fill = cat_fill

                row += 1
            row += 1  # gap between ages

        row += 1  # gap between tables

    auto_width(ws, min_width=6, max_width=20)
    ws.freeze_panes = f"F{6}"


def build_goods_demands_generic(wb, land_units, categories, goods_demands, production_recipes):
    """Sheet: Goods demands for generic (non-special) infantry/cav/art/aux only."""
    ws = wb.create_sheet("Goods (Generic)")

    ws.cell(row=1, column=1, value="Generic Unit Goods Demands").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Standard upgrade-chain units only (no specials). "
                  "iron (total) = resolved from produced goods via workshop recipes.").font = SUBTITLE_FONT

    best_units = get_best_generic_units(land_units, categories)

    # Collect only goods that these units actually use
    relevant_goods = set()
    for age_types in best_units.values():
        for t in age_types:
            unit_name = t["unit_name"]
            if unit_name == "-":
                continue
            u = next((u for u in land_units if u["name"] == unit_name), None)
            if not u:
                continue
            for field in ["construction_demand", "maintenance_demand"]:
                ref = u.get(field, "") or categories.get(u["category"], {}).get(field, "")
                for g in goods_demands.get(ref, {}):
                    relevant_goods.add(g)
    relevant_goods = sorted(relevant_goods)

    headers = ["Age", "Type", "Unit", "Demand Ref"] + relevant_goods + ["iron (total)"]

    row = 4
    for table_type, demand_field in [
        ("Construction", "construction_demand"),
        ("Maintenance", "maintenance_demand"),
    ]:
        ws.cell(row=row, column=1, value=f"{table_type} Demands").font = Font(bold=True, size=13)
        row += 1

        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        row += 1

        for age in AGE_ORDER:
            types = best_units[age]
            for t in types:
                unit_name = t["unit_name"]
                if unit_name == "-":
                    continue
                u = next((u for u in land_units if u["name"] == unit_name), None)
                if not u:
                    continue

                cat_label = CAT_LABELS.get(u["category"], u["category"])
                demand_ref = u.get(demand_field, "") or categories.get(u["category"], {}).get(demand_field, "")
                goods = goods_demands.get(demand_ref, {})

                # Resolve raw iron from all produced goods
                raw = resolve_raw_materials(goods, production_recipes) if goods else {}
                total_iron = round(raw.get("iron", 0), 4) or ""

                values = [
                    AGE_LABELS.get(age, "?"),
                    t["type_label"],
                    loc(unit_name),
                    demand_ref,
                ] + [goods.get(g, "") for g in relevant_goods] + [total_iron]

                for j, v in enumerate(values, 1):
                    cell = ws.cell(row=row, column=j, value=v)
                    cell.border = THIN_BORDER
                    if isinstance(v, float):
                        cell.number_format = NUM_FMT_3
                    cat_fill = CAT_FILLS.get(cat_label)
                    if cat_fill:
                        cell.fill = cat_fill

                row += 1
            row += 1  # gap between ages

        row += 1  # gap between tables

    auto_width(ws, min_width=6, max_width=20)
    ws.freeze_panes = "E6"


def pick_recipe(recipes, good, prefer_tier="workshop", prefer_variant="iron"):
    """Pick a single production recipe for a good.

    Prefers: workshop tier, iron-based variant. Falls back through tiers.
    Returns dict of { input_good: amount_per_unit_output } or None.
    """
    candidates = recipes.get(good, [])
    if not candidates:
        return None

    tier_order = [prefer_tier, "guild", "manufactory", "factory", "unknown"]

    for tier in tier_order:
        tier_recipes = [r for r in candidates if r["tier"] == tier]
        if not tier_recipes:
            continue
        # Prefer iron-based variant
        for r in tier_recipes:
            if prefer_variant in r["method"]:
                return {k: v / r["output"] for k, v in r["inputs"].items()}
        # Prefer livestock over wild_game for leather
        for r in tier_recipes:
            if "livestock" in r["method"] or "cotton" in r["method"]:
                return {k: v / r["output"] for k, v in r["inputs"].items()}
        # Fall back to first recipe without "ammunition" or "stone" or "obsidian"
        for r in tier_recipes:
            if not any(x in r["method"] for x in ["ammunition", "stone", "obsidian", "pre_columbian"]):
                return {k: v / r["output"] for k, v in r["inputs"].items()}
        # Last resort
        return {k: v / tier_recipes[0]["output"] for k, v in tier_recipes[0]["inputs"].items()}

    return None


def resolve_raw_materials(goods_needed, recipes, max_depth=3):
    """Resolve produced goods into raw materials recursively.

    goods_needed: { good_name: amount }
    Returns: { raw_material: total_amount }
    """
    raw = {}
    to_resolve = dict(goods_needed)

    for _ in range(max_depth):
        next_resolve = {}
        for good, amount in to_resolve.items():
            recipe = pick_recipe(recipes, good)
            if recipe is None:
                # It's a raw material (or unknown) — keep as-is
                raw[good] = raw.get(good, 0) + amount
            else:
                # Produced good — break down into inputs
                for input_good, input_amt in recipe.items():
                    needed = input_amt * amount
                    next_resolve[input_good] = next_resolve.get(input_good, 0) + needed
        if not next_resolve:
            break
        to_resolve = next_resolve

    # Any remaining unresolved go into raw
    for good, amount in to_resolve.items():
        raw[good] = raw.get(good, 0) + amount

    return raw


def build_raw_materials(wb, land_units, categories, goods_demands, production_recipes):
    """Sheet: Raw material breakdown for generic units per age."""
    ws = wb.create_sheet("Raw Materials")

    ws.cell(row=1, column=1, value="Raw Material Requirements (Generic Units)").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Produced goods resolved to raw materials using workshop-tier iron-based recipes. "
                  "Recursive: tools->iron, etc.").font = SUBTITLE_FONT

    best_units = get_best_generic_units(land_units, categories)

    # First pass: collect all raw materials that appear
    all_raws = set()
    unit_raws = {}  # (age, type_label, demand_field) -> raw dict
    for age in AGE_ORDER:
        for t in best_units[age]:
            if t["unit_name"] == "-":
                continue
            u = next((u for u in land_units if u["name"] == t["unit_name"]), None)
            if not u:
                continue
            for demand_field in ["construction_demand", "maintenance_demand"]:
                ref = u.get(demand_field, "") or categories.get(u["category"], {}).get(demand_field, "")
                goods = goods_demands.get(ref, {})
                if goods:
                    raw = resolve_raw_materials(goods, production_recipes)
                    unit_raws[(age, t["type_label"], demand_field)] = raw
                    all_raws.update(raw.keys())

    all_raws = sorted(all_raws)
    headers = ["Age", "Type", "Unit"] + all_raws

    row = 4
    for table_type, demand_field in [
        ("Construction", "construction_demand"),
        ("Maintenance", "maintenance_demand"),
    ]:
        ws.cell(row=row, column=1, value=f"{table_type} - Raw Materials").font = Font(bold=True, size=13)
        row += 1

        for i, h in enumerate(headers, 1):
            ws.cell(row=row, column=i, value=h)
        style_header_row(ws, row, len(headers))
        row += 1

        for age in AGE_ORDER:
            for t in best_units[age]:
                if t["unit_name"] == "-":
                    continue
                u = next((u for u in land_units if u["name"] == t["unit_name"]), None)
                if not u:
                    continue

                cat_label = CAT_LABELS.get(u["category"], u["category"])
                raw = unit_raws.get((age, t["type_label"], demand_field), {})

                values = [
                    AGE_LABELS.get(age, "?"),
                    t["type_label"],
                    loc(t["unit_name"]),
                ] + [round(raw.get(g, 0), 4) if raw.get(g, 0) else "" for g in all_raws]

                for j, v in enumerate(values, 1):
                    cell = ws.cell(row=row, column=j, value=v)
                    cell.border = THIN_BORDER
                    if isinstance(v, float):
                        cell.number_format = NUM_FMT_3
                    cat_fill = CAT_FILLS.get(cat_label)
                    if cat_fill:
                        cell.fill = cat_fill

                row += 1
            row += 1  # gap between ages

        row += 1  # gap between tables

    auto_width(ws, min_width=6, max_width=15)
    ws.freeze_panes = "D6"


def main():
    if not DATA_DIR.exists():
        print("No data/ directory found. Run scraper.py first.")
        return

    land_units, categories, age_progression, prices, combined_arms, goods_demands, production_recipes, localizations = load_data()

    global LOC
    LOC = localizations

    wb = Workbook()

    print("Building Army Meta sheet...")
    build_army_meta(wb, age_progression, categories, prices)

    print("Building Buildable Units sheet...")
    build_buildable_units(wb, land_units, categories, prices)

    print("Building Optimal Composition sheet...")
    build_optimal_composition(wb, land_units, categories, combined_arms)

    print("Building Optimal Composition (Gold) sheet...")
    build_optimal_composition_gold(wb, land_units, categories, combined_arms, prices)

    print("Building Optimal Composition (Iron) sheet...")
    build_optimal_composition_iron(wb, land_units, categories, combined_arms,
                                    goods_demands, production_recipes)

    print("Building Goods Demands sheet...")
    build_goods_demands(wb, land_units, categories, goods_demands)

    print("Building Goods (Generic) sheet...")
    build_goods_demands_generic(wb, land_units, categories, goods_demands, production_recipes)

    print("Building Raw Materials sheet...")
    build_raw_materials(wb, land_units, categories, goods_demands, production_recipes)

    print("Building Upgrade Chains sheet...")
    build_upgrade_chains(wb, land_units)

    print("Building Special Units sheet...")
    build_special_units(wb, land_units, categories)

    print("Building Category Stats sheet...")
    build_category_reference(wb, categories, prices)

    print("Building Light vs Heavy sheet...")
    build_light_vs_heavy(wb, land_units, categories)

    output_path = OUTPUT_DIR / "eu5_army_analysis.xlsx"

    # Close Excel if it has our file open
    if sys.platform == "win32":
        import time
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE"],
            capture_output=True, timeout=5,
        )
        time.sleep(1)

    wb.save(output_path)
    print(f"\nSaved to: {output_path}")

    # Open in Excel
    if sys.platform == "win32":
        os.startfile(output_path)


if __name__ == "__main__":
    main()
