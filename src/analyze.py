"""Analyze EU5 army unit data for optimal composition."""

import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from tabulate import tabulate

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

AGE_ORDER = [
    "age_1_traditions",
    "age_2_renaissance",
    "age_3_discovery",
    "age_4_reformation",
    "age_5_absolutism",
    "age_6_revolutions",
]
AGE_LABELS = {
    "age_1_traditions": "1. Traditions",
    "age_2_renaissance": "2. Renaissance",
    "age_3_discovery": "3. Discovery",
    "age_4_reformation": "4. Reformation",
    "age_5_absolutism": "5. Absolutism",
    "age_6_revolutions": "6. Revolutions",
}
CAT_LABELS = {
    "army_infantry": "Infantry",
    "army_cavalry": "Cavalry",
    "army_artillery": "Artillery",
    "army_auxiliary": "Auxiliary",
}


def load_data():
    with open(DATA_DIR / "land_units.json") as f:
        land_units = json.load(f)
    with open(DATA_DIR / "unit_categories.json") as f:
        categories = json.load(f)
    with open(DATA_DIR / "age_progression.json") as f:
        age_progression = json.load(f)
    return land_units, categories, age_progression


def print_section(title: str):
    print(f"\n{'='*80}")
    print(f"  {title}")
    print(f"{'='*80}\n")


def category_base_stats(categories: dict):
    """Show base category stats side-by-side."""
    print_section("CATEGORY BASE STATS")

    land_cats = ["army_infantry", "army_cavalry", "army_artillery", "army_auxiliary"]
    stats = [
        "damage_taken",
        "combat_speed",
        "initiative",
        "frontage",
        "flanking_ability",
        "secure_flanks_defense",
        "supply_weight",
        "attrition_loss",
        "food_storage_per_strength",
        "food_consumption_per_strength",
        "startup_amount",
    ]

    rows = []
    for stat in stats:
        row = {"Stat": stat}
        for cat in land_cats:
            row[CAT_LABELS[cat]] = categories.get(cat, {}).get(stat, "-")
        rows.append(row)

    print(tabulate(rows, headers="keys", tablefmt="simple_outline", numalign="right"))
    print()
    print("Key takeaways:")
    print("  - Cavalry: 0.75x damage_taken (tankier per strength), 2x flanking, 5x combat_speed")
    print("  - Artillery/Auxiliary: 1.25x damage_taken (fragile)")
    print("  - Cavalry: 2x food cost, +25% attrition")
    print("  - Artillery: 2x food cost, +50% attrition")


def age_progression_table(age_data: list):
    """Show how template stats scale across ages."""
    print_section("TEMPLATE STATS BY AGE (Base Power Progression)")

    df = pd.DataFrame(age_data)
    df["age_label"] = df["age"].map(AGE_LABELS)
    df["cat_label"] = df["category"].map(CAT_LABELS)

    for cat in ["army_infantry", "army_cavalry", "army_artillery", "army_auxiliary"]:
        cat_df = df[df["category"] == cat].sort_values(
            "age", key=lambda s: s.map({a: i for i, a in enumerate(AGE_ORDER)})
        )
        if cat_df.empty:
            continue

        print(f"  {CAT_LABELS[cat]}:")
        cols = ["age_label", "max_strength", "combat_power"]
        if cat == "army_artillery":
            cols += ["bombard_efficiency", "artillery_barrage"]

        display = cat_df[cols].rename(columns={"age_label": "Age"})
        print(tabulate(display, headers="keys", tablefmt="simple_outline", showindex=False, numalign="right"))
        print()

    # Effective power comparison (strength * combat_power)
    print_section("EFFECTIVE POWER PER UNIT (max_strength x combat_power)")

    rows = []
    for _, row in df.iterrows():
        rows.append({
            "Age": AGE_LABELS.get(row["age"], row["age"]),
            "Category": CAT_LABELS.get(row["category"], row["category"]),
            "max_strength": row["max_strength"],
            "combat_power": row["combat_power"],
            "effective_power": round(row["max_strength"] * row["combat_power"], 2),
        })

    power_df = pd.DataFrame(rows)
    pivot = power_df.pivot_table(
        index="Age", columns="Category", values="effective_power", sort=False
    )
    pivot = pivot.reindex(
        [AGE_LABELS[a] for a in AGE_ORDER],
        columns=["Infantry", "Cavalry", "Artillery", "Auxiliary"],
    )
    print(tabulate(pivot, headers="keys", tablefmt="simple_outline", numalign="right"))


def buildable_units_by_age(land_units: list):
    """Show all buildable (non-levy, non-template) units per age."""
    print_section("BUILDABLE UNITS BY AGE")

    units = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
    ]

    for age in AGE_ORDER:
        age_units = [u for u in units if u.get("age") == age]
        if not age_units:
            continue

        print(f"  {AGE_LABELS[age]}:")
        rows = []
        for u in sorted(age_units, key=lambda x: (x["category"], x["name"])):
            row = {
                "Unit": u["name"],
                "Cat": CAT_LABELS.get(u["category"], u["category"]),
                "Light": "yes" if u.get("light") else "",
                "Special": "yes" if u.get("is_special") else "",
                "Strength": u.get("max_strength", 0),
                "CombatPwr": u.get("combat_power", 0),
                "EffPower": round(u.get("max_strength", 0) * u.get("combat_power", 0), 2),
                "CbtSpeed": u.get("combat_speed", 0),
                "Init": u.get("initiative", 0),
                "Flank": u.get("flanking_ability", 0),
                "StrDmgTkn": u.get("strength_damage_taken", 0),
                "MrlDmgTkn": u.get("morale_damage_taken", 0),
                "StrDmgDn": u.get("strength_damage_done", 0),
                "MrlDmgDn": u.get("morale_damage_done", 0),
                "Upgrades": u.get("upgrades_to", ""),
            }
            rows.append(row)

        print(tabulate(rows, headers="keys", tablefmt="simple_outline", showindex=False, numalign="right"))
        print()


def upgrade_chains(land_units: list):
    """Show unit upgrade paths across ages."""
    print_section("UPGRADE CHAINS")

    units_by_name = {u["name"]: u for u in land_units}
    buildable = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
    ]

    # Find chain starters (units that nothing upgrades TO)
    upgraded_to = {u.get("upgrades_to") for u in buildable if u.get("upgrades_to")}
    starters = [u for u in buildable if u["name"] not in upgraded_to]

    for start in sorted(starters, key=lambda x: (x["category"], x["name"])):
        chain = []
        current = start
        visited = set()
        while current and current["name"] not in visited:
            visited.add(current["name"])
            age_label = AGE_LABELS.get(current.get("age", ""), "?")
            light_tag = " [L]" if current.get("light") else ""
            special_tag = " [S]" if current.get("is_special") else ""
            eff_power = round(
                current.get("max_strength", 0) * current.get("combat_power", 0), 2
            )
            chain.append(f"{current['name']}{light_tag}{special_tag} (EP:{eff_power}, {age_label})")
            next_name = current.get("upgrades_to", "")
            current = units_by_name.get(next_name) if next_name else None

        cat = CAT_LABELS.get(start["category"], start["category"])
        print(f"  [{cat}] {' -> '.join(chain)}")

    print()
    print("  [L] = Light unit, [S] = Special unit, EP = Effective Power")


def light_vs_heavy_comparison(land_units: list):
    """Compare light vs heavy variants within the same age and category."""
    print_section("LIGHT vs HEAVY COMPARISON (same age, same category)")

    buildable = [
        u for u in land_units
        if u.get("buildable", True)
        and not u.get("levy", False)
        and not u["name"].startswith("a_age_")
        and not u.get("is_special", False)
    ]

    for age in AGE_ORDER:
        age_units = [u for u in buildable if u.get("age") == age]
        for cat in ["army_infantry", "army_cavalry"]:
            cat_units = [u for u in age_units if u["category"] == cat]
            lights = [u for u in cat_units if u.get("light")]
            heavies = [u for u in cat_units if not u.get("light")]

            if not lights or not heavies:
                continue

            print(f"  {AGE_LABELS[age]} - {CAT_LABELS[cat]}:")
            rows = []
            for u in heavies + lights:
                row = {
                    "Unit": u["name"],
                    "Type": "Light" if u.get("light") else "Heavy",
                    "EffPower": round(u.get("max_strength", 0) * u.get("combat_power", 0), 2),
                    "CbtSpeed": u.get("combat_speed", 0),
                    "Init": u.get("initiative", 0),
                    "StrDmgTkn": u.get("strength_damage_taken", 0),
                    "MrlDmgTkn": u.get("morale_damage_taken", 0),
                    "StrDmgDn": u.get("strength_damage_done", 0),
                    "Flank": u.get("flanking_ability", 0),
                }
                rows.append(row)
            print(tabulate(rows, headers="keys", tablefmt="simple_outline", showindex=False, numalign="right"))
            print()


def special_units_analysis(land_units: list):
    """Analyze special/unique units and their advantages."""
    print_section("SPECIAL/UNIQUE UNITS")

    specials = [
        u for u in land_units
        if u.get("is_special", False)
        and not u.get("levy", False)
        and u.get("buildable", True)
    ]

    if not specials:
        print("  No buildable special units found.")
        return

    rows = []
    for u in sorted(specials, key=lambda x: (x.get("age", ""), x["category"], x["name"])):
        terrain_combat = u.get("terrain_combat", {})
        terrain_str = ", ".join(f"{k}:{v:+.2f}" for k, v in terrain_combat.items()) if terrain_combat else ""

        row = {
            "Unit": u["name"],
            "Age": AGE_LABELS.get(u.get("age", ""), "?"),
            "Cat": CAT_LABELS.get(u["category"], u["category"]),
            "EffPower": round(u.get("max_strength", 0) * u.get("combat_power", 0), 2),
            "CbtSpeed": u.get("combat_speed", 0),
            "Init": u.get("initiative", 0),
            "Terrain": terrain_str,
        }
        rows.append(row)

    print(tabulate(rows, headers="keys", tablefmt="simple_outline", showindex=False, numalign="right"))


def composition_analysis(age_data: list, categories: dict):
    """Analyze optimal army ratios per age."""
    print_section("ARMY COMPOSITION ANALYSIS")

    print("  Scoring methodology:")
    print("  - DPS = effective_power * combat_speed (sustained damage output per frontage)")
    print("  - Survivability = 1 / damage_taken (how efficiently the unit absorbs hits)")
    print("  - Efficiency = DPS * Survivability (damage dealt per damage received)")
    print()

    df = pd.DataFrame(age_data)

    for age in AGE_ORDER:
        age_df = df[df["age"] == age]
        if age_df.empty:
            continue

        print(f"  {AGE_LABELS[age]}:")
        rows = []
        for _, row in age_df.iterrows():
            cat = row["category"]
            cat_data = categories.get(cat, {})

            eff_power = row["max_strength"] * row["combat_power"]
            combat_speed = cat_data.get("combat_speed", 1)
            damage_taken = cat_data.get("damage_taken", 1.0)
            flanking = cat_data.get("flanking_ability", 1.0)

            dps = eff_power * combat_speed
            survivability = 1.0 / damage_taken if damage_taken > 0 else 0
            efficiency = dps * survivability

            rows.append({
                "Category": CAT_LABELS.get(cat, cat),
                "EffPower": round(eff_power, 2),
                "CbtSpeed": combat_speed,
                "DPS": round(dps, 2),
                "DmgTaken": damage_taken,
                "Survive": round(survivability, 2),
                "Efficiency": round(efficiency, 2),
                "Flanking": flanking,
            })

        print(tabulate(rows, headers="keys", tablefmt="simple_outline", showindex=False, numalign="right"))
        print()


def main():
    if not DATA_DIR.exists():
        print("No data/ directory found. Run scraper.py first.")
        return

    land_units, categories, age_progression = load_data()

    category_base_stats(categories)
    age_progression_table(age_progression)
    buildable_units_by_age(land_units)
    upgrade_chains(land_units)
    light_vs_heavy_comparison(land_units)
    special_units_analysis(land_units)
    composition_analysis(age_progression, categories)


if __name__ == "__main__":
    main()
