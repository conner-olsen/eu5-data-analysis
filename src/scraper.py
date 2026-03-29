"""Scrape EU5 unit data from vanilla game files and resolve inheritance."""

import json
import copy
from pathlib import Path

from parser import parse_directory

GAME_DIR = Path("C:/Steam/steamapps/common/Europa Universalis V/game")
COMMON_DIR = GAME_DIR / "in_game" / "common"
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data"

# Stats we care about for army analysis
NUMERIC_STATS = [
    "max_strength",
    "combat_power",
    "morale_damage_done",
    "morale_damage_taken",
    "strength_damage_done",
    "strength_damage_taken",
    "combat_speed",
    "initiative",
    "movement_speed",
    "frontage",
    "flanking_ability",
    "secure_flanks_defense",
    "bombard_efficiency",
    "artillery_barrage",
    "supply_weight",
    "attrition_loss",
    "food_storage_per_strength",
    "food_consumption_per_strength",
    "damage_taken",
    "build_time_modifier",
]

# Properties to carry forward from inheritance
CARRY_PROPS = NUMERIC_STATS + [
    "category",
    "age",
    "light",
    "buildable",
    "levy",
    "is_special",
    "default",
    "upgrades_to",
    "maintenance_demand",
    "construction_demand",
    "combat",
    "impact",
]


def scrape_categories() -> dict:
    """Scrape unit category base stats."""
    raw = parse_directory(COMMON_DIR / "unit_categories")
    categories = {}
    for name, data in raw.items():
        if not isinstance(data, dict):
            continue
        cat = {"name": name}
        for stat in NUMERIC_STATS:
            if stat in data:
                cat[stat] = data[stat]
        # Boolean flags
        for flag in ["is_army", "bombard", "assault", "is_garrison", "auxiliary", "transport"]:
            if flag in data and data[flag] is True:
                cat[flag] = True
        if "build_time" in data:
            cat["build_time"] = data["build_time"]
        if "ai_weight" in data:
            cat["ai_weight"] = data["ai_weight"]
        if "startup_amount" in data:
            cat["startup_amount"] = data["startup_amount"]
        categories[name] = cat
    return categories


def resolve_terrain_block(block) -> dict[str, float]:
    """Extract terrain modifiers from a combat/impact block."""
    if not isinstance(block, dict):
        return {}
    result = {}
    for k, v in block.items():
        if k.startswith("__"):
            continue
        if isinstance(v, (int, float)):
            result[k] = v
    return result


    # Flags that are per-unit identity, not inherited from templates
NO_INHERIT = {"buildable", "levy", "default", "is_special"}


def resolve_inheritance(all_units: dict) -> dict:
    """Resolve copy_from chains to produce final stat blocks."""
    resolved = {}
    resolving = set()  # cycle detection

    def resolve(name: str) -> dict:
        if name in resolved:
            return resolved[name]
        if name not in all_units:
            return {}
        if name in resolving:
            return all_units[name]  # cycle - return raw

        resolving.add(name)
        raw = all_units[name]

        if "copy_from" in raw and isinstance(raw["copy_from"], str):
            parent = resolve(raw["copy_from"])
            merged = copy.deepcopy(parent)
            # Remove non-inheritable flags from parent
            for flag in NO_INHERIT:
                merged.pop(flag, None)
            # Override with child values
            for key, val in raw.items():
                if key == "copy_from":
                    continue
                if key in ("combat", "impact") and isinstance(val, dict):
                    # Merge terrain blocks - child overrides parent per-terrain
                    if key not in merged:
                        merged[key] = {}
                    if isinstance(merged[key], dict):
                        merged[key].update(val)
                    else:
                        merged[key] = val
                else:
                    merged[key] = val
            merged["_parent"] = raw["copy_from"]
        else:
            merged = copy.deepcopy(raw)

        resolving.discard(name)
        resolved[name] = merged
        return merged

    for name in all_units:
        resolve(name)

    return resolved


def extract_unit_stats(name: str, data: dict, categories: dict) -> dict:
    """Extract relevant stats from a resolved unit definition."""
    unit = {"name": name}

    # Category
    cat_name = data.get("category", "")
    unit["category"] = cat_name

    # Get category base stats
    cat_stats = categories.get(cat_name, {})

    # Merge stats: category base + unit overrides
    for stat in NUMERIC_STATS:
        cat_val = cat_stats.get(stat, 0)
        unit_val = data.get(stat)
        if unit_val is not None:
            unit[stat] = unit_val
        elif cat_val:
            unit[stat] = cat_val

    # Also check nested modifier block for combat_power etc.
    modifier = data.get("modifier", {})
    if isinstance(modifier, dict):
        for stat in NUMERIC_STATS:
            if stat in modifier:
                unit[stat] = modifier[stat]

    # Boolean properties
    unit["light"] = data.get("light", False)
    unit["is_special"] = data.get("is_special", False)
    unit["buildable"] = data.get("buildable", True)
    unit["levy"] = data.get("levy", False)
    unit["default"] = data.get("default", False)

    # Age
    unit["age"] = data.get("age", "")

    # Upgrade path
    upgrades = data.get("upgrades_to", "")
    if isinstance(upgrades, list):
        unit["upgrades_to"] = upgrades[-1] if upgrades else ""
    else:
        unit["upgrades_to"] = upgrades or ""

    # Maintenance/construction (just store the reference name)
    unit["maintenance_demand"] = data.get("maintenance_demand", "")
    unit["construction_demand"] = data.get("construction_demand", "")

    # Terrain modifiers
    combat = data.get("combat", {})
    impact = data.get("impact", {})
    unit["terrain_combat"] = resolve_terrain_block(combat)
    unit["terrain_impact"] = resolve_terrain_block(impact)

    # Inheritance info
    unit["_parent"] = data.get("_parent", "")

    return unit


def determine_age(unit: dict) -> str:
    """Determine which age a unit belongs to based on its parent chain or age field."""
    if unit.get("age"):
        return unit["age"]
    # Infer from parent name
    parent = unit.get("_parent", "")
    for i in range(1, 7):
        if f"age_{i}" in parent:
            age_names = {
                1: "age_1_traditions",
                2: "age_2_renaissance",
                3: "age_3_discovery",
                4: "age_4_reformation",
                5: "age_5_absolutism",
                6: "age_6_revolutions",
            }
            return age_names[i]
    return "unknown"


def build_age_progression(units: list[dict]) -> list[dict]:
    """Build a table showing stat progression across ages for each category."""
    age_order = [
        "age_1_traditions",
        "age_2_renaissance",
        "age_3_discovery",
        "age_4_reformation",
        "age_5_absolutism",
        "age_6_revolutions",
    ]
    land_categories = ["army_infantry", "army_cavalry", "army_artillery", "army_auxiliary"]

    rows = []
    for cat in land_categories:
        for age in age_order:
            # Find the age template for this category+age
            template = None
            for u in units:
                if (
                    u["category"] == cat
                    and u["age"] == age
                    and not u["buildable"]
                    and not u["levy"]
                    and not u["is_special"]
                    and u["name"].startswith("a_age_")
                ):
                    template = u
                    break
            if template:
                rows.append({
                    "category": cat,
                    "age": age,
                    "max_strength": template.get("max_strength", 0),
                    "combat_power": template.get("combat_power", 0),
                    "bombard_efficiency": template.get("bombard_efficiency", 0),
                    "artillery_barrage": template.get("artillery_barrage", 0),
                })
    return rows


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Scraping unit categories...")
    categories = scrape_categories()

    print("Parsing unit type files...")
    raw_units = parse_directory(COMMON_DIR / "unit_types")

    # Filter to only dict entries (skip non-unit top-level keys)
    raw_units = {k: v for k, v in raw_units.items() if isinstance(v, dict)}

    print(f"  Found {len(raw_units)} unit definitions")

    print("Resolving inheritance chains...")
    resolved = resolve_inheritance(raw_units)

    print("Extracting stats...")
    units = []
    for name, data in resolved.items():
        unit = extract_unit_stats(name, data, categories)
        # Backfill age from parent chain
        if not unit["age"]:
            unit["age"] = determine_age(unit)
        units.append(unit)

    # Separate land vs naval
    land_cats = {"army_infantry", "army_cavalry", "army_artillery", "army_auxiliary"}
    land_units = [u for u in units if u["category"] in land_cats]
    naval_units = [u for u in units if u["category"] not in land_cats]

    # Build age progression from templates
    age_progression = build_age_progression(units)

    # Save outputs
    with open(OUTPUT_DIR / "unit_categories.json", "w") as f:
        json.dump(categories, f, indent=2)
    print(f"  Wrote unit_categories.json ({len(categories)} categories)")

    with open(OUTPUT_DIR / "land_units.json", "w") as f:
        json.dump(land_units, f, indent=2)
    print(f"  Wrote land_units.json ({len(land_units)} units)")

    with open(OUTPUT_DIR / "naval_units.json", "w") as f:
        json.dump(naval_units, f, indent=2)
    print(f"  Wrote naval_units.json ({len(naval_units)} units)")

    with open(OUTPUT_DIR / "age_progression.json", "w") as f:
        json.dump(age_progression, f, indent=2)
    print(f"  Wrote age_progression.json ({len(age_progression)} rows)")

    print("Done!")


if __name__ == "__main__":
    main()
