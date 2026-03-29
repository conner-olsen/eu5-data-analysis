"""Scrape EU5 unit data from vanilla game files and resolve inheritance."""

import json
import copy
from pathlib import Path

from parser import parse_directory, parse_file

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
        for demand in ["maintenance_demand", "construction_demand"]:
            if demand in data:
                cat[demand] = data[demand]
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


def scrape_unit_prices() -> dict:
    """Scrape unit build/reinforce/maintenance gold costs from prices/02_units.txt.

    Returns dict keyed by category name, e.g.:
    { "army_infantry": { "build_gold": 50, "reinforce_gold": 2.5, "maintenance_gold": 0.5 }, ... }
    """
    prices_dir = COMMON_DIR / "prices"
    raw = parse_directory(prices_dir)

    prices = {}
    for cat in ["army_infantry", "army_cavalry", "army_artillery", "army_auxiliary",
                 "navy_heavy_ship", "navy_light_ship", "navy_galley", "navy_transport"]:
        entry = {}
        for cost_type in ["build", "reinforce", "maintenance"]:
            key = f"{cat}_{cost_type}"
            if key in raw and isinstance(raw[key], dict):
                entry[f"{cost_type}_gold"] = raw[key].get("gold", 0)
                entry[f"{cost_type}_manpower"] = raw[key].get("manpower", raw[key].get("sailors", 0))
        if entry:
            prices[cat] = entry

    return prices


def scrape_goods_demands() -> dict:
    """Scrape unit goods demands from goods_demand/army_demands.txt.

    Returns dict keyed by demand name (e.g., "infantry_construction"),
    with goods quantities (excluding the "category" field).
    """
    raw = parse_directory(COMMON_DIR / "goods_demand")
    demands = {}
    for name, data in raw.items():
        if not isinstance(data, dict):
            continue
        goods = {k: v for k, v in data.items() if k != "category" and isinstance(v, (int, float))}
        if goods:
            demands[name] = goods
    return demands


def scrape_production_recipes() -> dict:
    """Scrape production recipes from building_types/production_*.txt files.

    Returns dict keyed by produced good, containing lists of recipes:
    { "firearms": [ { "method": "guns_workshop_iron_maintenance", "tier": "workshop",
                       "inputs": { "iron": 0.5, "tools": 0.3 }, "output": 1.0 }, ... ] }
    """
    building_dir = COMMON_DIR / "building_types"
    recipes_by_good = {}

    # Tier detection from category field
    TIER_MAP = {
        "guild_input": "guild",
        "workshop_input": "workshop",
        "manufactory_input": "manufactory",
        "factory_input": "factory",
    }
    SKIP_KEYS = {"produced", "output", "category", "debug_max_profit"}

    for filepath in sorted(building_dir.glob("production_*.txt")):
        raw = parse_file(filepath)
        # Walk all buildings in the file
        for building_name, building_data in raw.items():
            if not isinstance(building_data, dict):
                continue
            # unique_production_methods can appear multiple times (parsed as list)
            upm = building_data.get("unique_production_methods", {})
            if isinstance(upm, dict):
                upm_list = [upm]
            elif isinstance(upm, list):
                upm_list = upm
            else:
                continue

            for methods_block in upm_list:
                if not isinstance(methods_block, dict):
                    continue
                for method_name, method_data in methods_block.items():
                    if not isinstance(method_data, dict):
                        continue
                    produced = method_data.get("produced")
                    output_amt = method_data.get("output")
                    category = method_data.get("category", "")
                    if not produced or not output_amt:
                        continue

                    inputs = {}
                    for k, v in method_data.items():
                        if k not in SKIP_KEYS and isinstance(v, (int, float)):
                            inputs[k] = v

                    tier = TIER_MAP.get(category, "unknown")
                    recipe = {
                        "method": method_name,
                        "building": building_name,
                        "tier": tier,
                        "inputs": inputs,
                        "output": output_amt,
                    }
                    recipes_by_good.setdefault(produced, []).append(recipe)

    return recipes_by_good


def scrape_unit_localizations() -> dict:
    """Scrape display names for unit types from localization yml files.

    Loads all english loc files to resolve $key$ cross-references.
    Returns dict: { "a_footmen": "Footmen", "a_archers": "Archers", ... }
    """
    import re
    loc_dir = GAME_DIR / "main_menu" / "localization" / "english"

    # First pass: build a global lookup of ALL localization keys
    all_loc = {}
    for loc_file in sorted(loc_dir.glob("*_l_english.yml")):
        text = loc_file.read_text(encoding="utf-8-sig")
        for match in re.finditer(r'^\s+(\w+):\s*"([^"]*)"', text, re.MULTILINE):
            all_loc[match.group(1)] = match.group(2)

    # Resolve $key$ references (one pass is enough for single-depth refs)
    def resolve(value: str) -> str:
        def replacer(m):
            ref_key = m.group(1)
            return all_loc.get(ref_key, m.group(0))
        return re.sub(r'\$(\w+)\$', replacer, value)

    # Extract unit names (a_ and n_ prefixes) from the units file
    units_file = loc_dir / "units_l_english.yml"
    if not units_file.exists():
        return {}

    text = units_file.read_text(encoding="utf-8-sig")
    names = {}
    for match in re.finditer(r'^\s+([an]_\w+):\s*"([^"]*)"', text, re.MULTILINE):
        key = match.group(1)
        if key.endswith("_desc"):
            continue
        value = resolve(match.group(2))
        # Strip [Script(...)] calls that can't be resolved statically
        value = re.sub(r'\[.*?\]', '', value).strip()
        names[key] = value

    return names


def scrape_combined_arms() -> dict:
    """Scrape combined arms defines from auto_modifiers/country.txt."""
    raw = parse_directory(COMMON_DIR / "auto_modifiers")
    base = raw.get("country_base_values", {})
    return {
        "bonus_per_type": base.get("combined_bonus_per_type", 0),
        "min_percent": base.get("combined_arms_min_percent_for_bonus", 0),
        "max_threshold": base.get("combined_arms_max_threshold", 0),
    }


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Scraping unit categories...")
    categories = scrape_categories()

    print("Scraping unit prices...")
    prices = scrape_unit_prices()

    print("Scraping goods demands...")
    goods_demands = scrape_goods_demands()

    print("Scraping unit localizations...")
    localizations = scrape_unit_localizations()

    print("Scraping production recipes...")
    recipes = scrape_production_recipes()

    print("Scraping combined arms defines...")
    combined_arms = scrape_combined_arms()

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
    with open(OUTPUT_DIR / "localizations.json", "w") as f:
        json.dump(localizations, f, indent=2)
    print(f"  Wrote localizations.json ({len(localizations)} entries)")

    with open(OUTPUT_DIR / "production_recipes.json", "w") as f:
        json.dump(recipes, f, indent=2)
    total_recipes = sum(len(v) for v in recipes.values())
    print(f"  Wrote production_recipes.json ({len(recipes)} goods, {total_recipes} recipes)")

    with open(OUTPUT_DIR / "goods_demands.json", "w") as f:
        json.dump(goods_demands, f, indent=2)
    print(f"  Wrote goods_demands.json ({len(goods_demands)} demand types)")

    with open(OUTPUT_DIR / "combined_arms.json", "w") as f:
        json.dump(combined_arms, f, indent=2)
    print(f"  Wrote combined_arms.json ({combined_arms})")

    with open(OUTPUT_DIR / "unit_prices.json", "w") as f:
        json.dump(prices, f, indent=2)
    print(f"  Wrote unit_prices.json ({len(prices)} categories)")

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
