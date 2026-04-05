"""Scrape EU5 unit data from vanilla game files and resolve inheritance."""

import json
import copy
from pathlib import Path

from parser import parse_directory, parse_file

GAME_DIR = Path("C:/Steam/steamapps/common/Europa Universalis V/game")
COMMON_DIR = GAME_DIR / "in_game" / "common"
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data"

MARITIME_VALUES = {}  # populated in main() before unit extraction

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
    # Naval stats
    "hull_size",
    "cannons",
    "crew_size",
    "blockade_capacity",
    "transport_capacity",
    "anti_piracy_warfare",
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
        # Terrain combat/impact modifiers (used by naval categories like galley)
        for terrain_key in ["combat", "impact"]:
            if terrain_key in data and isinstance(data[terrain_key], dict):
                cat[terrain_key] = resolve_terrain_block(data[terrain_key])
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

    # Maritime presence (resolve string reference to numeric value)
    mp_ref = data.get("maritime_presence", "")
    if isinstance(mp_ref, str) and mp_ref:
        unit["maritime_presence"] = MARITIME_VALUES.get(mp_ref, 0)
    elif isinstance(mp_ref, (int, float)):
        unit["maritime_presence"] = mp_ref

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


def scrape_maritime_presence_values() -> dict:
    """Scrape maritime presence script value definitions.

    Returns: { "ship_small_maritime": 0.1, "ship_medium_maritime": 0.2, ... }
    """
    sv_file = GAME_DIR / "main_menu" / "common" / "script_values" / "default_values.txt"
    if not sv_file.exists():
        return {}
    raw = parse_file(sv_file)
    result = {}
    for key, val in raw.items():
        if key.startswith("ship_") and key.endswith("_maritime") and isinstance(val, (int, float)):
            result[key] = val
    return result


def scrape_combined_arms() -> dict:
    """Scrape combined arms defines from auto_modifiers/country.txt."""
    raw = parse_directory(COMMON_DIR / "auto_modifiers")
    base = raw.get("country_base_values", {})
    return {
        "bonus_per_type": base.get("combined_bonus_per_type", 0),
        "min_percent": base.get("combined_arms_min_percent_for_bonus", 0),
        "max_threshold": base.get("combined_arms_max_threshold", 0),
    }


def scrape_food_goods() -> dict:
    """Scrape all goods that have a food value.

    Returns: { "wheat": {"food_value": 8.0, "method": "farming", "price": 1.0}, ... }
    """
    raw = parse_directory(COMMON_DIR / "goods")
    result = {}
    for name, data in raw.items():
        if not isinstance(data, dict):
            continue
        food_val = data.get("food")
        if food_val is not None and isinstance(food_val, (int, float)):
            result[name] = {
                "food_value": food_val,
                "method": data.get("method", ""),
                "price": data.get("default_market_price", 1.0),
            }
    return result


def _extract_location_potential(pot: dict) -> dict:
    """Extract building requirements from a location_potential block.

    Returns dict with keys: rgo (list), vegetation (list), features (list like
    'is_coastal', 'has_river'), development_min (int or None).
    """
    reqs = {"rgo": [], "vegetation": [], "features": [], "development_min": None}
    if not isinstance(pot, dict):
        return reqs

    def walk(block):
        if not isinstance(block, dict):
            return
        for k, v in block.items():
            if k == "raw_material":
                vals = v if isinstance(v, list) else [v]
                for val in vals:
                    if isinstance(val, str):
                        reqs["rgo"].append(val.replace("goods:", ""))
            elif k == "vegetation":
                vals = v if isinstance(v, list) else [v]
                for val in vals:
                    if isinstance(val, str):
                        reqs["vegetation"].append(val)
            elif k in ("is_coastal", "has_river", "is_adjacent_to_lake") and v is True:
                reqs["features"].append(k)
            elif k == "development":
                if isinstance(v, dict) and v.get("__op__") == ">=" and isinstance(v.get("__value__"), (int, float)):
                    reqs["development_min"] = v["__value__"]
                elif isinstance(v, (int, float)):
                    reqs["development_min"] = v
            elif k == "OR":
                sub = v if isinstance(v, list) else [v]
                for s in sub:
                    walk(s)
            elif isinstance(v, dict):
                walk(v)

    walk(pot)
    return reqs


def scrape_food_buildings() -> dict:
    """Scrape buildings relevant to food production from building_types.

    Returns dict keyed by building name with requirements, production, and modifiers.
    """
    food_goods_data = scrape_food_goods()
    food_good_names = set(food_goods_data.keys())
    building_dir = COMMON_DIR / "building_types"

    # Only parse files that contain food-relevant buildings
    target_files = ["rural_buildings.txt", "common_buildings.txt"]
    result = {}

    # Food-related modifier keys we care about
    FOOD_MODIFIER_KEYS = {
        "local_monthly_food_modifier",
        "local_monthly_food",
        "local_food_capacity",
    }

    for filename in target_files:
        filepath = building_dir / filename
        if not filepath.exists():
            continue
        raw = parse_file(filepath)

        for bld_name, bld_data in raw.items():
            if not isinstance(bld_data, dict):
                continue
            if bld_data.get("is_special") is True:
                continue

            # Check if this building is food-relevant:
            # 1) produces a food good, 2) has food modifiers, 3) has rgo output modifiers for food goods
            produces = None
            food_mods = {}
            rgo_output_mods = {}
            inputs = {}

            # Check production methods
            upm = bld_data.get("unique_production_methods", {})
            if isinstance(upm, dict):
                upm_list = [upm]
            elif isinstance(upm, list):
                upm_list = upm
            else:
                upm_list = []

            for methods_block in upm_list:
                if not isinstance(methods_block, dict):
                    continue
                for method_name, method_data in methods_block.items():
                    if not isinstance(method_data, dict):
                        continue
                    produced = method_data.get("produced")
                    output_amt = method_data.get("output")
                    if produced and produced in food_good_names and output_amt:
                        produces = {"good": produced, "output_per_level": output_amt}
                    # Collect inputs for this method
                    skip = {"produced", "output", "category", "debug_max_profit"}
                    for k, v in method_data.items():
                        if k not in skip and isinstance(v, (int, float)):
                            inputs[k] = v

            # Check modifiers
            modifier = bld_data.get("modifier", {})
            if isinstance(modifier, dict):
                for key in FOOD_MODIFIER_KEYS:
                    if key in modifier:
                        food_mods[key] = modifier[key]
                # Check for rgo output modifiers (e.g., local_wheat_output_modifier)
                for k, v in modifier.items():
                    if k.startswith("local_") and k.endswith("_output_modifier"):
                        good_name = k[len("local_"):-len("_output_modifier")]
                        if good_name in food_good_names:
                            rgo_output_mods[good_name] = v

            # Only include if food-relevant
            if not produces and not food_mods and not rgo_output_mods:
                continue

            # Extract requirements
            pot = bld_data.get("location_potential", {})
            reqs = _extract_location_potential(pot)

            # Check development requirement from top-level 'allow' block too
            allow = bld_data.get("allow", {})
            if isinstance(allow, dict):
                dev_req = allow.get("development")
                if isinstance(dev_req, dict) and dev_req.get("__op__") == ">=" :
                    reqs["development_min"] = dev_req.get("__value__")

            # Max levels
            max_levels = bld_data.get("max_levels", 1)

            # Location rank availability
            ranks = []
            for rank in ["rural_settlement", "town", "city"]:
                if bld_data.get(rank) is True:
                    ranks.append(rank)
            if not ranks:
                ranks = ["rural_settlement"]  # default for rural buildings

            entry = {
                "max_levels": max_levels,
                "requirements": reqs,
                "ranks": ranks,
                "inputs": inputs if inputs else None,
            }
            if produces:
                entry["produces"] = produces
            if food_mods:
                entry["food_modifiers"] = food_mods
            if rgo_output_mods:
                entry["rgo_output_modifiers"] = rgo_output_mods

            result[bld_name] = entry

    return result


def scrape_building_caps() -> dict:
    """Scrape building cap formulas for rural_building_cap and irrigant_cap.

    Returns dict with numeric components for each cap.
    """
    filepath = COMMON_DIR / "script_values" / "building_caps.txt"
    if not filepath.exists():
        return {}
    raw = parse_file(filepath)

    caps = {}
    for cap_name in ["rural_building_cap", "irrigant_cap"]:
        data = raw.get(cap_name)
        if not isinstance(data, dict):
            continue
        cap = {"base": 0, "per_development": 0, "per_max_rgo_workers": 0, "if_river": 0}

        # The parser represents duplicate 'add' keys as a list
        adds = data.get("add", [])
        if isinstance(adds, dict):
            adds = [adds]
        elif not isinstance(adds, list):
            adds = []

        for add_block in adds:
            if not isinstance(add_block, dict):
                continue
            value = add_block.get("value", 0)
            multiply = add_block.get("multiply")
            desc = add_block.get("desc", "")

            if multiply is not None and isinstance(value, str):
                # Scaled value: value = development/max_rgo_workers, multiply = factor
                if value == "development":
                    cap["per_development"] = multiply
                elif value == "max_rgo_workers":
                    cap["per_max_rgo_workers"] = multiply
            elif isinstance(value, (int, float)) and multiply is None:
                # Check if this is a base value (not inside a conditional)
                if "BASE" in desc or (not cap["base"] and "RIVER" not in desc):
                    cap["base"] = value

        # Check for river bonus in 'if' blocks
        if_block = data.get("if", {})
        if isinstance(if_block, list):
            if_blocks = if_block
        else:
            if_blocks = [if_block]

        for ib in if_blocks:
            if not isinstance(ib, dict):
                continue
            limit = ib.get("limit", {})
            if isinstance(limit, dict) and limit.get("has_river") is True:
                river_add = ib.get("add", {})
                if isinstance(river_add, dict):
                    cap["if_river"] = river_add.get("value", 0)
                elif isinstance(river_add, (int, float)):
                    cap["if_river"] = river_add

        caps[cap_name] = cap

    return caps


def scrape_terrain_food_modifiers() -> dict:
    """Scrape food modifiers from vegetation, topography, and location_ranks.

    Returns dict with terrain categories and their food modifiers.
    """
    result = {}

    # Vegetation
    veg_raw = parse_directory(COMMON_DIR / "vegetation")
    veg = {}
    for name, data in veg_raw.items():
        if not isinstance(data, dict):
            continue
        loc_mod = data.get("location_modifier", {})
        if isinstance(loc_mod, dict):
            food_mod = loc_mod.get("local_monthly_food_modifier")
            if food_mod is not None:
                veg[name] = {"local_monthly_food_modifier": food_mod}
            else:
                veg[name] = {}
        else:
            veg[name] = {}
    result["vegetation"] = veg

    # Topography (land only)
    topo_raw = parse_directory(COMMON_DIR / "topography")
    topo = {}
    for name, data in topo_raw.items():
        if not isinstance(data, dict):
            continue
        # Skip naval/wasteland topographies
        if "ocean" in name or "lake" in name or "wasteland" in name or "narrows" in name or "salt_pans" in name or "atoll" in name or "inland_sea" in name:
            continue
        loc_mod = data.get("location_modifier", {})
        if isinstance(loc_mod, dict):
            food_mod = loc_mod.get("local_monthly_food_modifier")
            if food_mod is not None:
                topo[name] = {"local_monthly_food_modifier": food_mod}
            else:
                topo[name] = {}
        else:
            topo[name] = {}
    result["topography"] = topo

    # Location ranks
    rank_raw = parse_directory(COMMON_DIR / "location_ranks")
    ranks = {}
    for name, data in rank_raw.items():
        if not isinstance(data, dict):
            continue
        rank_mod = data.get("rank_modifier", {})
        if isinstance(rank_mod, dict):
            food_mod = rank_mod.get("local_monthly_food_modifier")
            if food_mod is not None:
                ranks[name] = {"local_monthly_food_modifier": food_mod}
            else:
                ranks[name] = {}
        else:
            ranks[name] = {}
    result["location_ranks"] = ranks

    return result


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

    print("Scraping maritime presence values...")
    global MARITIME_VALUES
    MARITIME_VALUES = scrape_maritime_presence_values()

    print("Scraping combined arms defines...")
    combined_arms = scrape_combined_arms()

    print("Scraping food goods...")
    food_goods = scrape_food_goods()

    print("Scraping food buildings...")
    food_buildings = scrape_food_buildings()

    print("Scraping building caps...")
    building_caps = scrape_building_caps()

    print("Scraping terrain food modifiers...")
    terrain_food = scrape_terrain_food_modifiers()

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

    with open(OUTPUT_DIR / "food_goods.json", "w") as f:
        json.dump(food_goods, f, indent=2)
    print(f"  Wrote food_goods.json ({len(food_goods)} goods)")

    with open(OUTPUT_DIR / "food_buildings.json", "w") as f:
        json.dump(food_buildings, f, indent=2)
    print(f"  Wrote food_buildings.json ({len(food_buildings)} buildings)")

    with open(OUTPUT_DIR / "building_caps.json", "w") as f:
        json.dump(building_caps, f, indent=2)
    print(f"  Wrote building_caps.json ({len(building_caps)} caps)")

    with open(OUTPUT_DIR / "terrain_food_modifiers.json", "w") as f:
        json.dump(terrain_food, f, indent=2)
    print(f"  Wrote terrain_food_modifiers.json")

    print("Done!")


if __name__ == "__main__":
    main()
