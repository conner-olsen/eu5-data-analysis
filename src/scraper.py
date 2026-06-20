"""Scrape EU5 unit data from vanilla game files and resolve inheritance."""

import json
import copy
from pathlib import Path

from parser import parse_directory, parse_file

GAME_DIR = Path("C:/Steam/steamapps/common/Europa Universalis V/game")
COMMON_DIR = GAME_DIR / "in_game" / "common"
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data"

MARITIME_VALUES = {}  # populated in main() before unit extraction
# Unit names from 2_unlocked_through_tech.txt (the generic tech-unlock progression every
# country shares). Units defined in any other file are unique/nation/culture/DLC-specific.
GENERIC_UNIT_NAMES = set()  # populated in main() before unit extraction

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

# Country-scope trade modifiers split by unit: PERCENT feed the headline total,
# FLAT are reported in their own columns.
TRADE_MODS_PERCENT = {
    "trade_income",
    "selling_efficiency",
    "export_efficiency",
    "import_efficiency",
    "foreign_export_from_market_efficiency",
    "merchant_maintenance_efficiency",
    "global_merchant_power",
    "merchant_power_from_maritime_modifier",
    "global_merchant_capacity_modifier",
    "trade_range_modifier",
    "global_trade_center_power",
    "global_trade_protection_factor",
    "trade_land_efficiency",
    "trade_sea_efficiency",
    "global_trade_through_owned_territory_efficiency",
    "global_trades_per_burgher",
}
TRADE_MODS_FLAT = {
    "merchant_power_from_maritime",
    "trade_range",
    "market_building_levels",
}
TRADE_MODS = TRADE_MODS_PERCENT | TRADE_MODS_FLAT

# Country-specific laws whose owner is named in the law id but gated by an
# unlock/variable trigger instead of has_or_had_tag.
LAW_TAG = {
    "french_colonial_ambition_focus": "FRA",
    "teu_amber_monopoly_law": "TEU",
    "kor_currency": "KOR",
}

# Country-specific government reforms gated by an unlock trigger instead of has_or_had_tag.
REFORM_TAG = {
    "kor_daedongbeob": "KOR",
    "kor_sadae_reform": "KOR",
    "signoria_of_venice": "VEN",
    "omani_overseas_ambitions_reform": "OMA",
}

# Country-specific cabinet actions gated by an unlock trigger instead of has_or_had_tag.
CABINET_TAG = {
    "oma_bolster_omani_navy": "OMA",
}

# The five EU5 government types (common/government_types/00_default.txt).
GOV_TYPES = {"monarchy", "republic", "theocracy", "steppe_horde", "tribe"}

# Potential keys that gate an entry to a culture/religion/region, making any government-type
# match incidental rather than a clean "any nation of this government type gets it".
_RESTRICTIVE_GATE_KEYS = {
    "culture", "primary_culture", "has_culture_group", "has_culture",
    "religion", "religious_group", "region",
}

# Natural government type assumed when a country forms a formable (formables have no 1337 start,
# and forming keeps your government in-game, so this is a modeling assumption). Default monarchy.
FORMABLE_GOV = {
    "NED": "republic",
    "TUS": "republic",
}

# Formable culture-group / specific-culture requirement from its formable_countries potential,
# used to data-drive which trade nations can form it. Religion and city-ownership gates are
# approximated (handled via FORMABLE_OVERRIDES where they matter).
FORMABLE_REQ = {
    "NED": {"groups": {"netherlandish_group"}},
    "GBR": {"groups": {"british_group"}},
    "SPA": {"groups": {"iberian_group"}},
    "TUS": {"cultures": {"tuscan"}},
    "GER": {"groups": {"german_group"}},
    "PLC": {"groups": {"polish_group", "lithuanian_group"}},
    "IRA": {"groups": {"iranian_group"}},
    "SCA": {"groups": {"scandinavian_group"}},
    "AIA": {"groups": {"arabian_group"}},
    "MSA": {"cultures": {"malay"}},
    "MOL": {"cultures": {"moldovan"}},
    "DAL": {"cultures": {"dalmatian", "croatian"}},
    "BAV": {"cultures": {"danube_bavarian"}},
    "NUS": {"cultures": {"javanese", "sundanese", "madurese"}},
}

# Hand-verified corrections that replace the data-driven former derivation for a formable.
# Tags are still intersected with the trade-nation set.
FORMABLE_OVERRIDES = {
    "SPA": ["CAS", "ARA"],  # iberian_group also contains Muslim Granada; restrict to Christian formers
    "NED": ["HOL"],         # low_franconian spans into the Rhineland; drop the Palatinate false match
}

# Building-based countries (country_type = building, e.g. the Hanseatic League) are not territorial
# states and cannot form territorial countries, so they are never formers.
NON_FORMER_TAGS = {"HSA"}


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
    unit["light"] = "light" in cat_name
    unit["is_special"] = data.get("is_special", False)
    unit["buildable"] = data.get("buildable", True)
    unit["levy"] = data.get("levy", False)
    unit["default"] = data.get("default", False)
    unit["generic"] = name in GENERIC_UNIT_NAMES

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
    land_categories = ["army_heavy_infantry", "army_light_infantry",
                       "army_heavy_cavalry", "army_light_cavalry",
                       "army_artillery", "army_auxiliary"]

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
    { "army_heavy_infantry": { "build_gold": 50, "reinforce_gold": 2.5, "maintenance_gold": 0.5 }, ... }

    Unit-specific entries in 02_units.txt are deltas added to the category base.
    They are returned under "unit_adds", keyed by unit type name.
    """
    prices_dir = COMMON_DIR / "prices"
    raw = parse_directory(prices_dir)

    categories = ["army_heavy_infantry", "army_light_infantry",
                  "army_heavy_cavalry", "army_light_cavalry",
                  "army_artillery", "army_auxiliary",
                  "navy_heavy_ship", "navy_light_ship", "navy_galley", "navy_transport"]

    prices = {}
    for cat in categories:
        entry = {}
        for cost_type in ["build", "reinforce", "maintenance"]:
            key = f"{cat}_{cost_type}"
            if key in raw and isinstance(raw[key], dict):
                entry[f"{cost_type}_gold"] = raw[key].get("gold", 0)
                entry[f"{cost_type}_manpower"] = raw[key].get("manpower", raw[key].get("sailors", 0))
        if entry:
            prices[cat] = entry

    units_raw = parse_file(prices_dir / "02_units.txt")
    unit_adds = {}
    for cost_type in ["build", "reinforce", "maintenance"]:
        suffix = f"_{cost_type}"
        for key, data in units_raw.items():
            if not key.endswith(suffix) or not isinstance(data, dict):
                continue
            name = key[: -len(suffix)]
            if name in categories:
                continue
            entry = unit_adds.setdefault(name, {})
            entry[f"{cost_type}_gold"] = data.get("gold", 0)
            entry[f"{cost_type}_manpower"] = data.get("manpower", data.get("sailors", 0))
    if unit_adds:
        prices["unit_adds"] = unit_adds

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
        "mills_input": "mills",
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

    # First pass: build a global lookup of ALL localization keys (base + DLC subdirs)
    all_loc = {}
    for loc_file in sorted(loc_dir.rglob("*_l_english.yml")):
        text = loc_file.read_text(encoding="utf-8-sig")
        for match in re.finditer(r'^\s+(\w+):\s*"([^"]*)"', text, re.MULTILINE):
            all_loc[match.group(1)] = match.group(2)

    # Resolve $key$ references (one pass is enough for single-depth refs)
    def resolve(value: str) -> str:
        def replacer(m):
            ref_key = m.group(1)
            return all_loc.get(ref_key, m.group(0))
        return re.sub(r'\$(\w+)\$', replacer, value)

    # Extract unit names (a_ and n_ prefixes) from every units file (base + DLC subdirs)
    names = {}
    for units_file in sorted(loc_dir.rglob("*units_l_english.yml")):
        text = units_file.read_text(encoding="utf-8-sig")
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
    """Scrape combined arms defines from auto_modifiers/country.txt plus per-age advance bonuses."""
    raw = parse_directory(COMMON_DIR / "auto_modifiers")
    base = raw.get("country_base_values", {})

    # Advances that raise combined_bonus_per_type, keyed by the age they unlock in
    advances = parse_directory(COMMON_DIR / "advances")
    bonus_advances_by_age = {}
    for data in advances.values():
        if not isinstance(data, dict):
            continue
        delta = data.get("combined_bonus_per_type")
        age = data.get("age")
        if isinstance(delta, (int, float)) and isinstance(age, str):
            bonus_advances_by_age[age] = bonus_advances_by_age.get(age, 0) + delta

    return {
        "bonus_per_type": base.get("combined_bonus_per_type", 0),
        "min_percent": base.get("combined_arms_min_percent_for_bonus", 0),
        "max_threshold": base.get("combined_arms_max_threshold", 0),
        "bonus_advances_by_age": bonus_advances_by_age,
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

    # Shared method definitions referenced via possible_production_methods
    shared_methods = parse_directory(COMMON_DIR / "production_methods")

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

            SKIP_KEYS = {"produced", "output", "category", "debug_max_profit"}

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
                    for k, v in method_data.items():
                        if k not in SKIP_KEYS and isinstance(v, (int, float)):
                            inputs[k] = v

            # Methods referenced by name; the first food-producing one wins
            ppm = bld_data.get("possible_production_methods", {})
            ppm_names = ppm.get("__bare_values__", []) if isinstance(ppm, dict) else []
            for method_name in ppm_names:
                method_data = shared_methods.get(method_name)
                if not isinstance(method_data, dict):
                    continue
                produced = method_data.get("produced")
                output_amt = method_data.get("output")
                if produces or not produced or produced not in food_good_names or not output_amt:
                    continue
                produces = {"good": produced, "output_per_level": output_amt}
                for k, v in method_data.items():
                    if k not in SKIP_KEYS and isinstance(v, (int, float)):
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


def scrape_forts() -> list:
    """Scrape fort buildings and their fort_level from building_types/forts.txt.

    Returns list of dicts sorted by fort_level: [{name, fort_level}, ...]
    """
    filepath = COMMON_DIR / "building_types" / "forts.txt"
    raw = parse_file(filepath)
    forts = []
    for name, data in raw.items():
        if not isinstance(data, dict):
            continue
        raw_mod = data.get("raw_modifier", {})
        if isinstance(raw_mod, dict) and "fort_level" in raw_mod:
            forts.append({
                "name": name,
                "fort_level": raw_mod["fort_level"],
            })
    forts.sort(key=lambda f: f["fort_level"])
    return forts


def scrape_pop_demands() -> dict:
    """Scrape per-pop-type goods demands from goods definitions and pop types.

    Reads demand_add / demand_multiply from each good's definition and
    pop_food_consumption from pop_types.  Returns resolved demand per pop per
    good, plus prices and food consumption.

    Returns: {
        "pop_types": { "nobles": {"food_consumption": 20.0}, ... },
        "goods": [
            { "name": "wine", "price": 2, "demands": {"nobles": 0.006, ...} },
            ...
        ]
    }
    """
    POP_TYPES = ["nobles", "clergy", "burghers", "soldiers", "laborers",
                 "peasants", "slaves", "tribesmen"]
    UPPER = {"nobles", "clergy", "burghers"}

    def resolve_add(add_block: dict) -> dict[str, float]:
        """Resolve demand_add to per-pop-type values (additive)."""
        result = {p: 0.0 for p in POP_TYPES}
        if not isinstance(add_block, dict):
            return result
        base_all = add_block.get("all", 0)
        base_upper = add_block.get("upper", 0)
        for p in POP_TYPES:
            val = base_all
            if p in UPPER:
                val += base_upper
            val += add_block.get(p, 0)
            result[p] = val
        return result

    def resolve_multiply(mul_block: dict) -> dict[str, float]:
        """Resolve demand_multiply to per-pop-type multipliers."""
        result = {p: 1.0 for p in POP_TYPES}
        if not isinstance(mul_block, dict):
            return result
        upper_mul = mul_block.get("upper", 1.0)
        for p in POP_TYPES:
            mul = 1.0
            if p in UPPER:
                mul *= upper_mul
            if p in mul_block:
                mul *= mul_block[p]
            result[p] = mul
        return result

    # Scrape pop types for food consumption
    pop_raw = parse_directory(COMMON_DIR / "pop_types")
    pop_info = {}
    for name in POP_TYPES:
        data = pop_raw.get(name, {})
        pop_info[name] = {
            "food_consumption": data.get("pop_food_consumption", 0),
        }

    # Scrape goods definitions
    goods_dir = COMMON_DIR / "goods"
    goods_list = []
    for filepath in sorted(goods_dir.glob("*.txt")):
        raw = parse_file(filepath)
        for name, data in raw.items():
            if not isinstance(data, dict):
                continue
            demand_add = data.get("demand_add")
            demand_multiply = data.get("demand_multiply")
            if not demand_add and not demand_multiply:
                continue  # no pop demand for this good

            price = data.get("default_market_price", 1.0)
            add_resolved = resolve_add(demand_add or {})
            mul_resolved = resolve_multiply(demand_multiply or {})

            demands = {}
            for p in POP_TYPES:
                demands[p] = add_resolved[p] * mul_resolved[p]

            # Skip goods where all demands are zero
            if all(v == 0 for v in demands.values()):
                continue

            goods_list.append({
                "name": name,
                "price": price,
                "demands": demands,
            })

    return {"pop_types": pop_info, "goods": goods_list}


def _load_script_values() -> dict:
    """Numeric named values from default_values.txt (resolves tokens like small_trade_efficiency_bonus)."""
    sv_file = GAME_DIR / "main_menu" / "common" / "script_values" / "default_values.txt"
    if not sv_file.exists():
        return {}
    raw = parse_file(sv_file)
    return {k: v for k, v in raw.items() if isinstance(v, (int, float)) and not isinstance(v, bool)}


def _resolve_mod(value, tokens, unresolved):
    """Resolve a modifier value to a float: raw numbers as-is, named tokens via the script-value map."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        if value in tokens:
            return float(tokens[value])
        unresolved.add(value)
        return None
    if isinstance(value, list):
        total, found = 0.0, False
        for v in value:
            r = _resolve_mod(v, tokens, unresolved)
            if r is not None:
                total += r
                found = True
        return total if found else None
    return None


def _collect_tags(node, out):
    """Walk a potential block adding every has_or_had_tag value; NOT subtrees invert meaning, so skip them."""
    if not isinstance(node, dict):
        return
    for key, val in node.items():
        kl = key.lower()
        if kl == "not":
            continue
        if kl == "has_or_had_tag":
            if isinstance(val, list):
                out.update(v for v in val if isinstance(v, str))
            elif isinstance(val, str):
                out.add(val)
        elif isinstance(val, dict):
            _collect_tags(val, out)
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    _collect_tags(item, out)


def _add_trade_mods(block, tokens, unresolved, acc):
    """Sum whitelisted trade modifiers from one block (advance top level or a country_modifier) into acc."""
    if not isinstance(block, dict):
        return
    for key, val in block.items():
        if key in TRADE_MODS:
            r = _resolve_mod(val, tokens, unresolved)
            if r is not None:
                acc[key] = acc.get(key, 0.0) + r


def _country_modifier_blocks(container):
    """Yield each country_modifier dict in a container (parser returns a list when the key repeats)."""
    cm = container.get("country_modifier")
    if isinstance(cm, dict):
        yield cm
    elif isinstance(cm, list):
        for item in cm:
            if isinstance(item, dict):
                yield item


def _collect_gov_types(node, out):
    """Walk a potential block adding every government_type = government_type:X value (NOT subtrees skipped)."""
    if not isinstance(node, dict):
        return
    for key, val in node.items():
        if key.lower() == "not":
            continue
        if key == "government_type":
            vals = val if isinstance(val, list) else [val]
            for v in vals:
                if isinstance(v, str) and v.split(":")[-1] in GOV_TYPES:
                    out.add(v.split(":")[-1])
        elif isinstance(val, dict):
            _collect_gov_types(val, out)
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    _collect_gov_types(item, out)


def _entry_gov_types(data):
    """Government types an entry is gated to: a top-level government/law_gov_group key plus any
    government_type in its potential."""
    out = set()
    for key in ("government", "law_gov_group"):
        v = data.get(key)
        if isinstance(v, str) and v in GOV_TYPES:
            out.add(v)
        elif isinstance(v, list):
            out.update(x for x in v if isinstance(x, str) and x in GOV_TYPES)
    _collect_gov_types(data.get("potential"), out)
    return out


def _has_restrictive_gate(node):
    """True if a potential gates by culture/religion/region, so its government-type match is incidental."""
    if not isinstance(node, dict):
        return False
    for key, val in node.items():
        if key.lower() in _RESTRICTIVE_GATE_KEYS:
            return True
        if isinstance(val, dict):
            if _has_restrictive_gate(val):
                return True
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict) and _has_restrictive_gate(item):
                    return True
    return False


def _collect_culture_keys(node, out):
    """Walk a potential collecting culture gate keys: 'culture:X', 'group:Y', 'lang:Z'. Skip NOT/NOR subtrees."""
    if not isinstance(node, dict):
        return
    for key, val in node.items():
        kl = key.lower()
        if kl in ("not", "nor", "nand"):
            continue
        if kl == "culture":
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str) and it.startswith("culture:"):
                    out.add("culture:" + it.split(":", 1)[1])
                elif isinstance(it, dict):
                    _collect_culture_keys(it, out)
        elif kl in ("has_culture", "merged_culture_group_contains_culture"):
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str):
                    out.add("culture:" + it.split(":")[-1])
        elif kl == "has_culture_group":
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str):
                    out.add("group:" + it.split(":")[-1])
        elif kl == "culture.language":
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str):
                    out.add("lang:" + it.split(":")[-1])
        elif isinstance(val, dict):
            _collect_culture_keys(val, out)
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    _collect_culture_keys(item, out)


def _collect_religion_keys(node, out):
    """Walk a potential collecting religion gate keys: 'religion:X', 'group:Y'. Skip NOT/NOR subtrees."""
    if not isinstance(node, dict):
        return
    for key, val in node.items():
        kl = key.lower()
        if kl in ("not", "nor", "nand"):
            continue
        if kl == "religion":
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str) and it.startswith("religion:"):
                    out.add("religion:" + it.split(":", 1)[1])
                elif isinstance(it, dict):
                    _collect_religion_keys(it, out)
        elif kl in ("religion.group", "religious_group", "religion_group"):
            for it in (val if isinstance(val, list) else [val]):
                if isinstance(it, str):
                    out.add("group:" + it.split(":")[-1])
        elif isinstance(val, dict):
            _collect_religion_keys(val, out)
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    _collect_religion_keys(item, out)


_FAITH_REGION_KEYS = {"religion", "religion.group", "religious_group", "region", "area"}


def _has_gate_key(node, keys):
    """True if any of keys appears anywhere in the potential (recursing dicts/lists)."""
    if not isinstance(node, dict):
        return False
    for k, v in node.items():
        if k.lower() in keys:
            return True
        if isinstance(v, dict):
            if _has_gate_key(v, keys):
                return True
        elif isinstance(v, list):
            for it in v:
                if isinstance(it, dict) and _has_gate_key(it, keys):
                    return True
    return False


def _iter_country_blocks():
    """Yield (tag, [(depth, code_line)]) for each top-level country block in the setup start file.

    The start file is one huge nested block the recursive parser mis-nests, so blocks are found by a
    brace-depth text scan; depth 1 is the country's own keys (capital, include), deeper is nested.
    """
    import re
    start_file = GAME_DIR / "main_menu" / "setup" / "start" / "10_countries.txt"
    if not start_file.exists():
        return
    header_re = re.compile(r'^\t+([A-Z][A-Z0-9_]{1,4})\s*=\s*\{')
    lines = start_file.read_text(encoding="utf-8-sig").split("\n")
    i, n = 0, len(lines)
    while i < n:
        m = header_re.match(lines[i])
        if not m:
            i += 1
            continue
        tag = m.group(1)
        code0 = lines[i].split("#", 1)[0]
        depth = code0.count("{") - code0.count("}")
        block = []
        i += 1
        while i < n and depth > 0:
            code = lines[i].split("#", 1)[0]
            block.append((depth, code))
            depth += code.count("{") - code.count("}")
            i += 1
        yield tag, block


def _template_gov_types() -> dict:
    """Map setup template name -> government type, following include chains (setup/templates/*.txt)."""
    tdir = GAME_DIR / "main_menu" / "setup" / "templates"
    if not tdir.exists():
        return {}
    raw = {}
    for f in sorted(tdir.glob("*.txt")):
        data = parse_file(f)
        gov = data.get("government")
        own = gov.get("type") if isinstance(gov, dict) else None
        incs = data.get("include")
        inc_list = incs if isinstance(incs, list) else ([incs] if incs is not None else [])
        raw[f.stem] = (own if isinstance(own, str) and own in GOV_TYPES else None,
                       [str(i).strip('"') for i in inc_list if isinstance(i, str)])
    resolved = {}

    def resolve(name, seen):
        if name in resolved:
            return resolved[name]
        if name not in raw or name in seen:
            return None
        own, incs = raw[name]
        out = own
        if not out:
            for inc in incs:
                out = resolve(inc, seen | {name})
                if out:
                    break
        resolved[name] = out
        return out

    for name in raw:
        resolve(name, set())
    return {k: v for k, v in resolved.items() if v}


def scrape_country_governments() -> dict:
    """Map country tag -> starting government type from the setup start file.

    The type is on each country's `government = { type = X }` when explicit, otherwise inherited from a
    setup template via `include` (e.g. catholic_monarchy -> monarchy).
    Returns { "FRA": "monarchy", "VEN": "republic", ... }.
    """
    import re
    templates = _template_gov_types()
    inc_re = re.compile(r'\binclude\s*=\s*"?([A-Za-z0-9_]+)"?')
    type_re = re.compile(r'\btype\s*=\s*(monarchy|republic|theocracy|tribe|steppe_horde)\b')
    out = {}
    for tag, block in _iter_country_blocks():
        if tag in out:
            continue
        gtype, includes = None, []
        for depth, code in block:
            if depth == 1:
                im = inc_re.search(code)
                if im:
                    includes.append(im.group(1))
            tm = type_re.search(code)
            if tm and gtype is None:
                gtype = tm.group(1)
        if not gtype:
            for inc in includes:
                if inc in templates:
                    gtype = templates[inc]
                    break
        if gtype:
            out[tag] = gtype
    return out


def _scrape_location_culture_religion():
    """Parse location_templates.txt single-line blocks -> ({location: culture}, {location: religion})."""
    import re
    f = GAME_DIR / "in_game" / "map_data" / "location_templates.txt"
    cul, rel = {}, {}
    if not f.exists():
        return cul, rel
    block_re = re.compile(r'^([a-z][a-z0-9_]*)\s*=\s*\{([^{}]*)\}', re.MULTILINE)
    cul_re = re.compile(r'\bculture\s*=\s*([a-z][a-z0-9_]*)')
    rel_re = re.compile(r'\breligion\s*=\s*([a-z][a-z0-9_]*)')
    for m in block_re.finditer(f.read_text(encoding="utf-8-sig")):
        loc, body = m.group(1), m.group(2)
        cm, rm = cul_re.search(body), rel_re.search(body)
        if cm:
            cul[loc] = cm.group(1)
        if rm:
            rel[loc] = rm.group(1)
    return cul, rel


def _scrape_culture_info():
    """Map culture -> {groups: set, language: str} from cultures/*.txt."""
    out = {}
    for name, data in parse_directory(COMMON_DIR / "cultures").items():
        if not isinstance(data, dict):
            continue
        groups = data.get("culture_groups")
        vals = groups.get("__bare_values__") if isinstance(groups, dict) else (groups if isinstance(groups, list) else None)
        lang = data.get("language")
        out[name] = {
            "groups": set(v for v in vals if isinstance(v, str)) if isinstance(vals, list) else set(),
            "language": lang if isinstance(lang, str) else "",
        }
    return out


def scrape_country_cultures() -> dict:
    """Map country tag -> {culture, groups, religion} via capital -> location culture/religion -> groups."""
    import re
    loc_cul, loc_rel = _scrape_location_culture_religion()
    cul_info = _scrape_culture_info()
    cap_re = re.compile(r'\bcapital\s*=\s*(?:location:)?([a-z][a-z0-9_]*)')
    out = {}
    for tag, block in _iter_country_blocks():
        if tag in out:
            continue
        capital = None
        for depth, code in block:
            if depth == 1:
                cm = cap_re.search(code)
                if cm:
                    capital = cm.group(1)
                    break
        culture = loc_cul.get(capital) if capital else None
        if not culture:
            continue
        info = cul_info.get(culture, {})
        out[tag] = {
            "culture": culture,
            "groups": info.get("groups", set()),
            "language": info.get("language", ""),
            "religion": loc_rel.get(capital, ""),
        }
    return out


def scrape_formable_formers(trade_tags, country_governments) -> dict:
    """Map each trade-bearing formable -> {gov, formers}.

    A trade nation with no 1337 start government is a formable. Formers are the trade nations whose
    capital culture / culture group satisfies FORMABLE_REQ, replaced by FORMABLE_OVERRIDES where set;
    formables with no in-list former get an empty list (rendered standalone downstream).
    """
    cultures = scrape_country_cultures()
    out = {}
    for ftag in sorted(trade_tags):
        if ftag in country_governments:
            continue  # real start country, not a formable
        if ftag in FORMABLE_OVERRIDES:
            formers = [t for t in FORMABLE_OVERRIDES[ftag] if t in trade_tags and t != ftag]
        else:
            req = FORMABLE_REQ.get(ftag, {})
            want_groups = req.get("groups", set())
            want_cultures = req.get("cultures", set())
            formers = []
            for tag in trade_tags:
                if tag == ftag or tag in NON_FORMER_TAGS:
                    continue
                info = cultures.get(tag)
                if not info:
                    continue
                if (want_groups and info["groups"] & want_groups) or (want_cultures and info["culture"] in want_cultures):
                    formers.append(tag)
        out[ftag] = {"gov": FORMABLE_GOV.get(ftag, "monarchy"), "formers": sorted(set(formers))}
    return out


def scrape_country_names() -> dict:
    """Map country tag to localized name from country_names_l_english.yml.

    Returns: { "FRA": "France", "ENG": "England", ... }
    """
    import re
    loc_file = (GAME_DIR / "main_menu" / "localization" / "english"
                / "country_names_l_english.yml")
    names = {}
    if not loc_file.exists():
        return names
    text = loc_file.read_text(encoding="utf-8-sig")
    for match in re.finditer(r'^\s+([A-Za-z0-9_]+):\d*\s*"([^"]*)"', text, re.MULTILINE):
        names[match.group(1)] = match.group(2)
    return names


def scrape_religions() -> dict:
    """Map religion -> {group, def_mods} from religions/*.txt (group plus trade mods in definition_modifier)."""
    tokens = _load_script_values()
    unresolved = set()
    out = {}
    for name, data in parse_directory(COMMON_DIR / "religions").items():
        if not isinstance(data, dict):
            continue
        group = data.get("group")
        mods = {}
        defmod = data.get("definition_modifier")
        if isinstance(defmod, dict):
            _add_trade_mods(defmod, tokens, unresolved, mods)
        out[name] = {"group": group if isinstance(group, str) else "", "def_mods": mods}
    return out


def scrape_trade_nations():
    """Sum country-attributable trade modifiers from advances, estate privileges, laws, government reforms, and cabinet actions.

    A nation gets a modifier when its source is gated to that tag (has_or_had_tag in the entry's
    potential) or, for the named laws/reforms/cabinet actions in LAW_TAG/REFORM_TAG/CABINET_TAG, by the
    entry id. A tagless entry gated only to a government type (and not culture/religion/region) goes to
    that type's bundle. Named magnitude tokens resolve via default_values.txt.

    Returns (nations, gov_bonuses):
      nations    = { "<TAG>": {"name": str, "mods": {modifier: total}, "entries": [{source, id, mods}]} }
      gov_bonuses = { "<gov_type>": {"mods": {modifier: total}, "entries": [{source, id, mods}]} }
    """
    tokens = _load_script_values()
    names = scrape_country_names()
    unresolved = set()
    excluded_laws = []
    excluded_reforms = []
    excluded_cabinet = []

    nations = {}
    gov_bonuses = {}
    culture_entries = []  # (culture_keys, source, id, mods) for tagless culture-gated entries
    religion_entries = []  # (religion_keys, source, id, mods) for tagless religion-gated entries

    def attribute(tag, source, entry_id, mods):
        nat = nations.setdefault(tag, {"name": names.get(tag, tag), "mods": {}, "entries": []})
        for mod, val in mods.items():
            nat["mods"][mod] = nat["mods"].get(mod, 0.0) + val
        nat["entries"].append({"source": source, "id": entry_id, "mods": mods})

    def gov_attribute(gov_types, source, entry_id, mods):
        for gt in gov_types:
            b = gov_bonuses.setdefault(gt, {"mods": {}, "entries": []})
            for mod, val in mods.items():
                b["mods"][mod] = b["mods"].get(mod, 0.0) + val
            b["entries"].append({"source": source, "id": entry_id, "mods": mods})

    def route_generic(data, source, entry_id, mods, excluded):
        """A tagless entry: route to the culture bundle if culture-gated, else the gov-type bundle, else exclude."""
        potential = data.get("potential")
        ckeys = set()
        _collect_culture_keys(potential, ckeys)
        if ckeys and not _has_gate_key(potential, _FAITH_REGION_KEYS):
            culture_entries.append((ckeys, source, entry_id, mods))
            return
        rkeys = set()
        _collect_religion_keys(potential, rkeys)
        if rkeys:
            religion_entries.append((rkeys, source, entry_id, mods))
            return
        if not _has_restrictive_gate(potential):
            gts = _entry_gov_types(data)
            if gts:
                gov_attribute(gts, source, entry_id, mods)
                return
        if excluded is not None:
            excluded.append(entry_id)

    # Advances: modifiers sit at the advance's top level.
    for adv_id, data in parse_directory(COMMON_DIR / "advances").items():
        if not isinstance(data, dict):
            continue
        mods = {}
        _add_trade_mods(data, tokens, unresolved, mods)
        if not mods:
            continue
        tags = set()
        _collect_tags(data.get("potential"), tags)
        if tags:
            for tag in tags:
                attribute(tag, "advance", adv_id, dict(mods))
        else:
            route_generic(data, "advance", adv_id, dict(mods), None)

    # Estate privileges: modifiers sit in country_modifier block(s) at the privilege's top level.
    for priv_id, data in parse_directory(COMMON_DIR / "estate_privileges").items():
        if not isinstance(data, dict):
            continue
        mods = {}
        for block in _country_modifier_blocks(data):
            _add_trade_mods(block, tokens, unresolved, mods)
        if not mods:
            continue
        tags = set()
        _collect_tags(data.get("potential"), tags)
        if tags:
            for tag in tags:
                attribute(tag, "privilege", priv_id, dict(mods))
        else:
            route_generic(data, "privilege", priv_id, dict(mods), None)

    # Laws: modifiers sit in country_modifier block(s) one level down, inside each named policy.
    for law_id, data in parse_directory(COMMON_DIR / "laws").items():
        if not isinstance(data, dict):
            continue
        law_tags = set()
        _collect_tags(data.get("potential"), law_tags)
        if law_id in LAW_TAG:
            law_tags.add(LAW_TAG[law_id])
        law_gov = _entry_gov_types(data)
        law_restrictive = _has_restrictive_gate(data.get("potential"))
        for policy_id, policy in data.items():
            if policy_id == "country_modifier" or not isinstance(policy, dict):
                continue
            mods = {}
            for block in _country_modifier_blocks(policy):
                _add_trade_mods(block, tokens, unresolved, mods)
            if not mods:
                continue
            entry_id = f"{law_id}.{policy_id}"
            tags = set()
            _collect_tags(policy.get("potential"), tags)
            tags |= law_tags
            if tags:
                for tag in tags:
                    attribute(tag, "law", entry_id, dict(mods))
            elif law_restrictive or _has_restrictive_gate(policy.get("potential")):
                excluded_laws.append(entry_id)
            else:
                gts = _entry_gov_types(policy) | law_gov
                if gts:
                    gov_attribute(gts, "law", entry_id, dict(mods))
                else:
                    excluded_laws.append(entry_id)

    # Government reforms: country_modifier block(s) at the reform top level, like privileges,
    # with a REFORM_TAG fallback for the named reforms gated by unlock triggers.
    for reform_id, data in parse_directory(COMMON_DIR / "government_reforms").items():
        if not isinstance(data, dict):
            continue
        mods = {}
        for block in _country_modifier_blocks(data):
            _add_trade_mods(block, tokens, unresolved, mods)
        if not mods:
            continue
        tags = set()
        _collect_tags(data.get("potential"), tags)
        if reform_id in REFORM_TAG:
            tags.add(REFORM_TAG[reform_id])
        if tags:
            for tag in tags:
                attribute(tag, "reform", reform_id, dict(mods))
        else:
            route_generic(data, "reform", reform_id, dict(mods), excluded_reforms)

    # Cabinet actions: country_modifier block at the action top level, like reforms,
    # with a CABINET_TAG fallback for the unlock-gated country actions.
    for cab_id, data in parse_directory(COMMON_DIR / "cabinet_actions").items():
        if not isinstance(data, dict):
            continue
        mods = {}
        for block in _country_modifier_blocks(data):
            _add_trade_mods(block, tokens, unresolved, mods)
        if not mods:
            continue
        tags = set()
        _collect_tags(data.get("potential"), tags)
        if cab_id in CABINET_TAG:
            tags.add(CABINET_TAG[cab_id])
        if tags:
            for tag in tags:
                attribute(tag, "cabinet", cab_id, dict(mods))
        else:
            route_generic(data, "cabinet", cab_id, dict(mods), excluded_cabinet)

    # Culture content: every real nation gets its starting culture's (and culture group's / language's)
    # trade bonuses, deduped by entry.
    cultures = scrape_country_cultures()
    for tag, nat in nations.items():
        info = cultures.get(tag)
        if not info:
            continue
        nkeys = {"culture:" + info["culture"]} | {"group:" + g for g in info["groups"]}
        if info.get("language"):
            nkeys.add("lang:" + info["language"])
        for ckeys, source, entry_id, mods in culture_entries:
            if ckeys & nkeys:
                for mod, val in mods.items():
                    nat["mods"][mod] = nat["mods"].get(mod, 0.0) + val
                nat["entries"].append({"source": "culture", "id": entry_id, "mods": mods})

    # Religion content: each religion's definition_modifier trade mods plus the religion- and
    # religion-group-gated entries collected above (group-gated content reaches every member religion).
    religions = scrape_religions()
    religion_bonuses = {}
    group_members = {}
    for rel, ri in religions.items():
        if ri["group"]:
            group_members.setdefault(ri["group"], []).append(rel)
        if ri["def_mods"]:
            b = religion_bonuses.setdefault(rel, {"mods": {}, "entries": []})
            for mod, val in ri["def_mods"].items():
                b["mods"][mod] = b["mods"].get(mod, 0.0) + val
            b["entries"].append({"source": "religion_def", "id": rel, "mods": ri["def_mods"]})
    for rkeys, source, entry_id, mods in religion_entries:
        targets = set()
        for k in rkeys:
            kind, _, val = k.partition(":")
            if kind == "religion":
                targets.add(val)
            elif kind == "group":
                targets.update(group_members.get(val, []))
        for rel in targets:
            b = religion_bonuses.setdefault(rel, {"mods": {}, "entries": []})
            for mod, val in mods.items():
                b["mods"][mod] = b["mods"].get(mod, 0.0) + val
            b["entries"].append({"source": source, "id": entry_id, "mods": mods})

    for label, excluded, src_map in (("law policies", excluded_laws, "LAW_TAG"),
                                     ("government reforms", excluded_reforms, "REFORM_TAG"),
                                     ("cabinet actions", excluded_cabinet, "CABINET_TAG")):
        if excluded:
            print(f"  [trade] {len(excluded)} trade {label} excluded as generic (no tag/gov type, not in {src_map}):")
            for x in sorted(set(excluded)):
                print(f"    - {x}")
    if unresolved:
        print(f"  [trade] {len(unresolved)} unresolved modifier tokens: {sorted(unresolved)}")

    return nations, gov_bonuses, religion_bonuses


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

    print("Scraping forts...")
    forts = scrape_forts()

    print("Scraping pop demands...")
    pop_demands = scrape_pop_demands()

    print("Scraping trade nations...")
    trade_nations, gov_bonuses, religion_bonuses = scrape_trade_nations()

    print("Scraping country governments...")
    country_governments = scrape_country_governments()

    print("Scraping formable formers...")
    formables = scrape_formable_formers(set(trade_nations), country_governments)

    by_religion = {t: i["religion"] for t, i in scrape_country_cultures().items() if i.get("religion")}

    print("Parsing unit type files...")
    raw_units = parse_directory(COMMON_DIR / "unit_types")

    # Filter to only dict entries (skip non-unit top-level keys)
    raw_units = {k: v for k, v in raw_units.items() if isinstance(v, dict)}

    print(f"  Found {len(raw_units)} unit definitions")

    global GENERIC_UNIT_NAMES
    GENERIC_UNIT_NAMES = {
        k for k, v in parse_file(COMMON_DIR / "unit_types" / "2_unlocked_through_tech.txt").items()
        if isinstance(v, dict)
    }

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

    # Separate land vs naval by the category's is_army flag
    land_cats = {name for name, c in categories.items() if c.get("is_army")}
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
    n_cats = len([k for k in prices if k != "unit_adds"])
    print(f"  Wrote unit_prices.json ({n_cats} categories, {len(prices.get('unit_adds', {}))} unit adds)")

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

    with open(OUTPUT_DIR / "forts.json", "w") as f:
        json.dump(forts, f, indent=2)
    print(f"  Wrote forts.json ({len(forts)} fort buildings)")

    with open(OUTPUT_DIR / "pop_demands.json", "w") as f:
        json.dump(pop_demands, f, indent=2)
    print(f"  Wrote pop_demands.json ({len(pop_demands['goods'])} goods)")

    with open(OUTPUT_DIR / "trade_nations.json", "w") as f:
        json.dump(trade_nations, f, indent=2)
    print(f"  Wrote trade_nations.json ({len(trade_nations)} nations)")

    with open(OUTPUT_DIR / "trade_gov_types.json", "w") as f:
        json.dump({"by_tag": country_governments, "bonuses": gov_bonuses, "formables": formables,
                   "religion_bonuses": religion_bonuses, "by_religion": by_religion}, f, indent=2)
    print(f"  Wrote trade_gov_types.json ({len(country_governments)} tags, {len(gov_bonuses)} gov types, "
          f"{len(formables)} formables, {len(religion_bonuses)} religions)")

    print("Done!")


if __name__ == "__main__":
    main()
