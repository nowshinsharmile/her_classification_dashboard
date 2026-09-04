from __future__ import annotations

import re
import warnings
from functools import lru_cache
from datetime import datetime
from io import BytesIO, StringIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt


# ============================================================
# SETTINGS
# ============================================================

warnings.filterwarnings("ignore", message=r"Print area cannot be set to Defined name.*", category=UserWarning, module=r"openpyxl\.reader\.workbook")

DEFAULT_EXCEL_PATH = r"C:\Users\nowsh\Downloads\Research\Fall 2025\What's In The Foods You Eat Search Tool 2021-2023.xlsx"
DEFAULT_SHEET_NAME = "Search Tool"
DEFAULT_FBCENC_CLASSIFICATION_FILENAME = "FBCENC_HER_Classifications.xlsx"
REFERENCE_MEASURE = "Quantity not specified"
MAX_VISIBLE_RESULTS = 40
FOOD_CODE, FOOD_DESCRIPTION, NUMBER_SERVINGS, MEASURE_DESCRIPTION, SERVING_WEIGHT = "Food Code", "Food Description", "Number of Servings", "Measure Description", "Serving Weight"
BASE_COLUMNS = [FOOD_CODE, FOOD_DESCRIPTION, NUMBER_SERVINGS, MEASURE_DESCRIPTION, SERVING_WEIGHT]
KEY_NUTRIENTS = ["Energy", "Protein", "Carbohydrate", "Total Fat", "Total Sugars", "Dietary Fiber"]
SUMMARY_NUTRIENTS = ["Energy", "Protein", "Carbohydrate", "Total Fat", "Total Sugars", "Dietary Fiber", "Water", "Cholesterol", "Total Saturated", "Total Monounsaturated", "Total Polyunsaturated", "Calcium", "Iron", "Magnesium", "Phosphorus", "Potassium", "Sodium", "Zinc", "Copper", "Selenium", "Vit A, RAE", "Retinol", "Carotene, Alpha", "Carotene, Beta", "Cryptoxanthin, Beta", "Lycopene", "Lutein + Zeaxanthin", "Vit E", "Vit E, added", "Vit D", "Vit K", "Vit C", "Thiamin", "Riboflavin", "Niacin", "Vit B6", "Folate, Total", "Folate, DFE", "Folic Acid", "Folate, Food", "Vit B12", "Vit B12, added", "Choline", "Alcohol", "Caffeine", "Theobromine"]
NUTRIENT_CATEGORIES = {
    "Macronutrients": ["Energy", "Protein", "Carbohydrate", "Total Fat", "Total Sugars", "Dietary Fiber", "Water", "Alcohol"],
    "Fatty acids": ["Total Saturated", "4:0", "6:0", "8:0", "10:0", "12:0", "14:0", "16:0", "18:0", "Total Monounsaturated", "16:1", "18:1", "20:1", "22:1", "Total Polyunsaturated", "18:2", "18:3", "18:4", "20:4", "20:5 n-3", "22:5 n-3", "22:6 n-3"],
    "Minerals": ["Calcium", "Iron", "Magnesium", "Phosphorus", "Potassium", "Sodium", "Zinc", "Copper", "Selenium"],
    "Vitamins": ["Vit A, RAE", "Retinol", "Carotene, Alpha", "Carotene, Beta", "Cryptoxanthin, Beta", "Lycopene", "Lutein + Zeaxanthin", "Vit E", "Vit E, added", "Vit D", "Vit K", "Vit C", "Thiamin", "Riboflavin", "Niacin", "Vit B6", "Folate, Total", "Folate, DFE", "Folic Acid", "Folate, Food", "Vit B12", "Vit B12, added", "Choline"],
    "Other": ["Cholesterol", "Alcohol", "Caffeine", "Theobromine"],
}
HER_CATEGORIES = ["Fruits and Vegetables", "Grains", "Protein", "Dairy", "Non-Dairy Alternatives", "Beverages", "Mixed Dishes", "Processed and Packaged Snacks", "Desserts", "Infant", "Condiments and Cooking Staples", "Miscellaneous Products"]
HER_LEVEL = {"Choose Often": 0, "Choose Sometimes": 1, "Choose Rarely": 2}
HER_ICONS = {"Choose Often": "🟢", "Choose Sometimes": "🟡", "Choose Rarely": "🔴", "Unranked": "⚫", "Not Ranked": "⚪", "Needs Review": "⚪", "Assorted": "◐", "Non Food": "▫"}
# Semantic colors used consistently in banners, load-composition cards, and charts.
HER_COLORS = {
    "Choose Often": "#2E7D32",
    "Choose Sometimes": "#FBC02D",
    "Choose Rarely": "#C62828",
    "Unranked": "#424242",
    "Not Ranked": "#BDBDBD",
    "Needs Review": "#757575",
    "Assorted": "#607D8B",
    "Non Food": "#90A4AE",
}
HER_TEXT_COLORS = {
    "Choose Often": "#FFFFFF",
    "Choose Sometimes": "#111111",
    "Choose Rarely": "#FFFFFF",
    "Unranked": "#FFFFFF",
    "Not Ranked": "#111111",
    "Needs Review": "#FFFFFF",
    "Assorted": "#FFFFFF",
    "Non Food": "#111111",
}

# Any description containing baby, infant or toddler is placed in Infant and forced Unranked.
INFANT_TERMS = ["baby", "infant", "toddler", "enfamil", "infamil", "similac", "gerber good start", "good start formula"]

# Protein powders are also forced Unranked.
PROTEIN_POWDER_TERMS = ["protein powder", "powdered protein", "whey protein", "casein protein","nutritional powder", "soy protein powder", "plant protein powder", "plant-based protein powder", "protein supplement powder", "protein shake powder"]

# These protein products are explicitly Choose Rarely under the HER guideline.
FORCED_RED_PROTEIN_TERMS = ["refried beans", "deli meat", "deli meats", "luncheon meat", "processed meat", "sausage", "bacon", "breaded chicken", "pastrami"]

# Flavored liquid milk remains Dairy even when the flavor word is also a dessert keyword.
FLAVORED_MILK_TERMS = ["chocolate milk", "strawberry milk", "vanilla milk", "flavored milk", "flavoured milk", "banana milk"]
MILK_CHOCOLATE_TERMS = ["milk chocolate"]


# Explicit whole-grain wording anywhere in the food description counts as whole grain.
# Grain names alone, such as quinoa, bulgur, millet, sorghum or teff, do not count.
WHOLE_GRAIN_PHRASES = ["whole grain", "whole wheat", "whole oat", "whole oats", "whole rye", "whole barley", "whole corn", "brown rice"]

# Category order matters because the first matching group is used.
CATEGORY_KEYWORDS = {
    "Desserts": ["ice cream", "milkshake", "milk shake", "frozen milk dessert", "frozen dairy dessert", "baklava", "frozen yogurt", "milk pudding", "rice pudding", "dairy dessert","sherbet","cobbler","churros","tiramisu","trifle","creme brulee", "chocolate", "cookie", "cake", "brownie", "pastry", "pie", "dessert", "donut", "doughnut", "cupcake", "pudding","ambrosia", "candy", "cheesecake", "custard", "flan", "gelato", "creamsicle", "fudgesicle", "waffle"],
    "Non-Dairy Alternatives": ["soy milk", "almond milk", "oat milk", "rice milk", "plant-based milk", "non-dairy milk", "nondairy milk", "vegan cheese", "plant-based cheese", "plant-based yogurt", "soy yogurt", "coconut milk", "agave"],
    "Processed and Packaged Snacks": ["chips", "cracker", "granola bar", "snack bar", "popcorn", "pretzel", "corn chip", "potato chip", "tortilla chip"],
    "Mixed Dishes": ["big mac","lo mein","pad thai","pizza", "sandwich", "burger", "burrito", "taco", "casserole", "stew", "soup", "frozen meal","cheeseburger","hamburger","chili", "macaroni and cheese", "lasagna", "pasta dish", "rice dish", "mixed dish"],
    "Beverages": ["beverage", "drink", "soda", "coffee", "tea", "water", "sports drink", "energy drink", "latte", "cappuccino", "lemonade", "fruit drink", "soft drink"],
    "Dairy": ["milk", "cheese", "yogurt", "yoghurt", "cottage cheese", "cream cheese", "buttermilk", "strawberry milk", "cream", "goat milk", "kefir"],
    "Protein": ["nuts","turtle","chickpeas","armadillo","beaver", "almonds","beans","beef", "pork", "chicken", "turkey", "fish","codfish", "seafood", "salmon", "tuna","goat", "shrimp", "crab","veal","moose","bear","caribou","duck","goose","quail","pepperoni","salami","spam","fish", "catfish", "clam", "clams", "cod", "flounder", "sole", "haddock", "halibut", "lobster", "perch", "roughy", "oyster", "oysters", "pollock", "trout", "rockfish", "scallop", "scallops", "swordfish", "tilapia", "egg", "sausage", "ham", "hot dog", "tofu", "tempeh", "bean", "lentil", "chickpea", "peanut butter", "nut", "almond", "cashew", "peanut", "abalone", "seed", "veggie burger", "bacon", "ribs", "lamb", "goat", "veal", "rabbit"],
    "Grains": ["bread", "rice", "pasta", "cereal", "oatmeal", "oats", "tortilla", "bagel", "roll", "bun", "grain", "noodle", "couscous", "quinoa"],
    "Fruits and Vegetables": ["watermelon","lime","papaya","zucchini","olive","okra","plum","pineapple","persimmon","raspberry","pumpkin","radish","mushroom","nectarine","apple","applesauce","artichoke","asparagus","avocado","apricot", "basil","blueberries","blueberry","cherries","blackberries","banana","squash", "green beans","mushroom","celery","tomato","sweeet potato","carrot", "orange", "berry", "berries", "fruit", "vegetable", "broccoli", "spinach", "carrot", "potato", "tomato", "peas", "corn", "cauliflower", "cabbage", "lettuce", "pepper", "onion", "melon", "grape", "mango", "peach", "pear"],
    "Condiments and Cooking Staples": ["seasoning","soy sauce", "spice", "salt", "sugar", "oil", "butter", "margarine", "dressing", "mayonnaise", "mustard", "ketchup", "syrup", "flour", "shortening", "cooking spray"],
}

# Fresh-produce matching uses whole words and rejects processed forms first.
FRESH_FRUIT_TERMS = ["apple", "apples", "avocado", "avocados", "banana", "bananas", "cantaloupe", "grapefruit", "grape", "grapes", "honeydew", "kiwifruit", "kiwi", "lemon", "lemons", "lime", "limes", "nectarine", "nectarines", "orange", "oranges", "peach", "peaches", "pear", "pears", "pineapple", "pineapples", "plum", "plums", "strawberry", "strawberries", "cherry", "cherries", "tangerine", "tangerines", "watermelon", "watermelons", "blueberry", "blueberries", "raspberry", "raspberries", "blackberry", "blackberries", "mango", "mangoes", "papaya", "papayas", "pomegranate", "pomegranates"]
FRESH_VEGETABLE_TERMS = ["artichoke", "asparagus", "beet", "beets", "broccoli", "cabbage", "carrot", "carrots", "cauliflower", "celery", "cucumber", "cucumbers", "eggplant", "garlic", "lettuce", "mushroom", "mushrooms", "okra", "onion", "onions", "pepper", "peppers", "potato", "potatoes", "spinach", "squash", "tomato", "tomatoes", "turnip", "turnips", "zucchini", "corn", "peas", "green beans", "snap peas"]
FRESH_FORM_TERMS = ["fresh", "raw", "whole", "uncooked", "plain"]
PROCESSED_PRODUCE_TERMS = ["juice", "juice drink", "fruit drink", "nectar", "smoothie", "dried", "dry", "dehydrated", "freeze dried", "fruit leather", "canned", "packed in syrup", "heavy syrup", "light syrup", "sweetened", "sugar added", "with sugar", "candied", "glazed", "jam", "jelly", "preserves", "marmalade", "spread", "sauce", "applesauce", "puree", "pulp", "concentrate", "pie", "cake", "cookie", "muffin", "bread", "pudding", "cobbler", "crisp", "tart", "pastry", "dessert", "flavored", "flavour", "flavor", "syrup", "soda", "yogurt", "ice cream", "frozen dessert", "cereal", "fried", "deep fried", "breaded", "chips", "chip", "pickled", "pickle", "relish", "soup", "stew", "baby food", "infant food", "formula"]

# Supplemental FDA-style fruit table. These rows are appended to the workbook data.
SUPPLEMENTAL_FRUIT_CSV = '''Food Description,Measure Description,Serving Weight,Energy,Calories from Fat,Total Fat,Total Fat (%DV),Sodium,Sodium (%DV),Potassium,Potassium (%DV),Carbohydrate,Carbohydrate (%DV),Dietary Fiber,Dietary Fiber (%DV),Total Sugars,Protein,Vitamin A (%DV),Vitamin C (%DV),Calcium (%DV),Iron (%DV)
Apple,1 large,242,130,0,0,0,0,0,260,7,34,11,5,20,25,1,2,8,2,2
"Avocado, California",1/5 medium,30,50,35,4.5,7,0,0,140,4,3,1,1,4,0,1,0,4,0,2
Banana,1 medium,126,110,0,0,0,0,0,450,13,30,10,3,12,19,1,2,15,0,2
Cantaloupe,1/4 medium,134,50,0,0,0,20,1,240,7,12,4,1,4,11,1,120,80,2,2
Grapefruit,1/2 medium,154,60,0,0,0,0,0,160,5,15,5,2,8,11,1,35,100,4,0
Grapes,3/4 cup,126,90,0,0,0,15,1,240,7,23,8,1,4,20,0,0,2,2,0
Honeydew Melon,1/10 medium melon,134,50,0,0,0,30,1,210,6,12,4,1,4,11,1,2,45,2,2
Kiwifruit,2 medium,148,90,10,1,2,0,0,450,13,20,7,4,16,13,1,2,240,4,2
Lemon,1 medium,58,15,0,0,0,0,0,75,2,5,2,2,8,2,0,0,40,2,0
Lime,1 medium,67,20,0,0,0,0,0,75,2,7,2,2,8,0,0,0,35,0,0
Nectarine,1 medium,140,60,5,0.5,1,0,0,250,7,15,5,2,8,11,1,8,15,0,2
Orange,1 medium,154,80,0,0,0,0,0,250,7,19,6,3,12,14,1,2,130,6,0
Peach,1 medium,147,60,0,0.5,1,0,0,230,7,15,5,2,8,13,1,6,15,0,2
Pear,1 medium,166,100,0,0,0,0,0,190,5,26,9,6,24,16,1,0,10,2,0
Pineapple,"2 slices, 3-inch diameter, 3/4-inch thick",112,50,0,0,0,10,0,120,3,13,4,1,4,10,1,2,50,2,2
Plums,2 medium,151,70,0,0,0,0,0,230,7,19,6,2,8,16,1,8,10,0,2
Strawberries,8 medium,147,50,0,0,0,0,0,170,5,11,4,2,8,8,1,0,160,2,2
Sweet Cherries,"21 cherries; 1 cup",140,100,0,0,0,0,0,350,10,26,9,1,4,16,1,2,15,2,2
Tangerine,1 medium,109,50,0,0,0,0,0,160,5,13,4,2,8,9,1,6,45,4,0
Watermelon,"1/18 medium melon; 2 cups diced",280,80,0,0,0,0,0,270,8,21,7,1,4,20,1,30,25,2,4'''


# Supplemental seafood table. All entries use the FDA serving size of 84 g / 3 oz.
SUPPLEMENTAL_SEAFOOD_CSV = '''Food Description,Measure Description,Serving Weight,Energy,Calories from Fat,Total Fat,Total Fat (%DV),Total Saturated,Total Saturated (%DV),Cholesterol,Cholesterol (%DV),Sodium,Sodium (%DV),Potassium,Potassium (%DV),Carbohydrate,Carbohydrate (%DV),Protein,Vitamin A (%DV),Vitamin C (%DV),Calcium (%DV),Iron (%DV)
Blue Crab,3 oz,84,100,10,1,2,0,0,95,32,330,14,300,9,0,0,20,0,4,10,4
Catfish,3 oz,84,130,60,6,9,2,10,50,17,40,2,230,7,0,0,17,0,0,0,0
"Clams, about 12 small",3 oz,84,110,15,1.5,2,0,0,80,27,95,4,470,13,6,2,17,10,0,8,30
Cod,3 oz,84,90,5,1,2,0,0,50,17,65,3,460,13,0,0,20,0,2,2,2
Flounder/Sole,3 oz,84,100,15,1.5,2,0,0,55,18,100,4,390,11,0,0,19,0,0,2,0
Haddock,3 oz,84,100,10,1,2,0,0,70,23,85,4,340,10,0,0,21,2,0,2,6
Halibut,3 oz,84,120,15,2,3,0,0,40,13,60,3,500,14,0,0,23,4,0,2,6
Lobster,3 oz,84,80,0,0.5,1,0,0,60,20,320,13,300,9,1,0,17,2,0,6,2
Ocean Perch,3 oz,84,110,20,2,3,0.5,3,45,15,95,4,290,8,0,0,21,0,2,10,4
Orange Roughy,3 oz,84,80,5,1,2,0,0,20,7,70,3,340,10,0,0,16,2,0,4,2
"Oysters, about 12 medium",3 oz,84,100,35,4,6,1,5,80,27,300,13,220,6,6,2,10,0,6,6,45
Pollock,3 oz,84,90,10,1,2,0,0,80,27,110,5,370,11,0,0,20,2,0,0,2
Rainbow Trout,3 oz,84,140,50,6,9,2,10,55,18,35,1,370,11,0,0,20,4,4,8,2
Rockfish,3 oz,84,110,15,2,3,0,0,40,13,70,3,440,13,0,0,21,4,0,2,2
Salmon Atlantic/Coho/Sockeye/Chinook,3 oz,84,200,90,10,15,2,10,70,23,55,2,430,12,0,0,24,4,4,2,2
Salmon Chum/Pink,3 oz,84,130,40,4,6,1,5,70,23,65,3,420,12,0,0,22,2,0,2,4
"Scallops, about 6 large or 14 small",3 oz,84,140,10,1,2,0,0,65,22,310,13,430,12,5,2,27,2,0,4,14
Shrimp,3 oz,84,100,10,1.5,2,0,0,170,57,240,10,220,6,0,0,21,4,4,6,10
Swordfish,3 oz,84,120,50,6,9,1.5,8,40,13,100,4,310,9,0,0,16,2,2,0,6
Tilapia,3 oz,84,110,20,2.5,4,1,5,75,25,30,1,360,10,0,0,22,0,2,0,2
Tuna,3 oz,84,130,15,1.5,2,0,0,50,17,40,2,480,14,0,0,26,2,2,2,4'''


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(page_title="HER Food Classification", page_icon="🥗", layout="wide")
st.markdown("""<style>
.block-container{max-width:1180px;padding-top:1.4rem;padding-bottom:3rem}
[data-testid="stSidebar"]{min-width:285px}
.food-title{font-size:2rem;font-weight:750;line-height:1.15;margin:.1rem 0 .2rem}
.food-subtitle{color:#6b7280;font-size:.95rem;margin-bottom:.8rem}
.search-note{color:#6b7280;font-size:.9rem;margin-top:-.3rem;margin-bottom:.8rem}
.alias-chip{display:inline-block;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:999px;padding:.2rem .55rem;margin:.15rem .25rem .15rem 0;font-size:.82rem;color:#374151}
.status-card{border:1px solid rgba(128,128,128,.25);border-radius:12px;padding:.85rem 1rem;min-height:112px}
.status-label{font-size:.82rem;color:#6b7280;text-transform:uppercase;letter-spacing:.03em}
.status-value{font-size:1.25rem;font-weight:700;margin:.2rem 0}
.status-rank{font-size:.92rem;font-weight:600}
div[data-testid="stMetric"]{border:1px solid rgba(128,128,128,.22);border-radius:10px;padding:.65rem .75rem}
</style>""", unsafe_allow_html=True)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(value: object) -> str:
    if pd.isna(value): return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ").replace("\r", " ")).strip()


def clean_header(value: object) -> str: return clean_text(value).strip('"').strip("'")


def normalize_column_name(value: object) -> str:
    text = clean_header(value)
    aliases = {"Food code": FOOD_CODE, "Food description": FOOD_DESCRIPTION, "Number servings": NUMBER_SERVINGS, "Number of servings": NUMBER_SERVINGS, "Measure description": MEASURE_DESCRIPTION, "Serving weight": SERVING_WEIGHT, "Total monounsaturated": "Total Monounsaturated", "Total polyunsaturated": "Total Polyunsaturated", "Total saturated": "Total Saturated"}
    return aliases.get(text, text)


def make_unique_columns(columns: list[str]) -> list[str]:
    counts, output = {}, []
    for column in columns:
        base = column or "Unnamed"; counts[base] = counts.get(base, -1) + 1; output.append(base if counts[base] == 0 else f"{base} [{counts[base]}]")
    return output


def is_group_heading(value: object) -> bool:
    text = clean_header(value).lower(); terms = ["saturated fatty acids", "monounsaturated fatty acids", "polyunsaturated fatty acids"]
    return any(term in text for term in terms) or bool(re.fullmatch(r"[-\s]+", text))


def normalize_food_code(value: object) -> str:
    text = clean_text(value)
    return text[:-2] if re.fullmatch(r"\d+\.0", text) else text


def format_number(value: object) -> str:
    if pd.isna(value): return "—"
    try: number = float(value)
    except (TypeError, ValueError): return clean_text(value)
    if number == 0: return "0"
    if abs(number) >= 1000: return f"{number:,.1f}".rstrip("0").rstrip(".")
    if abs(number) >= 100: return f"{number:.1f}".rstrip("0").rstrip(".")
    if abs(number) >= 10: return f"{number:.2f}".rstrip("0").rstrip(".")
    return f"{number:.3f}".rstrip("0").rstrip(".")


def scale_value(value: object, multiplier: float) -> object:
    if pd.isna(value): return np.nan
    try: return float(value) * multiplier
    except (TypeError, ValueError): return value


def excel_column_name(index: int) -> str:
    name = ""
    while index >= 0: index, remainder = divmod(index, 26); name = chr(65 + remainder) + name; index -= 1
    return name


def normalize_description_for_matching(description: str) -> str:
    text = clean_text(description).lower(); text = re.sub(r"[/_,;:()\[\]{}\-]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


EXCLUSION_ALIAS_PATTERNS = [
    r"\bother than\b",
    r"\bexcept(?: for)?\b",
    r"\bexcluding\b",
    r"\bexclude(?:s|d)?\b",
    r"\bnot including\b",
    r"\bdoes not include\b",
    r"\bwithout\b",
]


def split_food_description(description: str) -> tuple[str, list[str]]:
    """Split a source description into its canonical name and raw Includes fragments."""
    raw = clean_text(description)
    match = re.search(r"\s*\(Includes:\s*(.*?)\)\s*$", raw, flags=re.IGNORECASE)
    if not match:
        return raw, []
    canonical = raw[:match.start()].strip().rstrip(" ,;")
    fragments = [clean_text(part) for part in match.group(1).split(";") if clean_text(part)]
    return canonical or raw, fragments


def is_exclusion_alias(alias: str) -> bool:
    """True when an Includes fragment describes something the row explicitly excludes."""
    text = normalize_description_for_matching(alias)
    return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in EXCLUSION_ALIAS_PATTERNS)


def split_search_aliases(description: str) -> tuple[list[str], list[str]]:
    """Return (positive aliases, exclusion notes). Exclusion notes are never searchable aliases."""
    _, fragments = split_food_description(description)
    positive, exclusions = [], []
    for fragment in fragments:
        (exclusions if is_exclusion_alias(fragment) else positive).append(fragment)
    return positive, exclusions


def canonical_food_name(description: str) -> str:
    return split_food_description(description)[0]


def food_aliases(description: str) -> list[str]:
    return split_search_aliases(description)[0]


def food_exclusions(description: str) -> list[str]:
    return split_search_aliases(description)[1]


def compact_food_label(description: str, max_aliases: int = 2) -> str:
    canonical = canonical_food_name(description)
    aliases = food_aliases(description)
    if not aliases:
        return canonical
    alias_text = "; ".join(aliases[:max_aliases])
    if len(aliases) > max_aliases:
        alias_text += "; …"
    label = f"{canonical} — also: {alias_text}"
    return label if len(label) <= 150 else label[:147].rstrip() + "…"


def normalized_search_text(description: str, category: str = "") -> str:
    """Search only the canonical food name and positive aliases; never index exclusion notes."""
    canonical = canonical_food_name(description)
    aliases = food_aliases(description)
    return normalize_description_for_matching(" ".join([canonical, *aliases]))


def simple_singular(word: str) -> str:
    """Return a conservative singular form for basic s/es/ies plurals."""
    word = word.lower()
    if len(word) > 4 and word.endswith("ies") and word[-4] not in "aeiou": return word[:-3] + "y"
    if len(word) > 4 and word.endswith(("ches", "shes", "xes", "zes", "sses")): return word[:-2]
    if len(word) > 3 and word.endswith("s") and not word.endswith(("ss", "us", "is")): return word[:-1]
    return word


def plural_forms(word: str) -> set[str]:
    """Generate conservative singular/plural forms for one keyword token."""
    base = simple_singular(word)
    forms = {word.lower(), base}
    if len(base) > 1 and base.endswith("y") and base[-2] not in "aeiou": forms.add(base[:-1] + "ies")
    elif base.endswith(("s", "x", "z", "ch", "sh")): forms.add(base + "es")
    else: forms.add(base + "s")
    return {form for form in forms if form}


@lru_cache(maxsize=2048)
def plural_aware_phrase_pattern(phrase: str) -> str:
    """Build a whole-word regex that accepts s, es and ies on the phrase's final word."""
    normalized = normalize_description_for_matching(phrase)
    words = normalized.split()
    if not words: return r"(?!)"
    prefix = [re.escape(word) for word in words[:-1]]
    last_forms = sorted(plural_forms(words[-1]), key=len, reverse=True)
    last_pattern = "(?:" + "|".join(re.escape(form) for form in last_forms) + ")"
    body = r"\s+".join(prefix + [last_pattern])
    return r"(?<!\w)" + body + r"(?!\w)"


def contains_phrase(text: str, phrase: str) -> bool:
    """Match a phrase as whole words while accepting common final-word plurals."""
    normalized_text = normalize_description_for_matching(text)
    return re.search(plural_aware_phrase_pattern(phrase), normalized_text) is not None


def contains_any_phrase(description: str, phrases: list[str]) -> bool:
    text = normalize_description_for_matching(description)
    return any(contains_phrase(text, phrase) for phrase in phrases)


@lru_cache(maxsize=50000)
def whole_grain_is_first_in_description(description: str) -> bool:
    """Return True when explicit whole-grain wording appears anywhere in the food name."""
    return contains_any_phrase(description, WHOLE_GRAIN_PHRASES)


@lru_cache(maxsize=50000)
def is_fresh_produce_description(description: str) -> bool:
    """Apple/raw apple -> True; apple juice, applesauce, dried apple, apple pie -> False."""
    text = normalize_description_for_matching(description)
    if not text or any(contains_phrase(text, term) for term in PROCESSED_PRODUCE_TERMS): return False
    produce_terms = FRESH_FRUIT_TERMS + FRESH_VEGETABLE_TERMS
    matched = [term for term in produce_terms if contains_phrase(text, term)]
    if not matched: return False
    if any(contains_phrase(text, term) for term in FRESH_FORM_TERMS): return True
    allowed = {"large", "medium", "small", "extra", "mini", "seedless", "red", "green", "yellow", "golden", "white", "purple", "california", "domestic", "imported", "organic", "with", "skin", "without", "peeled", "unpeeled"}
    words, produce_words = set(re.findall(r"[a-z]+", text)), set()
    for term in matched: produce_words.update(re.findall(r"[a-z]+", term))
    return len(words - produce_words - allowed) == 0


# ============================================================
# EXCEL LOADING AND PREPARATION
# ============================================================

def detect_primary_header_row(raw_df: pd.DataFrame) -> int:
    expected, best_row, best_score = {"food code", "food description", "number of servings", "measure description", "serving weight"}, -1, -1
    for row_index in range(min(30, len(raw_df))):
        values = {clean_header(value).lower() for value in raw_df.iloc[row_index].tolist() if clean_header(value)}; score = len(expected.intersection(values))
        if score > best_score: best_row, best_score = row_index, score
        if score == len(expected): return row_index
    if best_score < 4: raise ValueError("Could not detect the Excel header row.")
    return best_row


def construct_columns(raw_df: pd.DataFrame, header_row: int) -> list[str]:
    top, lower = raw_df.iloc[header_row].tolist(), raw_df.iloc[header_row + 1].tolist()
    last_column = max(max((i for i, value in enumerate(top) if clean_header(value)), default=0), max((i for i, value in enumerate(lower) if clean_header(value)), default=0))
    columns = []
    for i in range(last_column + 1):
        top_name, lower_name = clean_header(top[i]), clean_header(lower[i]); name = top_name or lower_name if i < len(BASE_COLUMNS) else lower_name or (top_name if top_name and not is_group_heading(top_name) else "")
        columns.append(normalize_column_name(name))
    return make_unique_columns(columns)


def build_unit_mapping(raw_df: pd.DataFrame, header_row: int, columns: list[str]) -> dict[str, str]:
    values = raw_df.iloc[header_row + 2].tolist(); units = {column: clean_text(values[i]) if i < len(values) else "" for i, column in enumerate(columns)}; units[SERVING_WEIGHT] = "g"
    return units


def coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    result, protected = df.copy(), {FOOD_CODE, FOOD_DESCRIPTION, MEASURE_DESCRIPTION}
    for column in result.columns:
        if column in protected: continue
        nonblank = result[column].notna() & result[column].astype(str).str.strip().ne("")
        if nonblank.sum() == 0: continue
        converted = pd.to_numeric(result[column], errors="coerce")
        if converted.loc[nonblank].notna().mean() >= 0.70: result[column] = converted
    return result


def create_serving_label(row: pd.Series) -> str:
    number, measure, weight = row.get(NUMBER_SERVINGS, 1), clean_text(row.get(MEASURE_DESCRIPTION, "")), row.get(SERVING_WEIGHT, np.nan)
    try: number_float = float(number); number_text = str(int(number_float)) if number_float.is_integer() else f"{number_float:g}"
    except (TypeError, ValueError): number_text = clean_text(number) or "1"
    return f"{number_text} × {measure} ({float(weight):g} g)" if pd.notna(weight) else f"{number_text} × {measure}"


def build_supplemental_fruit_dataframe() -> pd.DataFrame:
    supplemental = pd.read_csv(StringIO(SUPPLEMENTAL_FRUIT_CSV)); supplemental[FOOD_CODE] = [f"FRESH_FRUIT_{i:03d}" for i in range(1, len(supplemental) + 1)]; supplemental[NUMBER_SERVINGS] = 1.0; supplemental["_data_source"] = "FDA Fresh Fruit Table"; supplemental["_fresh_produce"] = True; supplemental["_seafood_assumption"] = False
    return supplemental


def build_supplemental_seafood_dataframe() -> pd.DataFrame:
    """Seafood uses 0.5 g sugar and 0 g dietary fiber when those values are unavailable."""
    supplemental = pd.read_csv(StringIO(SUPPLEMENTAL_SEAFOOD_CSV)); supplemental[FOOD_CODE] = [f"SEAFOOD_{i:03d}" for i in range(1, len(supplemental) + 1)]; supplemental[NUMBER_SERVINGS] = 1.0; supplemental["Total Sugars"] = 0.5; supplemental["Dietary Fiber"] = 0.0; supplemental["_data_source"] = "FDA Seafood Nutrition Table"; supplemental["_fresh_produce"] = False; supplemental["_seafood_assumption"] = True
    return supplemental


@st.cache_data(show_spinner=False)
def load_food_data(excel_path: str, sheet_name: str) -> tuple[pd.DataFrame, dict[str, str], int]:
    file_path = Path(excel_path)
    if not file_path.exists(): raise FileNotFoundError(f"Excel file not found:\n{excel_path}")
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl"); header_row = detect_primary_header_row(raw_df); columns = construct_columns(raw_df, header_row); units = build_unit_mapping(raw_df, header_row, columns)
    df = raw_df.iloc[header_row + 3:, :len(columns)].copy(); df.columns = columns; df = df.dropna(how="all").reset_index(drop=True)
    missing = [column for column in BASE_COLUMNS if column not in df.columns]
    if missing: raise ValueError("Missing required columns: " + ", ".join(missing))
    df[FOOD_CODE] = df[FOOD_CODE].astype("string").ffill().map(normalize_food_code); df[FOOD_DESCRIPTION] = df[FOOD_DESCRIPTION].astype("string").ffill().map(clean_text); df[MEASURE_DESCRIPTION] = df[MEASURE_DESCRIPTION].map(clean_text)
    df = df[df[FOOD_CODE].ne("") & df[FOOD_DESCRIPTION].ne("") & df[MEASURE_DESCRIPTION].ne("")].copy(); df = coerce_numeric_columns(df)
    df[NUMBER_SERVINGS] = pd.to_numeric(df[NUMBER_SERVINGS], errors="coerce").fillna(1.0); df[SERVING_WEIGHT] = pd.to_numeric(df[SERVING_WEIGHT], errors="coerce"); df = df.copy()

    reference_mask = df[MEASURE_DESCRIPTION].astype(str).str.strip().str.casefold().eq(REFERENCE_MEASURE.casefold())
    df = df.loc[reference_mask].copy()
    if df.empty:
        raise ValueError(f'No workbook rows were found with Measure Description = "{REFERENCE_MEASURE}".')

    df["_data_source"] = "What's In The Foods You Eat 2021-2023"; df["_fresh_produce"] = df[FOOD_DESCRIPTION].map(is_fresh_produce_description); df["_seafood_assumption"] = False
    df = pd.concat([df, build_supplemental_fruit_dataframe(), build_supplemental_seafood_dataframe()], ignore_index=True, sort=False)
    units.update({"Calories from Fat": "kcal", "Total Fat (%DV)": "%DV", "Total Saturated (%DV)": "%DV", "Cholesterol (%DV)": "%DV", "Sodium (%DV)": "%DV", "Potassium (%DV)": "%DV", "Carbohydrate (%DV)": "%DV", "Dietary Fiber (%DV)": "%DV", "Vitamin A (%DV)": "%DV", "Vitamin C (%DV)": "%DV", "Calcium (%DV)": "%DV", "Iron (%DV)": "%DV"})

    # Precompute expensive description-based fields once inside the cached loader.
    unique_descriptions = pd.Series(df[FOOD_DESCRIPTION].dropna().astype(str).unique())
    category_map = {description: suggest_her_category(description) for description in unique_descriptions}
    whole_grain_map = {description: whole_grain_is_first_in_description(description) for description in unique_descriptions}
    df["_her_category"] = df[FOOD_DESCRIPTION].map(category_map)
    df["_whole_grain_detected"] = df[FOOD_DESCRIPTION].map(whole_grain_map).fillna(False)
    df["_display_name"] = df[FOOD_DESCRIPTION].map(canonical_food_name)
    df["_aliases"] = df[FOOD_DESCRIPTION].map(lambda value: "; ".join(food_aliases(value)))
    df["_exclusion_notes"] = df[FOOD_DESCRIPTION].map(lambda value: "; ".join(food_exclusions(value)))
    df["_search_text"] = df[FOOD_DESCRIPTION].map(normalized_search_text)

    helper = pd.DataFrame({"_food_label": df[FOOD_DESCRIPTION].map(compact_food_label), "_serving_label": df.apply(create_serving_label, axis=1), "_row_id": np.arange(len(df), dtype=np.int64)}, index=df.index)
    return pd.concat([df, helper], axis=1).reset_index(drop=True), units, header_row + 1


# ============================================================
# NUTRIENT DISPLAY
# ============================================================

def get_unit(nutrient: str, units: dict[str, str]) -> str: return clean_text(units.get(nutrient, ""))


def nutrient_table(row: pd.Series, columns: list[str], units: dict[str, str], multiplier: float) -> pd.DataFrame:
    records = []
    for nutrient in columns:
        if nutrient not in row.index: continue
        value = scale_value(row[nutrient], multiplier)
        if pd.isna(value): continue
        records.append({"Nutrient": nutrient, "Amount": format_number(value), "Unit": get_unit(nutrient, units)})
    return pd.DataFrame(records)


def all_nutrient_columns(df: pd.DataFrame) -> list[str]:
    excluded = {FOOD_CODE, FOOD_DESCRIPTION, NUMBER_SERVINGS, MEASURE_DESCRIPTION, SERVING_WEIGHT, "_food_label", "_serving_label", "_row_id", "_data_source", "_fresh_produce", "_seafood_assumption", "_her_category", "_whole_grain_detected", "_display_name", "_aliases", "_search_text"}
    return [column for column in df.columns if column not in excluded and not column.startswith("Unnamed")]


def render_metric_cards(row: pd.Series, units: dict[str, str], multiplier: float) -> None:
    available = [nutrient for nutrient in KEY_NUTRIENTS if nutrient in row.index]
    if not available: return
    for metric_column, nutrient in zip(st.columns(len(available)), available): metric_column.metric(nutrient, f"{format_number(scale_value(row[nutrient], multiplier))} {get_unit(nutrient, units)}".strip())


# ============================================================
# CATEGORY, EXCLUSION AND HER CLASSIFICATION
# ============================================================

@lru_cache(maxsize=50000)
def forced_unranked_reason(description: str) -> str:
    """Return the Unranked reason for infant/toddler items and protein powders."""
    if contains_any_phrase(description, INFANT_TERMS): return "Baby, infant and toddler foods or formulas are Unranked."
    if contains_any_phrase(description, PROTEIN_POWDER_TERMS): return "Protein powders and powdered protein supplements are Unranked."
    return ""

@lru_cache(maxsize=50000)
def forced_unranked_category(description: str) -> str:
    """Any baby, infant or toddler item belongs to the Infant category."""
    return "Infant" if contains_any_phrase(description, INFANT_TERMS) else "Miscellaneous Products"

@lru_cache(maxsize=50000)
def forced_red_protein_reason(description: str) -> str:
    """Return the explicit protein Choose Rarely rule triggered by the description."""
    for term in FORCED_RED_PROTEIN_TERMS:
        if contains_any_phrase(description, [term]): return f"{term.title()} is classified as Choose Rarely under the protein guideline."
    return ""

@lru_cache(maxsize=50000)
def suggest_her_category(description: str) -> str:
    if forced_unranked_reason(description): return forced_unranked_category(description)
    if forced_red_protein_reason(description): return "Protein"
    # Word order matters: "milk chocolate" is a chocolate/candy product, while
    # "chocolate milk" is flavored liquid milk.
    if contains_any_phrase(description, MILK_CHOCOLATE_TERMS): return "Desserts"
    if contains_any_phrase(description, FLAVORED_MILK_TERMS): return "Dairy"
    text = normalize_description_for_matching(description)
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(contains_phrase(text, keyword) for keyword in keywords): return category
    return "Miscellaneous Products"

def numeric_from_columns(row: pd.Series, columns: list[str], multiplier: float = 1.0) -> tuple[float, str]:
    for column in columns:
        if column not in row.index or pd.isna(row[column]): continue
        try: return float(row[column]) * multiplier, column
        except (TypeError, ValueError): continue
    return np.nan, ""


def classify_standard(value: float, green_max: float, yellow_max: float, red_min: float) -> str:
    if pd.isna(value): return "Not Ranked"
    if value <= green_max: return "Choose Often"
    if value <= yellow_max: return "Choose Sometimes"
    return "Choose Rarely" if value >= red_min else "Choose Sometimes"


def classify_sat_green_yellow(value: float, green_max: float) -> str:
    if pd.isna(value): return "Not Ranked"
    return "Choose Often" if value <= green_max else "Choose Sometimes"


def classify_zero_green(value: float, yellow_max: float, red_min: float) -> str:
    if pd.isna(value): return "Not Ranked"
    if value == 0: return "Choose Often"
    if value <= yellow_max: return "Choose Sometimes"
    return "Choose Rarely" if value >= red_min else "Choose Sometimes"


def worst_rank(ranks: list[str]) -> str:
    available = [rank for rank in ranks if rank in HER_LEVEL]
    return max(available, key=lambda rank: HER_LEVEL[rank]) if available else "Not Ranked"


def empty_not_ranked_classification(reason: str) -> dict[str, str]: return {"Saturated Fat": "Not Ranked", "Sodium": "Not Ranked", "Sugar": "Not Ranked", "Overall": "Not Ranked", "Rule": reason}

def empty_unranked_classification(reason: str) -> dict[str, str]: return {"Saturated Fat": "Unranked", "Sodium": "Unranked", "Sugar": "Unranked", "Overall": "Unranked", "Rule": reason}
def forced_red_protein_classification(reason: str) -> dict[str, str]: return {"Saturated Fat": "Choose Rarely", "Sodium": "Choose Rarely", "Sugar": "Choose Rarely", "Overall": "Choose Rarely", "Rule": reason}


def fresh_produce_classification() -> dict[str, str]: return {"Saturated Fat": "Choose Often", "Sodium": "Choose Often", "Sugar": "Choose Often", "Overall": "Choose Often", "Rule": "Plain fresh produce is Choose Often; naturally occurring total sugar does not lower its rank."}


def classify_her(category: str, saturated_fat: float, sodium: float, sugar: float, whole_grain_first: bool, juice_or_dried_fruit: bool) -> dict[str, str]:
    result = empty_not_ranked_classification("")
    if category == "Infant": return empty_unranked_classification("Baby, infant and toddler foods or formulas are Unranked.")
    if category in {"Condiments and Cooking Staples", "Miscellaneous Products"}: result["Rule"] = "This category is Not Ranked under the guideline."; return result
    if category == "Desserts": return {"Saturated Fat": "Choose Rarely", "Sodium": "Choose Rarely", "Sugar": "Choose Rarely", "Overall": "Choose Rarely", "Rule": "Milk-based and other desserts are automatically Choose Rarely."}
    if category == "Fruits and Vegetables": sat, sodium_rank, sugar_rank = classify_sat_green_yellow(saturated_fat, 2), classify_standard(sodium, 230, 479, 480), classify_standard(sugar, 12, 23, 24)
    elif category == "Grains": sat, sodium_rank, sugar_rank = classify_sat_green_yellow(saturated_fat, 2), classify_standard(sodium, 230, 479, 480), classify_standard(sugar, 6, 11, 12)
    elif category == "Protein": sat, sodium_rank, sugar_rank = classify_standard(saturated_fat, 2, 4.5, 5), classify_standard(sodium, 230, 479, 480), classify_standard(sugar, 6, 11, 12)
    elif category == "Dairy": sat, sodium_rank, sugar_rank = classify_standard(saturated_fat, 3, 6, 6.5), classify_standard(sodium, 230, 479, 480), classify_standard(sugar, 12, 23, 24)
    elif category == "Non-Dairy Alternatives": sat, sodium_rank, sugar_rank = classify_sat_green_yellow(saturated_fat, 2), classify_standard(sodium, 230, 479, 480), classify_standard(sugar, 6, 11, 12)
    elif category == "Beverages": sat, sodium_rank, sugar_rank = classify_zero_green(saturated_fat, 0, 1), classify_zero_green(sodium, 140, 141), classify_zero_green(sugar, 11, 12)
    elif category == "Mixed Dishes": sat, sodium_rank, sugar_rank = classify_standard(saturated_fat, 3, 6, 6.5), classify_standard(sodium, 480, 599, 600), classify_standard(sugar, 6, 11, 12)
    elif category == "Processed and Packaged Snacks": sat = "Not Ranked" if pd.isna(saturated_fat) else ("Choose Sometimes" if saturated_fat <= 2 else "Choose Rarely"); sodium_rank = "Not Ranked" if pd.isna(sodium) else ("Choose Sometimes" if sodium <= 140 else "Choose Rarely"); sugar_rank = "Not Ranked" if pd.isna(sugar) else ("Choose Sometimes" if sugar <= 6 else "Choose Rarely")
    else: return result
    overall = worst_rank([sat, sodium_rank, sugar_rank])
    if category == "Grains" and not whole_grain_first and overall == "Choose Often": overall, result["Rule"] = "Choose Sometimes", "Whole grain must be the first ingredient for a grain to qualify as Choose Often."
    if category == "Processed and Packaged Snacks":
        if not whole_grain_first: overall, result["Rule"] = "Choose Rarely", "Whole grain must be the first ingredient for a packaged grain snack to qualify as Choose Sometimes."
        elif overall == "Choose Often": overall = "Choose Sometimes"
    if category == "Fruits and Vegetables" and juice_or_dried_fruit and overall == "Choose Often": overall, result["Rule"] = "Choose Sometimes", "100% juice and plain dried fruit cannot receive a final Choose Often classification."
    result.update({"Saturated Fat": sat, "Sodium": sodium_rank, "Sugar": sugar_rank, "Overall": overall}); return result


def classify_food_description(description: str, category: str, saturated_fat: float, sodium: float, sugar: float, whole_grain_first: bool, juice_or_dried_fruit: bool, fresh_produce: bool = False) -> dict[str, str]:
    unranked_reason = forced_unranked_reason(description)
    if unranked_reason: return empty_unranked_classification(unranked_reason)
    red_reason = forced_red_protein_reason(description)
    if red_reason: return forced_red_protein_classification(red_reason)
    if fresh_produce: return fresh_produce_classification()
    return classify_her(category, saturated_fat, sodium, sugar, whole_grain_first, juice_or_dried_fruit)

def render_her_banner(rank: str) -> None:
    st.markdown(f'''<div style="background:{HER_COLORS.get(rank, HER_COLORS['Not Ranked'])};color:black;border-radius:10px;padding:18px 20px;margin:8px 0 18px"><div style="font-size:.95rem;font-weight:600">Overall HER Classification</div><div style="font-size:1.8rem;font-weight:800">{HER_ICONS.get(rank, '⚪')} {rank}</div></div>''', unsafe_allow_html=True)


# ============================================================
# BACKEND EXCEL SUMMARY
# ============================================================

def backend_rule_table() -> pd.DataFrame:
    rows = [
        ["Fruits and Vegetables", "Choose Often", "≤2 g", "≤230 mg", "≤12 g", "Fresh produce override: Choose Often; juice/dried fruit cannot finish green"], ["Fruits and Vegetables", "Choose Sometimes", ">2 g", "231–479 mg", "13–23 g", ""], ["Fruits and Vegetables", "Choose Rarely", "No separate red tier", "≥480 mg", "≥24 g", ""],
        ["Grains", "Choose Often", "≤2 g", "≤230 mg", "≤6 g", "Whole grain first ingredient required"], ["Grains", "Choose Sometimes", ">2 g", "231–479 mg", "7–11 g", ""], ["Grains", "Choose Rarely", "No separate red tier", "≥480 mg", "≥12 g", ""],
        ["Protein", "Choose Often", "≤2 g", "≤230 mg", "≤6 g", ""], ["Protein", "Choose Sometimes", ">2–4.5 g", "231–479 mg", "7–11 g", ""], ["Protein", "Choose Rarely", "≥5 g", "≥480 mg", "≥12 g", "Refried beans, deli meat, sausage, bacon and breaded chicken are forced red"],
        ["Dairy", "Choose Often", "≤3 g", "≤230 mg", "≤12 g", ""], ["Dairy", "Choose Sometimes", ">3–6 g", "231–479 mg", "13–23 g", ""], ["Dairy", "Choose Rarely", "≥6.5 g", "≥480 mg", "≥24 g", ""],
        ["Non-Dairy Alternatives", "Choose Often", "≤2 g", "≤230 mg", "≤6 g", ""], ["Non-Dairy Alternatives", "Choose Sometimes", ">2 g", "231–479 mg", "7–11 g", ""], ["Non-Dairy Alternatives", "Choose Rarely", "No separate red tier", "≥480 mg", "≥12 g", ""],
        ["Beverages", "Choose Often", "0 g", "0 mg", "0 g", ""], ["Beverages", "Choose Sometimes", "0 g", "1–140 mg", "1–11 g", ""], ["Beverages", "Choose Rarely", "≥1 g", "≥141 mg", "≥12 g", ""],
        ["Mixed Dishes", "Choose Often", "≤3 g", "≤480 mg", "≤6 g", ""], ["Mixed Dishes", "Choose Sometimes", ">3–6 g", "481–599 mg", "7–11 g", ""], ["Mixed Dishes", "Choose Rarely", "≥6.5 g", "≥600 mg", "≥12 g", ""],
        ["Processed and Packaged Snacks", "Choose Sometimes", "≤2 g", "≤140 mg", "≤6 g", "Whole-grain requirement must be reviewed"], ["Processed and Packaged Snacks", "Choose Rarely", ">2 g", "≥141 mg", "≥7 g", ""], ["Desserts", "Choose Rarely", "All", "All", "All", "Milk-based desserts including pudding, ice cream, frozen yogurt and milkshakes are desserts; flavored liquid milk remains Dairy"], ["Infant", "Unranked", "", "", "", "Baby, infant and toddler foods or formulas"], ["Protein Powder", "Unranked", "", "", "", "Forced exclusion"], ["Condiments and Cooking Staples", "Not Ranked", "", "", "", ""], ["Miscellaneous Products", "Not Ranked", "", "", "", ""]]
    return pd.DataFrame(rows, columns=["Food Category", "Tier", "Saturated Fat", "Sodium", "Added or Total Sugar", "Special Rule"])


def backend_keyword_table() -> pd.DataFrame:
    category_rows = [{"Rule Type": "Category", "Food Category": category, "Keyword": keyword, "Active": "Yes", "Additional Notes": ""} for category, keywords in CATEGORY_KEYWORDS.items() for keyword in keywords]
    infant_rows = [{"Rule Type": "Forced Unranked", "Food Category": "Infant", "Keyword": keyword, "Active": "Yes", "Additional Notes": "Any baby, infant or toddler item is Unranked"} for keyword in INFANT_TERMS]
    powder_rows = [{"Rule Type": "Forced Unranked", "Food Category": "Miscellaneous Products", "Keyword": keyword, "Active": "Yes", "Additional Notes": "Protein powder exclusion"} for keyword in PROTEIN_POWDER_TERMS]
    red_rows = [{"Rule Type": "Forced Choose Rarely", "Food Category": "Protein", "Keyword": keyword, "Active": "Yes", "Additional Notes": "Explicit HER protein red rule"} for keyword in FORCED_RED_PROTEIN_TERMS]
    return pd.DataFrame(category_rows + infant_rows + powder_rows + red_rows)

def backend_dictionary_table() -> pd.DataFrame:
    rows = [["Data Source", "Workbook or supplemental fruit table"], ["Fresh Produce", "Yes only when strict whole-word text rules identify plain fresh/raw produce"], ["Seafood Sugar Assumption", "Supplemental seafood uses 0.5 g sugar and 0 g dietary fiber"], ["Forced Protein Red", "Yes for refried beans, deli meat, sausage, bacon or breaded chicken"], ["Forced Unranked", "Yes for baby, infant or toddler food/formula and protein powder"], ["Automatic Category", "Description-based category suggestion with flavored-liquid-milk override to Dairy"], ["Automatic Overall Rank", "Final automatic result after overrides and category rules"], ["Manual Category", "Editable corrected category"], ["Manual Final Rank", "Editable reviewed rank"], ["Review Status", "Not Reviewed, Reviewed or Needs Research"], ["Reviewer Notes", "Free-text corrections and additions"]]
    return pd.DataFrame(rows, columns=["Field", "Description"])


@st.cache_data(show_spinner=False)
def build_backend_summary(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    for _, row in df.iterrows():
        description = row[FOOD_DESCRIPTION]; category = suggest_her_category(description); unranked_reason = forced_unranked_reason(description); red_reason = forced_red_protein_reason(description); fresh_produce = bool(row.get("_fresh_produce", False)); seafood_assumption = bool(row.get("_seafood_assumption", False)); data_source = clean_text(row.get("_data_source", ""))
        saturated_fat, saturated_source = numeric_from_columns(row, ["Total Saturated", "Saturated Fat", "Saturated fat"]); sodium, sodium_source = numeric_from_columns(row, ["Sodium"]); sugar, sugar_source = numeric_from_columns(row, ["Added Sugars", "Added Sugar", "Total Added Sugars", "Total Sugars"])
        if seafood_assumption and pd.isna(sugar): sugar, sugar_source = 0.5, "Assumed seafood sugar"
        whole_grain_first = whole_grain_is_first_in_description(description)
        classification = classify_food_description(description, category, saturated_fat, sodium, sugar, whole_grain_first, False, fresh_produce)
        warnings_list = []
        if unranked_reason: warnings_list.append(unranked_reason)
        if red_reason: warnings_list.append(red_reason)
        if seafood_assumption: warnings_list.append("Seafood sugar assumed at 0.5 g and dietary fiber at 0 g")
        if not unranked_reason and not red_reason and not fresh_produce and category in {"Grains", "Processed and Packaged Snacks"} and not whole_grain_first: warnings_list.append("Explicit whole-grain wording was not identified from the food name")
        if not unranked_reason and not red_reason and not fresh_produce and category == "Fruits and Vegetables": warnings_list.append("Juice or dried-fruit status not reviewed")
        if sugar_source == "Total Sugars" and not fresh_produce and not seafood_assumption: warnings_list.append("Total Sugars used because Added Sugars was unavailable")
        if category == "Miscellaneous Products" and not unranked_reason: warnings_list.append("Automatic category was not identified")
        if pd.isna(saturated_fat) and not fresh_produce and not unranked_reason and not red_reason: warnings_list.append("Saturated fat unavailable")
        if pd.isna(sodium) and not fresh_produce and not unranked_reason and not red_reason: warnings_list.append("Sodium unavailable")
        if pd.isna(sugar) and not fresh_produce and not unranked_reason and not red_reason: warnings_list.append("Sugar unavailable")
        records.append({"Food Code": row[FOOD_CODE], "Food Description": description, "Number of Servings": row[NUMBER_SERVINGS], "Measure Description": row[MEASURE_DESCRIPTION], "Serving Weight g": row[SERVING_WEIGHT], "Data Source": data_source, "Fresh Produce": "Yes" if fresh_produce else "No", "Seafood Sugar Assumption": "Yes" if seafood_assumption else "No", "Forced Unranked": "Yes" if unranked_reason else "No", "Forced Unranked Reason": unranked_reason, "Forced Protein Red": "Yes" if red_reason else "No", "Forced Protein Red Reason": red_reason, "Automatic Category": category, "Manual Category": "", "Whole Grain First Ingredient": "Yes" if whole_grain_first else "No", "100% Juice or Plain Dried Fruit": "", "Saturated Fat g": saturated_fat, "Saturated Fat Source": saturated_source, "Sodium mg": sodium, "Sodium Source": sodium_source, "Sugar g": sugar, "Sugar Source": sugar_source, "Automatic Saturated Fat Rank": classification["Saturated Fat"], "Automatic Sodium Rank": classification["Sodium"], "Automatic Sugar Rank": classification["Sugar"], "Automatic Overall Rank": classification["Overall"], "Automatic Special Rule": classification["Rule"], "Classification Warning": "; ".join(warnings_list), "Manual Final Rank": "", "Review Status": "Not Reviewed", "Reviewer Notes": ""})
    return pd.DataFrame(records)

def build_manual_review_sheet(summary_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["Food Code", "Food Description", "Measure Description", "Serving Weight g", "Data Source", "Fresh Produce", "Seafood Sugar Assumption", "Forced Unranked", "Forced Unranked Reason", "Forced Protein Red", "Forced Protein Red Reason", "Automatic Category", "Manual Category", "Whole Grain First Ingredient", "100% Juice or Plain Dried Fruit", "Saturated Fat g", "Sodium mg", "Sugar g", "Sugar Source", "Automatic Overall Rank", "Classification Warning", "Manual Final Rank", "Review Status", "Reviewer Notes"]
    review_df = summary_df[columns].copy(); needs_review = review_df["Classification Warning"].ne("") | review_df["Automatic Category"].eq("Miscellaneous Products") | review_df["Forced Unranked"].eq("Yes") | review_df["Forced Protein Red"].eq("Yes")
    return pd.concat([review_df[needs_review], review_df[~needs_review]], ignore_index=True)


def format_backend_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, freeze_columns: int = 0) -> None:
    workbook, worksheet = writer.book, writer.sheets[sheet_name]; header = workbook.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "border": 1, "text_wrap": True, "valign": "top"}); formats = {"Choose Often": workbook.add_format({"bg_color": "#C6E0B4"}), "Choose Sometimes": workbook.add_format({"bg_color": "#FFF2CC"}), "Choose Rarely": workbook.add_format({"bg_color": "#F4B084"}), "Unranked": workbook.add_format({"bg_color": "#A6A6A6"}), "Not Ranked": workbook.add_format({"bg_color": "#D9E1F2"})}
    for i, column in enumerate(df.columns): worksheet.write(0, i, column, header); lengths = df[column].fillna("").astype(str).head(1000).map(len); worksheet.set_column(i, i, min(max(max(len(column) + 2, int(lengths.max()) + 2 if len(lengths) else 10), 11), 45))
    worksheet.freeze_panes(1, freeze_columns); worksheet.autofilter(0, 0, len(df), max(len(df.columns) - 1, 0)); worksheet.set_row(0, 34)
    if df.empty: return
    for column_name in df.columns:
        if "Rank" not in column_name: continue
        i = df.columns.get_loc(column_name); cell_range = f"{excel_column_name(i)}2:{excel_column_name(i)}{len(df) + 1}"
        for value, fmt in formats.items(): worksheet.conditional_format(cell_range, {"type": "text", "criteria": "containing", "value": value, "format": fmt})


@st.cache_data(show_spinner=False)
def build_backend_workbook(df: pd.DataFrame) -> bytes:
    summary_df = build_backend_summary(df); review_df = build_manual_review_sheet(summary_df); rules_df = backend_rule_table(); keywords_df = backend_keyword_table(); dictionary_df = backend_dictionary_table()
    overview_df = pd.DataFrame({"Item": ["Generated At", "Food-Serving Records", "Unique Foods", "Choose Often", "Choose Sometimes", "Choose Rarely", "Unranked", "Not Ranked", "Fresh Produce", "Seafood Assumptions", "Forced Protein Red", "Forced Unranked", "Records With Review Warnings"], "Value": [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(summary_df), summary_df[[FOOD_CODE, FOOD_DESCRIPTION]].drop_duplicates().shape[0], int(summary_df["Automatic Overall Rank"].eq("Choose Often").sum()), int(summary_df["Automatic Overall Rank"].eq("Choose Sometimes").sum()), int(summary_df["Automatic Overall Rank"].eq("Choose Rarely").sum()), int(summary_df["Automatic Overall Rank"].eq("Unranked").sum()), int(summary_df["Automatic Overall Rank"].eq("Not Ranked").sum()), int(summary_df["Fresh Produce"].eq("Yes").sum()), int(summary_df["Seafood Sugar Assumption"].eq("Yes").sum()), int(summary_df["Forced Protein Red"].eq("Yes").sum()), int(summary_df["Forced Unranked"].eq("Yes").sum()), int(summary_df["Classification Warning"].ne("").sum())]})
    category_summary_df = summary_df.groupby(["Automatic Category", "Automatic Overall Rank"], dropna=False).size().reset_index(name="Serving Record Count").sort_values(["Automatic Category", "Automatic Overall Rank"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet, frame in [("Overview", overview_df), ("Category Summary", category_summary_df), ("Classification Summary", summary_df), ("Manual Review", review_df), ("Category Rules", rules_df), ("Category Keywords", keywords_df), ("Data Dictionary", dictionary_df)]: frame.to_excel(writer, sheet_name=sheet, index=False); format_backend_sheet(writer, sheet, frame, 2 if sheet in {"Classification Summary", "Manual Review", "Category Rules"} else 1 if sheet in {"Category Keywords", "Data Dictionary"} else 0)
        manual_sheet = writer.sheets["Manual Review"]; manual_columns = {column: index for index, column in enumerate(review_df.columns)}
        if len(review_df):
            first_row, last_row = 1, len(review_df)
            for column, source in [("Manual Category", HER_CATEGORIES), ("Whole Grain First Ingredient", ["Yes", "No", "Not Applicable", "Unknown"]), ("100% Juice or Plain Dried Fruit", ["Yes", "No", "Not Applicable", "Unknown"]), ("Manual Final Rank", ["Choose Often", "Choose Sometimes", "Choose Rarely", "Unranked", "Not Ranked"]), ("Review Status", ["Not Reviewed", "Reviewed", "Needs Research"])]: manual_sheet.data_validation(first_row, manual_columns[column], last_row, manual_columns[column], {"validate": "list", "source": source})
    output.seek(0); return output.getvalue()


# ============================================================
# INCOMING LOAD / COLOR COMPOSITION
# ============================================================

# Common broad incoming food names that have an established HER treatment.
# These are intentionally applied before USDA ambiguity checks in the load analyzer.
# They do not claim an exact USDA record unless one is explicitly given.
MANUAL_LOAD_CLASSIFICATIONS = {
    "salmon": {
        "category": "Protein",
        "overall": "Choose Often",
        "basis": "Broad salmon entry treated as plain salmon. HER protein classification: Choose Often.",
    },
    "canned tuna": {
        "category": "Protein",
        "overall": "Choose Sometimes",
        "basis": "Canned tuna is classified as Choose Sometimes in the HER food-category reference.",
    },
    "tuna canned": {
        "category": "Protein",
        "overall": "Choose Sometimes",
        "basis": "Canned tuna is classified as Choose Sometimes in the HER food-category reference.",
    },
    "tuna": {
        "category": "Protein",
        "overall": "Choose Sometimes",
        "basis": "Generic tuna entry uses the canned-tuna HER food-category classification unless a more specific tuna product is supplied.",
    },
    "brown rice": {
        "category": "Grains",
        "overall": "Choose Often",
        "basis": "Brown rice is treated as a whole grain and classified as Choose Often for load composition.",
    },
    "ice cream": {
        "category": "Desserts",
        "overall": "Choose Rarely",
        "basis": "Ice cream is a Dessert and is automatically Choose Rarely under the HER guideline.",
    },
    "bacon": {
        "category": "Protein",
        "overall": "Choose Rarely",
        "basis": "Bacon is explicitly classified as Choose Rarely under the HER protein guideline.",
    },
}

def manual_load_classification(query: str) -> dict[str, str] | None:
    q = normalize_description_for_matching(query)
    if not q:
        return None
    if q in MANUAL_LOAD_CLASSIFICATIONS:
        return MANUAL_LOAD_CLASSIFICATIONS[q]
    return None


def classify_row_for_load(row: pd.Series) -> tuple[str, str]:
    """Return (HER category, overall HER classification) for one stored food row."""
    description = row[FOOD_DESCRIPTION]
    category = suggest_her_category(description)
    fresh_produce = bool(row.get("_fresh_produce", False))
    whole_grain_first = bool(row.get("_whole_grain_detected", False))
    saturated_fat, _ = numeric_from_columns(row, ["Total Saturated", "Saturated Fat", "Saturated fat"])
    sodium, _ = numeric_from_columns(row, ["Sodium"])
    sugar, _ = numeric_from_columns(row, ["Added Sugars", "Added Sugar", "Total Added Sugars", "Total Sugars"])
    if bool(row.get("_seafood_assumption", False)) and pd.isna(sugar):
        sugar = 0.5
    result = classify_food_description(
        description, category, saturated_fat, sodium, sugar,
        whole_grain_first, False, fresh_produce
    )
    return category, result["Overall"]


def direct_her_rule_for_load(query: str) -> dict[str, str] | None:
    """Resolve broad incoming names when HER rules are decisive without a specific USDA row.

    This is intentionally conservative. It only returns a result when the description itself
    is enough to determine the HER outcome. Otherwise matching continues against USDA rows.
    """
    q = normalize_description_for_matching(query)
    if not q:
        return None

    # Forced exclusions / forced-red protein rules are decisive at the food-name level.
    unranked_reason = forced_unranked_reason(q)
    if unranked_reason:
        return {
            "category": forced_unranked_category(q),
            "overall": "Unranked",
            "basis": unranked_reason,
        }

    red_reason = forced_red_protein_reason(q)
    if red_reason:
        return {
            "category": "Protein",
            "overall": "Choose Rarely",
            "basis": red_reason,
        }

    # Desserts are automatically Choose Rarely under the HER logic, so a specific
    # ice-cream flavor/USDA row is not required for load-composition purposes.
    category = suggest_her_category(q)
    if category == "Desserts":
        return {
            "category": "Desserts",
            "overall": "Choose Rarely",
            "basis": "Desserts are automatically classified as Choose Rarely under the HER guideline.",
        }

    # Plain fresh produce can be resolved directly from a broad incoming description.
    if is_fresh_produce_description(q):
        return {
            "category": "Fruits and Vegetables",
            "overall": "Choose Often",
            "basis": "Plain fresh produce is classified as Choose Often.",
        }

    return None


def candidate_rows_for_query(food_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Return plausible USDA candidates using canonical names and positive aliases only."""
    q = normalize_description_for_matching(query)
    if not q:
        return food_df.iloc[0:0].copy()

    unique = food_df.drop_duplicates([FOOD_CODE, FOOD_DESCRIPTION]).copy()
    unique["_canon_norm"] = unique["_display_name"].map(normalize_description_for_matching)

    # Exact canonical name first.
    exact = unique[unique["_canon_norm"].eq(q)]
    if not exact.empty:
        return exact

    # Exact positive alias next.
    alias_exact = unique[
        unique["_aliases"].fillna("").apply(
            lambda x: any(normalize_description_for_matching(a) == q for a in str(x).split("; ") if a)
        )
    ]
    if not alias_exact.empty:
        return alias_exact

    # Then canonical/search contains. _search_text excludes negative/exclusion aliases.
    contains = unique[unique["_search_text"].str.contains(q, na=False, regex=False)]
    return contains


def classify_candidate_consensus(candidates: pd.DataFrame) -> tuple[str, str] | None:
    """Return a category/rank only if every plausible candidate agrees.

    This lets broad inputs such as 'bacon' or 'ice cream' resolve safely when all matching
    USDA rows lead to the same HER result, without pretending one specific USDA row was chosen.
    """
    if candidates.empty:
        return None
    outcomes = []
    for _, row in candidates.iterrows():
        category, overall = classify_row_for_load(row)
        outcomes.append((category, overall))
    unique_outcomes = sorted(set(outcomes))
    return unique_outcomes[0] if len(unique_outcomes) == 1 else None


def match_one_food(food_df: pd.DataFrame, query: str) -> tuple[pd.Series | None, str]:
    """Match one incoming item to a specific USDA row when that match is unambiguous."""
    q = normalize_description_for_matching(query)
    if not q:
        return None, "Blank food name"

    candidates = candidate_rows_for_query(food_df, query)
    if len(candidates) == 1:
        row = candidates.iloc[0]
        canon = normalize_description_for_matching(row["_display_name"])
        if canon == q:
            return row, "Exact food-name match"
        aliases = [normalize_description_for_matching(a) for a in str(row.get("_aliases", "")).split("; ") if a]
        if q in aliases:
            return row, "Exact alias match"
        return row, "Unique search match"

    if len(candidates) > 1:
        starts = candidates[candidates["_display_name"].map(normalize_description_for_matching).str.startswith(q, na=False)]
        if len(starts) == 1:
            return starts.iloc[0], "Unique name-prefix match"
        return None, f"{len(candidates)} possible USDA matches"

    return None, "No USDA match found"


# ============================================================
# FBCENC CLASSIFICATION CROSSWALK
# ============================================================

FBCENC_REQUIRED_COLUMNS = ["Item No", "Description", "Item Category", "HER Category New"]
FBCENC_HER_MAP = {
    "green": "Choose Often",
    "yellow": "Choose Sometimes",
    "red": "Choose Rarely",
    "unranked": "Unranked",
    "not ranked": "Not Ranked",
    "assorted": "Assorted",
    "non food": "Non Food",
    "non-food": "Non Food",
}


def normalize_fbcenc_rank(value: object) -> str:
    text = clean_text(value).lower()
    return FBCENC_HER_MAP.get(text, clean_text(value) or "Needs Review")


@st.cache_data(show_spinner=False)
def load_fbcenc_crosswalk(source) -> pd.DataFrame:
    """Load the FBCENC manually reviewed item-level HER crosswalk.

    The file is used as the first matching layer for incoming FBCENC loads.
    It does not require nutrient reclassification when an item number or exact
    FBCENC description already has a reviewed HER category.
    """
    frame = pd.read_excel(source)
    frame.columns = [clean_text(c) for c in frame.columns]
    missing = [c for c in FBCENC_REQUIRED_COLUMNS if c not in frame.columns]
    if missing:
        raise ValueError("FBCENC classification file is missing: " + ", ".join(missing))

    frame = frame[FBCENC_REQUIRED_COLUMNS].copy()
    frame["Item No"] = frame["Item No"].map(clean_text)
    frame["Description"] = frame["Description"].map(clean_text)
    frame["Item Category"] = frame["Item Category"].map(clean_text)
    frame["HER Category New"] = frame["HER Category New"].map(clean_text)
    frame["_item_norm"] = frame["Item No"].str.upper()
    frame["_description_norm"] = frame["Description"].map(normalize_description_for_matching)
    frame["_her_display"] = frame["HER Category New"].map(normalize_fbcenc_rank)
    frame = frame[(frame["Item No"].ne("")) | (frame["Description"].ne(""))].copy()
    return frame.reset_index(drop=True)


def match_fbcenc_crosswalk(crosswalk: pd.DataFrame | None, item_no: str, food: str) -> tuple[pd.Series | None, str]:
    if crosswalk is None or crosswalk.empty:
        return None, "No FBCENC crosswalk loaded"

    item_key = clean_text(item_no).upper()
    if item_key:
        hits = crosswalk[crosswalk["_item_norm"].eq(item_key)]
        if len(hits) == 1:
            return hits.iloc[0], "Exact FBCENC item-number match"
        if len(hits) > 1:
            # Item numbers can repeat across historical descriptions. If every reviewed
            # row agrees, use the reviewed result without pretending one row was unique.
            outcomes = hits[["Item Category", "_her_display"]].drop_duplicates()
            if len(outcomes) == 1:
                return hits.iloc[0], f"FBCENC item number matched {len(hits)} agreeing catalog rows"

    food_key = normalize_description_for_matching(food)
    if food_key:
        hits = crosswalk[crosswalk["_description_norm"].eq(food_key)]
        if len(hits) == 1:
            return hits.iloc[0], "Exact FBCENC description match"
        if len(hits) > 1:
            outcomes = hits[["Item Category", "_her_display"]].drop_duplicates()
            if len(outcomes) == 1:
                return hits.iloc[0], f"FBCENC description matched {len(hits)} agreeing catalog rows"

    return None, "No exact FBCENC catalog match"


def analyze_load(food_df: pd.DataFrame, incoming: pd.DataFrame, fbcenc_crosswalk: pd.DataFrame | None = None) -> pd.DataFrame:
    """Classify an incoming food list with a rule-aware, conservative hierarchy.

    Hierarchy:
      0) reviewed FBCENC crosswalk by Item No or exact description;
      1) common broad HER reference names (e.g., salmon, canned tuna, brown rice);
      2) exact/unique USDA row -> use its nutrients and HER logic;
      3) decisive HER name-level rule (e.g., bacon, ice cream, fresh produce);
      4) multiple USDA candidates that all agree on category + HER color -> use consensus;
      5) otherwise -> Needs Review.
    """
    records = []
    for _, r in incoming.iterrows():
        item_no = clean_text(r.get("Item No", ""))
        food = clean_text(r.get("Food", ""))
        pounds = pd.to_numeric(r.get("Pounds", np.nan), errors="coerce")
        if not food or pd.isna(pounds) or pounds <= 0:
            continue

        # Prefer the food bank's reviewed catalog classification when available.
        local_match, local_method = match_fbcenc_crosswalk(fbcenc_crosswalk, item_no, food)
        if local_match is not None:
            records.append({
                "Item No": item_no,
                "Input Food": food,
                "Pounds": float(pounds),
                "Matched Food": clean_text(local_match.get("Description", "")),
                "HER Category": clean_text(local_match.get("Item Category", "")) or "Needs Review",
                "HER Classification": clean_text(local_match.get("_her_display", "")) or "Needs Review",
                "Match Status": "FBCENC reviewed",
                "Match Note": local_method + "; uses HER Category New from the FBCENC classification workbook.",
            })
            continue

        # Common broad names with an established HER treatment are resolved before
        # USDA ambiguity checks. For example, "Salmon" means plain salmon here, not
        # eighteen preparation-specific USDA records.
        manual = manual_load_classification(food)
        if manual is not None:
            records.append({
                "Item No": item_no,
                "Input Food": food,
                "Pounds": float(pounds),
                "Matched Food": clean_text(food),
                "HER Category": manual["category"],
                "HER Classification": manual["overall"],
                "Match Status": "HER reference",
                "Match Note": manual["basis"],
            })
            continue

        # Next use a specific USDA row when the input identifies one unambiguously.
        matched, method = match_one_food(food_df, food)
        if matched is not None:
            category, overall = classify_row_for_load(matched)
            records.append({
                "Item No": item_no,
                "Input Food": food,
                "Pounds": float(pounds),
                "Matched Food": matched["_display_name"],
                "HER Category": category,
                "HER Classification": overall,
                "Match Status": "Matched",
                "Match Note": method,
            })
            continue

        # Next apply only HER rules that are decisive from the incoming name itself.
        direct = direct_her_rule_for_load(food)
        if direct is not None:
            records.append({
                "Item No": item_no,
                "Input Food": food,
                "Pounds": float(pounds),
                "Matched Food": "",
                "HER Category": direct["category"],
                "HER Classification": direct["overall"],
                "Match Status": "HER rule",
                "Match Note": direct["basis"],
            })
            continue

        # Finally, if several plausible USDA rows exist but every one gives the same HER
        # category and color, use that consensus without claiming a specific product match.
        candidates = candidate_rows_for_query(food_df, food)
        consensus = classify_candidate_consensus(candidates)
        if consensus is not None:
            category, overall = consensus
            records.append({
                "Item No": item_no,
                "Input Food": food,
                "Pounds": float(pounds),
                "Matched Food": "",
                "HER Category": category,
                "HER Classification": overall,
                "Match Status": "USDA consensus",
                "Match Note": f"{len(candidates)} plausible USDA records all classify as {overall} in {category}.",
            })
            continue

        records.append({
            "Item No": item_no,
            "Input Food": food,
            "Pounds": float(pounds),
            "Matched Food": "",
            "HER Category": "Needs Review",
            "HER Classification": "Needs Review",
            "Match Status": "Needs Review",
            "Match Note": method,
        })

    return pd.DataFrame(records)


def render_load_composition(food_df: pd.DataFrame) -> None:
    st.title("📦 Incoming Food Load Composition")
    st.caption("Analyze pounds by HER color and food category. FBCENC reviewed classifications are used first when available; common HER reference foods are resolved next, followed by USDA/HER matching.")

    # Load the reviewed FBCENC item-level crosswalk. Put FBCENC_HER_Classifications.xlsx
    # beside this script for automatic loading, or upload it here.
    fbcenc_crosswalk = None
    local_crosswalk = Path(__file__).resolve().with_name(DEFAULT_FBCENC_CLASSIFICATION_FILENAME)
    with st.expander("FBCENC reviewed classification crosswalk", expanded=True):
        st.caption("Preferred matching source for FBCENC items: Item No → exact FBCENC description → USDA/HER fallback.")
        crosswalk_upload = st.file_uploader(
            "FBCENC classification workbook",
            type=["xlsx", "xls"],
            key="fbcenc_crosswalk_upload",
            help="Expected columns: Item No, Description, Item Category, HER Category New",
        )
        try:
            if crosswalk_upload is not None:
                fbcenc_crosswalk = load_fbcenc_crosswalk(crosswalk_upload)
                st.success(f"Loaded {len(fbcenc_crosswalk):,} reviewed FBCENC catalog rows from the uploaded workbook.")
            elif local_crosswalk.exists():
                fbcenc_crosswalk = load_fbcenc_crosswalk(str(local_crosswalk))
                st.success(f"Loaded {len(fbcenc_crosswalk):,} reviewed FBCENC catalog rows from {local_crosswalk.name}.")
            else:
                st.info(f"Optional: place {DEFAULT_FBCENC_CLASSIFICATION_FILENAME} beside the script or upload it above.")
        except Exception as exc:
            st.error(f"Could not load the FBCENC classification workbook: {exc}")

    entry_tab, upload_tab, catalog_tab = st.tabs(["Enter / paste a load", "Upload load CSV or Excel", "FBCENC catalog preview"])
    incoming = None

    with entry_tab:
        # Deliberately varied demo load to exercise exact matches, aliases,
        # decisive HER rules, category consensus, and Needs Review behavior.
        # Demo uses several real descriptions/item numbers from the supplied FBCENC
        # classification workbook plus generic foods to test the USDA/HER fallback.
        # Pounds are simulated only for demonstrating load composition.
        toy = pd.DataFrame({
            "Item No": [
                "CV1043", "PTP024", "CD2013", "EN1618", "BV1009", "CD1812",
                "", "", "", "", "", "", "", "",
            ],
            "Food": [
                "1% LF MILK",
                "CHEWY GRANOLA BARS",
                "PEANUT BUTTER",
                "PROGRESSO POTATO, BROCCOLI AND CHEESE CHOWDER",
                "100% FRUIT JUICE",
                "ADVANTAGE CANOLA OIL",
                "Bacon",
                "Brown rice",
                "Ice cream",
                "Apple",
                "leche fresca",
                "Salmon",
                "Canned tuna",
                "Mystery donation box",
            ],
            "Pounds": [
                520.0, 180.0, 260.0, 340.0, 410.0, 150.0,
                220.0, 360.0, 175.0, 650.0, 300.0, 240.0, 195.0, 90.0,
            ],
        })
        incoming_editor = st.data_editor(
            toy,
            num_rows="dynamic",
            hide_index=True,
            width="stretch",
            column_config={
                "Item No": st.column_config.TextColumn("Item No", help="Optional FBCENC item number; preferred when available"),
                "Food": st.column_config.TextColumn("Food", help="FBCENC description, USDA food name, or searchable synonym"),
                "Pounds": st.column_config.NumberColumn("Pounds", min_value=0.0, step=1.0, format="%.1f"),
            },
            key="load_editor",
        )
        if st.button("Analyze entered load", type="primary", width="stretch", key="analyze_entered"):
            incoming = incoming_editor.copy()

    with upload_tab:
        uploaded = st.file_uploader("Upload a file", type=["csv", "xlsx", "xls"], key="load_upload")
        st.caption("Required: Food/Description and Pounds. Item No is optional but strongly preferred for FBCENC data.")
        if uploaded is not None:
            try:
                raw = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
                rename = {}
                for c in raw.columns:
                    lc = clean_text(c).lower()
                    if lc in {"item no", "item number", "item_no", "item #", "item#", "sku", "product code"}:
                        rename[c] = "Item No"
                    elif lc in {"food", "food name", "food description", "description", "item", "product description"}:
                        rename[c] = "Food"
                    elif lc in {"pounds", "pound", "lbs", "lb", "weight", "weight lb", "weight lbs", "ext gross weight", "ext_gross_weight"}:
                        rename[c] = "Pounds"
                raw = raw.rename(columns=rename)
                if {"Food", "Pounds"}.issubset(raw.columns):
                    if "Item No" not in raw.columns:
                        raw["Item No"] = ""
                    st.dataframe(raw[["Item No", "Food", "Pounds"]].head(50), hide_index=True, width="stretch")
                    if st.button("Analyze uploaded load", type="primary", width="stretch", key="analyze_uploaded"):
                        incoming = raw[["Item No", "Food", "Pounds"]].copy()
                else:
                    st.error("Could not identify both Food and Pounds columns.")
            except Exception as exc:
                st.error(f"Could not read the uploaded file: {exc}")

    with catalog_tab:
        if fbcenc_crosswalk is None or fbcenc_crosswalk.empty:
            st.info("Load the FBCENC classification workbook above to preview the reviewed catalog.")
        else:
            catalog_view = fbcenc_crosswalk[["Item No", "Description", "Item Category", "HER Category New"]].copy()
            search_catalog = st.text_input("Search FBCENC catalog", placeholder="milk, granola, item number...", key="catalog_search")
            if search_catalog.strip():
                q = normalize_description_for_matching(search_catalog)
                item_q = clean_text(search_catalog).upper()
                mask = catalog_view["Item No"].astype(str).str.upper().str.contains(item_q, na=False, regex=False) | catalog_view["Description"].map(normalize_description_for_matching).str.contains(q, na=False, regex=False)
                catalog_view = catalog_view[mask]
            st.caption(f"{len(catalog_view):,} catalog rows shown.")
            st.dataframe(catalog_view.head(500), hide_index=True, width="stretch", height=420)

    if incoming is not None:
        if "Item No" not in incoming.columns:
            incoming["Item No"] = ""
        st.session_state["load_analysis"] = analyze_load(food_df, incoming, fbcenc_crosswalk)

    analysis = st.session_state.get("load_analysis")
    if analysis is None or analysis.empty:
        return

    total_lb = float(analysis["Pounds"].sum())
    reviewed_lb = float(analysis.loc[analysis["Match Status"].eq("FBCENC reviewed"), "Pounds"].sum())
    comp = analysis.groupby("HER Classification", dropna=False)["Pounds"].sum().reset_index()
    comp["Percent"] = np.where(total_lb > 0, comp["Pounds"] / total_lb * 100, 0.0)
    order = {"Choose Often": 0, "Choose Sometimes": 1, "Choose Rarely": 2, "Assorted": 3, "Unranked": 4, "Not Ranked": 5, "Non Food": 6, "Needs Review": 7}
    comp["_order"] = comp["HER Classification"].map(order).fillna(99)
    comp = comp.sort_values("_order").drop(columns="_order")

    st.markdown("### Color composition")
    m1, m2 = st.columns(2)
    m1.metric("Total incoming load", f"{total_lb:,.1f} lb")
    m2.metric("Matched to reviewed FBCENC catalog", f"{reviewed_lb:,.1f} lb", help="Pounds classified directly from HER Category New in the FBCENC workbook")

    # Semantic cards. Avoid st.metric delta because Streamlit renders every positive
    # percentage in green, which is misleading for yellow/red/review classes.
    cols = st.columns(min(len(comp), 6))
    for col, (_, r) in zip(cols, comp.iterrows()):
        rank = str(r["HER Classification"])
        icon = HER_ICONS.get(rank, "⚪")
        bg = HER_COLORS.get(rank, HER_COLORS["Needs Review"])
        fg = HER_TEXT_COLORS.get(rank, "#FFFFFF")
        with col:
            st.markdown(
                f"""
                <div style="background:{bg};color:{fg};border-radius:14px;padding:18px 18px 16px;min-height:142px;
                            box-shadow:0 1px 2px rgba(0,0,0,.12);margin-bottom:8px">
                    <div style="font-size:1rem;font-weight:650;margin-bottom:12px">{icon} {rank}</div>
                    <div style="font-size:2rem;font-weight:750;line-height:1.05">{r['Pounds']:,.1f} lb</div>
                    <div style="font-size:1.05rem;font-weight:650;margin-top:12px">{r['Percent']:.1f}% of load</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # Interactive semantic-color chart. Clicking a bar returns the selected HER class.
    color_domain = ["Choose Often", "Choose Sometimes", "Choose Rarely", "Assorted", "Unranked", "Not Ranked", "Non Food", "Needs Review"]
    color_range = [HER_COLORS[x] for x in color_domain]
    point = alt.selection_point(fields=["HER Classification"], on="click", clear="dblclick", name="her_class")
    color_chart = (
        alt.Chart(comp)
        .mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5)
        .encode(
            x=alt.X("HER Classification:N", sort=color_domain, title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Pounds:Q", title="Pounds"),
            color=alt.Color(
                "HER Classification:N",
                scale=alt.Scale(domain=color_domain, range=color_range),
                legend=None,
            ),
            opacity=alt.condition(point, alt.value(1.0), alt.value(0.82)),
            tooltip=[
                alt.Tooltip("HER Classification:N", title="Classification"),
                alt.Tooltip("Pounds:Q", title="Pounds", format=",.1f"),
                alt.Tooltip("Percent:Q", title="% of load", format=".1f"),
            ],
        )
        .add_params(point)
        .properties(height=360)
    )
    # Render the semantic-color chart in a Streamlit-version-safe way.
    # Some Streamlit versions do not accept width="stretch" and/or do not
    # support chart selection callbacks. The drill-down buttons below work on
    # all supported versions.
    try:
        chart_event = st.altair_chart(
            color_chart,
            use_container_width=True,
            on_select="rerun",
            selection_mode="her_class",
            key="her_color_composition_chart",
        )
    except TypeError:
        chart_event = None
        st.altair_chart(color_chart, use_container_width=True)

    st.caption("Use the buttons below to open the foods and pounds behind each classification. On newer Streamlit versions, clicking a bar also opens the same drill-down.")

    selected_rank = None
    if chart_event is not None:
        try:
            selected_points = chart_event.selection.her_class
            if selected_points:
                selected_rank = selected_points[0].get("HER Classification")
        except Exception:
            selected_rank = None

    @st.dialog("Foods in selected HER classification", width="large")
    def show_her_items(rank: str) -> None:
        subset = analysis[analysis["HER Classification"].eq(rank)].copy()
        pounds = float(subset["Pounds"].sum())
        pct = (pounds / total_lb * 100) if total_lb > 0 else 0.0
        icon = HER_ICONS.get(rank, "⚪")
        st.markdown(f"### {icon} {rank}")
        st.write(f"**{pounds:,.1f} lb** · **{pct:.1f}%** of the incoming load · **{len(subset)} item(s)**")
        preferred = [
            "Input Food", "Pounds", "Matched Food", "HER Category",
            "HER Classification", "Match Status", "Match Note"
        ]
        visible = [c for c in preferred if c in subset.columns]
        if not visible:
            visible = subset.columns.tolist()
        st.dataframe(subset[visible], hide_index=True, width="stretch", height=min(500, 80 + 35 * len(subset)))

    if selected_rank:
        show_her_items(selected_rank)

    # Explicit buttons make the drill-down discoverable and also work if chart
    # selection is unavailable in an older Streamlit installation.
    button_cols = st.columns(min(len(comp), 6))
    for col, (_, r) in zip(button_cols, comp.iterrows()):
        rank = str(r["HER Classification"])
        with col:
            if st.button(f"View {rank}", key=f"view_rank_{rank}", width="stretch"):
                show_her_items(rank)

    display_comp = comp.rename(columns={"Percent": "% of Load"}).copy()
    st.dataframe(display_comp.style.format({"Pounds": "{:,.1f}", "% of Load": "{:.1f}%"}), width="stretch")

    needs_review = analysis[analysis["HER Classification"].eq("Needs Review")]
    if not needs_review.empty:
        st.warning(
            f"{len(needs_review)} item(s), totaling {needs_review['Pounds'].sum():,.1f} lb, need manual matching. "
            "They are kept out of the green/yellow/red totals rather than being guessed."
        )

    # Food-category composition by total incoming weight. Each item stays in one HER category;
    # mixed dishes are NOT split into estimated ingredient weights.
    st.markdown("### Food composition by weight")
    st.caption("Each incoming item contributes its full weight to one HER food category. Mixed dishes are kept as Mixed Dishes rather than being divided into ingredients.")

    category_analysis = analysis.copy()
    category_analysis["Food Composition Category"] = category_analysis["HER Category"].replace("", np.nan)
    category_analysis["Food Composition Category"] = category_analysis["Food Composition Category"].fillna("Needs Review")
    category = (
        category_analysis.groupby("Food Composition Category", dropna=False)["Pounds"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    category["Percent"] = np.where(total_lb > 0, category["Pounds"] / total_lb * 100, 0.0)

    if not category.empty:
        st.bar_chart(category.set_index("Food Composition Category")[["Percent"]], horizontal=True)
        display_category = category.rename(columns={"Food Composition Category": "Food Category", "Percent": "% of Total Load"})
        st.dataframe(
            display_category.style.format({"Pounds": "{:,.1f}", "% of Total Load": "{:.1f}%"}),
            hide_index=True,
            width="stretch",
        )

    st.markdown("### Item-level results")
    st.dataframe(analysis, hide_index=True, width="stretch")

# ============================================================
# CLEAN SEARCH-FIRST DASHBOARD
# ============================================================

with st.sidebar:
    st.header("Data")
    excel_path = st.text_input("Excel file path", value=DEFAULT_EXCEL_PATH)
    sheet_name = st.text_input("Worksheet name", value=DEFAULT_SHEET_NAME)
    if st.button("Reload data", width="stretch"):
        st.cache_data.clear()
        st.session_state.pop("backend_excel_bytes", None)
        st.rerun()

try:
    with st.spinner("Loading food data..."):
        food_df, nutrient_units, detected_header_row = load_food_data(excel_path, sheet_name)
except FileNotFoundError as error:
    st.error(str(error)); st.stop()
except PermissionError:
    st.error("The workbook could not be opened. Close it in Excel and try again."); st.stop()
except Exception as error:
    st.exception(error); st.stop()

unique_food_count = food_df[[FOOD_CODE, FOOD_DESCRIPTION]].drop_duplicates().shape[0]

app_mode = st.radio("Mode", ["Single food lookup", "Incoming load composition"], horizontal=True, key="app_mode")
if app_mode == "Incoming load composition":
    render_load_composition(food_df)
    st.stop()

with st.sidebar:
    st.success("Data loaded")
    st.caption(f"{unique_food_count:,} searchable foods")
    st.caption(f"Workbook rows use only: **{REFERENCE_MEASURE}**")
    st.caption("FDA supplemental fruit and seafood records retain their published reference serving.")

    with st.expander("Admin / backend", expanded=False):
        st.write("Generate the review workbook with classifications, warnings, rules, keywords, and manual-review fields.")
        if st.button("Generate backend workbook", width="stretch", type="primary"):
            with st.spinner("Building backend workbook..."):
                st.session_state["backend_excel_bytes"] = build_backend_workbook(food_df)
        if "backend_excel_bytes" in st.session_state:
            st.download_button(
                "Download backend workbook",
                data=st.session_state["backend_excel_bytes"],
                file_name="HER_food_classification_backend.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

st.title("🥗 HER Food Classification")
st.caption("Search by a food name, alternate name, or term in the source description. The dashboard uses one reference amount rather than asking users to choose a serving size.")

search_text = st.text_input(
    "Search for a food",
    placeholder="Try: leche fresca, whole milk, bacon, orange juice, chickpeas",
    key="food_search",
)
st.markdown('<div class="search-note">Alternative names inside “Includes:” are searchable. Example: <b>leche fresca</b> finds whole milk.</div>', unsafe_allow_html=True)

filter_col, code_col = st.columns([2.2, 1])
with filter_col:
    browse_category = st.selectbox("Category", ["All categories"] + HER_CATEGORIES)
with code_col:
    code_search = st.text_input("Food code", placeholder="Optional")

query = normalize_description_for_matching(search_text)
filtered_df = food_df.copy()
if browse_category != "All categories":
    filtered_df = filtered_df[filtered_df["_her_category"].eq(browse_category)]
if query:
    for term in query.split():
        filtered_df = filtered_df[filtered_df["_search_text"].str.contains(term, na=False, regex=False)]
if code_search.strip():
    filtered_df = filtered_df[filtered_df[FOOD_CODE].astype(str).str.contains(code_search.strip(), case=False, na=False, regex=False)]

active_filter = bool(query or code_search.strip() or browse_category != "All categories")
if not active_filter:
    st.info("Start typing a food name or synonym above. You can also browse one HER category.")
    st.stop()

food_options = (
    filtered_df[[FOOD_CODE, FOOD_DESCRIPTION, "_display_name", "_aliases", "_food_label", "_her_category", "_row_id"]]
    .drop_duplicates([FOOD_CODE, FOOD_DESCRIPTION])
    .sort_values(["_display_name", FOOD_CODE])
    .reset_index(drop=True)
)

if food_options.empty:
    st.warning("No foods matched your search.")
    st.stop()

match_count = len(food_options)
st.caption(f"{match_count:,} matching food{'s' if match_count != 1 else ''}.")
if match_count > MAX_VISIBLE_RESULTS:
    st.info(f"Showing the first {MAX_VISIBLE_RESULTS} matches. Add another word or choose a category to narrow the search.")
    food_options = food_options.head(MAX_VISIBLE_RESULTS)

selected_label = st.selectbox("Choose a result", food_options["_food_label"].tolist(), key="food_result")
selected_food = food_options[food_options["_food_label"].eq(selected_label)].iloc[0]
selected_food_code = selected_food[FOOD_CODE]
selected_food_description = selected_food[FOOD_DESCRIPTION]
selected_row = food_df.loc[food_df["_row_id"].eq(selected_food["_row_id"])].iloc[0]
selected_name = canonical_food_name(selected_food_description)
selected_aliases = food_aliases(selected_food_description)
selected_exclusions = food_exclusions(selected_food_description)

forced_reason = forced_unranked_reason(selected_food_description)
red_reason = forced_red_protein_reason(selected_food_description)
fresh_produce = bool(selected_row.get("_fresh_produce", False))
seafood_assumption = bool(selected_row.get("_seafood_assumption", False))
suggested_category = suggest_her_category(selected_food_description)
automatic_whole_grain_first = bool(selected_row.get("_whole_grain_detected", False))

with st.expander("Review classification inputs", expanded=False):
    st.caption("These controls are for manual review. Most users do not need to change them.")
    her_category = st.selectbox(
        "HER food category",
        HER_CATEGORIES,
        index=HER_CATEGORIES.index(suggested_category),
        help="The category is suggested from the food description.",
    )
    review_col_1, review_col_2 = st.columns(2)
    with review_col_1:
        whole_grain_first = st.checkbox(
            "Whole-grain wording appears in the food name",
            value=automatic_whole_grain_first,
            disabled=bool(forced_reason) or bool(red_reason) or fresh_produce or her_category not in {"Grains", "Processed and Packaged Snacks"},
        )
    with review_col_2:
        juice_or_dried_fruit = st.checkbox(
            "This is 100% juice or plain dried fruit",
            value=False,
            disabled=bool(forced_reason) or bool(red_reason) or fresh_produce or her_category != "Fruits and Vegetables",
        )

saturated_fat, saturated_column = numeric_from_columns(selected_row, ["Total Saturated", "Saturated Fat", "Saturated fat"])
sodium, sodium_column = numeric_from_columns(selected_row, ["Sodium"])
sugar, sugar_column = numeric_from_columns(selected_row, ["Added Sugars", "Added Sugar", "Total Added Sugars", "Total Sugars"])
if seafood_assumption and pd.isna(sugar):
    sugar, sugar_column = 0.5, "Assumed seafood sugar"

classification = classify_food_description(
    selected_food_description,
    her_category,
    saturated_fat,
    sodium,
    sugar,
    whole_grain_first,
    juice_or_dried_fruit,
    fresh_produce,
)

st.markdown(f'<div class="food-title">{selected_name}</div>', unsafe_allow_html=True)
header_parts = [her_category]
if selected_aliases:
    header_parts.append("Also known as: " + "; ".join(selected_aliases[:4]))
st.markdown(f'<div class="food-subtitle">{" &nbsp;•&nbsp; ".join(header_parts)}</div>', unsafe_allow_html=True)
if selected_aliases:
    alias_html = "".join(f'<span class="alias-chip">{alias}</span>' for alias in selected_aliases[:8])
    st.markdown(alias_html, unsafe_allow_html=True)

classification_tab, nutrition_tab, details_tab = st.tabs(["Classification", "Nutrition", "Details & source"])

with classification_tab:
    render_her_banner(classification["Overall"])

    def render_status_card(column, label: str, value: float, unit: str, rank: str) -> None:
        amount = "Unavailable" if pd.isna(value) else f"{format_number(value)} {unit}".strip()
        icon = HER_ICONS.get(rank, "⚪")
        html = f'<div class="status-card"><div class="status-label">{label}</div><div class="status-value">{amount}</div><div class="status-rank">{icon} {rank}</div></div>'
        with column:
            st.markdown(html, unsafe_allow_html=True)

    nutrient_cols = st.columns(3)
    render_status_card(nutrient_cols[0], "Saturated fat", saturated_fat, "g", classification["Saturated Fat"])
    render_status_card(nutrient_cols[1], "Sodium", sodium, "mg", classification["Sodium"])
    render_status_card(nutrient_cols[2], sugar_column or "Sugar", sugar, "g", classification["Sugar"])

    messages = []
    if forced_reason:
        messages.append(("warning", forced_reason))
    if red_reason:
        messages.append(("error", red_reason))
    if fresh_produce:
        messages.append(("success", "Plain fresh produce receives the fresh-produce override."))
    if seafood_assumption:
        messages.append(("info", "Supplemental seafood uses 0.5 g sugar and 0 g dietary fiber where those values are unavailable."))
    if sugar_column == "Total Sugars" and not forced_reason and not red_reason and not fresh_produce and not seafood_assumption:
        messages.append(("warning", "Added sugar is unavailable, so Total Sugars is used as the fallback."))
    missing_values = [name for name, value in [("saturated fat", saturated_fat), ("sodium", sodium), ("sugar", sugar)] if pd.isna(value)]
    if missing_values and not forced_reason and not red_reason and not fresh_produce:
        messages.append(("warning", "Unavailable classification values: " + ", ".join(missing_values) + "."))
    if classification["Rule"]:
        messages.append(("info", classification["Rule"]))

    for message_type, message in messages:
        getattr(st, message_type)(message)

    with st.expander("How was this classification determined?", expanded=False):
        st.write(f"**Category:** {her_category}")
        st.write(f"**Reference amount:** {create_serving_label(selected_row)}")
        st.write(f"**Saturated-fat source:** {saturated_column or 'Not found'}")
        st.write(f"**Sodium source:** {sodium_column or 'Not found'}")
        st.write(f"**Sugar source:** {sugar_column or 'Not found'}")
        st.write(f"**Whole-grain wording detected:** {'Yes' if whole_grain_first else 'No'}")
        st.write(f"**Fresh-produce override:** {'Yes' if fresh_produce else 'No'}")
        st.write(f"**Forced Unranked:** {'Yes' if forced_reason else 'No'}")
        st.write(f"**Forced protein Choose Rarely:** {'Yes' if red_reason else 'No'}")

with nutrition_tab:
    st.subheader("Nutrition for the reference amount")
    st.caption(create_serving_label(selected_row))
    render_metric_cards(selected_row, nutrient_units, 1.0)

    summary_df = nutrient_table(
        selected_row,
        [nutrient for nutrient in SUMMARY_NUTRIENTS if nutrient in food_df.columns],
        nutrient_units,
        1.0,
    )
    if summary_df.empty:
        st.info("No nutrient values were found for this record.")
    else:
        st.dataframe(summary_df, hide_index=True, width="stretch", height=520)

    with st.expander("Browse all nutrient fields", expanded=False):
        nutrient_filter = st.text_input("Filter nutrient names", placeholder="sodium, vitamin, fatty acid", key="nutrient_filter")
        all_df = nutrient_table(selected_row, all_nutrient_columns(food_df), nutrient_units, 1.0)
        if nutrient_filter.strip():
            all_df = all_df[all_df["Nutrient"].str.contains(nutrient_filter.strip(), case=False, na=False, regex=False)]
        st.dataframe(all_df, hide_index=True, width="stretch", height=500)

with details_tab:
    detail_left, detail_right = st.columns(2)
    with detail_left:
        st.markdown("#### Food record")
        st.write(f"**Display name:** {selected_name}")
        st.write(f"**Food code:** {selected_food_code}")
        st.write(f"**HER category:** {her_category}")
        st.write(f"**Data source:** {selected_row.get('_data_source', 'Unknown')}")
    with detail_right:
        st.markdown("#### Reference amount")
        st.write(f"**Measure:** {clean_text(selected_row.get(MEASURE_DESCRIPTION, ''))}")
        st.write(f"**Number of servings:** {format_number(selected_row.get(NUMBER_SERVINGS, np.nan))}")
        weight = selected_row.get(SERVING_WEIGHT, np.nan)
        st.write(f"**Serving weight:** {'Unavailable' if pd.isna(weight) else format_number(weight) + ' g'}")

    if selected_aliases:
        st.markdown("#### Searchable alternate names")
        st.write("; ".join(selected_aliases))

    with st.expander("Original source description", expanded=False):
        st.write(selected_food_description)

    with st.expander("HER guideline rule table", expanded=False):
        rule_view = backend_rule_table()
        st.dataframe(rule_view[rule_view["Food Category"].eq(her_category)], hide_index=True, width="stretch")
