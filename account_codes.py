"""
Single source of truth for HITOPS Account Codes.
Used by PRMakerWidget (UI dropdown), main.py (legacy dialog) and
menu_navigator.py (dropdown navigation by index).

IMPORTANT: The order of ACCOUNT_CODES must exactly match the order of items
shown in the HITOPS Account Code dropdown, otherwise index-based navigation
will select the wrong item.
"""

# Display form: "<code>/<korean description>" — used by UI dropdowns
ACCOUNT_CODES = [
    "0501030000/수선유지비",
    "0501030100/수선유지비",
    "0501030101/장비 자재비-QC",
    "0501030102/장비 자재비-ATC",
    "0501030103/장비 자재비-RS",
    "0501030104/장비 자재비-YT",
    "0501030105/장비 자재비-YC",
    "0501030106/장비 자재비-FL",
    "0501030107/장비 자재비-기타",
    "0501030108/수선유지비-외주수리-QC",
    "0501030109/수선유지비-외주수리-ATC",
    "0501030110/수선유지비-외주수리-RS",
    "0501030111/수선유지비-외주수리-YT",
    "0501030112/수선유지비-외주수리-YC",
    "0501030113/수선유지비-외주수리-FL",
    "0501030114/수선유지비-외주수리-기타",
    "0501030115/시설물-야드시설물(자재)",
    "0501030116/수선유지비-시설물-CFS시설물",
    "0501030117/시설물-전기시설물(자재)",
    "0501030118/시설물-외주수리",
    "0501030119/수선유지비_작업공구-야드공구",
    "0501030120/수선유지비_작업공구-정비공구",
    "0501030121/수선유지비_작업공구-CFS공구",
    "0501030122/수선유지비_작업공구-안전공구",
    "0501030123/수선유지비_작업공구-기타공구",
    "0501030124/수선유지비_작업소모품-야드소모품",
    "0501030125/작업소모품-정비소모품/공구",
    "0501030126/수선유지비_작업소모품-CFS소모품",
    "0501030127/수선유지비_작업소모품-안전소모품",
    "0501030128/수선유지비_작업소모품-기타소모품",
    "0501030129/수선유지비-CNTR",
    "0501030130/수선유지비-기타 (사고변상금등)",
    "0501030131/장비자재비-ECH",
    "0501030132/수선유지비-외주수리-ECH",
    "0501040106/동력비-윤활유",
]

# Prefix-only form: "<code>" — used by index-based dropdown navigation
ACCOUNT_CODE_PREFIXES = [c.split('/')[0].strip() for c in ACCOUNT_CODES]


def find_index_by_prefix(code_text):
    """Return the dropdown index for a given code/displayed string, or None."""
    if not code_text:
        return None
    prefix = code_text.split('/')[0].strip()
    try:
        return ACCOUNT_CODE_PREFIXES.index(prefix)
    except ValueError:
        return None
