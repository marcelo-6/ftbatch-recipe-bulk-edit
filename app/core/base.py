"""
Shared constants and utilities.
"""

NAMESPACE = "urn:Rockwell/MasterRecipe"
NSMAP = {None: NAMESPACE}

# Excel column configuration (fixed part)
EXCEL_COLUMNS = [
    "TagType",
    "Name",
    "FullPath",
    "Real",
    "Integer",
    "High",
    "Low",
    "String",
    "EnumerationSet",
    "EnumerationMember",
    "Defer",
    # FVL columns
    "FormulaValueLimit_Verification",
    "FormulaValueLimit_LowLowLowValue",
    "FormulaValueLimit_LowLowValue",
    "FormulaValueLimit_LowValue",
    "FormulaValueLimit_HighValue",
    "FormulaValueLimit_HighHighValue",
    "FormulaValueLimit_HighHighHighValue",
]
