import re

# The excel columns A or B getting convertet into numbers
def excel_col_to_number(col: str) -> int:
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - ord("A") + 1)
    return num

# Changes the Numbers back into excel columns 
def number_to_excel_col(n: int) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord("A")) + result
    return result

# Increases the Excel column by one
def next_column(col: str) -> str:
    return number_to_excel_col(excel_col_to_number(col) + 1)

# Reads an existing formula and increases the columns by one: =SUMME(AY7:AY11)  -> =SUMME(AZ7:AZ11)
def change_formula(formula: str) -> str:

    def repl(match):
        return next_column(match.group(1))
    
    #Pattern so the formal operater doesnt get interpreted as column
    pattern = r"([A-Z]+)(?=\d)"

    return re.sub(pattern, repl, formula)

# Increases the Excel column by two
def next_two_columns(col: str) -> str:
    return number_to_excel_col(excel_col_to_number(col) + 2)

# Reads an existing formula and increases the columns by two: =Volume_Report!CB12 -> =Volume_Report!CD12
def change_formula_two_columns(formula: str) -> str:
    def repl(match):
        return next_two_columns(match.group(1))
    
    # Pattern so the formula operator doesn't get interpreted as column
    pattern = r"([A-Z]+)(?=\d)"
    
    return re.sub(pattern, repl, formula)

def increment_working_days_reference(formular: str) -> str:

    current_formula = formular

    if not isinstance(current_formula, str) or not current_formula.startswith("="):
        return
    
    # Get the Pattern
    match = re.search(r"(.*?)(\d+)$", current_formula)

    if not match:
        return

    prefix = match.group(1)      
    number = int(match.group(2))

    # increase number by one
    new_number = number + 1

    new_formula = f"{prefix}{new_number}"

    return new_formula