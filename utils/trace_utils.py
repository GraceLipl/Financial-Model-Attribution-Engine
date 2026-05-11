from openpyxl.utils import range_boundaries, get_column_letter

from utils.workbook_utils import (
    lookup_column_period_from_map,
    resolve_label_with_hierarchy,
    parse_sheet_reference,
)


def trace_merged(wb_values, wb_formulas, sheet_name, cell_ref, all_col_to_period, max_depth=20, visited=None, depth=0, path=None):
    def is_terminal_cell(sheet_name, cell_ref, formula):
        col_letter = re.match(r"[A-Z]+", cell_ref).group(0)
        # Updated to use the new lookup function
        period = lookup_column_period_from_map(all_col_to_period, sheet_name, cell_ref)
        if not period:
            return True
        if sheet_name.lower().startswith("schedule"):
            return True
        if formula is None:
            return True
        if isinstance(formula, str) and "[" in formula:
            return True
        if re.match(r"^=[0-9.\-+*/ ()]+$", formula):
            return True
        return False

    if visited is None:
        visited = set()
    if path is None:
        path = []

    key = (sheet_name, cell_ref)
    if key in visited or depth > max_depth:
        return []

    visited.add(key)
    current_path = path + [f"{sheet_name}!{cell_ref}"]

    sheet_v = wb_values[sheet_name]
    sheet_f = wb_formulas[sheet_name]
    cell_v = sheet_v[cell_ref]

    try:
        cell_f = sheet_f[cell_ref]
        formula = cell_f.value if isinstance(cell_f.value, str) and cell_f.value.startswith("=") else None
        description = resolve_label_with_hierarchy(wb_values, sheet_name, cell_f.row) or f"Row {cell_f.row}"
    except KeyError:
        formula = None
        description = None

    value = cell_v.value
    external_ref_segment = ""
    is_external = False

    if formula and "[" in formula:
        match_ref = re.search(r"(\[[^\]]+\][^\+\-\*/\)^\n]*)", formula)
        external_ref_segment = match_ref.group(1) if match_ref else "[EXTERNAL]"
        is_external = True

    is_terminal = is_terminal_cell(sheet_name, cell_ref, formula)

    numeric_literals = []
    if not is_external and formula:
        # Remove quoted strings first
        cleaned_formula = re.sub(r'"[^"]*"', '', formula)
        
        # Find all cell references (including those with $ and sheet names)
        cell_refs = re.findall(r'(?<![0-9A-Za-z_])(?:\'[^\']*\'!\$?[A-Z]{1,3}\$?[0-9]{1,5}|[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,5}|\$?[A-Z]{1,3}\$?[0-9]{1,5})(?![0-9A-Za-z_])', cleaned_formula)
        
        # Remove all cell references from the formula to isolate numeric constants
        temp_formula = cleaned_formula
        for ref in cell_refs:
            temp_formula = temp_formula.replace(ref, '')
        
        # Extract numeric constants
        all_numbers = re.findall(r'[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?', temp_formula)
        filtered_constants = []
        for num_str in all_numbers:
            try:
                val = float(num_str)
                if val == 1.0 and (re.search(r'\(1\s*\+\s*[\w\d]', formula) or re.search(r'[\w\d]\s*\+\s*1\)', formula)):
                    continue
                if val == -1.0 and formula.strip().endswith("-1"):
                    continue
                filtered_constants.append(val)
            except:
                continue
        numeric_literals = filtered_constants

    # Updated period lookup - use the new function
    period = lookup_column_period_from_map(all_col_to_period, sheet_name, cell_ref)

    output = [{
        "sheet": sheet_name,
        "cell": cell_ref,
        "formula": formula,
        "value": value,
        "depth": depth,
        "description": description,
        "is_terminal": is_terminal or is_external,
        "external_ref_segment": external_ref_segment,
        "constants": numeric_literals,
        "path": current_path,
        "period": period
    }]

    if is_terminal or is_external:
        return output

    refs = []
    ranges = []
    if isinstance(formula, str):
        # Updated regex to handle quoted sheet names AND $ symbols for absolute references
        # Pattern explanation:
        # - (?:'[^']*'!\$?[A-Z]{1,3}\$?[0-9]{1,5}) - quoted sheet with absolute refs
        # - (?:[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,5}) - unquoted sheet with absolute refs  
        # - (?:\$?[A-Z]{1,3}\$?[0-9]{1,5}) - local cell with absolute refs
        refs = re.findall(r"(?<![0-9A-Za-z_])(?:'[^']*'!\$?[A-Z]{1,3}\$?[0-9]{1,5}|[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,5}|\$?[A-Z]{1,3}\$?[0-9]{1,5})(?![0-9A-Za-z_])", formula)
        refs = [ref.replace("$", "") for ref in refs]
        
        # Updated regex for ranges with quoted sheet names AND $ symbols
        ranges = re.findall(r"('[^']*'!\$?[A-Z]{1,3}\$?[0-9]{1,5}:\$?[A-Z]{1,3}\$?[0-9]{1,5}|[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,5}:\$?[A-Z]{1,3}\$?[0-9]{1,5}|\$?[A-Z]{1,3}\$?[0-9]{1,5}:\$?[A-Z]{1,3}\$?[0-9]{1,5})", formula)
        ranges = [rng.replace("$", "") for rng in ranges]

    for ref in refs:
        # Use the new parsing function
        other_sheet, other_cell = parse_sheet_reference(ref)
        if other_sheet is None:
            other_sheet = sheet_name
            other_cell = ref

        col_letter = re.match(r"[A-Z]+", other_cell).group(0)

        # Updated filtering logic
        if other_sheet.lower().startswith("schedule"):
            continue
        
        # Check if this sheet has any period mapping at all
        if other_sheet not in all_col_to_period:
            continue
            
        # Check if this specific column has a period
        if col_letter not in all_col_to_period[other_sheet]:
            continue

        # Check if the sheet exists before trying to trace
        if other_sheet not in wb_values:
            print(f"Warning: Sheet '{other_sheet}' not found in workbook")
            continue

        output += trace_merged(wb_values, wb_formulas, other_sheet, other_cell, all_col_to_period, max_depth, visited, depth + 1, current_path)

    for rng in ranges:
        # Use the new parsing function for ranges too
        other_sheet, cell_range = parse_sheet_reference(rng)
        if other_sheet is None:
            other_sheet = sheet_name
            cell_range = rng

        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                col_letter = get_column_letter(col)
                
                # Updated filtering logic for ranges
                if other_sheet.lower().startswith("schedule"):
                    continue
                
                # Check if this sheet has any period mapping at all
                if other_sheet not in all_col_to_period:
                    continue
                    
                # Check if this specific column has a period
                if col_letter not in all_col_to_period[other_sheet]:
                    continue
                
                # Check if the sheet exists before trying to trace
                if other_sheet not in wb_values:
                    print(f"Warning: Sheet '{other_sheet}' not found in workbook")
                    continue
                    
                cell = f"{col_letter}{row}"
                output += trace_merged(wb_values, wb_formulas, other_sheet, cell, all_col_to_period, max_depth, visited, depth + 1, current_path)

    return output



def is_effectively_zero(value, tolerance=1e-10):
    """Check if a value is effectively zero, accounting for floating-point precision"""
    if value is None:
        return True
    if isinstance(value, (int, float)):
        return abs(value) < tolerance
    return False

def is_error_value(value):
    """Check if a value is an Excel error (starts with #)"""
    if value is None:
        return False
    return str(value).startswith("#")



def is_terminal_input(item, trace, all_col_to_period):
    period = item["period"]
    if not period.endswith("E"):
        return False

    formula = item.get("formula")
    constants = item.get("constants", [])
    
    excluded_constants = {0.25, 0.33, 0.5, 0.67, 0.75, 1/4, 1/3, 1/2, 2/3, 3/4, 
                          1, -1, 2, 3, 4, 52, 90, 100, 
                          180, 270, 365, 1000, 10000, 100000, 1000000}
    meaningful_constants = [c for c in constants if c not in excluded_constants]
    has_meaningful_constants = bool(meaningful_constants)

    if not formula or "[" in formula:
        return True

    if re.match(r"^=[0-9.\-+*/ ()]+$", formula):
        return True

    if has_meaningful_constants:
        complex_keywords = ["IF", "CONCATENATE", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", 
                            "CHOOSE", "OFFSET", "INDIRECT", "SUMIF", "COUNTIF", "AVERAGEIF"]
        formula_upper = formula.upper()
        if any(keyword in formula_upper for keyword in complex_keywords):
            return False
        return True

    refs = re.findall(
        r'(?<![0-9A-Za-z_])(?:\'[^\']*\'!\$?[A-Z]{1,3}\$?[0-9]{1,5}|[A-Za-z0-9_]+!\$?[A-Z]{1,3}\$?[0-9]{1,5}|\$?[A-Z]{1,3}\$?[0-9]{1,5})(?![0-9A-Za-z_])',
        formula
    )

    for ref in refs:
        ref = ref.replace("$", "")
        ref_sheet, ref_cell = parse_sheet_reference(ref)
        if ref_sheet is None:
            ref_sheet = item['sheet']

        ref_found = False
        for ref_item in trace:
            if ref_item["sheet"] == ref_sheet and ref_item["cell"] == ref_cell:
                ref_found = True
                if ref_item["period"].endswith("E"):
                    return False
                break

        if not ref_found:
            ref_period = lookup_column_period_from_map(all_col_to_period, ref_sheet, ref_cell)
            if ref_period.endswith("E"):
                return False

    return True



import re
from collections import defaultdict

def get_terminal_cell_path(terminal_item):
    return f"{terminal_item['sheet']}!{terminal_item['cell']}"

def find_terminal_chain(filtered_trace, terminal_item):
    """
    Find the full path (chain of cell references) for a given terminal item.
    """
    terminal_path = get_terminal_cell_path(terminal_item)
    for item in filtered_trace:
        if get_terminal_cell_path(item) == terminal_path:
            return item["path"]
    return []

def filter_chain_by_period(chain_paths, terminal_item, all_col_to_period):
    """
    Keep only items in the same period as the terminal.
    Returns list of (sheet, cell) tuples AND the filtered path strings for printing.
    """
    terminal_period = terminal_item["period"]
    filtered_tuples = []
    filtered_paths = []
    
    for path in chain_paths:
        sheet, cell = path.split("!")
        period = lookup_column_period_from_map(all_col_to_period, sheet, cell)
        if period == terminal_period:
            filtered_tuples.append((sheet, cell))
            filtered_paths.append(path)
    
    return filtered_tuples, filtered_paths

def get_trace_lookup(trace):
    return {f"{item['sheet']}!{item['cell']}": item for item in trace}