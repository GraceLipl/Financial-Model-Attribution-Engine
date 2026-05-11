import re

from utils.trace_utils import (
    get_terminal_cell_path,
    get_trace_lookup,
    find_terminal_chain,
    filter_chain_by_period,
)


from utils.workbook_utils import (
    build_column_period_map,
    parse_sheet_reference,
    lookup_column_period_from_map, 
)



def get_nearest_parents_descriptions(filtered_chain, terminal_item, trace_lookup, num_parents=2):
    """
    From filtered chain, return up to two parent descriptions above the terminal cell.
    Also tries to find adjacent row descriptions if chain is too short.
    """
    terminal_path = get_terminal_cell_path(terminal_item)
    terminal_row = int(re.match(r"[A-Z]+(\d+)", terminal_item["cell"]).group(1))
    terminal_sheet = terminal_item["sheet"]
    terminal_col = re.match(r"([A-Z]+)\d+", terminal_item["cell"]).group(1)

    descriptions = []
    
    # First, try to get descriptions from the filtered chain
    for sheet, cell in reversed(filtered_chain):
        path = f"{sheet}!{cell}"
        if path == terminal_path:
            continue
        row = int(re.match(r"[A-Z]+(\d+)", cell).group(1))
        if sheet == terminal_sheet and row == terminal_row:
            continue
        desc = trace_lookup.get(path, {}).get("description", "")
        if desc and desc not in descriptions:
            descriptions.append(desc)
        if len(descriptions) == num_parents:
            break
    
    # If we don't have enough descriptions from chain, try adjacent rows
    if len(descriptions) < num_parents and len(filtered_chain) <= 1:
        print(f"DEBUG PARENTS: Chain too short ({len(filtered_chain)}), checking adjacent rows for {terminal_path}")
        
        # Check rows above the terminal (row-1, row-2, etc.)
        for offset in range(1, 5):  # Check up to 4 rows above
            if len(descriptions) >= num_parents:
                break
            adjacent_row = terminal_row - offset
            if adjacent_row > 0:
                adjacent_path = f"{terminal_sheet}!{terminal_col}{adjacent_row}"
                desc = trace_lookup.get(adjacent_path, {}).get("description", "")
                if desc and desc not in descriptions and desc != terminal_item.get("description", ""):
                    descriptions.append(desc)
                    print(f"DEBUG PARENTS: Found adjacent row description: {adjacent_path} -> '{desc}'")
    
    return list(reversed(descriptions))  # Return in order: parent1, parent2

def get_enriched_terminal_description(terminal_item, trace, all_col_to_period):
    """
    Enhanced version that returns both enriched description and chain info for printing.
    Format: branch → subbranch → terminal_description
    """
    trace_lookup = get_trace_lookup(trace)
    terminal_path = get_terminal_cell_path(terminal_item)

    full_chain = find_terminal_chain(trace, terminal_item)
    if not full_chain:
        return terminal_item.get("description", ""), None, None

    # Get both filtered tuples and paths for printing
    filtered_chain, filtered_paths = filter_chain_by_period(full_chain, terminal_item, all_col_to_period)
    
    if not filtered_chain:
        return terminal_item.get("description", ""), None, None

    parents = get_nearest_parents_descriptions(filtered_chain, terminal_item, trace_lookup)
    
    # Build enriched description: parents + terminal description (terminal at the end)
    terminal_desc = terminal_item.get("description", "")
    enriched_parts = parents + [terminal_desc] if terminal_desc else parents
    enriched_desc = " → ".join(p for p in enriched_parts if p)
    
    # If no enrichment occurred (no parents found), just return original description
    if not parents:
        enriched_desc = terminal_desc
    
    # Return enriched description, full chain info, and filtered chain info
    chain_info = {
        "full_chain": full_chain,
        "filtered_chain": filtered_paths,
        "full_length": len(full_chain),
        "filtered_length": len(filtered_paths)
    }
    
    return enriched_desc, chain_info, parents




import datetime
from openpyxl.utils import column_index_from_string, get_column_letter

def should_force_zero(value, description, number_format):
    """
    Returns True if the value and description indicate a growth/change input that should be forced to 0.
    """
    if not isinstance(value, (int, float)):
        return False

    if not description:
        return False

    desc_lower = description.lower()

    # Keywords for growth/change situationsa
    keywords = [
        "% change", "growth", "yoy", "yr/yr", "q/q", "mom", "y/y", "m/m",
        "variance", "var", "delta", "change", "chg", "chng",
        "inflation", "cpi", "ppi", "fx impact", "discount rate change",
        "wacc delta", "wacc change", "risk premium change",
        "margin delta", "cost increase rate", "tax change", "tax rate change"
    ]

    for kw in keywords:
        if kw in desc_lower:
            if "margin %" in desc_lower or "tax rate %" in desc_lower:
                return False
            return True

    return False



def extract_numeric_constants(formula):
    if not formula or not formula.startswith("="):
        return []

    # Remove quoted strings (e.g., text literals in formulas)
    cleaned = re.sub(r'"[^"]*"', '', formula)

    # Tokenize the formula into pieces split by operators, parentheses, commas, spaces
    tokens = re.split(r'[\+\-\*/\^\(\),\s]+', cleaned)

    constants = []
    for token in tokens:
        token = token.strip()
        # Skip if it's a cell reference like A1, AA123, etc.
        if re.match(r'^[A-Za-z]{1,3}[0-9]{1,5}$', token):
            continue
        try:
            constants.append(float(token))
        except ValueError:
            pass  # ignore non-numeric tokens
    return constants



def run_enrichment_for_all_terminals(terminal_inputs, trace, all_col_to_period):
        """
        Enrich and print descriptions for all terminal input cells with chain printing.
        """
        enriched_results = []

        for i, item in enumerate(terminal_inputs, 1):

            enriched_desc, chain_info, parents = get_enriched_terminal_description(item, trace, all_col_to_period)

            enriched_results.append({
                "sheet": item["sheet"],
                "cell": item["cell"],
                "period": item["period"],
                "value": item.get("value"),
                "depth": item.get("depth"),
                "original_desc": item.get("description", ""),
                "enriched_desc": enriched_desc,
                "chain_info": chain_info
            })

        enriched_count = 0
        for result in enriched_results:
            enriched = result["enriched_desc"]
            original = result["original_desc"]
            enriched_flag = "✅ ENRICHED" if ("→" in enriched and enriched != original) else "❌ NO ENRICHMENT"
            if enriched_flag == "✅ ENRICHED":
                enriched_count += 1

            chain_length = result["chain_info"]["filtered_length"] if result["chain_info"] else 0
        total = len(enriched_results)

        return enriched_results
    



def run_baseline_forcing(terminal_inputs, enriched_results, all_col_to_period, wb_values, wb_formulas):
    enriched_descriptions = {
        f"{result['sheet']}!{result['cell']}": result['enriched_desc']
        for result in enriched_results
    }

    sheet_names = {item['sheet'] for item in terminal_inputs}

    manual_today_date = datetime.date(2025, 4, 1)
    sheet_to_col_period = {}

    for sheet in sheet_names:
        if sheet in all_col_to_period:
            sheet_to_col_period[sheet] = all_col_to_period[sheet]
        else:
            sheet_to_col_period[sheet] = build_column_period_map(
                wb_values, sheet, start_row=1, today_date=manual_today_date, max_row=150
            )

    terminal_period_map = {}
    for item in terminal_inputs:
        sheet = item["sheet"]
        cell = item["cell"]
        period = item["period"]
        col = re.match(r"[A-Z]+", cell).group(0)
        terminal_period_map.setdefault(sheet, {})[col] = period

    baseline_inputs = []
    recorded_cells = set()

    for item in terminal_inputs:
        try:
            # --- simplified to skip inner debug comments ---
            sheet = item["sheet"]
            cell = item["cell"]
            row = int(re.match(r"[A-Z]+(\d+)", cell).group(1))
            desc = item.get("description", "").lower()
            col_period_map = sheet_to_col_period[sheet]

            hist_cols = [col for col, period in col_period_map.items() if not period.endswith("E")]
            forecast_cols = [col for col, period in col_period_map.items() if period.endswith("E")]
            hist_cols_to_use = hist_cols

            if forecast_cols:
                first_forecast_col = sorted(forecast_cols, key=column_index_from_string)[0]
                first_index = column_index_from_string(first_forecast_col)
                hist_cols_to_use = [
                    col for col in hist_cols if column_index_from_string(col) < first_index
                ] or hist_cols

            terminal_period = item["period"]
            if terminal_period and re.match(r'^\d{4}E$', terminal_period):
                year = int(terminal_period.replace('E', ''))
                prior_year_cols = [col for col, period in col_period_map.items() if period == str(year - 1)]
                latest_hist_col = prior_year_cols[0] if prior_year_cols else sorted(hist_cols_to_use, key=column_index_from_string)[-1]
            else:
                latest_hist_col = sorted(hist_cols_to_use, key=column_index_from_string)[-1]

            latest_hist_cell = f"{latest_hist_col}{row}"
            latest_hist_val = wb_values[sheet][latest_hist_cell].value
            original_val = wb_values[sheet][cell].value
            number_format = wb_values[sheet][cell].number_format

            if should_force_zero(original_val, item.get("description", ""), number_format):
                baseline_val, baseline_source = 0, "forced_zero_due_to_indicator"
            else:
                formula = item.get("formula") or ""
                ref_cells = re.findall(r"(?:'[^']+'|[^'!]+)![A-Z]+\d+|\$?[A-Z]+\$?\d+", formula)
                ref_cells = [ref.replace("$", "") for ref in ref_cells]

                if len(ref_cells) == 1:
                    ref = ref_cells[0]
                    ref_sheet, ref_cell = parse_sheet_reference(ref)
                    if ref_sheet is None:
                        ref_sheet = sheet
                    ref_period = lookup_column_period_from_map(all_col_to_period, ref_sheet, ref_cell)
                    baseline_val = (
                        wb_values[ref_sheet][ref_cell].value
                        if not ref_period.endswith("E")
                        else latest_hist_val
                    )
                    baseline_source = f"{ref_sheet}!{ref_cell}"
                else:
                    baseline_val, baseline_source = latest_hist_val, latest_hist_cell

            enriched_desc = enriched_descriptions.get(f"{sheet}!{cell}", item.get("description", ""))

            # Build record
            formula_obj = wb_formulas[sheet][cell]
            cell_formula = formula_obj.value if isinstance(formula_obj.value, str) and formula_obj.value.startswith("=") else ""
            period_display = item["period"]

            if (sheet, cell) in recorded_cells:
                continue
            recorded_cells.add((sheet, cell))

            baseline_inputs.append({
                "sheet": sheet,
                "cell": cell,
                "desc": enriched_desc,
                "original_value": original_val,
                "change_to": baseline_val,
                "type": "baseline_forcing",
                "info": f"Forced to {baseline_val} from {baseline_source}",
                "period": period_display,
                "formula": cell_formula
            })

        except Exception as e:
            continue
    return baseline_inputs

