
def run_attribution_analysis_complete(baseline_inputs, workbook_path, target_sheet, target_cell):
    """
    Complete attribution analysis: No-copy connection logic + Full Jupyter attribution logic
    """
    import xlwings as xw
    import time
    import os
    import re
    from datetime import datetime
    from collections import defaultdict
    import traceback

    # Store all output for return
    attribution_output = []
    summary = {}
    
    def capture_print(msg):
        """Capture print output for return to main function"""
        msg_str = str(msg)
        print(msg_str)  # Still print to console
        attribution_output.append(msg_str)

    # === Filter invalid baseline inputs (from Jupyter) ===
    baseline_inputs = [item for item in baseline_inputs if item["change_to"] is not None]
    
    capture_print("🚀 New Version!!!!!!!!!!!Starting complete attribution analysis...")
    capture_print("⚠️ Working directly on original file - will restore all changes at end")
    capture_print(f"📋 Processing {len(baseline_inputs)} baseline inputs")
    capture_print(f"🎯 Target: {target_sheet}!{target_cell}")
    capture_print(f"📂 Original file: {workbook_path}")



    # === STEP 1-2: Connect to existing Excel workbook (from working no-copy version) ===
    wb = None
    app = None
    original_values_backup = {}  # Store original values for restoration
    error_msg = False


    try:
        capture_print("📂 Connecting to Excel workbook...")
        
        # Strategy: Connect to already open workbook
        try:
            # First, try to find existing Excel apps
            apps = xw.apps
            if len(apps) > 0:
                app = apps[0]
                capture_print("   Found existing Excel application")
                
                # Look for the workbook by name
                workbook_name = os.path.basename(workbook_path)
                wb = None
                
                for book in app.books:
                    if book.name == workbook_name or book.fullname == workbook_path:
                        wb = book
                        capture_print(f"   Found open workbook: {book.name}")
                        break
                
                if wb is None:
                    # Open the file in existing Excel
                    capture_print(f"   Opening {workbook_name} in existing Excel...")
                    wb = app.books.open(workbook_path, update_links=False)
            else:
                raise Exception("No Excel apps found")
                
        except Exception as e:
            # Fallback: Create new Excel and open file
            capture_print(f"   Creating new Excel application... ({str(e)})")
            app = xw.App(visible=True, add_book=False)
            wb = app.books.open(workbook_path, update_links=False)
        
        if wb is None:
            raise Exception("Could not open workbook")
            
        capture_print("✅ Excel workbook connected successfully")
        
        # Configure Excel settings (non-critical)
        try:
            app.screen_updating = False
            app.enable_events = False
            app.display_alerts = False
            app.calculation = 'manual'
            capture_print("   Excel settings configured")
        except:
            capture_print("   ⚠️ Some Excel settings could not be configured")

        # === STEP 3: Get original forecast value (from working no-copy version) ===
        capture_print("📌 Reading original forecast value...")
        original_forecast_value = wb.sheets[target_sheet].range(target_cell).value
        
        if original_forecast_value is None:
            raise Exception(f"Target cell {target_sheet}!{target_cell} is empty or not accessible")
            
        capture_print(f"📌 Original Forecast Target Value: {original_forecast_value:.4f}")

        # === STEP 4: Backup original values AND FORMULAS ===
        capture_print("💾 Backing up original values and formulas...")
        
        for item in baseline_inputs:
            sheet_name = item["sheet"]
            cell_ref = item["cell"]
            
            try:
                rng = wb.sheets[sheet_name].range(cell_ref)
                original_val = rng.value  # Get the calculated value
                original_formula = rng.formula  # Get the formula (this was missing!)
                
                original_values_backup[f"{sheet_name}!{cell_ref}"] = {
                    "value": original_val,
                    "formula": original_formula  # Now backing up formulas too!
                }
                
                # Debug: Show what we're backing up for first few cells
                if len(original_values_backup) < 5:
                    if original_formula and original_formula.startswith('='):
                        capture_print(f"   Backing up {sheet_name}!{cell_ref}: formula='{original_formula}', value={original_val}")
                    else:
                        capture_print(f"   Backing up {sheet_name}!{cell_ref}: no formula, value={original_val}")
                        
            except Exception as e:
                capture_print(f"⚠️ Could not backup {sheet_name}!{cell_ref}: {e}")

        capture_print(f"💾 Backed up {len(original_values_backup)} cell values and formulas")

        # === STEP 5: Apply baseline changes (from Jupyter logic) ===
        capture_print("🔧 Applying baseline changes...")
        changes_applied = 0
        changes_verified = 0

        for i, item in enumerate(baseline_inputs):
            sheet, cell, change_to = item["sheet"], item["cell"], item["change_to"]
            original_value = item["original_value"]
            
            try:
                rng = wb.sheets[sheet].range(cell)
                
                # Check current value before change
                current_value = rng.value
                
                # Apply change
                rng.formula = None
                rng.value = change_to
                changes_applied += 1
                
                # Verify the change was applied
                new_value = rng.value
                if abs(new_value - change_to) < 1e-10:
                    changes_verified += 1
                
                # Show first few changes for debugging
                if i < 5:
                    capture_print(f"   {sheet}!{cell}: {original_value:.6f} → {current_value:.6f} → {new_value:.6f} (target: {change_to:.6f})")
                elif i == 5:
                    capture_print("   ...")
                
                if i % 10 == 0:  # Progress indicator
                    capture_print(f"   Applied {i+1}/{len(baseline_inputs)} changes...")
                    
            except Exception as e:
                capture_print(f"⚠️ Error applying baseline to {sheet}!{cell}: {e}")

        capture_print(f"📊 Changes applied: {changes_applied}/{len(baseline_inputs)}")
        capture_print(f"📊 Changes verified: {changes_verified}/{len(baseline_inputs)}")

        # === STEP 6: Force calculation (from Jupyter logic with no-copy robustness) ===
        capture_print("\n🔄 Forcing calculation after all baseline changes...")
        try:
            app.calculate()
            time.sleep(1)
            # Force a second calculation to ensure all dependencies are updated
            app.calculate()
            time.sleep(1)
            capture_print("✅ Forced calculations completed")
        except Exception as e:
            capture_print(f"⚠️ Error during forced calculation: {e}")

        # === STEP 7: Calculate baseline (from Jupyter logic) ===
        capture_print("\n🧮 Calculating baseline...")
        for i in range(5):  # Increased attempts
            try:
                capture_print(f"   Calculation attempt {i+1}...")
                app.calculate()
                time.sleep(2)
                
                # Check if target value changed
                current_target = wb.sheets[target_sheet].range(target_cell).value
                if current_target is not None:
                    capture_print(f"   Target after calc {i+1}: {current_target:.6f}")
                    
                    if abs(current_target - original_forecast_value) > 1e-10:
                        capture_print("   ✅ Target value changed, calculation successful")
                        break
                    else:
                        capture_print("   ⚠️ Target unchanged, trying again...")
                else:
                    capture_print(f"   ⚠️ Target is None after calculation {i+1}")
                    
            except Exception as e:
                capture_print(f"⚠️ Excel busy (attempt {i+1}): {e}")
                time.sleep(3)

        base_target_value = wb.sheets[target_sheet].range(target_cell).value
        if base_target_value is None:
            capture_print("❌ Cannot proceed - baseline target value is None")
            raise Exception("Baseline target value is None after calculations")
            
        capture_print(f"📌 Baseline Target Value: {base_target_value:.6f}")

        # === STEP 8: Diagnostic checks (from Jupyter logic) ===
        capture_print("\n🔍 Diagnostic: Verifying baseline changes are still applied...")
        still_applied = 0
        reverted_count = 0

        for i, item in enumerate(baseline_inputs[:10]):  # Check first 10 for speed
            sheet, cell, change_to = item["sheet"], item["cell"], item["change_to"]
            original_value = item["original_value"]
            
            try:
                current_value = wb.sheets[sheet].range(cell).value
                if current_value is not None:
                    if abs(current_value - change_to) < 1e-10:
                        still_applied += 1
                    elif abs(current_value - original_value) < 1e-10:
                        reverted_count += 1
                    
                    if i < 5:
                        capture_print(f"   {sheet}!{cell}: current={current_value:.6f}, should_be={change_to:.6f}, original={original_value:.6f}")
            except Exception as e:
                capture_print(f"⚠️ Error checking {sheet}!{cell}: {e}")

        capture_print(f"📊 First 10 cells: {still_applied} still changed, {reverted_count} reverted to original")

        # === STEP 9: Check calculation mode (from Jupyter logic) ===
        capture_print(f"\n🔍 Excel calculation mode: {app.calculation}")
        if app.calculation != 'manual':
            capture_print("⚠️ Calculation mode is not manual, setting to manual...")
            try:
                app.calculation = 'manual'
                capture_print("✅ Set to manual calculation")
            except Exception as e:
                capture_print(f"⚠️ Could not set manual calculation: {e}")

        # === STEP 10: Aggressive calculation if needed (from Jupyter logic) ===
        if abs(base_target_value - original_forecast_value) < 1e-10:
            capture_print("\n🔄 Target still unchanged, trying more aggressive calculation...")
            try:
                # Try calculating specific sheets
                wb.sheets[target_sheet].calculate()
                time.sleep(1)
                app.calculate()
                time.sleep(2)
                
                # Check again
                final_check = wb.sheets[target_sheet].range(target_cell).value
                capture_print(f"📌 Target after aggressive calc: {final_check:.6f}")
                
                if abs(final_check - original_forecast_value) > 1e-10:
                    base_target_value = final_check
                    capture_print("✅ Aggressive calculation worked!")
                else:
                    capture_print("⚠️ Target still unchanged - there may be an issue with the model or inputs")
                    
            except Exception as e:
                capture_print(f"⚠️ Error during aggressive calculation: {e}")

        # === STEP 11: Group baseline inputs (COMPLETE Jupyter logic) ===
        def extract_row_from_cell(cell):
            """Extract row number from Excel cell reference (e.g., 'A5' -> 5, 'BC123' -> 123)"""
            match = re.search(r'(\d+)', cell)
            return int(match.group(1)) if match else None

        def find_common_description_part(descriptions):
            """Find the common part among multiple descriptions"""
            if not descriptions:
                return ""
            if len(descriptions) == 1:
                return descriptions[0]
            
            # Find common prefix
            common_prefix = ""
            min_len = min(len(desc) for desc in descriptions)
            
            for i in range(min_len):
                chars = set(desc[i] for desc in descriptions)
                if len(chars) == 1:
                    common_prefix += list(chars)[0]
                else:
                    break
            
            # Find common suffix
            common_suffix = ""
            for i in range(1, min_len - len(common_prefix) + 1):
                chars = set(desc[-i] for desc in descriptions)
                if len(chars) == 1:
                    common_suffix = list(chars)[0] + common_suffix
                else:
                    break
            
            # Combine and clean up
            common_part = (common_prefix + common_suffix).strip()
            
            # If common part is too short or empty, use the first description
            if len(common_part) < 10:
                return descriptions[0]
            
            return common_part

        # Group by sheet and row instead of by enriched description
        grouped_inputs = defaultdict(list)
        for item in baseline_inputs:
            sheet = item["sheet"]
            cell = item["cell"]
            row = extract_row_from_cell(cell)
            
            if row is not None:
                # Group by sheet and row
                group_key = f"{sheet}_row_{row}"
                grouped_inputs[group_key].append(item)
            else:
                # Fallback: group by individual cell if row extraction fails
                group_key = f"{sheet}_{cell}"
                grouped_inputs[group_key].append(item)

        capture_print(f"\n📊 Grouped into {len(grouped_inputs)} categories (by sheet and row)")

        # === Create final grouped results with common descriptions ===
        final_grouped_inputs = {}
        for group_key, items in grouped_inputs.items():
            # Extract all descriptions from items in this group
            descriptions = [item['desc'] for item in items]
            
            # Find common description part
            common_desc = find_common_description_part(descriptions)
            
            # Create a more readable group name
            if len(items) == 1:
                final_key = common_desc
            else:
                sheet_name = items[0]['sheet']
                row_num = extract_row_from_cell(items[0]['cell'])
                final_key = f"{common_desc} (Row {row_num})"
            
            final_grouped_inputs[final_key] = items

        # === Show grouping summary ===
        capture_print("\n📋 Grouping Summary:")
        for i, (desc, items) in enumerate(final_grouped_inputs.items(), 1):
            cells_in_group = [f"{item['sheet']}!{item['cell']}" for item in items]
            cells_display = ", ".join(cells_in_group[:3])
            if len(cells_in_group) > 3:
                cells_display += f" (+{len(cells_in_group)-3} more)"
            capture_print(f"  {i:2d}. {desc:<50} | {len(items):2d} cells | {cells_display}")

        # === Diagnostic: Check for ungrouped inputs ===
        grouped_cells = {f"{item['sheet']}!{item['cell']}" for group in final_grouped_inputs.values() for item in group}
        input_cells = {f"{item['sheet']}!{item['cell']}" for item in baseline_inputs}
        missing_cells = input_cells - grouped_cells

        if missing_cells:
            capture_print(f"\n⚠️ {len(missing_cells)} baseline input cells were not grouped:")
            for cell in sorted(list(missing_cells)[:5]):  # Show first 5
                capture_print(f"  - {cell}")
            if len(missing_cells) > 5:
                capture_print(f"  ... and {len(missing_cells) - 5} more")
        else:
            capture_print("\n✅ All baseline inputs were grouped correctly.")

        # === STEP 12: COMPLETE Attribution (change-back by group) - ALL GROUPS ===
        results = []
        total_groups = len(final_grouped_inputs)

        for group_idx, (group_desc, items) in enumerate(final_grouped_inputs.items(), 1):
            capture_print(f"\n🔄 [{group_idx}/{total_groups}] Changing back group: {group_desc}")

            target_before = wb.sheets[target_sheet].range(target_cell).value

            # Revert each cell in this group to its original forecast value
            for item in items:
                sheet, cell = item["sheet"], item["cell"]
                backup_key = f"{sheet}!{cell}"
                
                if backup_key in original_values_backup:
                    backup_data = original_values_backup[backup_key]
                    original_val = backup_data["value"]
                    original_formula = backup_data["formula"]
                    
                    try:
                        rng = wb.sheets[sheet].range(cell)
                        
                        # Check what type of original content this cell had
                        if original_formula and isinstance(original_formula, str) and original_formula.startswith('='):
                            # Cell originally had a formula like "=SUM(A1:A10)"
                            rng.formula = original_formula
                        else:
                            # Cell originally had just a number (no formula)
                            rng.formula = None  # Clear any temporary formula
                            rng.value = original_val
                            
                    except Exception as e:
                        capture_print(f"⚠️ Error reverting {sheet}!{cell}: {e}")

            # Calculate after reverting this group
            for j in range(2):  # Reduced attempts for speed
                try:
                    app.calculate()
                    time.sleep(1)
                    break
                except Exception as e:
                    capture_print(f"⚠️ Excel busy (attempt {j+1}): {e}")
                    time.sleep(2)

            new_target_value = wb.sheets[target_sheet].range(target_cell).value
            if new_target_value is not None:
                delta = new_target_value - base_target_value
                capture_print(f"   🎯 Target: {target_before:.4f} → {new_target_value:.4f} (Δ = {delta:+.4f})")
            else:
                delta = 0
                capture_print(f"   🎯 Target became None - using delta = 0")

            results.append({
                "description": group_desc,  # Use the common description for the row group
                "cells": [f"{item['sheet']}!{item['cell']}" for item in items],
                "delta": delta
            })

            # Restore cells in this group back to baseline
            for item in items:
                try:
                    rng = wb.sheets[item["sheet"]].range(item["cell"])
                    rng.value = item["change_to"]
                except Exception as e:
                    capture_print(f"⚠️ Error restoring {item['sheet']}!{item['cell']}: {e}")

        # === STEP 13: Normalize results (from Jupyter logic) ===
        total_delta = original_forecast_value - base_target_value

        capture_print(f"\n✅ Total groups analyzed: {len(results)}")
        capture_print(f"📉 Total Δ Target: {total_delta:+.4f}")

        for r in results:
            if abs(total_delta) > 1e-6:
                r["normalized_pct"] = (r["delta"] / total_delta) * 100
            else:
                r["normalized_pct"] = 0.0

        # === Add residual if needed ===
        sum_of_group_deltas = sum(r["delta"] for r in results)
        residual_delta = total_delta - sum_of_group_deltas

        if abs(residual_delta) > 1e-6:
            residual_pct = residual_delta / total_delta * 100 if abs(total_delta) > 1e-6 else 0
            results.append({
                "description": "Interrelated effects",
                "cells": [],
                "delta": residual_delta,
                "normalized_pct": residual_pct
            })
            capture_print(f"\n⚠️ Added Interrelated effects: Δ={residual_delta:+.4f}, % of Total={residual_pct:+.2f}%")

        # === STEP 14: Final report (COMPLETE Jupyter format) ===
        capture_print("\n📊 Final Attribution Results (% of total Δ Target):")
        capture_print(f"{'Row-Based Group Description':<60} | {'% of Total':>10} | {'Δ Target':>12} | {'Cells'}")
        capture_print("=" * 120)

        for r in sorted(results, key=lambda x: abs(x['delta']), reverse=True):
            cell_list = ", ".join(r['cells'][:3])  # Show first 3 cells
            if len(r['cells']) > 3:
                cell_list += f" (+{len(r['cells'])-3} more)"
            
            capture_print(
                f"{r['description']:<60} | {r['normalized_pct']:+8.2f}% | "
                f"{r['delta']:+10.4f} | {cell_list}"
            )

        # === STEP 15: Check sum (from Jupyter logic) ===
        capture_print(f"\n✅ Sum of all deltas: {sum_of_group_deltas:+.4f}")
        capture_print(f"📊 Total change: {total_delta:+.4f}")
        capture_print(f"🎯 Residual: {residual_delta:+.4f}")

        capture_print(f"\n🎉 Attribution analysis complete!")

        # Create summary
        '''
        summary = {
            "total_delta": total_delta,
            "sum_deltas": sum_of_group_deltas,
            "residual": residual_delta,
            "original_value": original_forecast_value,
            "baseline_value": base_target_value,
            "groups_analyzed": len(results),
            "changes_applied": changes_applied,
            "changes_verified": changes_verified
        }
        '''
        # MOVE FILE WRITING HERE - BEFORE THE RETURN AND BEFORE THE FINALLY BLOCK
        capture_print("🔍 DEBUG: About to attempt file writing...")
        capture_print(f"🔍 DEBUG: attribution_output has {len(attribution_output)} lines")
        
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            result_file = os.path.join(desktop_path, "ARG_CONFIRMATION.txt")
            capture_print(f"\U0001f4dd Writing results to file: {result_file}")

            with open(result_file, "w", encoding="utf-8") as f:
                f.write("\U0001f3af ATTRIBUTION ANALYSIS RESULTS\n")
                f.write("="*60 + "\n")
                for line in attribution_output:
                    f.write(f"{line}\n")

            if os.path.exists(result_file):
                file_size = os.path.getsize(result_file)
                capture_print(f"\u2705 File written successfully ({file_size} bytes)")
            else:
                capture_print("\u274c File write failed: File does not exist after writing!")

        except Exception as file_error:
            capture_print(f"\u274c Error writing file: {file_error}")

    except Exception as e:
        error_msg = f"Error in attribution logic: {e}"
        capture_print(error_msg)
        capture_print(traceback.format_exc())

    finally:
        try:
            capture_print("RESTORING ALL ORIGINAL VALUES AND FORMULAS...")
            restored_count = 0
            formula_restored = 0
            value_restored = 0
            for cell_key, backup_data in original_values_backup.items():
                try:
                    sheet_name, cell_ref = cell_key.split('!')
                    rng = wb.sheets[sheet_name].range(cell_ref)
                    original_value = backup_data["value"]
                    original_formula = backup_data["formula"]
                    if original_formula and isinstance(original_formula, str) and original_formula.startswith('='):
                        rng.formula = original_formula
                        formula_restored += 1
                    else:
                        rng.formula = None
                        rng.value = original_value
                        value_restored += 1
                    restored_count += 1
                except Exception as e:
                    capture_print(f" Error restoring {cell_key}: {e}")
            capture_print(f"Restored {restored_count}/{len(original_values_backup)} cells")
            capture_print(f"Formulas restored: {formula_restored}")
            capture_print(f"Values restored: {value_restored}")
            app.calculate()
            final_target = wb.sheets[target_sheet].range(target_cell).value
            if final_target is not None:
                capture_print(f"\U0001f4cc Final target value: {final_target:.6f}")
            if app:
                app.screen_updating = True
                app.enable_events = True
                app.display_alerts = True
        except Exception as e:
            capture_print(f"Error in final restoration: {e}")

        capture_print("Cleanup completed - original file restored")

    return {
        "results": results if error_msg == False else [],
        "output": attribution_output,
        "summary": summary if error_msg == False else {},
        "error": error_msg,
        'original_target_value': original_forecast_value,
        'baseline_target_value': base_target_value
    }
