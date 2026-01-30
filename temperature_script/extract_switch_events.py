from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from typing import List
from highlight_registry import HighlightRegistry, SwitchSession, HighlightPoint

class ExtractSwitchEvents:
    def __init__(self, file_path: str, config: dict, registry: HighlightRegistry):
        self.file_path = file_path
        self.config = config
        self.registry = registry
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active

        self.digital_start_col: int = config.digitalStartCol
        self.pressure_col: int = config.pressureCol
        self.protected_headers = config.protectedHeaders

    def copy_fill(self, cell):
        """Copies a cell's fill style to a new cell."""
        return PatternFill(
            start_color=cell.fill.start_color.rgb,
            end_color=cell.fill.end_color.rgb,
            fill_type=cell.fill.fill_type
        ) if cell.fill else None

    def create_switch_events_sheet(self, sheet_name: str = "SwitchEvents"):
        # Checks if sheet 
        if sheet_name in self.wb.sheetnames:
            del self.wb[sheet_name]
        ws_out = self.wb.create_sheet(sheet_name)

        # Copies header rows from orignal sheet
        for col in range(1, self.ws.max_column + 1):
            source_cell = self.ws.cell(row=1, column=col)
            ws_out.cell(row=1, column=col).value = source_cell.value
            #ws_out.cell(row=1, column=col).fill = self.copy_fill(source_cell)
        
        # list of columns that have digital data (pressure switches)
        digital_cols = list(range(self.digital_start_col, self.ws.max_column + 1))

        """
        cell_values = {}
        for row in range(1, self.ws.max_row + 1):
            for col in range(1, self.ws.max_column + 1):
                cell_values[(row, col)] = self.ws.cell(row=row, column=col).value
        """
        # Gets all the rows from the registry
        # uses a set to remove duplicates
        all_rows_set = set()
        
        for col_sessions in self.registry.get_sessions_by_column().values():
            for session in col_sessions:
                all_rows_set.add(session.green_point.row)
                if session.yellow_point:
                    all_rows_set.add(session.yellow_point.row)
        
        all_rows = sorted(all_rows_set)

        out_row = 2
        event_rows = []           # Rows in current event block
        event_columns = set()     # Columns that have a GREEN in the block
        yellow_columns_in_event = set()  # Columns that have a YELLOW in the block

        # --- 3. Map (row, col) -> value for fast lookup ---
        row_col_values = {}
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
            for session in col_sessions:
                row_col_values[(session.green_point.row, col_idx)] = session.green_point.value
                if session.yellow_point:
                    row_col_values[(session.yellow_point.row, col_idx)] = session.yellow_point.value

        # --- 4. Process each row ---
        for row in all_rows:
            # Copy row to output
            for col_idx in range(1, self.ws.max_column + 1):
                new_cell = ws_out.cell(row=out_row, column=col_idx)
                orig_cell = self.ws.cell(row=row, column=col_idx)
                new_cell.fill = self.copy_fill(orig_cell)

                header_value = self.ws.cell(row=1, column=col_idx).value
                if header_value in self.protected_headers:
                    new_cell.value = orig_cell.value
                else:
                    # Only fill if this cell is recorded in the registry
                    new_cell.value = row_col_values.get((row, col_idx), None)

            out_row += 1
            event_rows.append(row)

            # Track GREEN/YELLOW columns in this event
            for col_idx in digital_cols:
                if (row, col_idx) in row_col_values:
                    val = row_col_values[(row, col_idx)]
                    session_type = None
                    # Determine session type by comparing with previous row
                    prev_val = self.ws.cell(row=row-1, column=col_idx).value if row > 1 else None
                    if prev_val == 1 and val == 0:
                        session_type = "GREEN"
                    elif prev_val == 0 and val == 1:
                        session_type = "YELLOW"

                    if session_type == "GREEN":
                        event_columns.add(col_idx)
                    elif session_type == "YELLOW" and col_idx in event_columns:
                        yellow_columns_in_event.add(col_idx)

            # --- Check if event block is complete ---
            if event_columns and event_columns == yellow_columns_in_event:
                # --- Differential Row ---
                diff_row = out_row
                ws_out.cell(row=diff_row, column=1, value="Differential")

                for col in digital_cols:
                    green_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows
                        if (r, col) in row_col_values and
                           row_col_values[(r, col)] == 0  # GREEN assumed as 0→1 transition
                    ]
                    yellow_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows
                        if (r, col) in row_col_values and
                           row_col_values[(r, col)] == 1  # YELLOW assumed as 1→0 transition
                    ]
                    if green_vals and yellow_vals:
                        ws_out.cell(row=diff_row, column=col, value=max(green_vals) - min(yellow_vals))

                out_row += 1  # Move past differential
                out_row += 1  # Blank row

                # Reset for next event
                event_rows.clear()
                event_columns.clear()
                yellow_columns_in_event.clear()

        # --- 5. Save workbook ---
        self.wb.save(self.file_path)
"""
from openpyxl import load_workbook
from typing import List
from openpyxl.styles import PatternFill

class ExtractSwitchEvents:
    def __init__(self, file_path: str, config: dict):
        self.file_path = file_path
        self.config = config
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active

        self.GREEN: str = config.highlightColors.green
        self.YELLOW: str = config.highlightColors.yellow
        self.digital_start_col: int = config.digitalStartCol
        self.pressure_col: int = config.pressureCol
        self.protected_headers = config.protectedHeaders

    def is_highlighted(self, cell) -> bool:
        return cell.fill and cell.fill.fill_type and cell.fill.start_color.rgb in (self.GREEN, self.YELLOW)

    def is_yellow(self, cell) -> bool:
        return cell.fill and cell.fill.fill_type and cell.fill.start_color.rgb == self.YELLOW

    # Creates a new colour fill with the same properties, gimick to transfer styles
    def copy_fill(self, cell):
        return PatternFill(
            start_color=cell.fill.start_color.rgb,
            end_color=cell.fill.end_color.rgb,
            fill_type=cell.fill.fill_type
        ) if cell.fill else None

    def create_switch_events_sheet(
        self,
        green_rows: List[int],
        yellow_rows: List[int],
        sheet_name: str = "SwitchEvents"
    ):
        # Deletes duplicate if it already exists
        if sheet_name in self.wb.sheetnames:
            del self.wb[sheet_name]

        ws_out = self.wb.create_sheet(sheet_name)

        # Copy header row form the old sheet to the new one
        for col in range(1, self.ws.max_column + 1):
            ws_out.cell(row=1, column=col).value = self.ws.cell(row=1, column=col).value
            ws_out.cell(row=1, column=col).fill = self.copy_fill(self.ws.cell(row=1, column=col))

        all_rows = sorted(set(green_rows) | set(yellow_rows))
        out_row = 2
        digital_cols = list(range(self.digital_start_col, self.ws.max_column + 1))

        event_rows = []        # rows belonging to the current event block
        event_columns = set()  # digital columns that have had a green→yellow transition
        yellow_columns_in_event = set()
        

        for row in all_rows:
            # Copy row to output with blanking unhighlighted digital cells
            for col_idx in range(1, self.ws.max_column + 1):
                
                new_cell = ws_out.cell(row=out_row, column=col_idx)
                orig_cell = self.ws.cell(row=row, column=col_idx)
                new_cell.fill = self.copy_fill(orig_cell)

                header_value = self.ws.cell(row=1, column=col_idx).value
                
                # Keeps the columns in protected headers
                if header_value in self.protected_headers:
                
                    new_cell.value = orig_cell.value
                
                else:
                    # For other columns, blank if not highlighted
                    if orig_cell.fill and orig_cell.fill.fill_type and orig_cell.fill.start_color.rgb in (self.GREEN, self.YELLOW):
                        new_cell.value = orig_cell.value
                    else:
                        new_cell.value = None

            out_row += 1
            event_rows.append(row)

            # Track digital transitions
            for col in digital_cols:
                
                cell_fill = self.ws.cell(row=row, column=col).fill
                
                if not (cell_fill and cell_fill.fill_type):
                    continue
                
                rgb = cell_fill.start_color.rgb
                
                if rgb == self.GREEN:
                
                    event_columns.add(col)
                
                elif rgb == self.YELLOW and col in event_columns:
                
                    yellow_columns_in_event.add(col)

           # Check if the event block is complete
            if event_columns and event_columns == yellow_columns_in_event:
                # --- Insert one differential row for all digital columns ---
                differential_row = out_row
                ws_out.cell(row=differential_row, column=1, value="Differential")  # label first column

                for col in digital_cols:
                    
                    green_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows
                        if self.ws.cell(r, col).fill and self.ws.cell(r, col).fill.fill_type and self.ws.cell(r, col).fill.start_color.rgb == self.GREEN
                    ]
                    
                    yellow_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows
                        if self.ws.cell(r, col).fill and self.ws.cell(r, col).fill.fill_type and self.ws.cell(r, col).fill.start_color.rgb == self.YELLOW
                    ]
                    
                    if green_vals and yellow_vals:
                        
                        ws_out.cell(row=differential_row, column=col, value=max(green_vals) - min(yellow_vals))
                    
                    else:
                        ws_out.cell(row=differential_row, column=col, value=None)  # save blank if no data

                out_row += 1  # move past differential row

                # --- Insert one blank row after differential ---
                out_row += 1

                # Reset for next event
                event_rows.clear()
                event_columns.clear()
                yellow_columns_in_event.clear()


        self.wb.save(self.file_path)

        """