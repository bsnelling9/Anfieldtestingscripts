from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from highlight_registry import HighlightRegistry

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

        
        cell_values = {}
        cell_fills = {}
        for row in range(1, self.ws.max_row + 1):
            
            for col in range(1, self.ws.max_column + 1):
                
                cell = self.ws.cell(row=row, column=col)
                cell_values[(row, col)] = cell.value
                cell_fills[(row, col)] = self.copy_fill(cell)
        
        # Gets all the rows from the registry
        # uses a set to remove duplicates
        all_rows_set = set()
        
        for col_sessions in self.registry.get_sessions_by_column().values():
            
            for session in col_sessions:
                
                all_rows_set.add(session.green_point.row)
                
                if session.yellow_point:
                    all_rows_set.add(session.yellow_point.row)
        
        all_rows = sorted(all_rows_set)

        # Pre compute the colours lookup into a dictionary to improve speed
        green_rows_per_col = {col: set() for col in digital_cols}
        yellow_rows_per_col = {col: set() for col in digital_cols}
        
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
            
            for session in col_sessions:
                green_rows_per_col[col_idx].add(session.green_point.row)
                
                if session.yellow_point:
                    yellow_rows_per_col[col_idx].add(session.yellow_point.row)

        # --- 3. Map (row, col) -> value for fast lookup ---
        row_col_values = {}
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
            for session in col_sessions:
                row_col_values[(session.green_point.row, col_idx)] = session.green_point.value
                if session.yellow_point:
                    row_col_values[(session.yellow_point.row, col_idx)] = session.yellow_point.value

        out_row = 2
        event_rows = []           # Rows in current event block
        event_columns = set()     # Columns that have a GREEN in the block
        yellow_columns_in_event = set()  # Columns that have a YELLOW in the block
        
        # --- 4. Process each row ---
        for row in all_rows:
            # Copy row to output
            for col_idx in range(1, self.ws.max_column + 1):
                new_cell = ws_out.cell(row=out_row, column=col_idx)
                new_cell.fill = cell_fills[(row, col_idx)]
                header_value = self.ws.cell(row=1, column=col_idx).value
                
                if header_value in self.protected_headers:

                    new_cell.value = cell_values[(row, col_idx)]
                else:
                    # Only fill if this cell is recorded in the registry
                    new_cell.value = row_col_values.get((row, col_idx), None)

            out_row += 1
            event_rows.append(row)

            # Track GREEN/YELLOW columns in this event
            for col_idx in digital_cols:
                if row in green_rows_per_col[col_idx]:
                    event_columns.add(col_idx)
                elif row in yellow_rows_per_col[col_idx] and col_idx in event_columns:
                    yellow_columns_in_event.add(col_idx)
            
            # --- Check if event block is complete ---
            if event_columns and event_columns == yellow_columns_in_event:
                
                diff_row = out_row
                ws_out.cell(row=diff_row, column=1, value="Differential")

                for col in digital_cols:
                    green_vals = [
                        cell_values[(r, self.pressure_col)]
                        for r in event_rows if r in green_rows_per_col[col]
                    ]
                    yellow_vals = [
                        cell_values[(r, self.pressure_col)]
                        for r in event_rows if r in yellow_rows_per_col[col]
                    ]
                    
                    if green_vals and yellow_vals:
                        ws_out.cell(row=diff_row, column=col, value=max(green_vals) - min(yellow_vals))

                out_row += 2  # Move past differential and blank row
                
                event_rows.clear()
                event_columns.clear()
                yellow_columns_in_event.clear()

        self.wb.save(self.file_path)