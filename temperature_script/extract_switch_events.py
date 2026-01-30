import pandas as pd
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

        # Columns where digital data starts, column containing pressure, and protected headers
        self.digital_start_col: int = config.digitalStartCol
        self.pressure_col: int = config.pressureCol
        self.protected_headers = config.protectedHeaders

    """
    This does not work now because the code is now using pandas
    def copy_fill(self, cell):
        return PatternFill(
            start_color=cell.fill.start_color.rgb,
            end_color=cell.fill.end_color.rgb,
            fill_type=cell.fill.fill_type
        ) if cell.fill else None
    """
    def create_switch_events_sheet(self, sheet_name: str = "SwitchEvents"):

        # Stores all the rows where the switch opens or closes
        all_rows_set = set()
        
        for col_sessions in self.registry.get_sessions_by_column().values():
            for session in col_sessions:
                all_rows_set.add(session.open_point.row)
                if session.close_point:
                    all_rows_set.add(session.close_point.row)
        all_rows = sorted(all_rows_set)

        # Stores the headers in an array, maybe could use a set as they're unique
        headers = [self.ws.cell(row=1, column=col).value for col in range(1, self.ws.max_column + 1)]

        # Initialize a dictionary of lists for each column
        data = {header: [] for header in headers}

        # this is a dictionary to quickly get a switch value (1/0) at a specified (row,col), which is a key
        # (97, 5): 0 where 97 is a row and 5 is a column (switch, TMA 7.0)
        # basically avoids the .cell() call to look for the cell everytime
        row_col_values = {}
        
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
            row_col_values[(session.open_point.row, col_idx)] = session.open_point.value
            if session.close_point:
                row_col_values[(session.close_point.row, col_idx)] = session.close_point.value        
        
        # This tracks open and closed rows
        digital_cols = list(range(self.digital_start_col, self.ws.max_column + 1))
        
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
                
                for session in col_sessions:
                    row_col_values[(session.open_point.row, col_idx)] = session.open_point.value
                    if session.close_point:
                        row_col_values[(session.close_point.row, col_idx)] = session.close_point.value

        # --- 3. Track open and closed rows for differential logic ---
        open_rows_per_col = {col: set() for col in digital_cols}
        close_rows_per_col = {col: set() for col in digital_cols}
        
        for col_idx, col_sessions in self.registry.get_sessions_by_column().items():
            
            for session in col_sessions:
                open_rows_per_col[col_idx].add(session.open_point.row)
                if session.close_point:
                    close_rows_per_col[col_idx].add(session.close_point.row)
        # --- 5. Fill data dictionary row by row ---

        event_rows = []
        event_columns = set()
        closed_columns_in_event = set()

        for row in all_rows:
            row_data = []
            
            for col_idx, header in enumerate(headers, start=1):
                if header in self.protected_headers:
                    value = self.ws.cell(row=row, column=col_idx).value
                elif col_idx in digital_cols:
                    # Only show 0/1 for switch points, blank otherwise
                    value = row_col_values.get((row, col_idx), None)
                else:
                    value = None
                row_data.append(value)

            for idx, val in enumerate(row_data):
                data[headers[idx]].append(val)

            # --- 5. Event tracking for differential ---
            event_rows.append(row)
           
            for col_idx in digital_cols:
                if row in open_rows_per_col[col_idx]:
                    event_columns.add(col_idx)
                elif row in close_rows_per_col[col_idx] and col_idx in event_columns:
                    closed_columns_in_event.add(col_idx)

            if event_columns and event_columns == closed_columns_in_event:
                diff_row_data = [None] * len(headers)
                diff_row_data[0] = "Differential"
                
                for col_idx in digital_cols:
                    open_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows if r in open_rows_per_col[col_idx]
                    ]
                    closed_vals = [
                        self.ws.cell(r, self.pressure_col).value
                        for r in event_rows if r in close_rows_per_col[col_idx]
                    ]
                    if open_vals and closed_vals:
                        diff_row_data[col_idx - 1] = max(open_vals) - min(closed_vals)

                for idx, val in enumerate(diff_row_data):
                    data[headers[idx]].append(val)

                # Blank row after differential
                for header in headers:
                    data[header].append(None)

                # Reset for next event block
                event_rows.clear()
                event_columns.clear()
                closed_columns_in_event.clear()


        # --- 6. Convert to DataFrame and write to Excel ---
        df = pd.DataFrame(data)
        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a') as writer:
            if sheet_name in writer.book.sheetnames:
                idx = writer.book.sheetnames.index(sheet_name)
                std = writer.book.worksheets[idx]
                writer.book.remove(std)
            df.to_excel(writer, sheet_name=sheet_name, index=False)