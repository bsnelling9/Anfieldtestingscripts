from openpyxl import load_workbook
from typing import List
from openpyxl.styles import PatternFill

class ExtractSwitchEvents:
    def __init__(self, file_path: str, config: dict):
        self.file_path = file_path
        self.config = config
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active

        self.GREEN: str = config["highlightColors"]["green"]
        self.YELLOW: str = config["highlightColors"]["yellow"]
        self.digital_start_col: int = config["digitalStartCol"]
        self.pressure_col: int = config["pressureCol"]
        self.protected_headers = config["protectedHeaders"]


    def is_highlighted(self, cell) -> bool:
        return cell.fill and cell.fill.fill_type and cell.fill.start_color.rgb in (self.GREEN, self.YELLOW)

    def is_yellow(self, cell) -> bool:
        return cell.fill and cell.fill.fill_type and cell.fill.start_color.rgb == self.YELLOW

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
        if sheet_name in self.wb.sheetnames:
            del self.wb[sheet_name]

        ws_out = self.wb.create_sheet(sheet_name)

        # Copy header with fill
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
                if header_value in self.protected_headers:
                    new_cell.value = orig_cell.value  # keep protected columns
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
                        ws_out.cell(row=differential_row, column=col, value=None)  # safe blank if no data

                out_row += 1  # move past differential row

                # --- Insert one blank row after differential ---
                out_row += 1

                # Reset for next event
                event_rows.clear()
                event_columns.clear()
                yellow_columns_in_event.clear()


        self.wb.save(self.file_path)


