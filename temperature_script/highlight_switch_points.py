from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from typing import List, Tuple

class HighlightSwitchPoints:
    def __init__(self, file_path: str, config: dict):
        self.file_path = file_path
        self.config = config
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active
        
        # Create fills from JSON
        self.green_fill = PatternFill(
            start_color=config["highlightColors"]["green"],
            end_color=config["highlightColors"]["green"],
            fill_type="solid"
        )

        self.yellow_fill = PatternFill(
            start_color=config["highlightColors"]["yellow"],
            end_color=config["highlightColors"]["yellow"],
            fill_type="solid"
        )
        self.protected_headers = config["protectedHeaders"]
    
    # Used as it could be a standalone function, but it makes sense to be apart of this class
    @staticmethod
    def is_yellow(self, cell) -> bool:
        return cell.fill and cell.fill.fill_type and cell.fill.start_color.rgb == self.yellow_fill.start_color.rgb


    def highlight_switch_points(self) -> Tuple[List[int], List[int]]:
        """
        Highlights 1→0 as green, 0→1 as yellow.
        Returns list of green and yellow row indices.
        """
        # Identify digital columns
        digital_cols = []

        for col in range(1, self.ws.max_column + 1):

            header = self.ws.cell(row=1, column=col).value

            if header and header not in self.protected_headers:

                digital_cols.append(col)

        green_rows = []
        yellow_rows = []

        for col in digital_cols:
            
            prev_val = self.ws.cell(row=2, column=col).value

            for row in range(3, self.ws.max_row + 1):

                cell = self.ws.cell(row=row, column=col)

                if prev_val == 1 and cell.value == 0:

                    cell.fill =  self.green_fill
                    green_rows.append(row)

                elif prev_val == 0 and cell.value == 1:

                    cell.fill = self.yellow_fill
                    yellow_rows.append(row)

                prev_val = cell.value

        self.wb.save(self.file_path)

        return green_rows, yellow_rows
