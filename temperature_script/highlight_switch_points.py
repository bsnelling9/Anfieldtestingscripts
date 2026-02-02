from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from highlight_registry import HighlightRegistry, HighlightPoint

class HighlightSwitchPoints:
    def __init__(self, file_path: str, config: dict, registry: HighlightRegistry):
        self.file_path = file_path
        self.config = config
        self.registry = registry  # <--- NEW: The shared data store
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active
        
        # Create fills from config
        self.green_fill = PatternFill(
            start_color=config.highlightColors.green,
            end_color=config.highlightColors.green,
            fill_type="solid"
        )

        self.yellow_fill = PatternFill(
            start_color=config.highlightColors.yellow,
            end_color=config.highlightColors.yellow,
            fill_type="solid"
        )
        self.protected_headers = config.protectedHeaders

    def highlight_switch_points(self):
        """
        Detects switch is Open 1→0 as green, and closed 0→1 as yellow.
        Highlights cells, and records the state in registry.
        """
        # Identify digital columns
        digital_cols = [
            col for col in range(1, self.ws.max_column + 1)
            if self.ws.cell(row=1, column=col).value not in self.protected_headers
        ]

        for col in digital_cols:
            header_name = self.ws.cell(row=1, column=col).value
            prev_val = self.ws.cell(row=2, column=col).value

            for row in range(3, self.ws.max_row + 1):
                cell = self.ws.cell(row=row, column=col)

                if prev_val == 1 and cell.value == 0:
                    cell.fill = self.green_fill
                    self.registry.add_point(HighlightPoint(row, col, True, header_name, cell.value))

                elif prev_val == 0 and cell.value == 1:
                    cell.fill = self.yellow_fill
                    self.registry.add_point(HighlightPoint(row, col, False, header_name, cell.value))

                prev_val = cell.value

        self.wb.save(self.file_path)