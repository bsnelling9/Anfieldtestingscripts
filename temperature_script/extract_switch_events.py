import pandas as pd
from openpyxl import load_workbook
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

    def create_switch_events_sheet(self, sheet_name: str = "SwitchEvents"):
        """
        Builds a SwitchEvents sheet by iterating switch SESSIONS directly
        instead of reconstructing events from rows.
        """
        # -----------------------------
        # 1. Headers & digital columns
        # -----------------------------
        all_data = list(self.ws.values)  # list of tuples
        headers = all_data[0]
        
        sessions_by_col = self.registry.get_sessions_by_column()

        digital_cols = [col for col in range(self.digital_start_col, self.ws.max_column + 1)
                if col in sessions_by_col]
        
        # Output structure for pandas
        data = {header: [] for header in headers}
       
        # -----------------------------
        # 2. Determine number of complete events
        #    (event index = same index across columns)
        # -----------------------------
        complete_event_count = min(
            len([s for s in sessions if s.is_complete])
            for sessions in sessions_by_col.values()
        )

        # -----------------------------
        # 3. Process each event (SESSION-DRIVEN)
        # -----------------------------
        for event_idx in range(complete_event_count):

            current_sessions = {
                col: sessions[event_idx]
                for col, sessions in sessions_by_col.items()
            }

            event_rows = sorted(
                {s.open_point.row for s in current_sessions.values()} |
                {s.close_point.row for s in current_sessions.values()}
            )

            """for sessions in sessions_by_col.values():
                session = sessions[event_idx]
                event_rows.add(session.open_point.row)
                event_rows.add(session.close_point.row)
            """
            # ---------------------------------
            # 3a. Write event rows
            # ---------------------------------
            for row in event_rows:
                
                row_values = []

                for col_idx, header in enumerate(headers, start=1):

                    if header in self.protected_headers:
                        value = all_data[row - 1][col_idx - 1]

                    elif col_idx in digital_cols:
                        # Only show 0/1 at actual switch points
                        value = self.registry.lookup.get((row, col_idx))
                       
                    else:
                        value = None

                    row_values.append(value)

                for idx, val in enumerate(row_values):
                    data[headers[idx]].append(val)

            # ---------------------------------
            # 3b. Differential row
            # ---------------------------------
            diff_row = [None] * len(headers)
            diff_row[0] = "Differential"

            for col_idx in digital_cols:
                session = sessions_by_col[col_idx][event_idx]

                open_pressure = self.ws.cell(
                    row=session.open_point.row,
                    column=self.pressure_col
                ).value

                close_pressure = self.ws.cell(
                    row=session.close_point.row,
                    column=self.pressure_col
                ).value

                if open_pressure is not None and close_pressure is not None:
                    diff_row[col_idx - 1] = open_pressure - close_pressure

            for idx, val in enumerate(diff_row):
                data[headers[idx]].append(val)

            # ---------------------------------
            # 3c. Blank separator row
            # ---------------------------------
            for header in headers:
                data[header].append(None)

        # -----------------------------
        # 4. Write to Excel
        # -----------------------------
        df = pd.DataFrame(data)

        with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="a") as writer:
            if sheet_name in writer.book.sheetnames:
                idx = writer.book.sheetnames.index(sheet_name)
                writer.book.remove(writer.book.worksheets[idx])

            df.to_excel(writer, sheet_name=sheet_name, index=False)

"""val = self.registry.lookup.get((row, col_idx))
                        for session in sessions_by_col.get(col_idx, []):
                            if session.open_point.row == row:
                                value = session.open_point.value
                                break
                            if session.close_point and session.close_point.row == row:
                                value = session.close_point.value
                                break"""
