from dataclasses import dataclass
from typing import Optional

@dataclass
class HighlightPoint:
    row: int
    col: int
    is_open: bool
    header: str
    value: int

@dataclass
class SwitchSession:
    open_point: HighlightPoint
    close_point: Optional[HighlightPoint] = None

    @property
    def is_complete(self) -> bool:
        return self.close_point is not None

class HighlightRegistry:
    def __init__(self):
        self.columns: dict[int, list[SwitchSession]] = {}

    def add_point(self, point: HighlightPoint):
        col = point.col
        if col not in self.columns:
            self.columns[col] = []

        if point.is_open:
            self.columns[col].append(SwitchSession(open_point=point))
        else:
            # Attach CLOSE to most recent incomplete OPEN
            for session in reversed(self.columns[col]):
                if session.close_point is None:
                    session.close_point = point
                    break

    def get_sessions_by_column(self):
        return self.columns