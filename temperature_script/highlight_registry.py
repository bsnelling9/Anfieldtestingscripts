from dataclasses import dataclass
from typing import Optional, Any

@dataclass
class HighlightPoint:
    row: int
    col: int
    color_type: str  # "GREEN" or "YELLOW"
    header: str
    value: Any

@dataclass
class SwitchSession:
    green_point: HighlightPoint
    yellow_point: Optional[HighlightPoint] = None

    @property
    def is_complete(self) -> bool:
        return self.yellow_point is not None

class HighlightRegistry:
    def __init__(self):
        self.columns: dict[int, list[SwitchSession]] = {}

    def add_point(self, point: HighlightPoint):
        col = point.col
        if col not in self.columns:
            self.columns[col] = []

        if point.color_type == "GREEN":
            session = SwitchSession(green_point=point)  # Only green_point required now
            self.columns[col].append(session)
        elif point.color_type == "YELLOW":
            # Assign to last incomplete GREEN session in this column
            for session in reversed(self.columns[col]):
                if not session.yellow_point:
                    session.yellow_point = point
                    break

    def get_sessions_by_column(self):
        return self.columns