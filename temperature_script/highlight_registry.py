from dataclasses import dataclass
from typing import Optional

@dataclass
class HighlightPoint:
    row: int
    col: int
    is_open: bool
    header: str
    value: int
    pressure: float

@dataclass
class SwitchSession:
    open_point: HighlightPoint
    close_point: Optional[HighlightPoint] = None

    @property
    def is_complete(self) -> bool:
        return self.close_point is not None

class HighlightRegistry:
    def __init__(self):
        self.sessions: dict[int, list[SwitchSession]] = {}
        
        self.active: dict[int, SwitchSession] = {}
        self.lookup: dict[tuple[int,int], HighlightPoint] = {}

    def record_event(self, point: HighlightPoint):
               
        key = (point.row, point.col)
        
        self.lookup[key] = point

        col = point.col
        if point.is_open:
            session = SwitchSession(open_point=point)
            self.active[col] = session
            self.sessions.setdefault(col, []).append(session)
        
        else:
            session = self.active.get(col)
            if session is not None:
                session.close_point = point
                del self.active[col]

    def get_sessions_by_column(self):
        return self.sessions