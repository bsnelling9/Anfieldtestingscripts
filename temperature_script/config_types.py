from dataclasses import dataclass, field
from typing import List, Optional


@dataclass(frozen=True)
class HighlightColors:
    green: str
    yellow: str


@dataclass(frozen=True)
class TMAConfig:
    model: str
    transducer: str
    pressure: int

    outputType: str
    resistor: Optional[float]
    outputMin: float
    outputMax: float

    digitalStartCol: int
    daqMetaData: int
    pressureCol: int

    highlightColors: HighlightColors
    protectedHeaders: List[str]

    transducers: dict = field(default_factory=dict)