from dataclasses import dataclass
from typing import List


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
    resistor: float
    outputMin: float
    outputMax: float

    digitalStartCol: int
    daqMetaData: int
    pressureCol: int

    highlightColors: HighlightColors
    protectedHeaders: List[str]
