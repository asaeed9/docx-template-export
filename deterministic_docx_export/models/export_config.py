from dataclasses import dataclass
from typing import Tuple, Optional

@dataclass(frozen=True)
class ListRenderConfig:
    max_visual_depth: Optional[int] = 3
    indent_inches_per_level: float = 0.25
    hanging_indent_inches: float = 0.25

    bullet_glyphs: Tuple[str, ...] = ("•", "◦", "‣")
    deep_bullet_strategy: str = "clamp_last"  # cycle | textual | clamp_last
