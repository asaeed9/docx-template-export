"""
Configuration model for list rendering behavior.

This module provides configuration options for controlling how bullet and
numbered lists are rendered in Word documents. The configuration allows
fine-tuning of visual appearance, indentation, and deep nesting behavior
while maintaining deterministic output.
"""
from dataclasses import dataclass
from typing import Tuple, Optional

@dataclass(frozen=True)
class ListRenderConfig:
    """
    Configuration for list rendering behavior in Word documents.
    
    This dataclass controls how bullet and numbered lists are rendered,
    including indentation, glyph selection, and deep nesting behavior.
    The configuration is immutable (frozen) to ensure deterministic behavior.
    
    The library uses manual list rendering (no Word numbering XML) to ensure
    cross-platform stability and deterministic output. All list formatting
    is controlled through paragraph indentation and manual glyph/number insertion.
    
    Attributes:
        max_visual_depth: Maximum visual nesting depth for lists. Beyond this
            depth, the deep_bullet_strategy determines how nesting is handled.
            Default: 20. Set to None for unlimited depth (not recommended
            due to Word layout constraints).
            
        indent_inches_per_level: Amount of left indentation (in inches) added
            for each nesting level. Level 1 has 0 indent, level 2 has this amount,
            level 3 has 2x this amount, etc. Default: 0.25 inches.
            
        hanging_indent_inches: Amount of negative (hanging) indent applied to
            list item paragraphs. This creates space for the bullet/number glyph
            while keeping the text aligned. Default: 0.25 inches (negative).
            
        paragraph_spacing_before_pt: Spacing (in points) added before each list
            item paragraph. This controls vertical spacing between list items.
            Default: 6.0 points.
            
        bullet_glyphs: Tuple of bullet glyph characters used for different
            nesting levels. The glyphs cycle or clamp based on max_visual_depth
            and deep_bullet_strategy. Default: ("•", "◦", "‣") for levels 1, 2, 3+.
            
        deep_bullet_strategy: Strategy for handling nesting beyond max_visual_depth.
            Options:
            - "clamp_last": Use the last glyph in bullet_glyphs for all deep levels
            - "cycle": Cycle through bullet_glyphs repeatedly
            - "textual": Use textual indicators like "---" for very deep levels
            Default: "clamp_last" (most stable and predictable).
    
    Example:
        ```python
        # Default configuration
        config = ListRenderConfig()
        
        # Custom configuration for deeper nesting
        config = ListRenderConfig(
            max_visual_depth=10,
            indent_inches_per_level=0.3,
            bullet_glyphs=("•", "◦", "‣", "▪"),
            deep_bullet_strategy="cycle"
        )
        ```
    
    Note:
        This configuration affects both bullet lists and numbered lists.
        Numbered lists use hierarchical numbering (1., 1.1., 1.1.1.) regardless
        of depth, but indentation follows the same rules as bullet lists.
    """
    max_visual_depth: Optional[int] = 20
    indent_inches_per_level: float = 0.25
    hanging_indent_inches: float = 0.25
    paragraph_spacing_before_pt: float = 6.0

    bullet_glyphs: Tuple[str, ...] = ("•", "◦", "‣")
    deep_bullet_strategy: str = "clamp_last"  # cycle | textual | clamp_last
