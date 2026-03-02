"""Utility functions for the birthday bot."""

import re
from typing import Tuple, Union, Optional, List
from datetime import datetime


def format_vietnamese_name(full_name: str, format_style: str = "short") -> str:
    """
    Format Vietnamese full name.
    
    Vietnamese names are: [Family Name] [Middle Names...] [Given Name]
    The last word is typically the personal/given name
    
    Args:
        full_name: Full Vietnamese name
        format_style: 
            - "short": "LAST_NAME GIVEN_NAME" (e.g., "TRANG NGUYỄN")
            - "full": Return as-is
            
    Returns:
        Formatted name
    """
    if not full_name or format_style == "full":
        return full_name
    
    parts = full_name.strip().split()
    if not parts:
        return full_name
    
    if format_style == "short" and len(parts) >= 2:
        # Take last word as given name, first word as family name
        given_name = parts[-1]  # Last word
        family_name = parts[0]  # First word
        return f"{given_name.title()} {family_name.title()}".upper()
    
    return full_name


def apply_text_transform(text: str, transform: str) -> str:
    """
    Apply text transformation.
    
    Supported transforms:
    - "vietnamese_name_short": Format Vietnamese name as "LAST FIRST"
    - "upper": Convert to UPPERCASE
    - "lower": Convert to lowercase
    - "title": Convert to Title Case
    - "none": No transform
    
    Args:
        text: Input text
        transform: Transform type
        
    Returns:
        Transformed text
    """
    if not transform or transform == "none":
        return text
    
    if transform == "vietnamese_name_short":
        return format_vietnamese_name(text, format_style="short")
    elif transform == "upper":
        return text.upper()
    elif transform == "lower":
        return text.lower()
    elif transform == "title":
        return text.title()
    else:
        return text



def parse_color(color_str: str) -> Tuple[int, int, int, int]:
    """
    Parse color string to RGBA tuple.
    
    Supports:
    - #RRGGBB hex format
    - rgba(r, g, b, a) where a is 0.0-1.0
    - rgb(r, g, b)
    
    Args:
        color_str: Color string
        
    Returns:
        Tuple of (R, G, B, A) where A is 0-255
        
    Raises:
        ValueError: If color format is invalid
    """
    color_str = color_str.strip()
    
    # Handle #RRGGBB
    if color_str.startswith('#'):
        hex_str = color_str.lstrip('#')
        if len(hex_str) == 6:
            try:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                return (r, g, b, 255)
            except ValueError:
                raise ValueError(f"Invalid hex color: {color_str}")
        else:
            raise ValueError(f"Hex color must be 6 digits: {color_str}")
    
    # Handle rgba(r, g, b, a)
    rgba_match = re.match(r'rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)', color_str)
    if rgba_match:
        r = int(rgba_match.group(1))
        g = int(rgba_match.group(2))
        b = int(rgba_match.group(3))
        a = float(rgba_match.group(4))
        if not (0 <= a <= 1):
            raise ValueError(f"Alpha must be 0.0-1.0: {a}")
        a = int(round(a * 255))
        return (r, g, b, a)
    
    # Handle rgb(r, g, b)
    rgb_match = re.match(r'rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', color_str)
    if rgb_match:
        r = int(rgb_match.group(1))
        g = int(rgb_match.group(2))
        b = int(rgb_match.group(3))
        return (r, g, b, 255)
    
    raise ValueError(f"Unsupported color format: {color_str}")


def wrap_text(text: str, max_width: int, font_obj) -> List[str]:
    """
    Wrap text to fit within max_width using PIL font.
    Breaks at word boundaries. Long words that exceed max_width are kept on their own line.
    
    Args:
        text: Text to wrap
        max_width: Maximum width in pixels
        font_obj: PIL Font object
        
    Returns:
        List of lines
    """
    if not text:
        return []
    
    # Try to use getbbox if available (Pillow >= 8.0)
    def get_text_width(s: str) -> int:
        try:
            bbox = font_obj.getbbox(s)
            return bbox[2] - bbox[0]
        except (AttributeError, TypeError):
            # Fallback for older Pillow versions
            return font_obj.getsize(s)[0]
    
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        test_line = ' '.join(current_line + [word])
        word_width = get_text_width(word)
        
        if get_text_width(test_line) <= max_width:
            # Word fits on current line
            current_line.append(word)
        elif word_width <= max_width:
            # Word fits on its own line
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
        else:
            # Word is too long - needs character-level breaking
            # First, finalize current line
            if current_line:
                lines.append(' '.join(current_line))
                current_line = []
            
            # Break long word into chunks that fit max_width
            chunk = ""
            for char in word:
                test_chunk = chunk + char
                if get_text_width(test_chunk) <= max_width:
                    chunk = test_chunk
                else:
                    if chunk:
                        lines.append(chunk)
                    chunk = char
            
            if chunk:
                current_line = [chunk]
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines


def format_placeholder(text: str, data: dict) -> str:
    """
    Format text with placeholders.
    
    Supports:
    - {field_name}: Direct field replacement
    - {field_name:%format}: Datetime formatting (e.g., {birthday:%d %b %Y})
    
    Args:
        text: Text with placeholders
        data: Dictionary of data
        
    Returns:
        Formatted text
    """
    # Handle datetime format placeholders like {birthday:%d %b %Y}
    def replace_datetime(match):
        field_name = match.group(1)
        fmt = match.group(2) if match.group(2) else None
        
        if field_name not in data:
            return match.group(0)  # Return original if field missing
        
        value = data[field_name]
        if fmt and isinstance(value, (datetime, type(None))):
            if value is None:
                return ""
            if isinstance(value, datetime):
                return value.strftime(fmt)
        
        return str(value) if fmt else str(data.get(field_name, match.group(0)))
    
    # Replace {field:%format} patterns
    text = re.sub(r'\{(\w+):([^}]*)\}', replace_datetime, text)
    
    # Replace remaining {field} patterns
    for key, value in data.items():
        if isinstance(value, datetime):
            text = text.replace(f"{{{key}}}", value.strftime("%Y-%m-%d"))
        else:
            text = text.replace(f"{{{key}}}", str(value))
    
    return text


def calculate_shrink_to_fit(
    text: str,
    max_width: int,
    max_height: int,
    font_obj,
    initial_size: int,
    min_size: int = 8,
    max_lines: Optional[int] = None,
    line_spacing: float = 1.2,
    font_path: Optional[str] = None
) -> Tuple[List[str], float]:
    """
    Iteratively shrink font size until text fits in bounds.
    
    Args:
        text: Text to fit
        max_width: Maximum width in pixels
        max_height: Maximum height in pixels
        font_obj: Base PIL Font object (will be recreated at different sizes)
        initial_size: Starting font size
        min_size: Minimum font size to try
        max_lines: Maximum number of lines allowed
        line_spacing: Line spacing multiplier
        font_path: Optional font file path for recreating fonts at different sizes
        
    Returns:
        Tuple of (wrapped_lines, final_font_size)
    """
    from PIL import ImageFont
    from pathlib import Path
    
    current_size = initial_size
    best_fit = None
    best_size = min_size
    
    while current_size >= min_size:
        try:
            # Try to create font at current size
            test_font = font_obj
            
            if font_path and Path(font_path).exists():
                try:
                    test_font = ImageFont.truetype(font_path, current_size)
                except Exception:
                    test_font = font_obj
            
            wrapped = wrap_text(text, max_width, test_font)
            
            # Check line count
            if max_lines and len(wrapped) > max_lines:
                current_size -= 1
                continue
            
            # Check height
            try:
                bbox = test_font.getbbox("Ag")
                line_height = (bbox[3] - bbox[1]) * line_spacing
            except (AttributeError, TypeError):
                line_height = test_font.getsize("Ag")[1] * line_spacing
            
            total_height = line_height * len(wrapped)
            
            if total_height <= max_height:
                best_fit = wrapped
                best_size = current_size
                break
        except Exception:
            pass
        
        current_size -= 1
    
    # If no fit found, still wrap the text at current font
    if best_fit is None:
        best_fit = wrap_text(text, max_width, font_obj)
        # If still too many lines, truncate to max_lines
        if max_lines and len(best_fit) > max_lines:
            best_fit = best_fit[:max_lines]
    
    return best_fit or [text], float(best_size)


def ensure_rgb_image(pil_image):
    """
    Ensure image is in RGB or RGBA mode.
    
    Args:
        pil_image: PIL Image object
        
    Returns:
        PIL Image in RGB or RGBA mode
    """
    if pil_image.mode in ('RGB', 'RGBA'):
        return pil_image
    elif pil_image.mode == 'P':
        # Palette mode
        return pil_image.convert('RGBA')
    else:
        return pil_image.convert('RGB')


def blend_text_with_stroke(
    draw_context,
    text: str,
    xy: Tuple[float, float],
    font_obj,
    fill_color: Tuple[int, int, int, int],
    stroke_enabled: bool = False,
    stroke_width: int = 2,
    stroke_color: Tuple[int, int, int, int] = (0, 0, 0, 255),
    anchor: Optional[str] = None
):
    """
    Draw text with optional stroke/outline.
    
    Args:
        draw_context: PIL ImageDraw context
        text: Text to draw
        xy: (x, y) position
        font_obj: PIL Font object
        fill_color: RGBA fill color
        stroke_enabled: Whether to draw stroke
        stroke_width: Stroke width in pixels
        stroke_color: RGBA stroke color
        anchor: Anchor point (e.g., 'lm', 'mm', 'rm')
    """
    if stroke_enabled and stroke_width > 0:
        draw_context.text(
            xy, text, font=font_obj, fill=stroke_color,
            stroke_width=stroke_width, anchor=anchor
        )
    
    draw_context.text(xy, text, font=font_obj, fill=fill_color, anchor=anchor)
