"""Card rendering engine."""

import logging
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import warnings

from . import utils

logger = logging.getLogger(__name__)


class RenderError(Exception):
    """Card rendering error."""
    pass


def load_font(font_path: str, size: int) -> ImageFont.FreeTypeFont:
    """
    Load TrueType font.
    
    Tries multiple strategies:
    1. Load from explicit file path
    2. Load from system fonts (Windows/Linux/macOS)
    3. Fall back to PIL default
    
    Args:
        font_path: Path to .ttf file or font name
        size: Font size in points
        
    Returns:
        PIL Font object
        
    Raises:
        RenderError: If font cannot be loaded
    """
    import sys
    from pathlib import Path
    
    font_file = Path(font_path)
    
    # Strategy 1: Try explicit file path
    if font_file.exists():
        try:
            return ImageFont.truetype(font_path, size)
        except Exception as e:
            logger.warning(f"Failed to load font from path {font_path}: {e}")
    
    # Strategy 2: Try system fonts
    system_font_paths = []
    
    if sys.platform == 'win32':
        # Windows system fonts
        system_font_paths = [
            Path('C:/Windows/Fonts/segoeui.ttf'),  # Segoe UI
            Path('C:/Windows/Fonts/arial.ttf'),    # Arial
            Path('C:/Windows/Fonts/calibri.ttf'),  # Calibri
        ]
    elif sys.platform == 'darwin':
        # macOS system fonts
        system_font_paths = [
            Path('/Library/Fonts/Arial.ttf'),
            Path('/System/Library/Fonts/Helvetica.ttc'),
        ]
    else:
        # Linux system fonts
        system_font_paths = [
            Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),
            Path('/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'),
        ]
    
    for sys_font in system_font_paths:
        if sys_font.exists():
            try:
                return ImageFont.truetype(str(sys_font), size)
            except Exception:
                continue
    
    # Strategy 3: Use PIL default with warning
    logger.warning(f"Font file not found: {font_path}, using default font")
    try:
        # Try to load a larger default font for better rendering
        return ImageFont.load_default(size=size)
    except TypeError:
        # Fallback for older PIL versions
        return ImageFont.load_default()


def load_template(template_path: str) -> Image.Image:
    """
    Load template PNG image.
    
    Args:
        template_path: Path to template PNG
        
    Returns:
        PIL Image in RGBA mode
        
    Raises:
        RenderError: If template cannot be loaded
    """
    template_file = Path(template_path)
    
    if not template_file.exists():
        raise RenderError(f"Template file not found: {template_path}")
    
    try:
        img = Image.open(template_file)
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        return img
    except Exception as e:
        raise RenderError(f"Failed to load template: {e}")


def load_photo(photo_path: str) -> Image.Image:
    """
    Load member photo from local file or URL.
    
    Args:
        photo_path: Path to photo file or full URL
        
    Returns:
        PIL Image in RGB or RGBA mode
        
    Raises:
        RenderError: If photo cannot be loaded
    """
    import io
    import urllib.request
    
    # Check if it's a URL
    if photo_path.startswith('http://') or photo_path.startswith('https://'):
        # Download from URL
        try:
            logger.info(f"Downloading photo from URL: {photo_path}")
            with urllib.request.urlopen(photo_path, timeout=10) as response:
                img_data = response.read()
            
            img = Image.open(io.BytesIO(img_data))
            return utils.ensure_rgb_image(img)
        except urllib.error.HTTPError as e:
            raise RenderError(f"HTTP error downloading photo {e.code}: {photo_path}")
        except urllib.error.URLError as e:
            raise RenderError(f"URL error downloading photo: {photo_path} - {e.reason}")
        except Exception as e:
            raise RenderError(f"Failed to load photo from URL {photo_path}: {e}")
    else:
        # Load from local file
        photo_file = Path(photo_path)
        
        if not photo_file.exists():
            raise RenderError(f"Photo file not found: {photo_path}")
        
        try:
            img = Image.open(photo_file)
            return utils.ensure_rgb_image(img)
        except Exception as e:
            raise RenderError(f"Failed to load photo: {e}")


def crop_to_square(img: Image.Image, crop_type: str = 'center_square') -> Image.Image:
    """
    Crop image to square.
    
    Args:
        img: Source image
        crop_type: Crop type (currently only 'center_square' supported)
        
    Returns:
        Cropped square image
    """
    width, height = img.size
    size = min(width, height)
    
    left = (width - size) // 2
    top = (height - size) // 2
    right = left + size
    bottom = top + size
    
    return img.crop((left, top, right, bottom))


def create_circle_mask(
    width: int,
    height: int,
    feather: int = 0,
    inset: int = 0
) -> Image.Image:
    """
    Create circular alpha mask.
    
    Args:
        width: Mask width
        height: Mask height
        feather: Blur amount for soft edges (0-10)
        inset: Shrink circle inward (in pixels)
        
    Returns:
        L-mode image with circle mask
    """
    mask = Image.new('L', (width, height), 0)
    draw = ImageDraw.Draw(mask)
    
    # Draw white circle (opaque)
    bbox = [inset, inset, width - inset - 1, height - inset - 1]
    draw.ellipse(bbox, fill=255)
    
    # Apply feather (soft edges)
    if feather > 0:
        mask = mask.filter(ImageFilter.GaussianBlur(radius=feather))
    
    return mask


def paste_photo_with_mask(
    canvas: Image.Image,
    photo: Image.Image,
    placement: Dict[str, int],
    crop_type: str = 'center_square',
    mask_cfg: Optional[Dict[str, Any]] = None
) -> Image.Image:
    """
    Paste photo onto canvas with circle mask.
    
    Args:
        canvas: Target canvas (RGBA)
        photo: Photo image
        placement: {x, y, width, height}
        crop_type: How to crop photo
        mask_cfg: {type, feather, inset}
        
    Returns:
        Canvas with photo pasted
    """
    if mask_cfg is None:
        mask_cfg = {'type': 'circle', 'feather': 0, 'inset': 0}
    
    # Crop photo to square
    photo_square = crop_to_square(photo, crop_type)
    
    # Resize to placement size
    photo_resized = photo_square.resize(
        (placement['width'], placement['height']),
        Image.Resampling.LANCZOS
    )
    
    # Create mask if circle type
    if mask_cfg.get('type') == 'circle':
        mask = create_circle_mask(
            placement['width'],
            placement['height'],
            feather=mask_cfg.get('feather', 0),
            inset=mask_cfg.get('inset', 0)
        )
        photo_resized.putalpha(mask)
    
    # Paste onto canvas
    canvas.paste(
        photo_resized,
        (placement['x'], placement['y']),
        photo_resized if photo_resized.mode == 'RGBA' else None
    )
    
    return canvas


def render_text_layer(
    canvas: Image.Image,
    layer: Dict[str, Any],
    member: Dict[str, Any],
    fonts_map: Dict[str, str],
    colors_map: Dict[str, str],
    base_path: Path,
    card_config: Optional[Dict[str, Any]] = None
) -> Image.Image:
    """
    Render a single text layer onto canvas.
    
    Args:
        canvas: RGBA canvas
        layer: Text layer configuration
        member: Member data
        fonts_map: Font reference to path mapping
        colors_map: Color reference to definition mapping
        base_path: Base path for resources
        card_config: Card configuration (for random greetings)
        
    Returns:
        Canvas with text rendered
    """
    # Get font
    font_ref = layer.get('font_ref', 'default')
    font_size = layer.get('size', 24)
    font_path = fonts_map.get(font_ref)
    
    if font_ref in fonts_map:
        font = load_font(fonts_map[font_ref], font_size)
    else:
        logger.warning(f"Font not found: {font_ref}, using default")
        font = ImageFont.load_default()
        font_path = None
    
    # Get color
    color_ref = layer.get('color_ref', 'black')
    color_str = colors_map.get(color_ref, '#000000')
    color = utils.parse_color(color_str)
    
    # Get stroke config
    stroke_cfg = layer.get('stroke', {})
    stroke_enabled = stroke_cfg.get('enabled', False)
    stroke_width = stroke_cfg.get('width', 2)
    stroke_color_ref = stroke_cfg.get('color_ref', 'black')
    stroke_color_str = colors_map.get(stroke_color_ref, '#000000')
    stroke_color = utils.parse_color(stroke_color_str)
    
    # Format text with placeholders
    text = utils.format_placeholder(layer['text'], member)
    
    # Apply text transformations (e.g., name formatting)
    text_transform = layer.get('transform', 'none')
    text = utils.apply_text_transform(text, text_transform)
    
    # Handle random greeting messages for custom_message layer
    if layer.get('id') == 'custom_message' and card_config:
        greeting_cfg = card_config.get('greeting_messages', {})
        if greeting_cfg.get('enabled') and greeting_cfg.get('use_random'):
            greeting_path = greeting_cfg.get('path', 'greeting_messages.json')
            greeting_file = base_path / greeting_path if base_path else Path(greeting_path)
            
            try:
                import json
                import random
                
                if greeting_file.exists():
                    with open(greeting_file, 'r', encoding='utf-8') as f:
                        greeting_data = json.load(f)
                        greetings = greeting_data.get('greetings', [])
                        if greetings:
                            text = random.choice(greetings)
                            logger.info(f"Using random greeting for {member.get('full_name', 'N/A')}")
            except Exception as e:
                logger.warning(f"Failed to load random greeting: {e}, using default text")
    
    # Get box
    box = layer['box']
    box_width = box['width']
    box_height = box['height']
    box_x = box['x']
    box_y = box['y']
    
    # Handle wrapping and fitting
    wrapped_lines = text.split('\n')
    
    if layer.get('wrap', True):
        # Wrap text to fit width
        wrapped_lines = []
        for line in text.split('\n'):
            wrapped_lines.extend(utils.wrap_text(line, box_width, font))
    
    # Handle shrink-to-fit
    fit_cfg = layer.get('fit', {})
    if fit_cfg.get('mode') == 'shrink_to_fit':
        wrapped_lines, _ = utils.calculate_shrink_to_fit(
            text if layer.get('wrap', True) else text.replace('\n', ' '),
            box_width,
            box_height,
            font,
            font_size,
            min_size=fit_cfg.get('min_size', 8),
            max_lines=fit_cfg.get('max_lines'),
            line_spacing=layer.get('line_spacing', 1.2),
            font_path=font_path
        )
    
    # Calculate text position based on alignment
    draw = ImageDraw.Draw(canvas)
    
    align = layer.get('align', 'left')
    valign = layer.get('valign', 'top')
    line_spacing = layer.get('line_spacing', 1.2)
    
    # Get line height
    try:
        bbox = font.getbbox("Ag")
        line_height = int((bbox[3] - bbox[1]) * line_spacing)
    except (AttributeError, TypeError):
        line_height = int(font.getsize("Ag")[1] * line_spacing)
    
    total_height = line_height * len(wrapped_lines)
    
    # Calculate starting Y
    if valign == 'top':
        start_y = box_y
    elif valign == 'middle':
        start_y = box_y + (box_height - total_height) // 2
    else:  # bottom
        start_y = box_y + box_height - total_height
    
    # Draw each line
    for i, line in enumerate(wrapped_lines):
        y = start_y + (i * line_height)
        
        if align == 'left':
            x = box_x
        elif align == 'center':
            try:
                line_bbox = font.getbbox(line)
                line_width = line_bbox[2] - line_bbox[0]
            except (AttributeError, TypeError):
                line_width = font.getsize(line)[0]
            x = box_x + (box_width - line_width) // 2
        else:  # right
            try:
                line_bbox = font.getbbox(line)
                line_width = line_bbox[2] - line_bbox[0]
            except (AttributeError, TypeError):
                line_width = font.getsize(line)[0]
            x = box_x + box_width - line_width
        
        # Draw text with optional stroke
        utils.blend_text_with_stroke(
            draw, line, (x, y), font, color,
            stroke_enabled=stroke_enabled,
            stroke_width=stroke_width,
            stroke_color=stroke_color
        )
    
    return canvas


def render_card(
    member: Dict[str, Any],
    card_config: Dict[str, Any],
    out_path: str,
    fonts_map: Dict[str, str],
    colors_map: Dict[str, str],
    base_path: Optional[Path] = None
) -> str:
    """
    Render birthday card for member.
    
    This is the main rendering function that:
    1. Loads template and resizes if needed
    2. Creates canvas
    3. Loads and prepares photo with circle mask
    4. Pastes photo onto canvas
    5. Pastes template over photo
    6. Renders all text layers
    
    Args:
        member: Member data dictionary
        card_config: Card configuration
        out_path: Output PNG file path
        fonts_map: Font reference to path mapping
        colors_map: Color reference to definition mapping
        base_path: Base path for relative paths
        
    Returns:
        Path to rendered card
        
    Raises:
        RenderError: If rendering fails
    """
    if base_path is None:
        base_path = Path.cwd()
    
    out_file = Path(out_path)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        # Step 1: Load and prepare template
        template_cfg = card_config['template']
        template_path = template_cfg['path']
        if not Path(template_path).is_absolute():
            template_path = str(base_path / template_path)
        
        template = load_template(template_path)
        
        # Determine canvas size
        target_width = template_cfg['size']['width']
        target_height = template_cfg['size']['height']
        
        resize_mode = template_cfg.get('resize_mode', 'force')
        if resize_mode == 'force':
            template = template.resize((target_width, target_height), Image.Resampling.LANCZOS)
            canvas_width, canvas_height = target_width, target_height
        else:  # keep
            canvas_width, canvas_height = template.size
        
        # Step 2: Create blank canvas
        canvas = Image.new('RGBA', (canvas_width, canvas_height), (0, 0, 0, 0))
        
        # Step 2.5: Paste template background first
        canvas.paste(template, (0, 0), template)
        
        # Step 3: Load and process photo
        photo_cfg = card_config['photo']
        photo_field = photo_cfg['source_field']
        
        photo_path = member.get(photo_field, '')
        if not photo_path:
            raise RenderError(f"Member {member['member_id']} missing {photo_field}")
        
        # Get base URL for photo downloads
        photo_base_url = card_config.get('photo_base_url', '')
        
        # Check if photo_path is already a full URL
        is_url = photo_path.startswith('http://') or photo_path.startswith('https://')
        
        # If using base URL and path is not absolute URL, construct full URL
        if photo_base_url and not is_url:
            # Normalize path: ensure single slash at start, no backslashes
            normalized_path = photo_path.replace('\\', '/').lstrip('/')
            photo_path = photo_base_url.rstrip('/') + '/' + normalized_path
            logger.info(f"Constructed photo URL: {photo_path}")
        elif not is_url and not Path(photo_path).is_absolute():
            # No base URL and not absolute, try as relative local path
            photo_path = str(base_path / photo_path)
        
        photo = load_photo(photo_path)
        
        # Step 4: Paste photo with mask on top of template
        placement = photo_cfg['placement']
        crop_type = photo_cfg.get('crop', {}).get('type', 'center_square')
        mask_cfg = photo_cfg.get('mask', {})
        
        canvas = paste_photo_with_mask(
            canvas, photo, placement, crop_type, mask_cfg
        )
        
        # Step 5: Calculate years with company and join year
        def calculate_company_tenure(member: Dict[str, Any]) -> tuple:
            """
            Calculate years with company and join year from Date Join company field.
            
            Returns:
                (years_with_company: str, join_year: str)
            """
            date_join_str = member.get('Date Join', '')
            if not date_join_str:
                return ("", "")
            
            try:
                # Parse the date - try multiple formats
                date_formats = ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d']
                date_join = None
                
                for fmt in date_formats:
                    try:
                        date_join = datetime.strptime(str(date_join_str).strip(), fmt)
                        break
                    except ValueError:
                        continue
                
                if date_join is None:
                    # Try parsing as text month (e.g. "Jan 15, 2020")
                    try:
                        date_join = datetime.strptime(str(date_join_str).strip(), '%b %d, %Y')
                    except ValueError:
                        try:
                            date_join = datetime.strptime(str(date_join_str).strip(), '%B %d, %Y')
                        except ValueError:
                            return ("", "")
                
                today = datetime.now()
                years_diff = today.year - date_join.year
                
                # Adjust if anniversary hasn't occurred this year
                if (today.month, today.day) < (date_join.month, date_join.day):
                    years_diff -= 1
                
                if years_diff < 0:
                    years_str = "0"
                else:
                    years_str = str(years_diff)
                
                join_year_str = str(date_join.year)
                
                return (years_str, join_year_str)
            except Exception as e:
                logger.warning(f"Failed to calculate company tenure: {e}")
                return ("", "")
        
        # Add calculated fields to member data
        member_data = member.copy()
        years_with_company, join_year = calculate_company_tenure(member)
        member_data['years_with_company'] = years_with_company
        member_data['join_year'] = join_year
        
        # Step 6: Render text layers
        for text_layer in card_config.get('texts', []):
            canvas = render_text_layer(
                canvas, text_layer, member_data, fonts_map, colors_map, base_path, card_config
            )
        
        # Save output
        canvas.save(out_path, 'PNG')
        logger.info(f"Rendered card for {member['member_id']} to {out_path}")
        
        return str(out_file.resolve())
    
    except RenderError:
        raise
    except Exception as e:
        raise RenderError(f"Failed to render card for {member['member_id']}: {e}")
