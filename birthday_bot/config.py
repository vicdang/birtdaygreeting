"""Configuration loading and validation."""

import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional
from dotenv import load_dotenv
import os

logger = logging.getLogger(__name__)


class ConfigError(Exception):
    """Configuration validation error."""
    pass


def load_env(env_file: Optional[str] = None):
    """
    Load environment variables from .env file.
    
    Args:
        env_file: Path to .env file (default: .env in project root)
    """
    if env_file:
        load_dotenv(env_file)
    else:
        # Try standard locations
        root = Path(__file__).parent.parent
        env_path = root / ".env"
        if env_path.exists():
            load_dotenv(str(env_path))
        else:
            logger.info("No .env file found, using system environment variables")


def get_required_env(key: str) -> str:
    """
    Get required environment variable.
    
    Args:
        key: Environment variable name
        
    Returns:
        Environment variable value
        
    Raises:
        ConfigError: If variable not set
    """
    value = os.getenv(key)
    if not value:
        raise ConfigError(f"Required environment variable not set: {key}")
    return value


def get_env(key: str, default: Optional[str] = None) -> Optional[str]:
    """
    Get optional environment variable.
    
    Args:
        key: Environment variable name
        default: Default value if not set
        
    Returns:
        Environment variable value or default
    """
    return os.getenv(key, default)


def load_card_config(config_path: str) -> Dict[str, Any]:
    """
    Load and validate card configuration from JSON file.
    
    Args:
        config_path: Path to card_config.json
        
    Returns:
        Configuration dictionary
        
    Raises:
        ConfigError: If config is invalid
    """
    config_file = Path(config_path)
    
    if not config_file.exists():
        raise ConfigError(f"Card config file not found: {config_path}")
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        raise ConfigError(f"Invalid JSON in {config_path}: {e}")
    
    _validate_card_config(config)
    return config


def _validate_card_config(config: Dict[str, Any]):
    """
    Validate card configuration structure.
    
    Args:
        config: Configuration dictionary
        
    Raises:
        ConfigError: If validation fails
    """
    required_keys = ['template', 'photo', 'texts']
    for key in required_keys:
        if key not in config:
            raise ConfigError(f"Missing required config key: {key}")
    
    # Validate template
    template = config['template']
    if 'path' not in template:
        raise ConfigError("template.path is required")
    if 'size' not in template:
        raise ConfigError("template.size is required")
    
    size = template['size']
    if 'width' not in size or 'height' not in size:
        raise ConfigError("template.size must have width and height")
    
    if 'resize_mode' not in template:
        template['resize_mode'] = 'force'
    
    if template['resize_mode'] not in ['force', 'keep']:
        raise ConfigError(f"Invalid resize_mode: {template['resize_mode']}")
    
    # Validate photo
    photo = config['photo']
    if 'source_field' not in photo:
        raise ConfigError("photo.source_field is required")
    if 'placement' not in photo:
        raise ConfigError("photo.placement is required")
    
    placement = photo['placement']
    required_box_keys = ['x', 'y', 'width', 'height']
    for key in required_box_keys:
        if key not in placement:
            raise ConfigError(f"photo.placement.{key} is required")
    
    # Validate crop
    if 'crop' in photo:
        if 'type' not in photo['crop']:
            photo['crop']['type'] = 'center_square'
    else:
        photo['crop'] = {'type': 'center_square'}
    
    # Validate mask
    if 'mask' in photo:
        if 'type' not in photo['mask']:
            photo['mask']['type'] = 'circle'
    else:
        photo['mask'] = {'type': 'circle'}
    
    # Apply defaults for mask
    if photo['mask'].get('type') == 'circle':
        if 'feather' not in photo['mask']:
            photo['mask']['feather'] = 0
        if 'inset' not in photo['mask']:
            photo['mask']['inset'] = 0
    
    # Validate fonts
    if 'fonts' not in config:
        config['fonts'] = {}
    
    # Validate colors
    if 'colors' not in config:
        config['colors'] = {}
    
    # Validate texts
    texts = config['texts']
    if not isinstance(texts, list):
        raise ConfigError("texts must be an array")
    
    for i, text_layer in enumerate(texts):
        if 'id' not in text_layer:
            raise ConfigError(f"texts[{i}] missing id")
        if 'text' not in text_layer:
            raise ConfigError(f"texts[{i}] missing text")
        if 'box' not in text_layer:
            raise ConfigError(f"texts[{i}] missing box")
        
        box = text_layer['box']
        for key in required_box_keys:
            if key not in box:
                raise ConfigError(f"texts[{i}].box.{key} is required")
        
        # Set defaults
        if 'font_ref' not in text_layer:
            text_layer['font_ref'] = 'default'
        if 'color_ref' not in text_layer:
            text_layer['color_ref'] = 'black'
        if 'wrap' not in text_layer:
            text_layer['wrap'] = True
        if 'align' not in text_layer:
            text_layer['align'] = 'left'
        if 'valign' not in text_layer:
            text_layer['valign'] = 'top'
        if 'line_spacing' not in text_layer:
            text_layer['line_spacing'] = 1.2
        
        # Validate fit
        if 'fit' in text_layer:
            fit = text_layer['fit']
            if 'mode' not in fit:
                fit['mode'] = 'none'
            if 'min_size' not in fit:
                fit['min_size'] = 8
            if 'max_lines' not in fit:
                fit['max_lines'] = None
        else:
            text_layer['fit'] = {'mode': 'none', 'min_size': 8, 'max_lines': None}
        
        # Validate stroke
        if 'stroke' in text_layer:
            stroke = text_layer['stroke']
            if 'enabled' not in stroke:
                stroke['enabled'] = False
            if 'width' not in stroke:
                stroke['width'] = 2
            if 'color_ref' not in stroke:
                stroke['color_ref'] = 'black'


def resolve_config_fonts(config: Dict[str, Any], base_path: Path) -> Dict[str, str]:
    """
    Resolve font file paths relative to config base path.
    
    Args:
        config: Card configuration
        base_path: Base path for relative paths
        
    Returns:
        Dictionary mapping font_ref to resolved file paths
    """
    fonts = {}
    
    for font_ref, font_cfg in config.get('fonts', {}).items():
        if isinstance(font_cfg, str):
            # Simple string path
            path = base_path / font_cfg
        elif isinstance(font_cfg, dict):
            path_str = font_cfg.get('path', '')
            path = base_path / path_str
        else:
            logger.warning(f"Invalid font config for {font_ref}: {font_cfg}")
            continue
        
        fonts[font_ref] = str(path.resolve())
    
    return fonts


def resolve_config_colors(config: Dict[str, Any]) -> Dict[str, str]:
    """
    Resolve color references.
    
    Args:
        config: Card configuration
        
    Returns:
        Dictionary mapping color_ref to color definitions
    """
    colors = config.get('colors', {})
    
    # Ensure some defaults
    defaults = {
        'black': '#000000',
        'white': '#FFFFFF',
    }
    
    for key, value in defaults.items():
        if key not in colors:
            colors[key] = value
    
    return colors
