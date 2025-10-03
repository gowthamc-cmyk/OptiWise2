"""
Standalone CSV parsers for OptiWise beam saw optimization tool.
Pure Python implementation without pandas dependency.
"""

import csv
import logging
from typing import List, Dict, Any
from data_models import MaterialDetails, Part

logger = logging.getLogger(__name__)


def load_parts_data(filepath: str) -> tuple[List[Part], List[str]]:
    """
    Load parts data from CSV file and create Part objects.
    
    Args:
        filepath: Path to the cutlist CSV file
        
    Returns:
        List of Part objects
    """
    parts = []
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            
            for index, row in enumerate(reader):
                try:
                    # Skip empty rows or rows that are just headers
                    if not row or not any(row.values()):
                        continue
                    
                    # Handle both old and new format column names
                    # Try new format first
                    part_id_col = 'ORDER ID / UNIQUE CODE' if 'ORDER ID / UNIQUE CODE' in row else 'Part ID'
                    length_col = 'CUT LENGTH' if 'CUT LENGTH' in row else 'Length (mm)' if 'Length (mm)' in row else 'Length'
                    width_col = 'CUT WIDTH' if 'CUT WIDTH' in row else 'Width (mm)' if 'Width (mm)' in row else 'Width'
                    qty_col = 'QTY' if 'QTY' in row else 'Quantity'
                    material_col = 'MATERIAL TYPE' if 'MATERIAL TYPE' in row else 'Material'
                    grains_col = 'GRAINS' if 'GRAINS' in row else 'Grain Sensitive' if 'Grain Sensitive' in row else 'Grain'
                    
                    # Skip if this looks like a header row
                    part_id_value = row.get(part_id_col, '').strip()
                    material_value = row.get(material_col, '').strip()
                    if (part_id_value.lower() in ['part id', 'part_id', 'partid', 'order id / unique code'] or 
                        not part_id_value or 
                        material_value.lower() in ['material', 'material type']):
                        continue
                    
                    # Extract data with flexible column names
                    part_id = row[part_id_col].strip()
                    length = float(row[length_col])
                    width = float(row[width_col])
                    quantity = int(row[qty_col])
                    
                    # Handle material using the dynamic column mapping
                    material_string = row.get(material_col, '').strip()
                    if not material_string:
                        logger.warning(f"No material found for part {part_id} in row {index + 2}")
                        continue
                    
                    # Handle grains using the dynamic column mapping
                    grain_val = row.get(grains_col, '0')
                    if isinstance(grain_val, str):
                        grain_val = grain_val.lower()
                        if grain_val in ['yes', 'true', '1', 'grain sensitive']:
                            grain_sensitive = 1
                        else:
                            grain_sensitive = 0
                    else:
                        grain_sensitive = int(grain_val)
                    
                    # Extract additional fields for new format (handle missing fields gracefully)
                    def safe_str(value):
                        return str(value).strip() if value and str(value).lower() != 'nan' else ''
                    
                    client_name = safe_str(row.get('CLIENT NAME', ''))
                    room_type = safe_str(row.get('ROOM TYPE', ''))
                    sub_category = safe_str(row.get('SUB CATEGORY', ''))
                    panel_name = safe_str(row.get('PANEL NAME', ''))
                    full_description = safe_str(row.get('FULL NAME DESCRIPTION', ''))
                    
                    # Create material details with error handling
                    try:
                        material_details = MaterialDetails(material_string)
                    except Exception as mat_error:
                        logger.error(f"Failed to parse material '{material_string}': {mat_error}")
                        continue
                    
                    # Create parts (one for each quantity)
                    for q in range(quantity):
                        unique_part_id = f"{part_id}_{q+1}" if quantity > 1 else part_id
                        part = Part(
                            part_id=unique_part_id,
                            requested_length=length,
                            requested_width=width,
                            quantity=1,  # Individual part instance
                            material_details=material_details,
                            grains=grain_sensitive,
                            original_part_index=index,
                            client_name=client_name,
                            room_type=room_type,
                            sub_category=sub_category,
                            panel_name=panel_name,
                            full_description=full_description
                        )
                        
                        # Store original CSV data for accurate Excel export
                        original_data = {}
                        for key, value in row.items():
                            if key and value is not None:
                                original_data[key] = str(value).strip()
                        part.original_data = original_data
                        
                        parts.append(part)
                        
                except (ValueError, KeyError) as e:
                    logger.error(f"Error processing row {index + 2}: {e}. Row data: {row}")
                    continue
                    
    except FileNotFoundError:
        logger.error(f"File not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error reading parts file: {e}")
        raise
    
    logger.info(f"Loaded {len(parts)} parts from {filepath}")
    return parts, []


def load_core_materials_config(filepath: str) -> tuple[Dict[str, Dict[str, Any]], List[str]]:
    """
    Load core materials configuration from CSV file.
    
    Args:
        filepath: Path to the core materials CSV file
        
    Returns:
        Dictionary mapping core names to their properties
    """
    core_materials = {}
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            logger.info(f"Core materials CSV columns: {reader.fieldnames}")
            
            for row in reader:
                try:
                    # Skip empty rows or header rows
                    if not row or not any(row.values()):
                        continue
                    
                    # Skip if this looks like a header row
                    # Find the core name column with flexible naming - prioritize 'Name' for core materials
                    core_name_key = next((key for key in row.keys() if 'name' in key.lower() and key.lower() != 'filename'), 
                                       next((key for key in row.keys() if 'core' in key.lower() and 'material' in key.lower()), 
                                       next((key for key in row.keys() if 'core' in key.lower()), 
                                       next((key for key in row.keys() if 'material' in key.lower()), None))))
                    
                    if not core_name_key:
                        logger.warning(f"No core material column found in row: {list(row.keys())}")
                        continue
                        
                    core_name_value = row.get(core_name_key, '').strip()
                    # Skip if this looks like a header row (value matches column name) or is empty
                    if (not core_name_value or 
                        core_name_value.lower() in ['core name', 'core_name', 'corename', 'core material', 'core_material'] or
                        core_name_value == core_name_key):
                        continue
                    
                    core_name = core_name_value
                    
                    # Handle different column name variations with flexible defaults
                    length_key = next((key for key in row.keys() if 'length' in key.lower()), None)
                    width_key = next((key for key in row.keys() if 'width' in key.lower()), None)
                    thickness_key = next((key for key in row.keys() if 'thickness' in key.lower()), None)
                    price_key = next((key for key in row.keys() if 'price' in key.lower()), None)
                    grade_key = next((key for key in row.keys() if 'grade' in key.lower()), None)
                    
                    # Extract values with fallbacks
                    length = float(row[length_key]) if length_key and row.get(length_key) else 2500.0
                    width = float(row[width_key]) if width_key and row.get(width_key) else 1250.0
                    thickness = float(row[thickness_key]) if thickness_key and row.get(thickness_key) else 16.0
                    price = float(row[price_key]) if price_key and row.get(price_key) else 850.0
                    grade = int(row[grade_key]) if grade_key and row.get(grade_key) else 1
                    
                    core_materials[core_name] = {
                        'standard_length': length,
                        'standard_width': width,
                        'thickness': thickness,
                        'price_per_sqm': price,
                        'grade_level': grade
                    }
                    
                except (ValueError, KeyError) as e:
                    logger.error(f"Error processing core material row: {e}")
                    continue
                    
    except FileNotFoundError:
        logger.error(f"File not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error reading core materials file: {e}")
        raise
    
    logger.info(f"Loaded {len(core_materials)} core materials from {filepath}")
    return core_materials, []


def load_laminates_config(filepath: str) -> tuple[Dict[str, float], List[str]]:
    """
    Load laminates configuration from CSV file.
    
    Args:
        filepath: Path to the laminates CSV file
        
    Returns:
        Dictionary mapping laminate names to prices
    """
    laminates = {}
    
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            logger.info(f"Laminates CSV columns: {reader.fieldnames}")
            
            for row in reader:
                try:
                    # Skip empty rows or header rows
                    if not row or not any(row.values()):
                        continue
                    
                    # Skip if this looks like a header row
                    # Find the laminate name column with flexible naming - try 'Laminate' first
                    name_key = next((key for key in row.keys() if 'laminate' in key.lower() and key.lower() != 'laminate name'), 
                                  next((key for key in row.keys() if 'name' in key.lower()), None))
                    
                    if not name_key:
                        logger.warning(f"No laminate name column found in row: {list(row.keys())}")
                        continue
                        
                    laminate_name_value = row.get(name_key, '').strip()
                    if (laminate_name_value.lower() in ['laminate name', 'laminate_name', 'laminatename', 'name', 'laminate'] or 
                        not laminate_name_value):
                        continue
                    
                    # Handle different column name variations for laminates
                    price_key = next((key for key in row.keys() if 'price' in key.lower()), None)
                    
                    if not price_key:
                        logger.warning(f"No price column found in row: {list(row.keys())}")
                        continue
                    
                    laminate_name = laminate_name_value
                    price = float(row[price_key]) if row.get(price_key) else 100.0
                    laminates[laminate_name] = price
                    
                except (ValueError, KeyError) as e:
                    logger.error(f"Error processing laminate row: {e}")
                    continue
                    
    except FileNotFoundError:
        logger.error(f"File not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error reading laminates file: {e}")
        raise
    
    logger.info(f"Loaded {len(laminates)} laminates from {filepath}")
    return laminates, []


def filter_parts_with_known_materials(parts_list: List[Part], core_db: Dict, laminate_db: Dict) -> tuple[List[Part], List[str]]:
    """
    Filter parts to only include those with known materials.
    
    Args:
        parts_list: List of Part objects
        core_db: Core materials database
        laminate_db: Laminates database
        
    Returns:
        Tuple of (filtered_parts_list, list_of_skipped_part_ids)
    """
    filtered_parts = []
    skipped_parts = []
    
    for part in parts_list:
        # Check if core material exists
        if part.material_details.core_name not in core_db:
            skipped_parts.append(f"{part.id} (unknown core: {part.material_details.core_name})")
            continue
            
        # Check if laminates exist
        if (part.material_details.top_laminate_name not in laminate_db or
            part.material_details.bottom_laminate_name not in laminate_db):
            missing_laminates = []
            if part.material_details.top_laminate_name not in laminate_db:
                missing_laminates.append(part.material_details.top_laminate_name)
            if part.material_details.bottom_laminate_name not in laminate_db:
                missing_laminates.append(part.material_details.bottom_laminate_name)
            skipped_parts.append(f"{part.id} (unknown laminates: {', '.join(missing_laminates)})")
            continue
            
        # Part has all known materials
        filtered_parts.append(part)
    
    return filtered_parts, skipped_parts


def validate_data_consistency(parts_list: List[Part], core_db: Dict, laminate_db: Dict) -> tuple[bool, List[str]]:
    """
    Validate consistency between parts data and material databases.
    
    Args:
        parts_list: List of Part objects
        core_db: Core materials database
        laminate_db: Laminates database
        
    Returns:
        Tuple of (is_valid, list_of_error_messages)
    """
    errors = []
    warnings = []
    
    # Check if databases are empty
    if not core_db:
        errors.append("Core materials database is empty")
    if not laminate_db:
        errors.append("Laminates database is empty")
    if not parts_list:
        errors.append("Parts list is empty")
    
    if errors:
        return False, errors
    
    # Collect all materials used in parts
    used_cores = set()
    used_laminates = set()
    
    for part in parts_list:
        used_cores.add(part.material_details.core_name)
        used_laminates.add(part.material_details.top_laminate_name)
        used_laminates.add(part.material_details.bottom_laminate_name)
    
    # Check for missing core materials - report as warnings, not errors
    missing_cores = used_cores - set(core_db.keys())
    if missing_cores:
        warnings.append(f"Core materials not found in database: {', '.join(missing_cores)}. Parts with these materials will be skipped.")
    
    # Check for missing laminates - report as warnings, not errors
    missing_laminates = used_laminates - set(laminate_db.keys())
    if missing_laminates:
        warnings.append(f"Laminates not found in database: {', '.join(missing_laminates)}. Parts with these materials will be skipped.")
    
    # Add warnings to errors list for display purposes
    errors.extend(warnings)
    
    # Check for parts with invalid dimensions
    invalid_parts = []
    for part in parts_list:
        if part.requested_length <= 0 or part.requested_width <= 0:
            invalid_parts.append(part.id)
    
    if invalid_parts:
        errors.append(f"Parts with invalid dimensions: {', '.join(invalid_parts)}")
    
    # Check for parts larger than any available board
    oversized_parts = []
    max_board_length = max(core['standard_length'] for core in core_db.values()) if core_db else 0
    max_board_width = max(core['standard_width'] for core in core_db.values()) if core_db else 0
    
    for part in parts_list:
        # Check if part fits on largest board (considering rotation)
        fits_normal = (part.requested_length <= max_board_length and 
                      part.requested_width <= max_board_width)
        fits_rotated = (part.requested_width <= max_board_length and 
                       part.requested_length <= max_board_width)
        
        if not fits_normal and not fits_rotated:
            oversized_parts.append(f"{part.id} ({part.requested_length}×{part.requested_width}mm)")
        elif not fits_normal and fits_rotated and part.grains == 1:
            # Part only fits when rotated but cannot be rotated
            oversized_parts.append(f"{part.id} (grain-sensitive, {part.requested_length}×{part.requested_width}mm)")
    
    if oversized_parts:
        errors.append(f"Parts too large for available boards: {', '.join(oversized_parts)}")
    
    is_valid = len(errors) == 0
    return is_valid, errors


def parse_csv_from_string(csv_content: str) -> List[Dict[str, str]]:
    """
    Parse CSV content from string format.
    
    Args:
        csv_content: CSV data as string
        
    Returns:
        List of dictionaries representing CSV rows
    """
    try:
        lines = csv_content.strip().split('\n')
        if len(lines) < 2:
            return []
        
        reader = csv.DictReader(lines)
        return list(reader)
        
    except Exception as e:
        logger.error(f"Error parsing CSV string: {e}")
        return []


def validate_csv_format(csv_content: str, expected_columns: List[str]) -> tuple[bool, str]:
    """
    Validate CSV format and required columns.
    
    Args:
        csv_content: CSV data as string
        expected_columns: List of required column names
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    try:
        lines = csv_content.strip().split('\n')
        if len(lines) < 2:
            return False, "CSV must have at least header and one data row"
        
        reader = csv.DictReader(lines)
        headers = reader.fieldnames
        
        if not headers:
            return False, "No headers found in CSV"
        
        # Check for required columns
        missing_columns = [col for col in expected_columns if col not in headers]
        if missing_columns:
            return False, f"Missing required columns: {', '.join(missing_columns)}"
        
        # Try to read first data row
        try:
            next(reader)
        except StopIteration:
            return False, "No data rows found"
        except Exception as e:
            return False, f"Error reading data: {e}"
        
        return True, "Valid CSV format"
        
    except Exception as e:
        return False, f"Error validating CSV: {e}"