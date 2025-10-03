"""
Input parsers for OptiWise beam saw optimization tool.
Handles reading and processing CSV data files.
"""

import pandas as pd
import logging
from typing import List, Dict, Any
from data_models import MaterialDetails, Part

logger = logging.getLogger(__name__)


def load_parts_data(filepath: str) -> List[Part]:
    """
    Load parts data from CSV file and create Part objects.
    
    Args:
        filepath: Path to the cutlist CSV file
        
    Returns:
        List of Part objects
        
    Expected CSV columns (new format):
        - CLIENT NAME: Customer/client name
        - ORDER ID / UNIQUE CODE: Unique identifier for the part
        - SL NO.: Serial number
        - ROOM TYPE: Room category
        - SUB CATEGORY: Sub-category
        - TYPE: Type classification
        - PANEL NAME: Panel identifier
        - FULL NAME DESCRIPTION: Full description
        - QTY: Number of pieces needed
        - GROOVE: Groove specification
        - CUT LENGTH: Part length in millimeters
        - CUT WIDTH: Part width in millimeters
        - FINISHED THICKNESS: Thickness specification
        - MATERIAL TYPE: Material specification string
        - EB1, EB2, EB3, EB4: Edge band specifications
        - GRAINS: 1 if grain-sensitive, 0 if grain-free
        - REMARKS: Additional notes
    """
    parts_list = []
    
    try:
        # Read CSV file
        df = pd.read_csv(filepath)
        
        # Log the actual columns for debugging
        logger.info(f"Parts CSV columns: {list(df.columns)}")
        
        # Map old format columns to new format (backward compatibility)
        old_to_new_mapping = {
            'Part ID': 'ORDER ID / UNIQUE CODE',
            'Length (mm)': 'CUT LENGTH',
            'Width (mm)': 'CUT WIDTH', 
            'Length': 'CUT LENGTH',
            'Width': 'CUT WIDTH',
            'Quantity': 'QTY',
            'Material': 'MATERIAL TYPE',
            'Grain Sensitive': 'GRAINS',
            'Grain': 'GRAINS'
        }
        
        # Apply column mapping for backward compatibility
        for old_name, new_name in old_to_new_mapping.items():
            if old_name in df.columns and new_name not in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Validate required columns (new format)
        required_columns = ['ORDER ID / UNIQUE CODE', 'CUT LENGTH', 'CUT WIDTH', 'QTY', 'MATERIAL TYPE', 'GRAINS']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Missing required columns in parts CSV: {missing_columns}")
            logger.error(f"Available columns: {list(df.columns)}")
            raise ValueError(f"Missing required columns in parts CSV: {missing_columns}")
        
        logger.info(f"Loading {len(df)} parts from {filepath}")
        logger.info(f"Final columns after mapping: {list(df.columns)}")
        
        # Process each row
        for index, row in df.iterrows():
            try:
                # Extract data from row (new format)
                original_part_id = str(row['ORDER ID / UNIQUE CODE']).strip()
                length = float(row['CUT LENGTH'])
                width = float(row['CUT WIDTH'])
                quantity = int(row['QTY'])
                material_type = str(row['MATERIAL TYPE']).strip()
                grains = int(row['GRAINS'])
                
                # Extract additional fields for enhanced display (handle NaN values)
                def safe_str(value):
                    """Safely convert value to string, handling NaN"""
                    if pd.isna(value):
                        return ''
                    return str(value).strip()
                
                client_name = safe_str(row.get('CLIENT NAME', ''))
                room_type = safe_str(row.get('ROOM TYPE', ''))
                sub_category = safe_str(row.get('SUB CATEGORY', ''))
                panel_name = safe_str(row.get('PANEL NAME', ''))
                full_description = safe_str(row.get('FULL NAME DESCRIPTION', ''))
                
                # Validate data
                if length <= 0 or width <= 0:
                    logger.warning(f"Invalid dimensions for part {original_part_id}: {length}x{width}")
                    continue
                
                if quantity <= 0:
                    logger.warning(f"Invalid quantity for part {original_part_id}: {quantity}")
                    continue
                
                if grains not in [0, 1]:
                    logger.warning(f"Invalid grains value for part {original_part_id}: {grains}")
                    continue
                
                # Create MaterialDetails with error handling
                try:
                    material_details = MaterialDetails(material_type)
                except ValueError as e:
                    logger.error(f"Skipping part {original_part_id} due to invalid material: {e}")
                    continue
                
                # Create individual Part objects for each quantity
                for i in range(quantity):
                    part_id = f"{original_part_id}_{i+1}" if quantity > 1 else original_part_id
                    
                    part = Part(
                        part_id=part_id,
                        requested_length=length,
                        requested_width=width,
                        quantity=1,  # Each Part object represents one piece
                        material_details=material_details,
                        grains=grains,
                        original_part_index=index,
                        client_name=client_name,
                        room_type=room_type,
                        sub_category=sub_category,
                        panel_name=panel_name,
                        full_description=full_description
                    )
                    
                    # Store original CSV data for accurate Excel export
                    original_data = {}
                    for col in ['CLIENT NAME', 'ORDER ID / UNIQUE CODE', 'SL NO.', 'ROOM TYPE', 'SUB CATEGORY', 'TYPE', 
                               'PANEL NAME', 'FULL NAME DESCRIPTION', 'QTY', 'GROOVE', 'CUT LENGTH', 'CUT WIDTH', 
                               'FINISHED THICKNESS', 'MATERIAL TYPE', 'EB1', 'EB2', 'EB3', 'EB4', 'GRAINS', 'REMARKS']:
                        original_data[col] = safe_str(row.get(col, ''))
                    part.original_data = original_data
                    
                    parts_list.append(part)
                    
            except Exception as e:
                logger.error(f"Error processing row {index} in parts CSV: {e}")
                continue
        
        logger.info(f"Successfully loaded {len(parts_list)} individual parts")
        return parts_list
        
    except FileNotFoundError:
        logger.error(f"Parts CSV file not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error loading parts data from {filepath}: {e}")
        raise


def load_core_materials_config(filepath: str) -> Dict[str, Dict[str, Any]]:
    """
    Load core materials configuration from CSV file.
    
    Args:
        filepath: Path to the core materials CSV file
        
    Returns:
        Dictionary mapping core name to its properties
        
    Expected CSV columns:
        - Name: Core material name (e.g., '18MR')
        - Thickness (mm): Thickness in millimeters
        - Price per SqM: Price per square meter
        - Standard Length (mm): Standard board length
        - Standard Width (mm): Standard board width
        - Grade Level: Quality/cost grade level (higher = better/more expensive)
    """
    try:
        # Read CSV file
        df = pd.read_csv(filepath)
        
        # Map column names to handle different formats
        core_column_mapping = {
            'Core Name': 'Name',
            'Name': 'Name'
        }
        
        # Apply column mapping
        for old_name, new_name in core_column_mapping.items():
            if old_name in df.columns and new_name not in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Validate required columns
        required_columns = ['Name', 'Thickness (mm)', 'Price per SqM', 
                           'Standard Length (mm)', 'Standard Width (mm)', 'Grade Level']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns in core materials CSV: {missing_columns}")
        
        core_db = {}
        
        logger.info(f"Loading {len(df)} core materials from {filepath}")
        
        # Process each row
        for index, row in df.iterrows():
            try:
                name = str(row['Name']).strip()
                thickness = int(row['Thickness (mm)'])
                price_per_sqm = float(row['Price per SqM'])
                standard_length = float(row['Standard Length (mm)'])
                standard_width = float(row['Standard Width (mm)'])
                grade_level = int(row['Grade Level'])
                
                # Validate data
                if thickness <= 0:
                    logger.warning(f"Invalid thickness for core {name}: {thickness}")
                    continue
                
                if price_per_sqm < 0:
                    logger.warning(f"Invalid price for core {name}: {price_per_sqm}")
                    continue
                
                if standard_length <= 0 or standard_width <= 0:
                    logger.warning(f"Invalid dimensions for core {name}: {standard_length}x{standard_width}")
                    continue
                
                core_db[name] = {
                    'thickness': thickness,
                    'price_per_sqm': price_per_sqm,
                    'standard_length': standard_length,
                    'standard_width': standard_width,
                    'grade_level': grade_level
                }
                
            except Exception as e:
                logger.error(f"Error processing core material row {index}: {e}")
                continue
        
        logger.info(f"Successfully loaded {len(core_db)} core materials")
        return core_db
        
    except FileNotFoundError:
        logger.error(f"Core materials CSV file not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error loading core materials data from {filepath}: {e}")
        raise


def load_laminates_config(filepath: str) -> Dict[str, float]:
    """
    Load laminates configuration from CSV file.
    
    Args:
        filepath: Path to the laminates CSV file
        
    Returns:
        Dictionary mapping laminate name to price per square meter
        
    Expected CSV columns:
        - Name: Laminate name (e.g., '2614 SF')
        - Price per SqM: Price per square meter
    """
    try:
        # Read CSV file
        df = pd.read_csv(filepath)
        
        # Map column names to handle different formats
        laminate_column_mapping = {
            'Laminate Name': 'Name',
            'Name': 'Name'
        }
        
        # Apply column mapping
        for old_name, new_name in laminate_column_mapping.items():
            if old_name in df.columns and new_name not in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # Validate required columns
        required_columns = ['Name', 'Price per SqM']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns in laminates CSV: {missing_columns}")
        
        laminate_db = {}
        
        logger.info(f"Loading {len(df)} laminates from {filepath}")
        
        # Process each row
        for index, row in df.iterrows():
            try:
                name = str(row['Name']).strip()
                price_per_sqm = float(row['Price per SqM'])
                
                # Validate data
                if price_per_sqm < 0:
                    logger.warning(f"Invalid price for laminate {name}: {price_per_sqm}")
                    continue
                
                laminate_db[name] = price_per_sqm
                
            except Exception as e:
                logger.error(f"Error processing laminate row {index}: {e}")
                continue
        
        logger.info(f"Successfully loaded {len(laminate_db)} laminates")
        return laminate_db
        
    except FileNotFoundError:
        logger.error(f"Laminates CSV file not found: {filepath}")
        raise
    except Exception as e:
        logger.error(f"Error loading laminates data from {filepath}: {e}")
        raise


def validate_data_consistency(parts_list: List[Part], core_db: Dict, laminate_db: Dict) -> Dict[str, Any]:
    """
    Validate data consistency across loaded datasets.
    
    Args:
        parts_list: List of Part objects
        core_db: Core materials database
        laminate_db: Laminates database
        
    Returns:
        Dictionary with validation results and statistics
    """
    validation_results = {
        'total_parts': len(parts_list),
        'unique_materials': set(),
        'missing_cores': set(),
        'missing_laminates': set(),
        'valid_parts': 0,
        'invalid_parts': 0
    }
    
    try:
        for part in parts_list:
            material = part.material_details
            validation_results['unique_materials'].add(material.full_material_string)
            
            # Check if core exists in database
            if material.core_name not in core_db:
                validation_results['missing_cores'].add(material.core_name)
                validation_results['invalid_parts'] += 1
                continue
            
            # Check if top laminate exists in database
            if material.top_laminate_name not in laminate_db:
                validation_results['missing_laminates'].add(material.top_laminate_name)
                validation_results['invalid_parts'] += 1
                continue
            
            # Check if bottom laminate exists in database
            if material.bottom_laminate_name not in laminate_db:
                validation_results['missing_laminates'].add(material.bottom_laminate_name)
                validation_results['invalid_parts'] += 1
                continue
            
            validation_results['valid_parts'] += 1
        
        # Log validation results
        if validation_results['missing_cores']:
            logger.warning(f"Missing core materials: {validation_results['missing_cores']}")
        
        if validation_results['missing_laminates']:
            logger.warning(f"Missing laminates: {validation_results['missing_laminates']}")
        
        logger.info(f"Validation complete: {validation_results['valid_parts']} valid parts, "
                   f"{validation_results['invalid_parts']} invalid parts")
        
        return validation_results
        
    except Exception as e:
        logger.error(f"Error during data validation: {e}")
        raise