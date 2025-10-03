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
    Load parts data from Excel file and create Part objects.
    
    Args:
        filepath: Path to the cutlist Excel file
        
    Returns:
        List of Part objects
        
    Expected Excel columns:
        - Part ID: Unique identifier for the part
        - Length (mm): Part length in millimeters  
        - Width (mm): Part width in millimeters
        - Quantity: Number of pieces needed
        - Material Type: Material specification string
        - Grains: 1 if grain-sensitive, 0 if grain-free
    """
    parts_list = []
    
    try:
        # Read Excel file
        df = pd.read_excel(filepath)
        
        # Validate required columns
        required_columns = ['Part ID', 'Length (mm)', 'Width (mm)', 'Quantity', 'Material Type', 'Grains']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns in parts Excel file: {missing_columns}")
        
        logger.info(f"Loading {len(df)} parts from {filepath}")
        
        # Process each row
        for index, row in df.iterrows():
            try:
                # Extract data from row
                original_part_id = str(row['Part ID']).strip()
                length = float(row['Length (mm)'])
                width = float(row['Width (mm)'])
                quantity = int(row['Quantity'])
                material_type = str(row['Material Type']).strip()
                grains = int(row['Grains'])
                
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
                        original_part_index=int(index)
                    )
                    
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
    Load core materials configuration from Excel file.
    
    Args:
        filepath: Path to the core materials Excel file
        
    Returns:
        Dictionary mapping core name to its properties
        
    Expected Excel columns:
        - Name: Core material name (e.g., '18MR')
        - Thickness (mm): Thickness in millimeters
        - Price per SqM: Price per square meter
        - Standard Length (mm): Standard board length
        - Standard Width (mm): Standard board width
        - Grade Level: Quality/cost grade level (higher = better/more expensive)
    """
    try:
        # Read Excel file
        df = pd.read_excel(filepath)
        
        # Validate required columns
        required_columns = ['Name', 'Thickness (mm)', 'Price per SqM', 
                           'Standard Length (mm)', 'Standard Width (mm)', 'Grade Level']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns in core materials Excel file: {missing_columns}")
        
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
    Load laminates configuration from Excel file.
    
    Args:
        filepath: Path to the laminates Excel file
        
    Returns:
        Dictionary mapping laminate name to price per square meter
        
    Expected Excel columns:
        - Name: Laminate name (e.g., '2614 SF')
        - Price per SqM: Price per square meter
    """
    try:
        # Read Excel file
        df = pd.read_excel(filepath)
        
        # Validate required columns
        required_columns = ['Name', 'Price per SqM']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns in laminates Excel file: {missing_columns}")
        
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
            
            # Check if laminate exists in database
            if material.laminate_name not in laminate_db:
                validation_results['missing_laminates'].add(material.laminate_name)
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
