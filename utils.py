"""
Utility functions for OptiWise beam saw optimization tool.
"""

import logging
import os
from typing import Any, Dict
import streamlit as st


def setup_logging(log_level: str = "INFO") -> None:
    """
    Set up logging configuration for the application.
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    """
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    logging.basicConfig(
        level=getattr(logging, log_level.upper()),
        format=log_format,
        handlers=[
            logging.StreamHandler(),
        ]
    )
    
    # Set specific logger levels
    logging.getLogger('matplotlib').setLevel(logging.WARNING)
    logging.getLogger('PIL').setLevel(logging.WARNING)


def validate_file_upload(uploaded_file, expected_extensions: list) -> bool:
    """
    Validate uploaded file type and size.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        expected_extensions: List of allowed file extensions
        
    Returns:
        True if file is valid, False otherwise
    """
    if uploaded_file is None:
        return False
    
    # Check file extension
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    if file_extension not in expected_extensions:
        st.error(f"Invalid file type. Expected: {', '.join(expected_extensions)}")
        return False
    
    # Check file size (max 50MB)
    max_size = 50 * 1024 * 1024  # 50MB in bytes
    if uploaded_file.size > max_size:
        st.error("File size too large. Maximum size is 50MB.")
        return False
    
    return True


def format_currency(amount: float) -> str:
    """
    Format currency amount for display in Indian Rupees.
    
    Args:
        amount: Amount to format
        
    Returns:
        Formatted currency string with Rupee symbol
    """
    return f"₹{amount:,.2f}"


def format_area(area_mm2: float) -> str:
    """
    Format area for display with appropriate units.
    
    Args:
        area_mm2: Area in square millimeters
        
    Returns:
        Formatted area string
    """
    if area_mm2 >= 1_000_000:
        return f"{area_mm2 / 1_000_000:.2f} m²"
    elif area_mm2 >= 1_000:
        return f"{area_mm2 / 1_000:.1f} cm²"
    else:
        return f"{area_mm2:.0f} mm²"


def format_percentage(value: float) -> str:
    """
    Format percentage for display.
    
    Args:
        value: Percentage value (0-100)
        
    Returns:
        Formatted percentage string
    """
    return f"{value:.1f}%"


def create_download_link(file_bytes: bytes, filename: str, mime_type: str) -> str:
    """
    Create a download link for file bytes.
    
    Args:
        file_bytes: File content as bytes
        filename: Name for the downloaded file
        mime_type: MIME type of the file
        
    Returns:
        Base64 encoded download link
    """
    import base64
    
    b64 = base64.b64encode(file_bytes).decode()
    return f'<a href="data:{mime_type};base64,{b64}" download="{filename}">Download {filename}</a>'


def display_optimization_metrics(boards, unplaced_parts, initial_cost, final_cost):
    """
    Display optimization metrics in Streamlit columns.
    
    Args:
        boards: List of optimized boards
        unplaced_parts: List of unplaced parts
        initial_cost: Initial cost estimate
        final_cost: Final optimized cost
    """
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Boards Used", len(boards))
    
    with col2:
        total_parts = sum(len(board.parts_on_board) for board in boards) + len(unplaced_parts)
        placed_parts = total_parts - len(unplaced_parts)
        st.metric("Parts Placed", f"{placed_parts}/{total_parts}")
    
    with col3:
        if boards:
            avg_utilization = sum(board.get_utilization_percentage() for board in boards) / len(boards)
            st.metric("Avg Utilization", format_percentage(avg_utilization))
        else:
            st.metric("Avg Utilization", "0.0%")
    
    with col4:
        cost_savings = initial_cost - final_cost
        savings_pct = (cost_savings / initial_cost * 100) if initial_cost > 0 else 0
        st.metric("Cost Savings", format_currency(cost_savings), 
                 delta=f"{savings_pct:.1f}%" if cost_savings >= 0 else None)


def display_board_summary(boards):
    """
    Display board summary table in Streamlit.
    
    Args:
        boards: List of optimized boards
    """
    if not boards:
        st.info("No boards to display.")
        return
    
    board_data = []
    for board in boards:
        board_data.append({
            'Board ID': board.id,
            'Material': board.material_details.full_material_string,
            'Dimensions': f"{board.total_length}×{board.total_width}mm",
            'Parts': len(board.parts_on_board),
            'Utilization': format_percentage(board.get_utilization_percentage()),
            'Remaining Area': format_area(board.get_remaining_area())
        })
    
    st.dataframe(board_data, use_container_width=True)


def display_error_summary(validation_results):
    """
    Display data validation error summary.
    
    Args:
        validation_results: Dictionary with validation results
    """
    if validation_results['invalid_parts'] > 0:
        st.warning(f"Found {validation_results['invalid_parts']} invalid parts that will be skipped:")
        
        if validation_results['missing_cores']:
            st.error(f"Missing core materials: {', '.join(validation_results['missing_cores'])}")
        
        if validation_results['missing_laminates']:
            st.error(f"Missing laminates: {', '.join(validation_results['missing_laminates'])}")
    
    st.info(f"Processing {validation_results['valid_parts']} valid parts out of {validation_results['total_parts']} total parts.")


def get_material_options(core_db: Dict[str, Any]) -> list:
    """
    Get list of available core materials sorted by grade level.
    
    Args:
        core_db: Core materials database
        
    Returns:
        List of core material names sorted by grade level
    """
    if not core_db:
        return []
    
    # Sort by grade level (ascending)
    sorted_cores = sorted(core_db.items(), key=lambda x: x[1].get('grade_level', 0))
    return [core_name for core_name, _ in sorted_cores]


def validate_upgrade_sequence(upgrade_sequence: str, core_db: Dict[str, Any]) -> tuple:
    """
    Validate user-provided upgrade sequence.
    
    Args:
        upgrade_sequence: Comma-separated string of core material names
        core_db: Core materials database
        
    Returns:
        Tuple of (is_valid, error_message, cleaned_sequence)
    """
    if not upgrade_sequence.strip():
        return True, "", ""
    
    # Parse and validate
    core_names = [name.strip() for name in upgrade_sequence.split(',') if name.strip()]
    
    # Check if all cores exist in database
    missing_cores = [name for name in core_names if name not in core_db]
    if missing_cores:
        return False, f"Unknown core materials: {', '.join(missing_cores)}", ""
    
    # Check for duplicates
    if len(core_names) != len(set(core_names)):
        return False, "Duplicate core materials in upgrade sequence", ""
    
    cleaned_sequence = ', '.join(core_names)
    return True, "", cleaned_sequence
