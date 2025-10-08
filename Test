"""
Report generation for OptiWise beam saw optimization tool.
Generates PDF cutting layouts and Excel reports.
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages
from typing import List, Dict, Any
import logging
import io
import base64

logger = logging.getLogger(__name__)

from data_models import Board, Part
from collections import Counter


def calculate_core_material_summary(boards: List[Board], core_db: Dict = None) -> List[Dict]:
    """Calculate summary of boards by core material type with costs and utilization."""
    core_summary = {}
    
    for board in boards:
        core_material_key = board.material_details.core_name  # Use core name instead of full material string
        board_area_sqm = (board.total_length * board.total_width) / 1_000_000
        board_area_sqft = board_area_sqm * 10.764  # Convert to square feet
        
        # Calculate actual utilized area (parts placed)
        utilized_area_sqm = 0.0
        for part in board.parts_on_board:
            actual_length = getattr(part, 'actual_length', None) or part.requested_length
            actual_width = getattr(part, 'actual_width', None) or part.requested_width
            part_area = (actual_length * actual_width) / 1_000_000
            utilized_area_sqm += part_area
        
        utilized_area_sqft = utilized_area_sqm * 10.764  # Convert to square feet
        
        if core_material_key not in core_summary:
            # Get unit price from core database (only core material cost, not including laminates)
            unit_price_sqm = 0.0
            if core_db and core_material_key in core_db:
                core_info = core_db[core_material_key]
                if isinstance(core_info, dict) and 'price_per_sqm' in core_info:
                    unit_price_sqm = core_info['price_per_sqm']
            
            unit_price_sqft = unit_price_sqm / 10.764  # Convert to per sqft (1 sqm = 10.764 sqft)
            
            core_summary[core_material_key] = {
                'core_material': core_material_key,
                'board_count': 0,
                'total_standard_area_sqft': 0.0,  # Total board area available in sqft
                'total_utilized_area_sqft': 0.0,  # Actual area used by parts in sqft
                'unit_price_per_sqft': unit_price_sqft,
                'total_cost': 0.0
            }
        
        core_summary[core_material_key]['board_count'] += 1
        core_summary[core_material_key]['total_standard_area_sqft'] += board_area_sqft
        core_summary[core_material_key]['total_utilized_area_sqft'] += utilized_area_sqft
        core_summary[core_material_key]['total_cost'] += board_area_sqft * core_summary[core_material_key]['unit_price_per_sqft']
    
    # Calculate wastage percentages
    for summary in core_summary.values():
        if summary['total_standard_area_sqft'] > 0:
            summary['wastage_area_sqft'] = summary['total_standard_area_sqft'] - summary['total_utilized_area_sqft']
            summary['wastage_percentage'] = (summary['wastage_area_sqft'] / summary['total_standard_area_sqft']) * 100
            summary['utilization_percentage'] = (summary['total_utilized_area_sqft'] / summary['total_standard_area_sqft']) * 100
        else:
            summary['wastage_area_sqft'] = 0.0
            summary['wastage_percentage'] = 0.0
            summary['utilization_percentage'] = 0.0
    
    return list(core_summary.values())


def calculate_laminate_type_summary(boards: List[Board], laminate_db: Dict = None) -> List[Dict]:
    """Calculate summary of laminates by type with costs and utilization."""
    laminate_usage = {}
    
    for board in boards:
        # Process top and bottom laminates separately
        top_laminate = board.material_details.top_laminate_name
        bottom_laminate = board.material_details.bottom_laminate_name
        
        board_area_sqm = (board.total_length * board.total_width) / 1_000_000
        board_area_sqft = board_area_sqm * 10.764  # Convert to square feet
        
        # Calculate actual laminate area needed for parts
        actual_laminate_area_sqm = 0.0
        for part in board.parts_on_board:
            actual_length = getattr(part, 'actual_length', None) or part.requested_length
            actual_width = getattr(part, 'actual_width', None) or part.requested_width
            part_area = (actual_length * actual_width) / 1_000_000
            actual_laminate_area_sqm += part_area
        
        actual_laminate_area_sqft = actual_laminate_area_sqm * 10.764  # Convert to square feet
        
        # Process top laminate
        for laminate_name, position in [(top_laminate, 'Top'), (bottom_laminate, 'Bottom')]:
            laminate_key = f"{laminate_name} ({position})"
            
            if laminate_key not in laminate_usage:
                # Get unit price from laminate database (per square meter)
                unit_price_sqm = 0.0
                if laminate_db and laminate_name in laminate_db:
                    unit_price_sqm = laminate_db[laminate_name]
                
                unit_price_sqft = unit_price_sqm / 10.764  # Convert to per sqft (1 sqm = 10.764 sqft)
                
                laminate_usage[laminate_key] = {
                    'laminate_type': laminate_key,
                    'laminate_count': 0,  # Count of laminate sheets
                    'total_standard_area_sqft': 0.0,  # Total laminate area purchased in sqft
                    'total_utilized_area_sqft': 0.0,  # Actual laminate area used in sqft
                    'unit_price_per_sqft': unit_price_sqft,
                    'total_cost': 0.0
                }
            
            # Each board uses 1 laminate sheet per side
            laminate_usage[laminate_key]['laminate_count'] += 1
            laminate_usage[laminate_key]['total_standard_area_sqft'] += board_area_sqft
            laminate_usage[laminate_key]['total_utilized_area_sqft'] += actual_laminate_area_sqft
            laminate_usage[laminate_key]['total_cost'] += board_area_sqft * laminate_usage[laminate_key]['unit_price_per_sqft']
    
    # Calculate wastage percentages
    for summary in laminate_usage.values():
        if summary['total_standard_area_sqft'] > 0:
            summary['wastage_area_sqft'] = summary['total_standard_area_sqft'] - summary['total_utilized_area_sqft']
            summary['wastage_percentage'] = (summary['wastage_area_sqft'] / summary['total_standard_area_sqft']) * 100
            summary['utilization_percentage'] = (summary['total_utilized_area_sqft'] / summary['total_standard_area_sqft']) * 100
        else:
            summary['wastage_area_sqft'] = 0.0
            summary['wastage_percentage'] = 0.0
            summary['utilization_percentage'] = 0.0
    
    return list(laminate_usage.values())


def generate_cutting_layout_pdf(boards: List[Board], output_path=None, order_name: str = "") -> bytes:
    """
    Generate PDF with cutting layouts for all boards.
    
    Args:
        boards: List of Board objects with placed parts
        output_path: Optional file path to save PDF (if None, returns bytes)
        order_name: Order name to include in report header
        
    Returns:
        PDF content as bytes
    """
    try:
        # Create PDF in memory
        pdf_buffer = io.BytesIO()
        
        with PdfPages(pdf_buffer) as pdf:
            for board in boards:
                fig, ax = plt.subplots(1, 1, figsize=(11.7, 8.3))  # A4 landscape
                
                # Set up the plot
                ax.set_xlim(0, board.total_length)
                ax.set_ylim(0, board.total_width)
                ax.set_aspect('equal')
                
                # Create title with order name if provided
                title_lines = []
                if order_name:
                    title_lines.append(f'Order: {order_name}')
                title_lines.extend([
                    f'Cutting Layout - {board.id}',
                    f'Material: {board.material_details.full_material_string}',
                    f'Board Size: {board.total_length}mm x {board.total_width}mm',
                    f'Utilization: {board.get_utilization_percentage():.1f}%',
                    f'Symbols: ⬆️ = Upgraded Material, 🔄 = Rotated Part'
                ])
                
                ax.set_title('\n'.join(title_lines), fontsize=9, pad=20)
                
                # Draw board outline
                board_rect = patches.Rectangle(
                    (0, 0), board.total_length, board.total_width,
                    linewidth=2, edgecolor='black', facecolor='lightgray', alpha=0.3
                )
                ax.add_patch(board_rect)
                
                # Color palette for parts
                colors = plt.cm.Set3(range(len(board.parts_on_board)))
                
                # Draw parts
                for i, part in enumerate(board.parts_on_board):
                    # Get safe attributes with defaults first, ensure numeric
                    x_pos = float(getattr(part, 'x_pos', 0.0) or 0.0)
                    y_pos = float(getattr(part, 'y_pos', 0.0) or 0.0)
                    
                    if x_pos is not None and y_pos is not None:
                        # Choose border color based on upgrade status
                        is_upgraded = getattr(part, 'is_upgraded', False)
                        border_color = 'red' if is_upgraded else 'black'
                        border_width = 2 if is_upgraded else 1
                        
                        # Get dimensions with defaults, ensure numeric
                        actual_length = float(getattr(part, 'actual_length', part.requested_length) or part.requested_length)
                        actual_width = float(getattr(part, 'actual_width', part.requested_width) or part.requested_width)
                        
                        # Draw part rectangle
                        part_rect = patches.Rectangle(
                            (x_pos, y_pos),
                            actual_length, actual_width,
                            linewidth=border_width, edgecolor=border_color,
                            facecolor=colors[i % len(colors)], alpha=0.7
                        )
                        ax.add_patch(part_rect)
                        
                        # Add part label with safe numeric calculations
                        center_x = float(x_pos) + float(actual_length) / 2
                        center_y = float(y_pos) + float(actual_width) / 2
                        
                        # Enhanced visual indicators for rotation and upgrades
                        rotation_text = " ↻" if getattr(part, 'rotated', False) else ""
                        upgrade_text = " ⬆" if getattr(part, 'is_upgraded', False) else ""
                        part_id = getattr(part, 'id', getattr(part, 'part_id', 'Unknown'))
                        label = f"{part_id}{rotation_text}{upgrade_text}\n{actual_length:.0f}×{actual_width:.0f}"
                        
                        ax.text(center_x, center_y, label,
                               ha='center', va='center', fontsize=8,
                               bbox=dict(boxstyle="round,pad=0.3", facecolor='white', alpha=0.8))
                
                # Add basic cutting guides for small datasets only
                if len(boards) <= 10:  # Only for small datasets to improve performance
                    cut_lines_x = set()
                    cut_lines_y = set()
                    
                    for part in board.parts_on_board:
                        x_pos = float(getattr(part, 'x_pos', 0.0) or 0.0)
                        y_pos = float(getattr(part, 'y_pos', 0.0) or 0.0)
                        actual_length = float(getattr(part, 'actual_length', part.requested_length) or part.requested_length)
                        actual_width = float(getattr(part, 'actual_width', part.requested_width) or part.requested_width)
                        
                        if x_pos > 0:
                            cut_lines_x.add(x_pos)
                        cut_lines_x.add(x_pos + actual_length)
                        
                        if y_pos > 0:
                            cut_lines_y.add(y_pos)
                        cut_lines_y.add(y_pos + actual_width)
                    
                    # Draw cut lines
                    for x in cut_lines_x:
                        if 0 < x < board.total_length:
                            ax.axvline(x=x, color='red', linestyle='-', alpha=0.5, linewidth=0.5)
                    
                    for y in cut_lines_y:
                        if 0 < y < board.total_width:
                            ax.axhline(y=y, color='red', linestyle='-', alpha=0.5, linewidth=0.5)
                
                # Add legend
                legend_elements = []
                for i, part in enumerate(board.parts_on_board):
                    legend_elements.append(
                        patches.Patch(color=colors[i % len(colors)], 
                                    label=f"{part.id} ({part.actual_length:.0f}x{part.actual_width:.0f})")
                    )
                
                if legend_elements:
                    ax.legend(handles=legend_elements, loc='center left', bbox_to_anchor=(1, 0.5))
                
                # Set labels and minimal grid
                ax.set_xlabel('Length (mm)')
                ax.set_ylabel('Width (mm)')
                ax.grid(True, alpha=0.2, linestyle='-', linewidth=0.5)
                
                # Invert y-axis to match typical cutting layout orientation
                ax.invert_yaxis()
                
                plt.tight_layout()
                pdf.savefig(fig, bbox_inches='tight')
                plt.close(fig)
        
        # Get PDF bytes
        pdf_bytes = pdf_buffer.getvalue()
        pdf_buffer.close()
        
        # Save to file if path provided
        if output_path and isinstance(output_path, str):
            with open(output_path, 'wb') as f:
                f.write(pdf_bytes)
            logger.info(f"Cutting layout PDF saved to {output_path}")
        
        return pdf_bytes
        
    except Exception as e:
        logger.error(f"Error generating cutting layout PDF: {e}")
        raise


def generate_optimized_cutlist_excel(boards: List[Board], unplaced_parts: List[Part], 
                                   upgrade_summary: List[Dict], initial_cost: float, 
                                   final_cost: float, core_db=None, 
                                   laminate_db=None, output_path=None,
                                   order_name: str = "") -> bytes:
    """
    Generate Excel report with optimized cutlist and summary information.
    
    Args:
        boards: List of boards with placed parts
        unplaced_parts: List of parts that couldn't be placed
        upgrade_summary: List of material upgrade information
        initial_cost: Initial cost estimate
        final_cost: Final optimized cost
        output_path: Optional file path to save Excel file
        
    Returns:
        Excel content as bytes
    """
    try:
        # Create Excel file in memory
        excel_buffer = io.BytesIO()
        
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            
            # Sheet 1: Optimized Cutlist
            cutlist_data = []
            for board in boards:
                for part in board.parts_on_board:
                    cutlist_data.append({
                        'Part ID': part.id,
                        'Board ID': part.assigned_board_id,
                        'Original Length (mm)': part.requested_length,
                        'Original Width (mm)': part.requested_width,
                        'Actual Length (mm)': part.actual_length,
                        'Actual Width (mm)': part.actual_width,
                        'X Position (mm)': part.x_pos,
                        'Y Position (mm)': part.y_pos,
                        'Rotated': 'Yes' if part.rotated else 'No',
                        'Original Material': str(part.material_details.full_material_string),
                        'Assigned Material': str(part.assigned_material_details.full_material_string if part.assigned_material_details else part.material_details.full_material_string),
                        'Material Upgraded': 'Yes' if part.is_upgraded else 'No',
                        'Grain Direction': 'Sensitive' if part.grains == 1 else 'Free'
                    })
            
            cutlist_df = pd.DataFrame(cutlist_data)
            cutlist_df.to_excel(writer, sheet_name='Optimized Cutlist', index=False)
            
            # Sheet 2: Board Summary
            board_summary_data = []
            for board in boards:
                board_summary_data.append({
                    'Board ID': board.id,
                    'Material': str(board.material_details.full_material_string),
                    'Board Length (mm)': board.total_length,
                    'Board Width (mm)': board.total_width,
                    'Total Board Area (mm²)': board.total_length * board.total_width,
                    'Parts Count': len(board.parts_on_board),
                    'Utilization (%)': board.get_utilization_percentage(),
                    'Remaining Area (mm²)': board.get_remaining_area(),
                    'Available Offcuts': len(board.available_rectangles)
                })
            
            board_summary_df = pd.DataFrame(board_summary_data)
            board_summary_df.to_excel(writer, sheet_name='Board Summary', index=False)
            
            # Sheet 3: Material Upgrades
            upgrade_data = []
            if upgrade_summary:
                try:
                    # Handle different upgrade_summary formats
                    if isinstance(upgrade_summary, dict):
                        # Enhanced optimization format
                        upgrades_by_material = upgrade_summary.get('upgrades_by_material', {})
                        for upgrade_path, count in upgrades_by_material.items():
                            if ' -> ' in str(upgrade_path):
                                original, upgraded = str(upgrade_path).split(' -> ', 1)
                                upgrade_data.append({
                                    'Original Material': str(original),
                                    'Upgraded Material': str(upgraded),
                                    'Parts Count': int(count)
                                })
                    elif isinstance(upgrade_summary, list):
                        if upgrade_summary and isinstance(upgrade_summary[0], dict):
                            # New dictionary format with Part ID, Original Material, Upgraded Material
                            upgrades_by_material = {}
                            for upgrade_dict in upgrade_summary:
                                orig_mat = str(upgrade_dict.get('Original Material', 'Unknown'))
                                upg_mat = str(upgrade_dict.get('Upgraded Material', 'Unknown'))
                                upgrade_path = f"{orig_mat} -> {upg_mat}"
                                upgrades_by_material[upgrade_path] = upgrades_by_material.get(upgrade_path, 0) + 1
                            
                            for upgrade_path, count in upgrades_by_material.items():
                                if ' -> ' in upgrade_path:
                                    original, upgraded = upgrade_path.split(' -> ', 1)
                                    upgrade_data.append({
                                        'Original Material': str(original),
                                        'Upgraded Material': str(upgraded),
                                        'Parts Count': int(count)
                                    })
                        else:
                            # Standard optimization format (list of tuples)
                            for item in upgrade_summary:
                                if isinstance(item, (list, tuple)) and len(item) >= 3:
                                    original_material, upgraded_material, count = item[:3]
                                    upgrade_data.append({
                                        'Original Material': str(original_material) if original_material is not None else 'Unknown',
                                        'Upgraded Material': str(upgraded_material) if upgraded_material is not None else 'Unknown',
                                        'Parts Count': int(count) if isinstance(count, (int, float)) else 0
                                    })
                except (ValueError, TypeError, AttributeError) as e:
                    logger.warning(f"Error processing upgrade summary: {e}")
                    pass
                
                if upgrade_data:
                    upgrade_df = pd.DataFrame(upgrade_data)
                    upgrade_df.to_excel(writer, sheet_name='Material Upgrades', index=False)
                else:
                    # Create empty sheet with headers
                    empty_upgrade_df = pd.DataFrame(columns=['Original Material', 'Upgraded Material', 'Parts Count'])
                    empty_upgrade_df.to_excel(writer, sheet_name='Material Upgrades', index=False)
            else:
                # Create empty sheet with headers
                empty_upgrade_df = pd.DataFrame(columns=['Original Material', 'Upgraded Material', 'Parts Count'])
                empty_upgrade_df.to_excel(writer, sheet_name='Material Upgrades', index=False)
            
            # Sheet 4: Unplaced Parts
            if unplaced_parts:
                unplaced_data = []
                for part in unplaced_parts:
                    unplaced_data.append({
                        'Part ID': part.id,
                        'Length (mm)': part.requested_length,
                        'Width (mm)': part.requested_width,
                        'Material': str(part.material_details.full_material_string),
                        'Grain Direction': 'Sensitive' if part.grains == 1 else 'Free',
                        'Reason': 'Could not fit on any available board'
                    })
                
                unplaced_df = pd.DataFrame(unplaced_data)
                unplaced_df.to_excel(writer, sheet_name='Unplaced Parts', index=False)
            else:
                # Create empty sheet
                empty_unplaced_df = pd.DataFrame(columns=['Part ID', 'Length (mm)', 'Width (mm)', 'Material', 'Grain Direction', 'Reason'])
                empty_unplaced_df.to_excel(writer, sheet_name='Unplaced Parts', index=False)
            
            # Sheet 5: Cost Analysis
            cost_data = [
                ['Metric', 'Value'],
                ['Initial Cost Estimate', f'₹{initial_cost:.2f}'],
                ['Final Optimized Cost', f'₹{final_cost:.2f}'],
                ['Cost Savings', f'₹{initial_cost - final_cost:.2f}'],
                ['Savings Percentage', f'{((initial_cost - final_cost) / initial_cost * 100):.1f}%' if initial_cost > 0 else '0.0%'],
                ['Total Parts', str(len(cutlist_data) + len(unplaced_parts))],
                ['Successfully Placed', str(len(cutlist_data))],
                ['Unplaced Parts', str(len(unplaced_parts))],
                ['Total Boards Used', str(len(boards))],
                ['Material Upgrades', str(len(upgrade_summary) if upgrade_summary else 0)],
                ['Average Board Utilization', f'{sum(board.get_utilization_percentage() for board in boards) / len(boards):.1f}%' if boards else '0.0%']
            ]
            
            cost_df = pd.DataFrame(cost_data[1:], columns=cost_data[0])
            cost_df.to_excel(writer, sheet_name='Cost Analysis', index=False)
            
            # Sheet 6: Core Material Summary
            core_summary = calculate_core_material_summary(boards, core_db)
            if core_summary:
                core_summary_data = []
                for item in core_summary:
                    core_summary_data.append({
                        'Core Material': item['core_material'],
                        'Board Count': item['board_count'],
                        'Standard Area (sqft)': f"{item['total_standard_area_sqft']:.2f}",
                        'Utilized Area (sqft)': f"{item['total_utilized_area_sqft']:.2f}",
                        'Wastage Area (sqft)': f"{item['wastage_area_sqft']:.2f}",
                        'Utilization %': f"{item['utilization_percentage']:.1f}%",
                        'Wastage %': f"{item['wastage_percentage']:.1f}%",
                        'Unit Price (₹/sqft)': f"₹{item['unit_price_per_sqft']:.2f}",
                        'Total Cost (₹)': f"₹{item['total_cost']:.2f}"
                    })
                
                core_summary_df = pd.DataFrame(core_summary_data)
                core_summary_df.to_excel(writer, sheet_name='Core Material Summary', index=False)
            else:
                empty_core_df = pd.DataFrame(columns=['Core Material', 'Board Count', 'Standard Area (sqft)', 'Utilized Area (sqft)', 'Wastage Area (sqft)', 'Utilization %', 'Wastage %', 'Unit Price (₹/sqft)', 'Total Cost (₹)'])
                empty_core_df.to_excel(writer, sheet_name='Core Material Summary', index=False)
            
            # Sheet 7: Laminate Type Summary
            laminate_summary = calculate_laminate_type_summary(boards, laminate_db)
            if laminate_summary:
                laminate_summary_data = []
                for item in laminate_summary:
                    laminate_summary_data.append({
                        'Laminate Type': item['laminate_type'],
                        'Laminate Count': item['laminate_count'],
                        'Standard Area (sqft)': f"{item['total_standard_area_sqft']:.2f}",
                        'Utilized Area (sqft)': f"{item['total_utilized_area_sqft']:.2f}",
                        'Wastage Area (sqft)': f"{item['wastage_area_sqft']:.2f}",
                        'Utilization %': f"{item['utilization_percentage']:.1f}%",
                        'Wastage %': f"{item['wastage_percentage']:.1f}%",
                        'Unit Price (₹/sqft)': f"₹{item['unit_price_per_sqft']:.2f}",
                        'Total Cost (₹)': f"₹{item['total_cost']:.2f}"
                    })
                
                laminate_summary_df = pd.DataFrame(laminate_summary_data)
                laminate_summary_df.to_excel(writer, sheet_name='Laminate Summary', index=False)
            else:
                empty_laminate_df = pd.DataFrame(columns=['Laminate Type', 'Laminate Count', 'Standard Area (sqft)', 'Utilized Area (sqft)', 'Wastage Area (sqft)', 'Utilization %', 'Wastage %', 'Unit Price (₹/sqft)', 'Total Cost (₹)'])
                empty_laminate_df.to_excel(writer, sheet_name='Laminate Summary', index=False)
            
            # Sheet 6: Offcuts Available
            offcuts_data = []
            for board in boards:
                for offcut in board.available_rectangles:
                    if offcut.get_area() > 10000:  # Only list significant offcuts (>100cm²)
                        offcuts_data.append({
                            'Offcut ID': offcut.id,
                            'Source Board': offcut.source_board_id,
                            'X Position (mm)': offcut.x,
                            'Y Position (mm)': offcut.y,
                            'Length (mm)': offcut.length,
                            'Width (mm)': offcut.width,
                            'Area (mm²)': offcut.get_area(),
                            'Material': str(offcut.material_details.full_material_string)
                        })
            
            if offcuts_data:
                offcuts_df = pd.DataFrame(offcuts_data)
                offcuts_df = offcuts_df.sort_values('Area (mm²)', ascending=False)
                offcuts_df.to_excel(writer, sheet_name='Available Offcuts', index=False)
            else:
                empty_offcuts_df = pd.DataFrame(columns=['Offcut ID', 'Source Board', 'X Position (mm)', 'Y Position (mm)', 'Length (mm)', 'Width (mm)', 'Area (mm²)', 'Material'])
                empty_offcuts_df.to_excel(writer, sheet_name='Available Offcuts', index=False)
        
        # Get Excel bytes
        excel_bytes = excel_buffer.getvalue()
        excel_buffer.close()
        
        # Save to file if path provided
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(excel_bytes)
            logger.info(f"Optimized cutlist Excel saved to {output_path}")
        
        return excel_bytes
        
    except Exception as e:
        logger.error(f"Error generating Excel report: {e}")
        raise


def generate_material_usage_report(boards: List[Board], core_db: Dict, 
                                 laminate_db: Dict) -> pd.DataFrame:
    """
    Generate detailed material usage report.
    
    Args:
        boards: List of boards used in optimization
        core_db: Core materials database
        laminate_db: Laminates database
        
    Returns:
        DataFrame with material usage statistics
    """
    try:
        material_usage = {}
        
        for board in boards:
            material_key = (board.material_details.laminate_name,
                           board.material_details.core_name,
                           board.material_details.thickness)
            
            if material_key not in material_usage:
                material_usage[material_key] = {
                    'laminate_name': board.material_details.laminate_name,
                    'core_name': board.material_details.core_name,
                    'thickness': board.material_details.thickness,
                    'boards_used': 0,
                    'total_area': 0,
                    'utilized_area': 0,
                    'total_cost': 0
                }
            
            usage = material_usage[material_key]
            usage['boards_used'] += 1
            
            board_area = board.total_length * board.total_width
            usage['total_area'] += board_area
            usage['utilized_area'] += board_area - board.get_remaining_area()
            
            # Calculate cost
            board_cost = board.material_details.get_cost_per_sqm(laminate_db, core_db)
            board_area_sqm = board_area / 1_000_000
            usage['total_cost'] += board_cost * board_area_sqm
        
        # Convert to DataFrame
        usage_data = []
        for usage in material_usage.values():
            utilization_pct = (usage['utilized_area'] / usage['total_area'] * 100) if usage['total_area'] > 0 else 0
            
            usage_data.append({
                'Material': f"{usage['laminate_name']}_{usage['core_name']}_{usage['laminate_name']}",
                'Thickness (mm)': usage['thickness'],
                'Boards Used': usage['boards_used'],
                'Total Area (m²)': usage['total_area'] / 1_000_000,
                'Utilized Area (m²)': usage['utilized_area'] / 1_000_000,
                'Utilization (%)': utilization_pct,
                'Total Cost': usage['total_cost']
            })
        
        usage_df = pd.DataFrame(usage_data)
        usage_df = usage_df.sort_values('Total Cost', ascending=False)
        
        return usage_df
        
    except Exception as e:
        logger.error(f"Error generating material usage report: {e}")
        return pd.DataFrame()
