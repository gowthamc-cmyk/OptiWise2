"""
Simple report generation for OptiWise without pandas/matplotlib dependencies.
Creates basic text and CSV reports for optimization results.
"""

import csv
import io
from typing import List, Dict, Any
from data_models import Board, Part


def generate_cutting_layout_text(boards: List[Board], order_name: str = "") -> str:
    """
    Generate text-based cutting layout report.
    
    Args:
        boards: List of Board objects with placed parts
        order_name: Order name to include in report header
        
    Returns:
        Formatted text report
    """
    report_lines = []
    
    # Header
    if order_name:
        report_lines.append(f"CUTTING LAYOUT REPORT - ORDER: {order_name}")
    else:
        report_lines.append("CUTTING LAYOUT REPORT")
    
    report_lines.append("=" * 60)
    report_lines.append("")
    
    # Summary
    total_boards = len(boards)
    total_parts = sum(len(board.parts_on_board) for board in boards)
    avg_utilization = sum(board.get_utilization_percentage() for board in boards) / total_boards if total_boards > 0 else 0
    
    report_lines.append("SUMMARY:")
    report_lines.append(f"Total Boards: {total_boards}")
    report_lines.append(f"Total Parts: {total_parts}")
    report_lines.append(f"Average Utilization: {avg_utilization:.1f}%")
    report_lines.append("")
    
    # Individual board details
    for i, board in enumerate(boards, 1):
        report_lines.append(f"BOARD {i}: {board.id}")
        report_lines.append(f"Material: {board.material_details}")
        report_lines.append(f"Size: {board.total_length}mm x {board.total_width}mm")
        report_lines.append(f"Utilization: {board.get_utilization_percentage():.1f}%")
        report_lines.append(f"Parts Count: {len(board.parts_on_board)}")
        report_lines.append("")
        
        if board.parts_on_board:
            report_lines.append("PARTS ON BOARD:")
            report_lines.append("Part ID".ljust(20) + "Dimensions".ljust(15) + "Position".ljust(15) + "Notes")
            report_lines.append("-" * 70)
            
            for part in board.parts_on_board:
                part_id = str(part.id)[:19]
                dimensions = f"{part.requested_length}x{part.requested_width}"
                position = f"({getattr(part, 'x_pos', 0):.0f},{getattr(part, 'y_pos', 0):.0f})"
                
                notes = []
                if getattr(part, 'rotated', False):
                    notes.append("Rotated")
                if getattr(part, 'is_upgraded', False):
                    notes.append("Upgraded")
                if part.grains == 1:
                    notes.append("Grain-sensitive")
                
                notes_str = ", ".join(notes) if notes else ""
                
                report_lines.append(
                    part_id.ljust(20) + 
                    dimensions.ljust(15) + 
                    position.ljust(15) + 
                    notes_str
                )
        
        report_lines.append("")
        report_lines.append("-" * 60)
        report_lines.append("")
    
    return "\n".join(report_lines)


def generate_optimized_cutlist_csv(boards: List[Board], unplaced_parts: List[Part], 
                                 upgrade_summary: List[Dict], initial_cost: float, 
                                 final_cost: float, order_name: str = "") -> str:
    """
    Generate CSV report with optimized cutlist.
    
    Returns:
        CSV content as string
    """
    output = io.StringIO()
    
    # Write header information
    if order_name:
        output.write(f"# OptiWise Optimization Report - Order: {order_name}\n")
    else:
        output.write("# OptiWise Optimization Report\n")
    
    output.write(f"# Initial Cost: ₹{initial_cost:.2f}\n")
    output.write(f"# Final Cost: ₹{final_cost:.2f}\n")
    output.write(f"# Cost Savings: ₹{initial_cost - final_cost:.2f}\n")
    output.write(f"# Total Boards: {len(boards)}\n")
    output.write(f"# Unplaced Parts: {len(unplaced_parts)}\n")
    output.write("#\n")
    
    # Main cutlist
    writer = csv.writer(output)
    writer.writerow([
        'Part ID', 'Board ID', 'Original Length (mm)', 'Original Width (mm)',
        'X Position (mm)', 'Y Position (mm)', 'Rotated', 'Original Material',
        'Assigned Material', 'Material Upgraded', 'Grain Direction'
    ])
    
    for board in boards:
        for part in board.parts_on_board:
            writer.writerow([
                part.id,
                getattr(part, 'assigned_board_id', board.id),
                part.requested_length,
                part.requested_width,
                getattr(part, 'x_pos', 0),
                getattr(part, 'y_pos', 0),
                'Yes' if getattr(part, 'rotated', False) else 'No',
                str(part.material_details),
                str(getattr(part, 'assigned_material_details', part.material_details)),
                'Yes' if getattr(part, 'is_upgraded', False) else 'No',
                'Sensitive' if part.grains == 1 else 'Free'
            ])
    
    # Add empty line and unplaced parts section
    if unplaced_parts:
        output.write("\n# UNPLACED PARTS\n")
        writer.writerow(['Part ID', 'Length (mm)', 'Width (mm)', 'Material', 'Grain Direction', 'Reason'])
        for part in unplaced_parts:
            writer.writerow([
                part.id,
                part.requested_length,
                part.requested_width,
                str(part.material_details),
                'Sensitive' if part.grains == 1 else 'Free',
                'Could not fit on any available board'
            ])
    
    # Add board summary section
    output.write("\n# BOARD SUMMARY\n")
    writer.writerow(['Board ID', 'Material', 'Length (mm)', 'Width (mm)', 'Parts Count', 'Utilization (%)', 'Remaining Area (mm²)'])
    for board in boards:
        writer.writerow([
            board.id,
            str(board.material_details),
            board.total_length,
            board.total_width,
            len(board.parts_on_board),
            f"{board.get_utilization_percentage():.1f}",
            f"{board.get_remaining_area():.0f}"
        ])
    
    return output.getvalue()


def generate_material_summary_csv(boards: List[Board]) -> str:
    """
    Generate material-wise summary as CSV.
    
    Returns:
        CSV content as string
    """
    output = io.StringIO()
    writer = csv.writer(output)
    
    # Calculate material summary
    material_summary = {}
    
    for board in boards:
        material_key = str(board.material_details)
        
        if material_key not in material_summary:
            material_summary[material_key] = {
                'board_count': 0,
                'total_board_area': 0,
                'utilized_area': 0,
                'wastage_area': 0,
                'parts_placed': 0
            }
        
        board_area = board.total_length * board.total_width / 1_000_000
        utilized_area = (board_area * board.get_utilization_percentage() / 100)
        wastage_area = board_area - utilized_area
        
        summary = material_summary[material_key]
        summary['board_count'] += 1
        summary['total_board_area'] += board_area
        summary['utilized_area'] += utilized_area
        summary['wastage_area'] += wastage_area
        summary['parts_placed'] += len(board.parts_on_board)
    
    # Write material summary
    writer.writerow(['Material', 'Board Count', 'Total Area (m²)', 'Utilized Area (m²)', 
                    'Waste Area (m²)', 'Utilization (%)', 'Parts Placed'])
    
    for material, summary in material_summary.items():
        utilization = (summary['utilized_area'] / summary['total_board_area'] * 100) if summary['total_board_area'] > 0 else 0
        writer.writerow([
            material,
            summary['board_count'],
            f"{summary['total_board_area']:.2f}",
            f"{summary['utilized_area']:.2f}",
            f"{summary['wastage_area']:.2f}",
            f"{utilization:.1f}",
            summary['parts_placed']
        ])
    
    return output.getvalue()


def generate_upgrade_summary_csv(upgrade_summary: List[Dict]) -> str:
    """
    Generate material upgrade summary as CSV.
    
    Returns:
        CSV content as string
    """
    output = io.StringIO()
    writer = csv.writer(output)
    
    writer.writerow(['Original Material', 'Upgraded Material', 'Parts Count'])
    
    if upgrade_summary:
        # Process different upgrade summary formats
        upgrade_counts = {}
        
        for item in upgrade_summary:
            if isinstance(item, dict):
                orig_mat = item.get('original_material', 'Unknown')
                upg_mat = item.get('upgraded_material', 'Unknown')
                key = f"{orig_mat} -> {upg_mat}"
                upgrade_counts[key] = upgrade_counts.get(key, 0) + 1
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                orig_mat, upg_mat = item[:2]
                key = f"{orig_mat} -> {upg_mat}"
                upgrade_counts[key] = upgrade_counts.get(key, 0) + 1
        
        for upgrade_path, count in upgrade_counts.items():
            if ' -> ' in upgrade_path:
                original, upgraded = upgrade_path.split(' -> ', 1)
                writer.writerow([original, upgraded, count])
    
    return output.getvalue()


def create_comprehensive_report_package(boards: List[Board], unplaced_parts: List[Part],
                                      upgrade_summary: List[Dict], initial_cost: float,
                                      final_cost: float, order_name: str = "") -> Dict[str, str]:
    """
    Create a comprehensive package of all reports.
    
    Returns:
        Dictionary with report names as keys and content as values
    """
    reports = {}
    
    # Text cutting layout
    reports['cutting_layout.txt'] = generate_cutting_layout_text(boards, order_name)
    
    # CSV optimization report
    reports['optimization_report.csv'] = generate_optimized_cutlist_csv(
        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, order_name
    )
    
    # Material summary
    reports['material_summary.csv'] = generate_material_summary_csv(boards)
    
    # Upgrade summary
    if upgrade_summary:
        reports['upgrade_summary.csv'] = generate_upgrade_summary_csv(upgrade_summary)
    
    return reports