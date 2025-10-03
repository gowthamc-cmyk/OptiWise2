"""
Enhanced PDF cutting layout generator for OptiWise.
Creates visual cutting layouts with upgrade indicators, rotation symbols, and proper formatting.
"""

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.colors as mcolors
import numpy as np
from typing import List, Dict, Any, Tuple, Optional
import logging
import io
from data_models import Board, Part

logger = logging.getLogger(__name__)

class PDFLayoutGenerator:
    """Generate visual PDF cutting layouts with enhanced indicators."""
    
    def __init__(self):
        # Exact color palette matching the reference image
        self.colors = [
            '#B2DFDB',  # Light teal (like top-left part)
            '#FFF9C4',  # Light yellow (like top-right part) 
            '#F8BBD9',  # Light pink (like bottom parts)
            '#C8E6C9',  # Light green
            '#E1BEE7',  # Light purple (like center parts)
            '#FFCCBC',  # Light orange
            '#DCEDC8',  # Light lime
            '#FFCDD2',  # Light red
            '#D1C4E9',  # Light deep purple
            '#B3E5FC',  # Light cyan
            '#F0F4C3',  # Light lime yellow
            '#FFAB91',  # Light deep orange
            '#CE93D8',  # Light purple
            '#90CAF9',  # Light blue
            '#C5E1A5'   # Light light green
        ]
        self.border_colors = [
            '#000000',  # Black borders for normal parts
            '#000000', '#000000', '#000000', '#000000',
            '#000000', '#000000', '#000000', '#000000',
            '#000000', '#000000', '#000000', '#000000',
            '#000000', '#000000'
        ]
    
    def generate_cutting_layouts_pdf(self, boards: List[Board], order_name: str = "OptiWise", 
                                   output_path: Optional[str] = None) -> bytes:
        """Generate comprehensive PDF cutting layouts with visual indicators."""
        
        pdf_buffer = io.BytesIO()
        
        try:
            with PdfPages(pdf_buffer) as pdf:
                for board_idx, board in enumerate(boards):
                    self._create_board_layout_page(pdf, board, board_idx, order_name)
            
            pdf_bytes = pdf_buffer.getvalue()
            pdf_buffer.close()
            
            if output_path:
                try:
                    with open(output_path, 'wb') as f:
                        f.write(pdf_bytes)
                    logger.info(f"PDF cutting layout saved to {output_path}")
                except Exception as e:
                    logger.warning(f"Could not save PDF to file: {e}")
            
            return pdf_bytes
            
        except Exception as e:
            logger.error(f"Error generating PDF layout: {e}")
            raise
    
    def _create_board_layout_page(self, pdf: PdfPages, board: Board, board_idx: int, order_name: str):
        """Create a single board layout page with fixed positioning and no overlaps."""
        
        # Create figure with proper proportions
        fig, ax = plt.subplots(1, 1, figsize=(11, 14))  # Taller figure for clean separation
        fig.patch.set_facecolor('white')
        
        # Adjust subplot positioning to prevent overlaps
        plt.subplots_adjust(left=0.08, right=0.92, top=0.88, bottom=0.35)  # More space at bottom for table
        
        # Extract board information
        board_material = self._extract_material_info(board)
        board_size = f"{board.total_length:.0f} mm x {board.total_width:.0f} mm"
        utilization = board.get_utilization_percentage()
        
        # Enhanced header with CLIENT NAME support
        # Try to get CLIENT NAME from first part if order_name is empty
        if not order_name and board.parts_on_board:
            client_name = getattr(board.parts_on_board[0], 'client_name', '')
            if client_name:
                order_name = client_name
        
        header_lines = [
            f"Order: {order_name}",
            f"Cutting Layout - Board {board_idx + 1}_{board_material['core']}",
            f"Material: {board_material['full']}",
            f"Board Size: {board_size}",
            f"Utilization: {utilization:.1f}%",
            f"Symbols: ↑ = Upgraded Material, ↻ = Rotated Part"
        ]
        
        # Position header text higher for the moved board layout
        header_y_start = 0.96
        for i, line in enumerate(header_lines):
            weight = 'bold' if i < 3 else 'normal'
            size = 10 if i < 3 else 9
            fig.text(0.5, header_y_start - i*0.02, line,  # Reduced spacing from 0.025 to 0.02
                    ha='center', va='top', fontsize=size, fontweight=weight)
        
        # Set up exact coordinate system like reference
        ax.set_xlim(0, board.total_length)
        ax.set_ylim(0, board.total_width)
        ax.set_aspect('equal')
        
        # Simple clean board outline - no shadows
        board_rect = patches.Rectangle(
            (0, 0), board.total_length, board.total_width,
            linewidth=2, edgecolor='black', facecolor='white'
        )
        ax.add_patch(board_rect)
        
        # Create legend data
        legend_data = []
        
        # Place parts on the board
        for i, part in enumerate(board.parts_on_board):
            part_info = self._place_part_on_layout(ax, part, i, board)
            if part_info:
                legend_data.append(part_info)
        
        # No cutting guide lines - clean layout like reference
        
        # Create comprehensive parts summary table
        self._add_parts_summary_table(fig, ax, board)
        
        # Skip legend creation to prevent overlaps - table provides all info
        # self._create_legend(fig, ax, legend_data)
        
        # Axis formatting matching reference exactly
        ax.set_xlabel('Length (mm)', fontsize=10)
        ax.set_ylabel('Width (mm)', fontsize=10)
        
        # Clean white background with minimal grid
        ax.set_facecolor('white')
        ax.grid(False)  # No grid like reference
        
        # Simple black borders only
        for spine in ax.spines.values():
            spine.set_edgecolor('black')
            spine.set_linewidth(1)
        
        # Keep standard coordinate system to prevent overlapping display
        # ax.invert_yaxis()  # Removed to fix overlapping panels issue
        
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight', dpi=300, facecolor='white')
        plt.close(fig)
    
    def _extract_material_info(self, board: Board) -> Dict[str, str]:
        """Extract material information from board."""
        try:
            if hasattr(board, 'material_details') and board.material_details:
                if hasattr(board.material_details, 'full_material_string'):
                    full_material = board.material_details.full_material_string
                    # Extract core material (BWP, HDHMR, MDF, etc.)
                    parts = full_material.split('_')
                    core = next((part for part in parts if any(c in part.upper() for c in ['BWP', 'HDHMR', 'MDF', 'MR'])), 'Unknown')
                    return {'full': full_material, 'core': core}
                else:
                    return {'full': str(board.material_details), 'core': 'Unknown'}
            return {'full': 'Unknown Material', 'core': 'Unknown'}
        except Exception:
            return {'full': 'Unknown Material', 'core': 'Unknown'}
    
    def _place_part_on_layout(self, ax, part: Part, index: int, board: Board) -> Dict:
        """Place a single part on the cutting layout with visual indicators."""
        try:
            # Get part position
            x_pos, y_pos = self._get_part_position(part)
            
            # Get part dimensions
            length, width = self._get_part_dimensions(part)
            
            # Check for upgrades and rotation
            is_upgraded = self._is_part_upgraded(part, board)
            # Handle both 'rotated' (from TEST algorithms) and 'is_rotated' (from other algorithms)
            is_rotated = getattr(part, 'rotated', False) or getattr(part, 'is_rotated', False)
            
            # Select colors
            face_color = self.colors[index % len(self.colors)]
            edge_color = self.border_colors[index % len(self.border_colors)]
            
            # Red border for upgraded materials, black for normal - matching reference
            if is_upgraded:
                edge_color = 'red'
                linewidth = 2
            else:
                edge_color = 'black'
                linewidth = 1
            
            # Simple rectangle with no shadows - clean reference style
            part_rect = patches.Rectangle(
                (x_pos, y_pos), length, width,
                linewidth=linewidth, edgecolor=edge_color,
                facecolor=face_color, alpha=1.0
            )
            ax.add_patch(part_rect)
            
            # Enhanced part labeling with new format fields
            part_id = getattr(part, 'id', f'Part-{index+1}')
            room_type = getattr(part, 'room_type', '')
            panel_name = getattr(part, 'panel_name', '')
            
            symbols = ""
            if is_upgraded:
                symbols += " ↑"
            if is_rotated:
                symbols += " ↻"
            
            # Simplified part labeling to prevent overlapping
            lines = []
            
            # Always include ORDER ID and dimensions (most critical info)
            lines.append(f"{part_id}{symbols}")
            lines.append(f"{length:.0f}×{width:.0f}")
            
            # Add ROOM TYPE and PANEL NAME only if there's space and they exist
            if length > 400 and width > 300:  # Larger parts can show more info
                if room_type and len(room_type) <= 10:  # Short room types only
                    lines.insert(1, room_type[:10])  # Truncate if needed
                if panel_name and len(panel_name) <= 12:  # Short panel names only
                    lines.insert(-1, panel_name[:12])  # Insert before dimensions
            
            center_x = x_pos + length / 2
            center_y = y_pos + width / 2
            
            # Create label with appropriate font sizes
            label_text = "\n".join(lines)
            
            # Dynamic font sizing based on part size and text length
            if length > 600 and width > 400:
                font_size = 8
            elif length > 400 and width > 250:
                font_size = 7
            elif length > 250 and width > 150:
                font_size = 6
            else:
                font_size = 5  # Very small parts get minimal text
            
            # Further reduce font size if too many lines
            if len(lines) > 3:
                font_size = max(4, font_size - 1)
            
            ax.text(center_x, center_y, label_text,
                   ha='center', va='center', fontsize=font_size, fontweight='bold',
                   bbox=dict(boxstyle="round,pad=0.2", facecolor='white', 
                            edgecolor='black', linewidth=0.5, alpha=0.9))
            
            # Return legend information
            return {
                'id': part_id,
                'dimensions': f"({length:.0f}x{width:.0f})",
                'color': face_color,
                'edge_color': edge_color,
                'upgraded': is_upgraded,
                'rotated': is_rotated
            }
            
        except Exception as e:
            logger.warning(f"Error placing part {getattr(part, 'id', 'unknown')}: {e}")
            return {
                'id': f'Part-{index+1}',
                'dimensions': '(0x0)',
                'color': self.colors[0],
                'edge_color': self.border_colors[0],
                'upgraded': False,
                'rotated': False
            }
    
    def _get_part_position(self, part: Part) -> Tuple[float, float]:
        """Extract part position coordinates from optimization algorithm."""
        # Primary position attributes from optimization algorithm
        if hasattr(part, 'x_pos') and hasattr(part, 'y_pos'):
            x = getattr(part, 'x_pos', 0)
            y = getattr(part, 'y_pos', 0)
            if x is not None and y is not None:
                # Debug log for coordinate verification
                if hasattr(part, 'id'):
                    logging.debug(f"PDF: Part {part.id} using x_pos={x}, y_pos={y}")
                return float(x), float(y)
        
        # Secondary fallback attributes
        position_attrs = [
            ('x', 'y'), ('position_x', 'position_y'), ('pos_x', 'pos_y'),
            ('placement_x', 'placement_y'), ('left', 'bottom'), ('start_x', 'start_y')
        ]
        
        for x_attr, y_attr in position_attrs:
            if hasattr(part, x_attr) and hasattr(part, y_attr):
                x = getattr(part, x_attr, 0)
                y = getattr(part, y_attr, 0)
                if x is not None and y is not None:
                    return float(x), float(y)
        
        return 0.0, 0.0
    
    def _get_part_dimensions(self, part: Part) -> Tuple[float, float]:
        """Get part dimensions as actually placed (considering rotation)."""
        # For TEST algorithms, use actual placed dimensions which already account for rotation
        if hasattr(part, 'actual_length') and hasattr(part, 'actual_width'):
            actual_length = getattr(part, 'actual_length', None)
            actual_width = getattr(part, 'actual_width', None)
            if actual_length is not None and actual_width is not None:
                return float(actual_length), float(actual_width)
        
        # Fallback to requested dimensions with rotation logic
        length = float(getattr(part, 'requested_length', 0))
        width = float(getattr(part, 'requested_width', 0))
        
        # Check if part is rotated (handle both attribute names)
        is_rotated = getattr(part, 'rotated', False) or getattr(part, 'is_rotated', False)
        if is_rotated:
            return width, length  # Swap dimensions for rotated parts
        
        return length, width
    
    def _is_part_upgraded(self, part: Part, board: Board) -> bool:
        """Check if part material was upgraded."""
        try:
            # Only check the direct upgrade flag set during optimization
            # This is the authoritative source of upgrade information
            if hasattr(part, 'is_upgraded') and part.is_upgraded:
                return True
            
            # Alternative: check if part has upgraded_from attribute
            if hasattr(part, 'upgraded_from') and part.upgraded_from:
                return True
                
            # If no explicit upgrade flags are set, part was not upgraded
            return False
        except Exception:
            return False
    
    def _add_major_cutting_guides(self, ax, board: Board):
        """Add only essential cutting guide lines for shop floor use."""
        try:
            # Only add major structural cutting lines
            major_cuts_x = set([0, board.total_length])
            major_cuts_y = set([0, board.total_width])
            
            # Find major structural cuts that span the full width/height
            part_positions = []
            for part in board.parts_on_board:
                x_pos, y_pos = self._get_part_position(part)
                length, width = self._get_part_dimensions(part)
                part_positions.append((x_pos, y_pos, x_pos + length, y_pos + width))
            
            # Add only cuts that create major divisions
            for x_pos, y_pos, x_end, y_end in part_positions:
                # Vertical cuts that span significant height
                if y_pos <= 50 and y_end >= board.total_width - 50:  # Nearly full height
                    major_cuts_x.add(x_pos)
                    major_cuts_x.add(x_end)
                
                # Horizontal cuts that span significant width
                if x_pos <= 50 and x_end >= board.total_length - 50:  # Nearly full width
                    major_cuts_y.add(y_pos)
                    major_cuts_y.add(y_end)
            
            # Draw only essential vertical cuts
            for x in sorted(major_cuts_x):
                if 0 < x < board.total_length:
                    ax.axvline(x=x, color='darkred', linestyle='--', alpha=0.4, linewidth=1)
            
            # Draw only essential horizontal cuts
            for y in sorted(major_cuts_y):
                if 0 < y < board.total_width:
                    ax.axhline(y=y, color='darkred', linestyle='--', alpha=0.4, linewidth=1)
                    
        except Exception as e:
            logger.warning(f"Error adding major cutting guides: {e}")
    
    def _add_parts_summary_table(self, fig, ax, board: Board):
        """Add parts summary table with specified fields."""
        try:
            if not hasattr(board, 'parts_on_board') or not board.parts_on_board:
                return
            
            # Prepare table data
            table_data = []
            headers = ['CLIENT NAME', 'ORDER ID / UNIQUE CODE', 'ROOM TYPE', 'SUB CATEGORY', 
                      'PANEL NAME', 'CUT\nLENGTH', 'CUT\nWIDTH', 'ORIGINAL MATERIAL']
            
            for part in board.parts_on_board:
                # Extract data from original_data or fallback to part attributes
                def get_field(field_name, fallback=''):
                    if hasattr(part, 'original_data') and part.original_data:
                        return str(part.original_data.get(field_name, fallback))
                    return str(getattr(part, field_name.lower().replace(' ', '_'), fallback))
                
                client_name = get_field('CLIENT NAME', 'N/A')[:15]
                order_id = get_field('ORDER ID / UNIQUE CODE', getattr(part, 'id', 'N/A'))[:20]
                room_type = get_field('ROOM TYPE', '')[:12] or 'N/A'
                sub_category = get_field('SUB CATEGORY', '')[:12] or 'N/A'
                panel_name = get_field('PANEL NAME', '')[:15] or 'N/A'
                cut_length = f"{getattr(part, 'requested_length', 0):.0f}"
                cut_width = f"{getattr(part, 'requested_width', 0):.0f}"
                
                # Get original material from MATERIAL TYPE field
                original_material = get_field('MATERIAL TYPE', '')[:20] or 'N/A'
                
                row = [client_name, order_id, room_type, sub_category, panel_name, cut_length, cut_width, original_material]
                table_data.append(row)
            
            # Create table with improved positioning to prevent overlaps
            if table_data:
                # Position table at bottom with more space - move further down
                table_ax = fig.add_axes([0.03, 0.01, 0.94, 0.25])  # Start lower, wider, taller
                table_ax.axis('off')
                
                # Create table with optimized column widths for full material name display
                table = table_ax.table(
                    cellText=table_data,
                    colLabels=headers,
                    cellLoc='center',
                    loc='center',
                    bbox=[0, 0, 1, 1],
                    colWidths=[0.13, 0.15, 0.11, 0.11, 0.13, 0.04, 0.04, 0.29]  # Larger ORIGINAL MATERIAL 29%
                )
                
                # Style the table with better spacing and larger font
                table.auto_set_font_size(False)
                table.set_fontsize(7)  # Smaller font to fit more content
                table.scale(1, 1.5)  # More vertical space for readability
                
                # Header styling
                for i in range(len(headers)):
                    table[(0, i)].set_facecolor('#CCCCCC')
                    table[(0, i)].set_text_props(weight='bold')
                
                # Alternate row colors
                for i in range(1, len(table_data) + 1):
                    for j in range(len(headers)):
                        if i % 2 == 0:
                            table[(i, j)].set_facecolor('#F5F5F5')
                        else:
                            table[(i, j)].set_facecolor('white')
                
        except Exception as e:
            logger.warning(f"Error adding parts summary table: {e}")
    
    def _create_legend(self, fig, ax, legend_data: List[Dict]):
        """Create legend matching reference design exactly."""
        if not legend_data:
            return
            
        # Create simple legend elements matching reference
        legend_elements = []
        legend_labels = []
        
        for item in legend_data:
            # Simple colored patch matching reference style
            edge_color = 'red' if item['upgraded'] else 'black'
            linewidth = 2 if item['upgraded'] else 1
            
            patch = patches.Patch(
                facecolor=item['color'],
                edgecolor=edge_color,
                linewidth=linewidth
            )
            legend_elements.append(patch)
            
            # Clean label format matching reference exactly
            symbols = ""
            if item['upgraded']:
                symbols += " ↑"
            if item['rotated']:
                symbols += " ↻"
            
            label = f"{item['id']}{symbols} ({item['dimensions']})"
            legend_labels.append(label)
        
        # Simple legend positioning like reference
        legend = ax.legend(legend_elements, legend_labels,
                          loc='center left', bbox_to_anchor=(1.02, 0.5),
                          fontsize=9, frameon=True, fancybox=False, shadow=False)
        
        # Clean legend frame like reference
        legend.get_frame().set_facecolor('white')
        legend.get_frame().set_edgecolor('black')
        legend.get_frame().set_linewidth(1)


# Integration function for existing codebase
def generate_cutting_layout_pdf(boards: List[Board], order_name: str = "OptiWise", 
                               output_path: Optional[str] = None) -> bytes:
    """Generate PDF cutting layouts using the enhanced generator."""
    generator = PDFLayoutGenerator()
    return generator.generate_cutting_layouts_pdf(boards, order_name, output_path)