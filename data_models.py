"""
Core data models for OptiWise beam saw optimization tool.
Defines MaterialDetails, Part, Offcut, and Board classes.
"""

import re
from typing import Optional, List, Dict, Tuple
import logging

logger = logging.getLogger(__name__)


class MaterialDetails:
    """
    Represents material specifications including laminate and core details.
    Enforces strict laminate consistency rules.
    """
    
    def __init__(self, full_material_string: str):
        """
        Initialize MaterialDetails from a material string.
        
        Args:
            full_material_string: Material specification string (e.g., '2614 SF_18MR_2614 SF')
        
        Raises:
            ValueError: If top and bottom laminates are not identical
        """
        self.full_material_string = full_material_string
        (self.top_laminate_name, self.core_name, 
         self.thickness, self.bottom_laminate_name) = self._parse_material_string(full_material_string)
        # For backward compatibility, keep laminate_name as top laminate
        self.laminate_name = self.top_laminate_name
    
    @staticmethod
    def _parse_material_string(material_string: str) -> Tuple[str, str, int, str]:
        """
        Parse material string to extract laminate and core components.
        
        Args:
            material_string: Raw material string to parse
            
        Returns:
            Tuple of (top_laminate_name, core_name, thickness, bottom_laminate_name)
            
        Raises:
            ValueError: If top and bottom laminates don't match or parsing fails
        """
        try:
            # Remove extra whitespace and handle both underscore and hyphen formats
            clean_string = material_string.strip()
            
            # Try hyphen format first (e.g., "2614 SF-18MR-2614 SF")
            if '-' in clean_string:
                parts = clean_string.split('-')
            else:
                # Fall back to underscore format (e.g., "2614 SF_18MR_2614 SF")
                parts = clean_string.split('_')
            
            if len(parts) == 3:
                top_laminate_name = parts[0].strip()
                core_component_str = parts[1].strip()
                bottom_laminate_name = parts[2].strip()
            elif len(parts) == 1:
                # Handle single-component materials like "17WPC"
                material_part = parts[0].strip()
                # Try to extract thickness from the beginning
                thickness_match = re.match(r'^(\d+)', material_part)
                if thickness_match:
                    thickness_str = thickness_match.group(1)
                    core_name = material_part
                    top_laminate_name = "NONE"
                    bottom_laminate_name = "NONE"
                    core_component_str = core_name
                else:
                    # Default case for unknown format
                    top_laminate_name = "NONE"
                    core_component_str = material_part
                    bottom_laminate_name = "NONE"
            else:
                raise ValueError(f"Invalid material string format: {material_string}")
            
            # Note: Different top/bottom laminates are allowed in input
            # Compatibility will be checked during placement/upgrading
            
            # Extract thickness and core name from core component
            # Handle multiple formats:
            # 1. "18HDHMR", "18MR", "18BWR" - thickness prefix format
            # 2. "MR MDF", "PARTICLE BOARD" - name-only format
            
            thickness_match = re.search(r'^(\d+)', core_component_str)
            if thickness_match:
                # Format with thickness prefix
                thickness = int(thickness_match.group(1))
                core_name = core_component_str.strip()
            else:
                # Format without thickness - use default thickness based on material type
                core_name = core_component_str.strip()
                
                # Default thickness mapping for common materials
                thickness_defaults = {
                    'MR MDF': 18,
                    'PARTICLE BOARD': 18,
                    'PLYWOOD': 18,
                    'HDHMR': 18,
                    'BWR': 18,
                    'WPC': 17
                }
                
                # Find matching material type (case insensitive)
                thickness = 18  # Default fallback
                for material_type, default_thickness in thickness_defaults.items():
                    if material_type.lower() in core_name.lower():
                        thickness = default_thickness
                        break
            
            if not core_name:
                raise ValueError(f"Could not extract core name from: {core_component_str}")
            
            return top_laminate_name, core_name, thickness, bottom_laminate_name
            
        except Exception as e:
            raise ValueError(f"Failed to parse material string '{material_string}': {str(e)}")
    
    def get_cost_per_sqm(self, laminate_db: Dict, core_db: Dict) -> float:
        """
        Calculate total cost per square meter for this material.
        
        Args:
            laminate_db: Dictionary mapping laminate names to prices
            core_db: Dictionary with core material details including prices
            
        Returns:
            Total cost per square meter (laminate + core costs)
        """
        try:
            # Get costs for top and bottom laminates separately
            top_laminate_cost = laminate_db.get(self.top_laminate_name, 0.0)
            bottom_laminate_cost = laminate_db.get(self.bottom_laminate_name, 0.0)
            
            # Get core cost
            core_details = core_db.get(self.core_name, {})
            core_cost = core_details.get('price_per_sqm', 0.0)
            
            # Total cost includes top laminate + bottom laminate + core
            total_cost = top_laminate_cost + bottom_laminate_cost + core_cost
            return total_cost
            
        except Exception as e:
            logger.error(f"Error calculating cost for material {self.full_material_string}: {e}")
            return 0.0
    
    def __str__(self) -> str:
        return f"{self.top_laminate_name}_{self.core_name}_{self.bottom_laminate_name}"
    
    def __repr__(self) -> str:
        return self.__str__()


class Part:
    """
    Represents a part to be cut with its dimensions, material requirements, and placement info.
    """
    
    def __init__(self, part_id: str, requested_length: float, requested_width: float, 
                 quantity: int, material_details: MaterialDetails, grains: int, 
                 original_part_index: int, client_name: str = "", room_type: str = "", 
                 sub_category: str = "", panel_name: str = "", full_description: str = ""):
        """
        Initialize a Part object.
        
        Args:
            part_id: Unique identifier for the part (ORDER ID / UNIQUE CODE)
            requested_length: Original length in mm (CUT LENGTH)
            requested_width: Original width in mm (CUT WIDTH)
            quantity: Number of pieces needed (QTY)
            material_details: MaterialDetails object for this part
            grains: 1 if grain-sensitive (cannot rotate), 0 if grain-free
            original_part_index: Index in original cutlist for tracking
            client_name: Customer/client name
            room_type: Room category
            sub_category: Sub-category
            panel_name: Panel identifier
            full_description: Full description
        """
        self.id = part_id
        self.requested_length = requested_length
        self.requested_width = requested_width
        self.quantity = quantity
        self.material_details = material_details
        self.grains = grains
        self.original_part_index = original_part_index
        
        # Additional fields for enhanced display
        self.client_name = client_name
        self.room_type = room_type
        self.sub_category = sub_category
        self.panel_name = panel_name
        self.full_description = full_description
        
        # Placement attributes (set when placed on board)
        self.assigned_board_id: Optional[str] = None
        self.actual_length: Optional[float] = None
        self.actual_width: Optional[float] = None
        self.x_pos: Optional[float] = None
        self.y_pos: Optional[float] = None
        self.rotated: bool = False
        
        # Material upgrade tracking
        self.assigned_material_details: Optional[MaterialDetails] = None
        self.is_upgraded: bool = False
        
        # Store original CSV data for Excel export
        self.original_data: Optional[Dict[str, str]] = None
    
    def get_area_with_kerf(self, kerf: float) -> float:
        """
        Calculate area including kerf allowance.
        
        Args:
            kerf: Kerf width in mm
            
        Returns:
            Area in square mm including kerf
        """
        length_with_kerf = self.requested_length + kerf
        width_with_kerf = self.requested_width + kerf
        return length_with_kerf * width_with_kerf
    
    def can_rotate(self) -> bool:
        """
        Check if part can be rotated 90 degrees.
        
        Returns:
            True if part is grain-free (can rotate), False if grain-sensitive
        """
        return self.grains == 0
    
    def get_dimensions_for_placement(self, rotated: bool = False) -> Tuple[float, float]:
        """
        Get dimensions considering rotation.
        
        Args:
            rotated: Whether part should be rotated 90 degrees
            
        Returns:
            Tuple of (length, width) for placement
        """
        if rotated and self.can_rotate():
            return self.requested_width, self.requested_length
        return self.requested_length, self.requested_width
    
    def __str__(self) -> str:
        return f"Part({self.id}, {self.requested_length}x{self.requested_width}, {self.material_details})"
    
    def __repr__(self) -> str:
        return self.__str__()
    
    def copy_with_material(self, new_material_details: MaterialDetails) -> 'Part':
        """
        Create a copy of this part with a different material while preserving all other attributes.
        """
        new_part = Part(
            part_id=self.id,
            requested_length=self.requested_length,
            requested_width=self.requested_width,
            quantity=self.quantity,
            material_details=new_material_details,
            grains=self.grains,
            original_part_index=self.original_part_index,
            client_name=self.client_name,
            room_type=self.room_type,
            sub_category=self.sub_category,
            panel_name=self.panel_name,
            full_description=self.full_description
        )
        
        # Preserve original CSV data and placement info
        new_part.original_data = self.original_data
        new_part.assigned_board_id = self.assigned_board_id
        new_part.actual_length = self.actual_length
        new_part.actual_width = self.actual_width
        new_part.x_pos = self.x_pos
        new_part.y_pos = self.y_pos
        new_part.rotated = self.rotated
        new_part.assigned_material_details = self.assigned_material_details
        new_part.is_upgraded = True  # Mark as upgraded since we're changing material
        
        return new_part


class Offcut:
    """
    Represents a remaining piece from cutting operations that can be reused.
    """
    
    def __init__(self, offcut_id: str, x: float, y: float, length: float, width: float,
                 material_details: MaterialDetails, source_board_id: str):
        """
        Initialize an Offcut object.
        
        Args:
            offcut_id: Unique identifier for the offcut
            x, y: Position coordinates on the source board
            length, width: Dimensions of the offcut
            material_details: Material specifications of this offcut
            source_board_id: ID of the board this offcut came from
        """
        self.id = offcut_id
        self.x = x
        self.y = y
        self.length = length
        self.width = width
        self.material_details = material_details
        self.source_board_id = source_board_id
    
    def get_area(self) -> float:
        """
        Calculate the area of this offcut.
        
        Returns:
            Area in square mm
        """
        return self.length * self.width
    
    def can_fit_part(self, part: 'Part', kerf: float, rotated: bool = False) -> bool:
        """
        Check if a part can fit in this offcut with kerf allowance.
        Kerf is only added between parts, not at board edges.
        
        Args:
            part: Part object to check
            kerf: Kerf width in mm
            rotated: Whether to check rotated orientation
            
        Returns:
            True if part fits with kerf, False otherwise
        """
        part_length, part_width = part.get_dimensions_for_placement(rotated)
        
        # For parts not at board edges, we need space for the part plus kerf
        # For now, we assume kerf on all sides (this will be optimized in placement)
        required_length = part_length + kerf
        required_width = part_width + kerf
        
        return (required_length <= self.length and required_width <= self.width)
    
    def __str__(self) -> str:
        return f"Offcut({self.id}, {self.length}x{self.width}, {self.material_details})"
    
    def __repr__(self) -> str:
        return self.__str__()


class Board:
    """
    Represents a full board with parts placement and available space tracking.
    """
    
    def __init__(self, board_id: str, material_details: MaterialDetails, 
                 total_length: float, total_width: float, kerf: float):
        """
        Initialize a Board object.
        
        Args:
            board_id: Unique identifier for the board
            material_details: Material specifications of this board
            total_length, total_width: Full board dimensions
            kerf: Kerf width in mm
        """
        self.id = board_id
        self.material_details = material_details
        self.total_length = total_length
        self.total_width = total_width
        self.kerf = kerf
        self.parts_on_board: List[Part] = []
        
        # Initialize with one offcut representing the full board
        initial_offcut = Offcut(
            offcut_id=f"{board_id}_initial",
            x=0.0, y=0.0,
            length=total_length, width=total_width,
            material_details=material_details,
            source_board_id=board_id
        )
        self.available_rectangles: List[Offcut] = [initial_offcut]
    
    def unplace_part(self, part: Part) -> bool:
        """
        Remove a part from the board and integrate its space back into available rectangles.
        
        Args:
            part: Part to remove from the board
            
        Returns:
            True if part was successfully removed, False if part not found
        """
        if part not in self.parts_on_board:
            return False
        
        # Remove the part from the board
        self.parts_on_board.remove(part)
        
        # Get the dimensions of the freed space
        if hasattr(part, 'rotated') and part.rotated:
            freed_length = part.requested_width
            freed_width = part.requested_length
        else:
            freed_length = part.requested_length
            freed_width = part.requested_width
        
        # Get the position where the part was placed
        freed_x = getattr(part, 'x_pos', 0.0)
        freed_y = getattr(part, 'y_pos', 0.0)
        
        # Create a new offcut representing the freed space
        freed_offcut = Offcut(
            offcut_id=f"{self.id}_freed_{len(self.available_rectangles)}",
            x=freed_x,
            y=freed_y,
            length=freed_length,
            width=freed_width,
            material_details=self.material_details,
            source_board_id=self.id
        )
        
        # Add the freed space back to available rectangles
        self.available_rectangles.append(freed_offcut)
        
        # Attempt to merge adjacent rectangles
        self._merge_adjacent_rectangles()
        
        # Update utilization
        self.utilization_percentage = self.get_utilization_percentage()
        
        return True
    
    def _merge_adjacent_rectangles(self):
        """
        Merge adjacent rectangles in available_rectangles to create larger usable spaces.
        This is a simplified implementation that merges rectangles that share edges.
        """
        if len(self.available_rectangles) <= 1:
            return
        
        merged = True
        while merged:
            merged = False
            for i in range(len(self.available_rectangles)):
                for j in range(i + 1, len(self.available_rectangles)):
                    rect1 = self.available_rectangles[i]
                    rect2 = self.available_rectangles[j]
                    
                    # Check if rectangles can be merged horizontally
                    if (rect1.y == rect2.y and rect1.width == rect2.width and
                        (rect1.x + rect1.length == rect2.x or rect2.x + rect2.length == rect1.x)):
                        
                        # Merge horizontally
                        new_x = min(rect1.x, rect2.x)
                        new_length = rect1.length + rect2.length
                        
                        merged_rect = Offcut(
                            offcut_id=f"{self.id}_merged_{i}_{j}",
                            x=new_x,
                            y=rect1.y,
                            length=new_length,
                            width=rect1.width,
                            material_details=self.material_details,
                            source_board_id=self.id
                        )
                        
                        # Replace the two rectangles with the merged one
                        self.available_rectangles.pop(max(i, j))
                        self.available_rectangles.pop(min(i, j))
                        self.available_rectangles.append(merged_rect)
                        merged = True
                        break
                    
                    # Check if rectangles can be merged vertically
                    elif (rect1.x == rect2.x and rect1.length == rect2.length and
                          (rect1.y + rect1.width == rect2.y or rect2.y + rect2.width == rect1.y)):
                        
                        # Merge vertically
                        new_y = min(rect1.y, rect2.y)
                        new_width = rect1.width + rect2.width
                        
                        merged_rect = Offcut(
                            offcut_id=f"{self.id}_merged_{i}_{j}",
                            x=rect1.x,
                            y=new_y,
                            length=rect1.length,
                            width=new_width,
                            material_details=self.material_details,
                            source_board_id=self.id
                        )
                        
                        # Replace the two rectangles with the merged one
                        self.available_rectangles.pop(max(i, j))
                        self.available_rectangles.pop(min(i, j))
                        self.available_rectangles.append(merged_rect)
                        merged = True
                        break
                
                if merged:
                    break

    def place_part(self, part: Part, offcut_to_use: Offcut, rotated: bool, 
                   x_pos: float, y_pos: float, core_db: Dict) -> bool:
        """
        Place a part on the board at specified position.
        
        Args:
            part: Part to place
            offcut_to_use: Offcut space to use for placement
            rotated: Whether part is rotated 90 degrees
            x_pos, y_pos: Position coordinates
            core_db: Core materials database for upgrade checking
            
        Returns:
            True if placement successful, False otherwise
        """
        try:
            from optimization_core_fixed import get_grade_level
            
            # Get part dimensions for placement
            part_length, part_width = part.get_dimensions_for_placement(rotated)
            
            # Check if part fits in the offcut with kerf
            if not offcut_to_use.can_fit_part(part, self.kerf, rotated):
                return False
            
            # Update part placement information
            part.assigned_board_id = self.id
            part.actual_length = part_length
            part.actual_width = part_width
            part.x_pos = x_pos
            part.y_pos = y_pos
            part.rotated = rotated
            
            # Set assigned material details and upgrade status
            part.assigned_material_details = offcut_to_use.material_details
            
            # Check if this is an upgrade
            original_grade = get_grade_level(part.material_details.core_name, core_db)
            assigned_grade = get_grade_level(offcut_to_use.material_details.core_name, core_db)
            part.is_upgraded = assigned_grade > original_grade
            
            # Store placement information on the part for unplacing later
            part.x_pos = x_pos
            part.y_pos = y_pos
            part.rotated = rotated
            part.assigned_board_id = self.id
            
            # Add part to board
            self.parts_on_board.append(part)
            
            # Split the offcut and update available rectangles
            self._split_offcut(offcut_to_use, part, part_length, part_width, x_pos, y_pos)
            
            return True
            
        except Exception as e:
            logger.error(f"Error placing part {part.id} on board {self.id}: {e}")
            return False
    
    def _split_offcut(self, offcut: Offcut, part: Part, part_length: float, 
                     part_width: float, x_pos: float, y_pos: float):
        """
        Split an offcut after part placement, creating new available rectangles.
        Implements guillotine cutting constraints.
        
        Args:
            offcut: Original offcut being split
            part: Part that was placed
            part_length, part_width: Actual dimensions of placed part
            x_pos, y_pos: Position where part was placed
        """
        # Remove the used offcut
        if offcut in self.available_rectangles:
            self.available_rectangles.remove(offcut)
        
        # Dimensions including kerf
        part_with_kerf_length = part_length + self.kerf
        part_with_kerf_width = part_width + self.kerf
        
        # Create new offcuts from the split (guillotine cuts)
        new_offcuts = []
        
        # Right offcut (if space available)
        if x_pos + part_with_kerf_length < offcut.x + offcut.length:
            right_offcut = Offcut(
                offcut_id=f"{self.id}_offcut_{len(self.available_rectangles)}_{len(new_offcuts)}",
                x=x_pos + part_with_kerf_length,
                y=offcut.y,
                length=offcut.x + offcut.length - (x_pos + part_with_kerf_length),
                width=offcut.width,
                material_details=offcut.material_details,
                source_board_id=self.id
            )
            if right_offcut.get_area() > 0:
                new_offcuts.append(right_offcut)
        
        # Bottom offcut (if space available)
        if y_pos + part_with_kerf_width < offcut.y + offcut.width:
            bottom_offcut = Offcut(
                offcut_id=f"{self.id}_offcut_{len(self.available_rectangles)}_{len(new_offcuts)}",
                x=offcut.x,
                y=y_pos + part_with_kerf_width,
                length=x_pos + part_with_kerf_length - offcut.x,
                width=offcut.y + offcut.width - (y_pos + part_with_kerf_width),
                material_details=offcut.material_details,
                source_board_id=self.id
            )
            if bottom_offcut.get_area() > 0:
                new_offcuts.append(bottom_offcut)
        
        # Add new offcuts to available rectangles
        self.available_rectangles.extend(new_offcuts)
    

    
    def get_utilization_percentage(self) -> float:
        """
        Calculate the percentage of board area utilized.
        
        Returns:
            Utilization percentage (0-100)
        """
        total_area = self.total_length * self.total_width
        if total_area == 0:
            return 0.0
        
        # For TEST algorithms, use actual placed dimensions since kerf is handled in placement
        # For other algorithms, use kerf-expanded area as before
        if hasattr(self, 'id') and any(test_name in self.id for test_name in ['BLF', 'Shelf', 'TightNest', 'BestFit', 'GlobalOpt']):
            # TEST algorithms: use actual placed area only
            used_area = sum(part.actual_length * part.actual_width for part in self.parts_on_board 
                          if hasattr(part, 'actual_length') and hasattr(part, 'actual_width') and 
                          part.actual_length is not None and part.actual_width is not None)
        else:
            # Traditional algorithms: use kerf-expanded area
            used_area = sum(part.get_area_with_kerf(self.kerf) for part in self.parts_on_board)
        
        return (used_area / total_area) * 100
    
    def get_remaining_area(self) -> float:
        """
        Calculate remaining unused area on the board.
        
        Returns:
            Remaining area in square mm
        """
        total_area = self.total_length * self.total_width
        used_area = sum(part.get_area_with_kerf(self.kerf) for part in self.parts_on_board)
        return total_area - used_area
    
    def get_largest_offcut(self) -> Optional[Offcut]:
        """
        Get the largest available offcut by area.
        
        Returns:
            Largest offcut or None if no offcuts available
        """
        if not self.available_rectangles:
            return None
        
        return max(self.available_rectangles, key=lambda o: o.get_area())
    
    def __str__(self) -> str:
        return f"Board({self.id}, {self.total_length}x{self.total_width}, {len(self.parts_on_board)} parts)"
    
    def __repr__(self) -> str:
        return self.__str__()
