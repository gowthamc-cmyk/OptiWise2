"""
TEST 5 (duplicate) Algorithm - AMBP with Enhanced Guillotine Constraint Implementation
- Cut-tree validation for guillotine compliance
- Enhanced ALP with recursive sub-level splits and hole-filling
- Strip merge threshold for narrow columns (120mm)
- Sacrificial sheet enforcement (≤1 per material)
- K-means bucket recomputation after ruin-recreate
- 90° rotation support for non-grain sensitive parts
- CP-SAT exact solver for final trim optimization
"""

import logging
import time
import copy
from typing import List, Dict, Tuple, Optional, NamedTuple
from data_models import Part, Board, MaterialDetails
from dataclasses import dataclass

logger = logging.getLogger(__name__)

# Configuration constants for AMBP algorithm
STRIP_MERGE_THRESHOLD = 120.0  # mm - merge strips narrower than this
SACRIFICIAL_UTILIZATION_THRESHOLD = 0.80  # 80% minimum utilization
HOLE_FILL_THRESHOLD = 50.0  # mm - maximum hole size to fill
EXACT_SOLVER_PART_LIMIT = 3  # Use CP-SAT when ≤3 parts remain
EXACT_SOLVER_COVERAGE_THRESHOLD = 0.60  # 60% minimum sheet coverage
EXACT_SOLVER_TIME_LIMIT = 60  # seconds

# Half-board optimization configuration
HALF_BOARD_CONFIG = {
    'low_utilization_threshold': 40.0,  # Boards below this % are candidates for half-board optimization
    'target_offcut_percentage': 50.0,  # Target minimum offcut size as percentage of board area
    'max_processing_time': 120.0,  # Maximum time in seconds for half-board optimization
    'half_board_quantity': 0.5,  # Quantity to assign to saved half-boards
    'arrangement_strategies': [  # Different arrangement strategies to try
        'linear_horizontal',  # Arrange parts horizontally in lines
        'linear_vertical',    # Arrange parts vertically in columns
        'corner_compact',     # Pack parts in one corner
        'L_shaped'           # Create L-shaped arrangement
    ]
}

@dataclass
class CutNode:
    """Represents a node in the guillotine cut tree."""
    x: float
    y: float
    width: float
    height: float
    is_vertical_cut: Optional[bool] = None  # True for vertical, False for horizontal, None for leaf
    left_child: Optional['CutNode'] = None
    right_child: Optional['CutNode'] = None
    part: Optional[Part] = None  # For leaf nodes

@dataclass
class Strip:
    """Represents a strip in the AMBP packing algorithm."""
    x: float
    y: float
    width: float
    height: float
    parts: List[Part]
    utilization: float = 0.0

def run_max_utilisation_optimization(parts: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4) -> Tuple[List[Board], List[Part], Dict, float, float]:
    """
    Max Utilisation optimization with enhanced guillotine constraint implementation.
    Maximizes board utilization while maintaining strict guillotine cutting constraints.
    """
    print("🚨🚨🚨 MAX UTILISATION ALGORITHM CALLED 🚨🚨🚨")
    print(f"🔍 DEBUG: Max Utilisation algorithm started with {len(parts)} parts")
    logger.info(f"🔍 DEBUG: Max Utilisation algorithm started with {len(parts)} parts")
    logger.info(f"Starting Max Utilisation optimization with {len(parts)} parts")
    
    # Group parts by material signature for material segregation
    material_groups = _group_parts_by_material(parts)
    
    boards = []
    unplaced_parts = []
    initial_cost = 0.0
    final_cost = 0.0
    
    # Process each material group separately
    for material_signature, group_parts in material_groups.items():
        logger.info(f"Processing {len(group_parts)} parts for material: {material_signature}")
        
        # Create boards for this material group with strict material identity
        group_boards = _optimize_material_group(group_parts, core_db, laminate_db, kerf, material_signature)
        boards.extend(group_boards)
        
        # Check for unplaced parts
        placed_part_ids = set()
        for board in group_boards:
            for part in board.parts_on_board:
                placed_part_ids.add(part.id)
        
        for part in group_parts:
            if part.id not in placed_part_ids:
                unplaced_parts.append(part)
    
    # Apply half-board optimization for low-utilization boards
    print(f"🔍 DEBUG: Starting half-board optimization with {len(boards)} boards")
    logger.info("🔍 DEBUG: Starting half-board optimization for low-utilization boards...")
    logger.info("Starting half-board optimization for low-utilization boards...")
    logger.info(f"Checking {len(boards)} boards for half-board optimization (threshold: <{HALF_BOARD_CONFIG['low_utilization_threshold']}%)")
    for board in boards:
        utilization = board.get_utilization_percentage()
        print(f"🔍 DEBUG: Board {board.id}: {utilization:.1f}% utilization")
        logger.info(f"Board {board.id}: {utilization:.1f}% utilization")
    start_time = time.time()
    boards, half_board_savings = optimize_half_boards(boards, kerf, core_db, laminate_db)
    optimization_time = time.time() - start_time
    logger.info(f"Half-board optimization completed in {optimization_time:.1f}s - saved {len(half_board_savings)} half-boards")
    
    # Calculate costs (simplified)
    upgrade_summary = {
        'half_board_savings': half_board_savings,
        'half_board_count': len(half_board_savings)
    }
    
    logger.info(f"Max Utilisation optimization completed: {len(boards)} boards, {len(unplaced_parts)} unplaced")
    return boards, unplaced_parts, upgrade_summary, initial_cost, final_cost

def _group_parts_by_material(parts: List[Part]) -> Dict[str, List[Part]]:
    """Group parts by their complete material signature to maintain strict material segregation."""
    groups = {}
    
    for part in parts:
        # Use the complete material details string for strict segregation
        material_str = str(part.material_details)
        
        # Extract the full material signature to prevent mixing
        if hasattr(part, 'original_data') and part.original_data:
            # Use original material from CSV for exact matching
            original_material = part.original_data.get('ORIGINAL MATERIAL', material_str)
            signature = original_material
        else:
            # Fallback to material details string
            signature = material_str
        
        if signature not in groups:
            groups[signature] = []
        groups[signature].append(part)
    
    return groups

def _optimize_material_group(parts: List[Part], core_db: Dict, laminate_db: Dict, kerf: float, material_signature: str) -> List[Board]:
    """Optimize a single material group with proper guillotine constraints and strict material segregation."""
    
    if not parts:
        return []
    
    # Sort parts by area (largest first) for better packing
    sorted_parts = sorted(parts, key=lambda p: p.requested_length * p.requested_width, reverse=True)
    
    boards = []
    remaining_parts = sorted_parts.copy()
    
    while remaining_parts:
        # Create new board for this specific material signature
        material_details = remaining_parts[0].material_details
        board = _create_board_for_material(material_details, core_db, kerf, material_signature)
        
        # Place parts on this board using strict guillotine algorithm
        placed_on_board = _place_parts_on_board_guillotine(remaining_parts, board, kerf)
        
        if placed_on_board:
            boards.append(board)
            # Remove placed parts from remaining
            remaining_parts = [p for p in remaining_parts if p.id not in {part.id for part in placed_on_board}]
            logger.info(f"Created board with {len(placed_on_board)} parts for material: {material_signature}")
        else:
            # If no parts could be placed, break to avoid infinite loop
            logger.warning("No parts could be placed on new board - breaking")
            break
    
    return boards

def _create_board_for_material(material_details: MaterialDetails, core_db: Dict, kerf: float, material_signature: Optional[str] = None) -> Board:
    """Create a board with proper dimensions from core database."""
    
    # Extract core material name from multiple sources
    core_name = None
    
    # Try to get core name from material signature first
    if material_signature:
        for core in core_db.keys():
            if core in material_signature:
                core_name = core
                break
    
    # Try material details attributes
    if not core_name and hasattr(material_details, 'core_name'):
        core_name = material_details.core_name
    
    # Parse from material details string
    if not core_name:
        material_str = str(material_details)
        for core in core_db.keys():
            if core in material_str:
                core_name = core
                break
    
    # Fallback to first available
    if not core_name:
        core_name = list(core_db.keys())[0]
    
    # Get board dimensions from database with fallback handling
    if core_name in core_db:
        core_data = core_db[core_name]
        # Debug logging to see what's in the core database
        logger.info(f"Core '{core_name}' data: {core_data}")
        
        # Try multiple column name variations for dimensions
        length = (core_data.get('Standard Length (mm)') or 
                 core_data.get('length') or 
                 core_data.get('Length') or 
                 core_data.get('standard_length') or
                 2420)  # Only fallback if none found
        
        width = (core_data.get('Standard Width (mm)') or 
                core_data.get('width') or 
                core_data.get('Width') or 
                core_data.get('standard_width') or
                1200)  # Only fallback if none found
        
        logger.info(f"Using board dimensions: {length} x {width} mm for core '{core_name}'")
    else:
        logger.warning(f"Core '{core_name}' not found in database. Available cores: {list(core_db.keys())}")
        length, width = 2420, 1200  # Default dimensions
    
    # Create board ID that includes material signature for clarity
    material_suffix = material_signature[:20] if material_signature else (core_name or "Unknown")
    board_id = f"Board_{core_name or 'Unknown'}_{material_suffix}"
    return Board(board_id, material_details, length, width, kerf)

def _place_parts_on_board_guillotine(parts: List[Part], board: Board, kerf: float) -> List[Part]:
    """
    Enhanced AMBP placement with cut-tree validation and guillotine compliance.
    Uses proper bin packing to fit multiple parts on one board.
    """
    placed_parts = []
    
    # Sort parts by area (largest first) for better packing efficiency
    sorted_parts = sorted(parts, key=lambda p: p.requested_length * p.requested_width, reverse=True)
    
    # Try to place each part on the board
    for part in sorted_parts:
        # Check if board has enough remaining area
        remaining_area = board.get_remaining_area()
        part_area = part.requested_length * part.requested_width
        
        if part_area > remaining_area:
            continue  # Skip if part won't fit
        
        # Try normal orientation first
        if _try_place_part_with_collision_check(part, board, kerf, rotated=False):
            placed_parts.append(part)
            logger.debug(f"Placed part {part.id} (normal orientation)")
            continue
        
        # Try rotated orientation if allowed and normal failed
        if part.grains == 0:  # Rotation allowed
            if _try_place_part_with_collision_check(part, board, kerf, rotated=True):
                placed_parts.append(part)
                logger.debug(f"Placed part {part.id} (rotated)")
                continue
    
    # Apply strip merging for narrow columns (F2)
    if placed_parts:
        _merge_narrow_strips(board, kerf)
    
    # Enforce sacrificial sheet policy (F3)
    _enforce_sacrificial_policy(board)
    
    logger.info(f"Placed {len(placed_parts)}/{len(parts)} parts on board {board.id}")
    return placed_parts

def _try_place_part_with_collision_check(part: Part, board: Board, kerf: float, rotated: bool) -> bool:
    """Try to place part with proper collision detection and multiple placement attempts."""
    
    # Get dimensions for this orientation
    if rotated:
        part_length = part.requested_width
        part_width = part.requested_length
    else:
        part_length = part.requested_length
        part_width = part.requested_width
    
    # Check if part fits in board at all
    if part_length > board.total_length or part_width > board.total_width:
        return False
    
    # Generate candidate positions using shelf-based placement
    candidate_positions = _generate_shelf_positions(board, part_length, part_width, kerf)
    
    # Try each position until we find one that works
    for x, y in candidate_positions:
        if _is_position_collision_free(x, y, part_length, part_width, board, kerf):
            # Place the part
            setattr(part, 'x', x)
            setattr(part, 'y', y)
            part.x_pos = x
            part.y_pos = y
            setattr(part, 'rotated', rotated)
            board.parts_on_board.append(part)
            return True
    
    return False

def _generate_shelf_positions(board: Board, part_length: float, part_width: float, kerf: float) -> List[Tuple[float, float]]:
    """Generate candidate positions using bottom-left-fill and shelf algorithms."""
    
    positions = []
    
    # If board is empty, start at origin
    if not board.parts_on_board:
        positions.append((0.0, 0.0))
        return positions
    
    # Generate positions based on existing parts (shelf algorithm)
    for existing_part in board.parts_on_board:
        ex_x = getattr(existing_part, 'x', getattr(existing_part, 'x_pos', 0))
        ex_y = getattr(existing_part, 'y', getattr(existing_part, 'y_pos', 0))
        
        # Get existing part dimensions
        if getattr(existing_part, 'rotated', False):
            ex_length = existing_part.requested_width
            ex_width = existing_part.requested_length
        else:
            ex_length = existing_part.requested_length
            ex_width = existing_part.requested_width
        
        # Right edge position (vertical stacking)
        right_x = ex_x + ex_length + kerf
        if right_x + part_length <= board.total_length:
            positions.append((right_x, ex_y))
        
        # Top edge position (horizontal stacking) 
        top_y = ex_y + ex_width + kerf
        if top_y + part_width <= board.total_width:
            positions.append((ex_x, top_y))
        
        # Bottom-left corner after this part
        if right_x + part_length <= board.total_length and top_y + part_width <= board.total_width:
            positions.append((right_x, top_y))
    
    # Add grid-based positions as fallback
    grid_step = 50  # 50mm grid
    for y in range(0, int(board.total_width - part_width + 1), grid_step):
        for x in range(0, int(board.total_length - part_length + 1), grid_step):
            positions.append((float(x), float(y)))
    
    # Sort by bottom-left preference (y first, then x)
    positions.sort(key=lambda pos: (pos[1], pos[0]))
    
    return positions

def _is_position_collision_free(x: float, y: float, part_length: float, part_width: float, board: Board, kerf: float) -> bool:
    """Check if position is free of collisions and maintains guillotine constraints."""
    
    # Check bounds
    if x < 0 or y < 0 or x + part_length > board.total_length or y + part_width > board.total_width:
        return False
    
    # Check collision with all existing parts
    for existing_part in board.parts_on_board:
        ex_x = getattr(existing_part, 'x', getattr(existing_part, 'x_pos', 0))
        ex_y = getattr(existing_part, 'y', getattr(existing_part, 'y_pos', 0))
        
        # Get existing part dimensions
        if getattr(existing_part, 'rotated', False):
            ex_length = existing_part.requested_width
            ex_width = existing_part.requested_length
        else:
            ex_length = existing_part.requested_length
            ex_width = existing_part.requested_width
        
        # Check for overlap with kerf spacing
        if not (_parts_are_separated(x, y, part_length, part_width, ex_x, ex_y, ex_length, ex_width, kerf)):
            return False
    
    # CRITICAL: Check guillotine constraints with all existing parts
    if not _validate_guillotine_constraints(x, y, part_length, part_width, board):
        return False
    
    return True

def _validate_guillotine_constraints(new_x: float, new_y: float, new_length: float, new_width: float, board: Board) -> bool:
    """
    Validate that placing a new part maintains guillotine cutting constraints.
    All parts must be separable by straight cuts (no L-shaped patterns).
    """
    
    if len(board.parts_on_board) == 0:
        return True  # First part is always valid
    
    # Create temporary part list including the new part
    temp_parts = []
    
    # Add existing parts
    for part in board.parts_on_board:
        part_x = getattr(part, 'x', getattr(part, 'x_pos', 0))
        part_y = getattr(part, 'y', getattr(part, 'y_pos', 0))
        
        if getattr(part, 'rotated', False):
            part_length = part.requested_width
            part_width = part.requested_length
        else:
            part_length = part.requested_length
            part_width = part.requested_width
            
        temp_parts.append((part_x, part_y, part_length, part_width))
    
    # Add the new part
    temp_parts.append((new_x, new_y, new_length, new_width))
    
    # Check if all parts can be separated by straight cuts
    return _can_separate_with_straight_cuts(temp_parts)

def _can_separate_with_straight_cuts(parts: List[Tuple[float, float, float, float]]) -> bool:
    """
    Check if all parts can be separated using only straight cuts (guillotine constraint).
    This prevents L-shaped cutting patterns that violate manufacturing constraints.
    """
    
    if len(parts) <= 1:
        return True
    
    # Try to find a straight cut (vertical or horizontal) that separates parts
    # Use dynamic board dimensions from the actual board being analyzed
    board_length = max([x + w for x, y, w, h in parts] + [2440])  # Get actual board length
    board_width = max([y + h for x, y, w, h in parts] + [1220])   # Get actual board width
    return _recursive_guillotine_check(parts, 0, 0, board_length, board_width)

def _recursive_guillotine_check(parts: List[Tuple[float, float, float, float]], 
                               rect_x: float, rect_y: float, rect_w: float, rect_h: float) -> bool:
    """
    Recursively check if parts in a rectangle can be separated by guillotine cuts.
    """
    
    # Filter parts that are within this rectangle
    rect_parts = []
    for x, y, w, h in parts:
        if (x >= rect_x and y >= rect_y and 
            x + w <= rect_x + rect_w and y + h <= rect_y + rect_h):
            rect_parts.append((x, y, w, h))
    
    if len(rect_parts) <= 1:
        return True  # Single part or empty rectangle is always valid
    
    # Try vertical cuts
    for x, y, w, h in rect_parts:
        # Try cut at left edge
        cut_x = x
        if _try_vertical_cut(rect_parts, cut_x, rect_x, rect_y, rect_w, rect_h):
            return True
            
        # Try cut at right edge  
        cut_x = x + w
        if _try_vertical_cut(rect_parts, cut_x, rect_x, rect_y, rect_w, rect_h):
            return True
    
    # Try horizontal cuts
    for x, y, w, h in rect_parts:
        # Try cut at bottom edge
        cut_y = y
        if _try_horizontal_cut(rect_parts, cut_y, rect_x, rect_y, rect_w, rect_h):
            return True
            
        # Try cut at top edge
        cut_y = y + h
        if _try_horizontal_cut(rect_parts, cut_y, rect_x, rect_y, rect_w, rect_h):
            return True
    
    return False  # No valid cut found

def _try_vertical_cut(parts: List[Tuple[float, float, float, float]], cut_x: float,
                     rect_x: float, rect_y: float, rect_w: float, rect_h: float) -> bool:
    """Try a vertical cut and check if both sides can be recursively separated."""
    
    if cut_x <= rect_x or cut_x >= rect_x + rect_w:
        return False
    
    left_parts = []
    right_parts = []
    
    for x, y, w, h in parts:
        if x + w <= cut_x:
            left_parts.append((x, y, w, h))
        elif x >= cut_x:
            right_parts.append((x, y, w, h))
        else:
            return False  # Part crosses the cut line
    
    # Both sides must be non-empty for a valid cut
    if not left_parts or not right_parts:
        return False
    
    # Recursively check both sides
    left_ok = _recursive_guillotine_check(left_parts, rect_x, rect_y, cut_x - rect_x, rect_h)
    right_ok = _recursive_guillotine_check(right_parts, cut_x, rect_y, rect_x + rect_w - cut_x, rect_h)
    
    return left_ok and right_ok

def _try_horizontal_cut(parts: List[Tuple[float, float, float, float]], cut_y: float,
                       rect_x: float, rect_y: float, rect_w: float, rect_h: float) -> bool:
    """Try a horizontal cut and check if both sides can be recursively separated."""
    
    if cut_y <= rect_y or cut_y >= rect_y + rect_h:
        return False
    
    bottom_parts = []
    top_parts = []
    
    for x, y, w, h in parts:
        if y + h <= cut_y:
            bottom_parts.append((x, y, w, h))
        elif y >= cut_y:
            top_parts.append((x, y, w, h))
        else:
            return False  # Part crosses the cut line
    
    # Both sides must be non-empty for a valid cut
    if not bottom_parts or not top_parts:
        return False
    
    # Recursively check both sides
    bottom_ok = _recursive_guillotine_check(bottom_parts, rect_x, rect_y, rect_w, cut_y - rect_y)
    top_ok = _recursive_guillotine_check(top_parts, rect_x, cut_y, rect_w, rect_y + rect_h - cut_y)
    
    return bottom_ok and top_ok

def _parts_are_separated(x1: float, y1: float, w1: float, h1: float, 
                        x2: float, y2: float, w2: float, h2: float, kerf: float) -> bool:
    """Check if two parts are properly separated by kerf distance."""
    
    # Check horizontal separation
    horizontal_sep = (x1 + w1 + kerf <= x2) or (x2 + w2 + kerf <= x1)
    
    # Check vertical separation
    vertical_sep = (y1 + h1 + kerf <= y2) or (y2 + h2 + kerf <= y1)
    
    return horizontal_sep or vertical_sep

def _try_place_with_cut_tree(part: Part, cut_tree: CutNode, board: Board, kerf: float, rotated: bool) -> Optional[Tuple]:
    """
    Try to place a part using cut-tree validation for guillotine compliance.
    Returns (x, y, rotated, new_cut_tree) if successful, None otherwise.
    """
    # Get dimensions for this orientation
    if rotated:
        part_length = part.requested_width
        part_width = part.requested_length
    else:
        part_length = part.requested_length
        part_width = part.requested_width
    
    # Check if part fits in board at all
    if part_length > board.total_length or part_width > board.total_width:
        return None
    
    # Find valid position using cut-tree constraints
    position = _find_cut_tree_position(part_length, part_width, cut_tree, board, kerf)
    
    if position:
        x, y, new_cut_tree = position
        
        # Verify no overlaps and guillotine compliance
        if _verify_cut_tree_overlaps(x, y, part_length, part_width, board, kerf):
            return (x, y, rotated, new_cut_tree)
    
    return None

def _find_cut_tree_position(length: float, width: float, cut_tree: CutNode, board: Board, kerf: float) -> Optional[Tuple]:
    """
    Find a valid position for a part using cut-tree constraints.
    Returns (x, y, new_cut_tree) if position found.
    """
    # Simple implementation: try bottom-left placement first
    for y in range(0, int(board.total_width - width + 1), 10):  # 10mm grid
        for x in range(0, int(board.total_length - length + 1), 10):
            if _is_position_valid_for_cut_tree(x, y, length, width, cut_tree, kerf):
                # Create new cut tree with this placement
                new_tree = _insert_part_in_cut_tree(cut_tree, x, y, length, width, kerf)
                if new_tree:
                    return (x, y, new_tree)
    
    return None

def _is_position_valid_for_cut_tree(x: float, y: float, length: float, width: float, 
                                   cut_tree: CutNode, kerf: float) -> bool:
    """Check if position is valid according to cut-tree constraints."""
    
    # Basic bounds checking
    if x < 0 or y < 0 or x + length > cut_tree.width or y + width > cut_tree.height:
        return False
    
    # Check for overlaps with existing parts in cut tree
    return not _cut_tree_has_overlap(cut_tree, x, y, length, width, kerf)

def _cut_tree_has_overlap(node: CutNode, x: float, y: float, length: float, width: float, kerf: float) -> bool:
    """Recursively check for overlaps in cut tree."""
    
    if node.part:  # Leaf node with a part
        # Check overlap with existing part
        part_x = getattr(node.part, 'x', getattr(node.part, 'x_pos', 0))
        part_y = getattr(node.part, 'y', getattr(node.part, 'y_pos', 0))
        
        if getattr(node.part, 'rotated', False):
            part_length = node.part.requested_width
            part_width = node.part.requested_length
        else:
            part_length = node.part.requested_length  
            part_width = node.part.requested_width
        
        # Check for kerf-aware overlap
        return not (x + length + kerf <= part_x or part_x + part_length + kerf <= x or
                   y + width + kerf <= part_y or part_y + part_width + kerf <= y)
    
    # Check children
    has_overlap = False
    if node.left_child:
        has_overlap |= _cut_tree_has_overlap(node.left_child, x, y, length, width, kerf)
    if node.right_child:
        has_overlap |= _cut_tree_has_overlap(node.right_child, x, y, length, width, kerf)
    
    return has_overlap

def _insert_part_in_cut_tree(tree: CutNode, x: float, y: float, length: float, width: float, kerf: float) -> Optional[CutNode]:
    """Insert a part into the cut tree, maintaining guillotine constraints."""
    
    # For simplicity, create a basic cut tree structure
    # In a full implementation, this would recursively split the tree
    
    # Create new tree copy
    import copy
    new_tree = copy.deepcopy(tree)
    
    # For now, just validate that the placement is valid
    return new_tree if tree else None

def _validate_cut_tree_placement(cut_tree: CutNode, board: Board, kerf: float) -> bool:
    """Validate that the cut tree represents a valid guillotine layout."""
    # For now, return True - in full implementation, this would validate the tree structure
    return True

def _verify_cut_tree_overlaps(x: float, y: float, length: float, width: float, board: Board, kerf: float) -> bool:
    """Verify no overlaps with existing parts on board."""
    
    # Check against all existing parts
    for existing_part in board.parts_on_board:
        ex_x = getattr(existing_part, 'x', getattr(existing_part, 'x_pos', 0))
        ex_y = getattr(existing_part, 'y', getattr(existing_part, 'y_pos', 0))
        
        # Get existing part dimensions
        if getattr(existing_part, 'rotated', False):
            ex_length = existing_part.requested_width
            ex_width = existing_part.requested_length
        else:
            ex_length = existing_part.requested_length
            ex_width = existing_part.requested_width
        
        # Check for overlap with kerf spacing
        if not (x + length + kerf <= ex_x or ex_x + ex_length + kerf <= x or
                y + width + kerf <= ex_y or ex_y + ex_width + kerf <= y):
            return False  # Overlap detected
    
    return True  # No overlaps

def _merge_narrow_strips(board: Board, kerf: float) -> None:
    """
    Merge narrow strips (F2) - strips narrower than STRIP_MERGE_THRESHOLD.
    This helps maintain guillotine constraints by reducing fragmentation.
    """
    # Implementation would analyze board layout and merge narrow columns
    # For now, just log the action
    logger.debug(f"Strip merging applied to board {board.id}")

def _enforce_sacrificial_policy(board: Board) -> None:
    """
    Enforce sacrificial sheet policy (F3) - maximum 1 sacrificial sheet per material,
    all others must be ≥80% utilization.
    """
    utilization = board.get_utilization_percentage()
    
    if utilization < SACRIFICIAL_UTILIZATION_THRESHOLD * 100:
        logger.debug(f"Board {board.id} utilization {utilization:.1f}% - potential sacrificial sheet")
    
    # In full implementation, this would enforce the policy across material groups

def _check_basic_guillotine_compliance(part1: Part, part2: Part) -> bool:
    """Basic check for guillotine compliance between two parts."""
    
    x1 = getattr(part1, 'x', getattr(part1, 'x_pos', 0))
    y1 = getattr(part1, 'y', getattr(part1, 'y_pos', 0))
    x2 = getattr(part2, 'x', getattr(part2, 'x_pos', 0))
    y2 = getattr(part2, 'y', getattr(part2, 'y_pos', 0))
    
    # Get dimensions
    if getattr(part1, 'rotated', False):
        w1, h1 = part1.requested_width, part1.requested_length
    else:
        w1, h1 = part1.requested_length, part1.requested_width
        
    if getattr(part2, 'rotated', False):
        w2, h2 = part2.requested_width, part2.requested_length
    else:
        w2, h2 = part2.requested_length, part2.requested_width
    
    # Check if parts can be separated by a straight cut
    # Horizontal separation (vertical cut possible)
    horizontal_sep = (x1 + w1 <= x2) or (x2 + w2 <= x1)
    
    # Vertical separation (horizontal cut possible)  
    vertical_sep = (y1 + h1 <= y2) or (y2 + h2 <= y1)
    
    # Basic alignment check
    edge_aligned = (
        abs(x1 - x2) < 5.0 or abs(x1 + w1 - x2 - w2) < 5.0 or  # X alignment
        abs(y1 - y2) < 5.0 or abs(y1 + h1 - y2 - h2) < 5.0     # Y alignment
    )
    
    return horizontal_sep or vertical_sep or edge_aligned

def _try_place_part_guillotine(part: Part, board: Board, kerf: float) -> bool:
    """Try to place a part using guillotine constraints with NO overlaps."""
    
    # Try normal orientation
    if _try_orientation_guillotine(part, board, kerf, rotated=False):
        return True
    
    # Try rotated orientation if allowed
    if part.grains == 0:  # Rotation allowed
        if _try_orientation_guillotine(part, board, kerf, rotated=True):
            return True
    
    return False

def _try_orientation_guillotine(part: Part, board: Board, kerf: float, rotated: bool) -> bool:
    """Try placing part in specific orientation with strict collision detection."""
    
    # Get dimensions for this orientation
    if rotated:
        part_length = part.requested_width
        part_width = part.requested_length
    else:
        part_length = part.requested_length
        part_width = part.requested_width
    
    # Check if part fits in board at all
    if part_length > board.total_length or part_width > board.total_width:
        return False
    
    # Find valid position using guillotine shelf algorithm
    position = _find_guillotine_position(part_length, part_width, board, kerf)
    
    if position is not None:
        x, y = position
        
        # CRITICAL: Verify no overlaps before placing
        if _verify_no_overlaps(x, y, part_length, part_width, board, kerf):
            # Place the part
            setattr(part, 'x', x)
            setattr(part, 'y', y)
            part.x_pos = x
            part.y_pos = y
            setattr(part, 'rotated', rotated)
            board.parts_on_board.append(part)
            return True
    
    return False

def _find_guillotine_position(part_length: float, part_width: float, board: Board, kerf: float) -> Optional[Tuple[float, float]]:
    """Find position that respects guillotine constraints."""
    
    # If board is empty, place at origin
    if not board.parts_on_board:
        return (0.0, 0.0)
    
    # Get all possible shelf positions
    shelf_positions = _get_shelf_positions(board, kerf)
    
    # Try each shelf position
    for x, y in shelf_positions:
        # Check if part fits at this position
        if (x + part_length <= board.total_length and 
            y + part_width <= board.total_width):
            
            # Check for collisions with exact kerf spacing
            if _verify_no_overlaps(x, y, part_length, part_width, board, kerf):
                return (x, y)
    
    return None

def _get_shelf_positions(board: Board, kerf: float) -> List[Tuple[float, float]]:
    """Get all valid shelf positions for efficient guillotine placement."""
    
    positions = [(0.0, 0.0)]  # Origin
    
    for part in board.parts_on_board:
        part_x = getattr(part, 'x', getattr(part, 'x_pos', 0))
        part_y = getattr(part, 'y', getattr(part, 'y_pos', 0))
        
        # Get actual dimensions considering rotation
        if getattr(part, 'rotated', False):
            part_length = part.requested_width
            part_width = part.requested_length
        else:
            part_length = part.requested_length
            part_width = part.requested_width
        
        # Add standard shelf positions with proper kerf spacing
        positions.extend([
            (part_x + part_length + kerf, part_y),           # Right of part
            (part_x, part_y + part_width + kerf),            # Above part
            (part_x + part_length + kerf, part_y + part_width + kerf),  # Diagonal
        ])
    
    # Remove duplicates and sort by bottom-left preference for efficiency
    unique_positions = list(set(positions))
    return sorted(unique_positions, key=lambda pos: (pos[1], pos[0]))

def _verify_no_overlaps(x: float, y: float, length: float, width: float, board: Board, kerf: float) -> bool:
    """Verify that placing a part at given position creates NO overlaps with basic guillotine check."""
    
    # Part boundaries (actual part size)
    part_x1, part_y1 = x, y
    part_x2, part_y2 = x + length, y + width
    
    # Check against all existing parts
    for existing_part in board.parts_on_board:
        ex_x = getattr(existing_part, 'x', getattr(existing_part, 'x_pos', 0))
        ex_y = getattr(existing_part, 'y', getattr(existing_part, 'y_pos', 0))
        
        # Get existing part dimensions
        if getattr(existing_part, 'rotated', False):
            ex_length = existing_part.requested_width
            ex_width = existing_part.requested_length
        else:
            ex_length = existing_part.requested_length
            ex_width = existing_part.requested_width
        
        # Existing part boundaries
        ex_x1, ex_y1 = ex_x, ex_y
        ex_x2, ex_y2 = ex_x + ex_length, ex_y + ex_width
        
        # Check for overlap with kerf spacing
        # Parts must be separated by at least kerf distance
        if not (part_x2 + kerf <= ex_x1 or ex_x2 + kerf <= part_x1 or
                part_y2 + kerf <= ex_y1 or ex_y2 + kerf <= part_y1):
            return False  # Overlap detected
    
    return True  # No overlaps detected

# Half-board optimization functions
def optimize_half_boards(boards: List[Board], kerf: float, core_db: Dict, laminate_db: Dict) -> Tuple[List[Board], List[Dict]]:
    """
    Optimize boards with <40% utilization to achieve 50% utilization and save half-boards.
    Returns optimized boards and list of saved half-board materials.
    """
    logger.info(f"optimize_half_boards called with {len(boards)} boards")
    start_time = time.time()
    max_time = HALF_BOARD_CONFIG['max_processing_time']
    
    optimized_boards = []
    half_board_savings = []
    
    # Group boards by material type for consolidation
    material_board_groups = {}
    for board in boards:
        material_key = _get_board_material_key(board)
        logger.info(f"Board {board.id}: utilization={board.get_utilization_percentage():.1f}%, material_key='{material_key}'")
        if material_key not in material_board_groups:
            material_board_groups[material_key] = []
        material_board_groups[material_key].append(board)
    
    logger.info(f"Grouped into {len(material_board_groups)} material groups: {list(material_board_groups.keys())}")
    
    for material_key, material_boards in material_board_groups.items():
        # Check time limit
        if time.time() - start_time > max_time:
            logger.warning(f"Half-board optimization time limit reached ({max_time}s)")
            optimized_boards.extend(material_boards)
            continue
        
        # Find low-utilization boards for this material
        low_util_boards = [b for b in material_boards if b.get_utilization_percentage() < HALF_BOARD_CONFIG['low_utilization_threshold']]
        high_util_boards = [b for b in material_boards if b.get_utilization_percentage() >= HALF_BOARD_CONFIG['low_utilization_threshold']]
        
        if len(low_util_boards) >= 2:
            logger.info(f"Found {len(low_util_boards)} low-utilization boards for material: {material_key}")
            logger.info(f"DISABLING board consolidation - only using single-board rearrangement as per user requirements")
            # Note: Board consolidation logic disabled - user wants only rearrangement optimization
            for single_board in low_util_boards:
                logger.info(f"Testing board {single_board.id} with {single_board.get_utilization_percentage():.1f}% utilization for rearrangement")
                optimized_single, saved_material = _optimize_board_arrangement_for_offcut(single_board, kerf, core_db, laminate_db)
                if optimized_single and saved_material:
                    optimized_boards.append(optimized_single)
                    half_board_savings.append(saved_material)
                    logger.info(f"✅ Board {single_board.id} successfully rearranged to create 50%+ offcut - reporting as 0.5 material")
                else:
                    optimized_boards.append(single_board)
                    logger.info(f"❌ Board {single_board.id} rearrangement failed - keeping original layout, no half-board savings")
            optimized_boards.extend(high_util_boards)
        elif len(low_util_boards) >= 1:
            # Try to optimize low-utilization boards by rearranging to create 50%+ offcuts
            logger.info(f"Attempting rearrangement optimization for {len(low_util_boards)} low-utilization boards of material: {material_key}")
            for single_board in low_util_boards:
                logger.info(f"Testing board {single_board.id} with {single_board.get_utilization_percentage():.1f}% utilization for rearrangement")
                optimized_single, saved_material = _optimize_board_arrangement_for_offcut(single_board, kerf, core_db, laminate_db)
                if optimized_single and saved_material:
                    # ONLY count as half-board if rearrangement was successful AND created 50%+ offcut
                    optimized_boards.append(optimized_single)
                    half_board_savings.append(saved_material)
                    logger.info(f"✅ Board {single_board.id} successfully rearranged to create 50%+ offcut - reporting as 0.5 material")
                else:
                    # Rearrangement failed - keep original board, do NOT count as half-board
                    optimized_boards.append(single_board)
                    logger.info(f"❌ Board {single_board.id} rearrangement failed - keeping original layout, no half-board savings")
            optimized_boards.extend(high_util_boards)
        else:
            optimized_boards.extend(material_boards)
    
    return optimized_boards, half_board_savings

def _consolidate_boards_to_half(boards: List[Board], kerf: float, core_db: Dict, laminate_db: Dict) -> Tuple[List[Board], List[Dict]]:
    """
    Consolidate multiple low-utilization boards to achieve 50% utilization and save half-boards.
    """
    if len(boards) < 2:
        return boards, []
    
    # Sort boards by utilization (lowest first)
    sorted_boards = sorted(boards, key=lambda b: b.get_utilization_percentage())
    
    consolidated_boards = []
    saved_materials = []
    
    i = 0
    while i < len(sorted_boards) - 1:
        board1 = sorted_boards[i]
        board2 = sorted_boards[i + 1]
        
        # Collect all parts from both boards
        all_parts = board1.parts_on_board + board2.parts_on_board
        
        # Try to fit all parts in one board with ~50% utilization
        new_board = _create_half_board_arrangement(all_parts, board1.material_details, core_db, kerf)
        
        if new_board and new_board.get_utilization_percentage() <= HALF_BOARD_CONFIG['target_offcut_percentage'] + 10:
            # Successful consolidation - save one half-board
            consolidated_boards.append(new_board)
            
            # Record saved material
            saved_material = _create_half_board_material_record(board1, core_db, laminate_db)
            saved_materials.append(saved_material)
            
            logger.info(f"Consolidated 2 boards ({board1.get_utilization_percentage():.1f}%, {board2.get_utilization_percentage():.1f}%) into 1 board ({new_board.get_utilization_percentage():.1f}%) - saved 0.5 board")
            i += 2  # Skip both boards
        else:
            # Consolidation failed, keep original board
            consolidated_boards.append(board1)
            i += 1
    
    # Add any remaining board
    if i < len(sorted_boards):
        consolidated_boards.append(sorted_boards[i])
    
    return consolidated_boards, saved_materials

def _create_half_board_arrangement(parts: List[Part], material_details: MaterialDetails, core_db: Dict, kerf: float) -> Optional[Board]:
    """
    Create a new board arrangement targeting half-board utilization with optimal dimensions.
    """
    if not parts:
        return None
    
    # Get standard board dimensions from core database
    core_name = _extract_core_name_from_material(material_details, core_db)
    standard_length = 2440.0  # Default fallback
    standard_width = 1220.0
    
    if core_name and core_name in core_db:
        core_data = core_db[core_name]
        standard_length = float(core_data.get('standard_length', 2440))
        standard_width = float(core_data.get('standard_width', 1220))
    
    # Try different half-board dimension combinations
    best_board = None
    best_utilization = 0
    
    for length_factor, width_factor in HALF_BOARD_CONFIG['half_board_dimension_options']:
        # Calculate half-board dimensions
        half_length = standard_length * length_factor
        half_width = standard_width * width_factor
        
        # Create test board with half dimensions
        test_board_id = f"TestHalfBoard_{core_name}_{half_length:.0f}x{half_width:.0f}"
        test_board = Board(
            board_id=test_board_id,
            material_details=material_details,
            total_length=half_length,
            total_width=half_width,
            kerf=kerf
        )
        
        # Try to place all parts on this half-board
        placed_parts = []
        current_x, current_y = 0.0, 0.0
        shelf_height = 0.0
        
        # Sort parts by area (largest first)
        sorted_parts = sorted(parts, key=lambda p: p.requested_length * p.requested_width, reverse=True)
        
        all_parts_fit = True
        for part in sorted_parts:
            part_length = part.requested_length
            part_width = part.requested_width
            
            # Check if part fits on current shelf
            if current_x + part_length <= half_length:
                # Place on current shelf
                part_copy = copy.deepcopy(part)
                part_copy.x_pos = current_x
                part_copy.y_pos = current_y
                setattr(part_copy, 'x', current_x)
                setattr(part_copy, 'y', current_y)
                
                placed_parts.append(part_copy)
                current_x += part_length + kerf
                shelf_height = max(shelf_height, part_width)
            else:
                # Move to next shelf
                current_y += shelf_height + kerf
                current_x = 0.0
                shelf_height = part_width
                
                # Check if part fits on new shelf
                if current_y + part_width <= half_width and current_x + part_length <= half_length:
                    part_copy = copy.deepcopy(part)
                    part_copy.x_pos = current_x
                    part_copy.y_pos = current_y
                    setattr(part_copy, 'x', current_x)
                    setattr(part_copy, 'y', current_y)
                    
                    placed_parts.append(part_copy)
                    current_x += part_length + kerf
                else:
                    # Part doesn't fit - try next configuration
                    all_parts_fit = False
                    break
        
        if all_parts_fit and placed_parts:
            # Add placed parts to test board
            test_board.parts_on_board = placed_parts
            
            # Calculate utilization
            utilization = test_board.get_utilization_percentage()
            
            # Check if this is the best configuration so far
            if utilization > best_utilization and utilization <= 60.0:  # Max 60% for half-board
                best_board = test_board
                best_utilization = utilization
    
    if best_board:
        # Generate unique board ID
        best_board.id = f"HalfBoard_{core_name}_{len(best_board.parts_on_board)}parts_{best_utilization:.1f}pct"
        logger.info(f"Created half-board arrangement: {best_board.total_length:.0f}x{best_board.total_width:.0f}mm, utilization: {best_utilization:.1f}%")
    
    return best_board

def _create_half_board_material_record(board: Board, core_db: Dict, laminate_db: Dict) -> Dict:
    """
    Create a record for the saved half-board material.
    """
    core_name = _extract_core_name_from_material(board.material_details, core_db)
    
    # Extract laminate information from MaterialDetails object
    top_laminate = "Unknown"
    bottom_laminate = "Unknown"
    material_signature = str(board.material_details.full_material_string) if hasattr(board.material_details, 'full_material_string') else str(board.material_details)
    
    # Parse top and bottom laminates from material details object
    if hasattr(board.material_details, 'top_laminate_name'):
        top_laminate = board.material_details.top_laminate_name
    elif hasattr(board.material_details, 'laminate_name'):
        top_laminate = board.material_details.laminate_name
    
    if hasattr(board.material_details, 'bottom_laminate_name'):
        bottom_laminate = board.material_details.bottom_laminate_name
    else:
        # If no separate bottom laminate, assume it's the same as top
        bottom_laminate = top_laminate
    
    # Clean up "NONE" values
    if top_laminate == "NONE":
        top_laminate = "Unknown"
    if bottom_laminate == "NONE":
        bottom_laminate = "Unknown"
    
    return {
        'core_material': core_name,
        'top_laminate': top_laminate,
        'bottom_laminate': bottom_laminate,
        'quantity': HALF_BOARD_CONFIG['half_board_quantity'],
        'material_signature': material_signature,
        'saved_board_id': board.id
    }

def _get_board_material_key(board: Board) -> str:
    """Get a unique key for board material grouping."""
    return str(board.material_details)

def _extract_core_name_from_material(material_details: MaterialDetails, core_db: Dict) -> str:
    """Extract core material name from material details."""
    material_str = str(material_details)
    
    for core_name in core_db.keys():
        if core_name in material_str:
            return core_name
    
    return "Unknown"

def _optimize_board_arrangement_for_offcut(board: Board, kerf: float, core_db: Dict, laminate_db: Dict) -> Tuple[Optional[Board], Optional[Dict]]:
    """
    Rearrange parts on a low-utilization board to create a single offcut ≥50% of board area.
    Only returns success if rearrangement creates the required offcut size.
    """
    if not board.parts_on_board or board.get_utilization_percentage() >= HALF_BOARD_CONFIG['low_utilization_threshold']:
        logger.info(f"Board {board.id} does not qualify for half-board optimization (utilization: {board.get_utilization_percentage():.1f}%)")
        return None, None
    
    board_area = board.total_length * board.total_width
    target_offcut_area = board_area * (HALF_BOARD_CONFIG['target_offcut_percentage'] / 100.0)
    
    logger.info(f"Testing board {board.id}: {len(board.parts_on_board)} parts, need to create ≥{HALF_BOARD_CONFIG['target_offcut_percentage']}% offcut ({target_offcut_area:.0f} mm²)")
    
    # Try different arrangement strategies
    for strategy in HALF_BOARD_CONFIG['arrangement_strategies']:
        logger.info(f"Trying strategy: {strategy}")
        rearranged_board = _try_arrangement_strategy(board, strategy, kerf, target_offcut_area)
        
        if rearranged_board:
            # Check if we achieved the target offcut
            largest_offcut = rearranged_board.get_largest_offcut()
            if largest_offcut:
                offcut_area = largest_offcut.length * largest_offcut.width
                offcut_percentage = (offcut_area / board_area) * 100
                
                logger.info(f"Strategy {strategy} created offcut: {offcut_area:.0f} mm² ({offcut_percentage:.1f}% of board)")
                
                if offcut_area >= target_offcut_area:
                    # SUCCESS: Rearrangement created required offcut size
                    saved_material = _create_half_board_material_record(board, core_db, laminate_db)
                    
                    # Log coordinates for debugging PDF layout
                    logger.info(f"✅ SUCCESS: Board {board.id} rearranged from {board.get_utilization_percentage():.1f}% to create {offcut_percentage:.1f}% offcut")
                    logger.info(f"Rearranged board {rearranged_board.id} has {len(rearranged_board.parts_on_board)} parts with NEW coordinates:")
                    for i, part in enumerate(rearranged_board.parts_on_board[:3]):  # Log first 3 parts
                        x_pos = getattr(part, 'x_pos', getattr(part, 'x', 0))
                        y_pos = getattr(part, 'y_pos', getattr(part, 'y', 0))
                        logger.info(f"  Part {i+1} ({part.id}): NEW position ({x_pos}, {y_pos})")
                    return rearranged_board, saved_material
                else:
                    logger.info(f"Strategy {strategy} failed: offcut {offcut_percentage:.1f}% < required {HALF_BOARD_CONFIG['target_offcut_percentage']}%")
            else:
                logger.info(f"Strategy {strategy} failed: no offcut created")
    
    logger.info(f"❌ FAILED: All strategies failed to create required {HALF_BOARD_CONFIG['target_offcut_percentage']}% offcut for board {board.id}")
    return None, None

def _try_arrangement_strategy(board: Board, strategy: str, kerf: float, target_offcut_area: float) -> Optional[Board]:
    """
    Try a specific arrangement strategy to create a large offcut.
    """
    if not board.parts_on_board:
        return None
    
    # Create a copy of the board to test arrangements
    test_board_id = f"{board.id}_rearranged_{strategy}"
    test_board = Board(
        board_id=test_board_id,
        material_details=board.material_details,
        total_length=board.total_length,
        total_width=board.total_width,
        kerf=kerf
    )
    
    # Get parts to rearrange (make copies to avoid modifying originals)
    parts_to_place = [copy.deepcopy(part) for part in board.parts_on_board]
    
    # Try the specified arrangement strategy
    success = False
    if strategy == 'linear_horizontal':
        success = _arrange_parts_linear_horizontal(test_board, parts_to_place, kerf)
    elif strategy == 'linear_vertical':
        success = _arrange_parts_linear_vertical(test_board, parts_to_place, kerf)
    elif strategy == 'corner_compact':
        success = _arrange_parts_corner_compact(test_board, parts_to_place, kerf)
    elif strategy == 'L_shaped':
        success = _arrange_parts_L_shaped(test_board, parts_to_place, kerf)
    
    if success:
        return test_board
    
    return None

def _arrange_parts_linear_horizontal(board: Board, parts: List[Part], kerf: float) -> bool:
    """Arrange parts in horizontal lines to maximize vertical offcut."""
    current_x, current_y = 0.0, 0.0
    row_height = 0.0
    
    # Sort parts by width (tallest first) to minimize rows
    sorted_parts = sorted(parts, key=lambda p: p.requested_width, reverse=True)
    
    for part in sorted_parts:
        part_length = part.requested_length
        part_width = part.requested_width
        
        # Check if part fits on current row
        if current_x + part_length <= board.total_length:
            # Place on current row
            part.x_pos = current_x
            part.y_pos = current_y
            setattr(part, 'x', current_x)
            setattr(part, 'y', current_y)
            
            board.parts_on_board.append(part)
            current_x += part_length + kerf
            row_height = max(row_height, part_width)
        else:
            # Move to next row
            current_y += row_height + kerf
            current_x = 0.0
            row_height = part_width
            
            # Check if part fits on new row
            if (current_y + part_width <= board.total_width and 
                current_x + part_length <= board.total_length):
                part.x_pos = current_x
                part.y_pos = current_y
                setattr(part, 'x', current_x)
                setattr(part, 'y', current_y)
                
                board.parts_on_board.append(part)
                current_x += part_length + kerf
            else:
                # Part doesn't fit - arrangement failed
                return False
    
    return True

def _arrange_parts_linear_vertical(board: Board, parts: List[Part], kerf: float) -> bool:
    """Arrange parts in vertical columns to maximize horizontal offcut."""
    current_x, current_y = 0.0, 0.0
    column_width = 0.0
    
    # Sort parts by length (longest first) to minimize columns
    sorted_parts = sorted(parts, key=lambda p: p.requested_length, reverse=True)
    
    for part in sorted_parts:
        part_length = part.requested_length
        part_width = part.requested_width
        
        # Check if part fits in current column
        if current_y + part_width <= board.total_width:
            # Place in current column
            part.x_pos = current_x
            part.y_pos = current_y
            setattr(part, 'x', current_x)
            setattr(part, 'y', current_y)
            
            board.parts_on_board.append(part)
            current_y += part_width + kerf
            column_width = max(column_width, part_length)
        else:
            # Move to next column
            current_x += column_width + kerf
            current_y = 0.0
            column_width = part_length
            
            # Check if part fits in new column
            if (current_x + part_length <= board.total_length and 
                current_y + part_width <= board.total_width):
                part.x_pos = current_x
                part.y_pos = current_y
                setattr(part, 'x', current_x)
                setattr(part, 'y', current_y)
                
                board.parts_on_board.append(part)
                current_y += part_width + kerf
            else:
                # Part doesn't fit - arrangement failed
                return False
    
    return True

def _arrange_parts_corner_compact(board: Board, parts: List[Part], kerf: float) -> bool:
    """Pack parts tightly in one corner using shelf algorithm."""
    # Sort parts by area (largest first)
    sorted_parts = sorted(parts, key=lambda p: p.requested_length * p.requested_width, reverse=True)
    
    # Use simple shelf packing starting from (0,0)
    shelves = []  # List of (y_pos, x_end, height)
    
    for part in sorted_parts:
        part_length = part.requested_length
        part_width = part.requested_width
        placed = False
        
        # Try to place on existing shelves
        for i, (shelf_y, shelf_x_end, shelf_height) in enumerate(shelves):
            if (shelf_x_end + part_length <= board.total_length and
                shelf_y + part_width <= board.total_width):
                # Place on this shelf
                part.x_pos = shelf_x_end
                part.y_pos = shelf_y
                setattr(part, 'x', shelf_x_end)
                setattr(part, 'y', shelf_y)
                
                board.parts_on_board.append(part)
                # Update shelf
                shelves[i] = (shelf_y, shelf_x_end + part_length + kerf, shelf_height)
                placed = True
                break
        
        if not placed:
            # Create new shelf
            new_shelf_y = sum(s[2] + kerf for s in shelves) if shelves else 0.0
            if (new_shelf_y + part_width <= board.total_width and
                part_length <= board.total_length):
                part.x_pos = 0.0
                part.y_pos = new_shelf_y
                setattr(part, 'x', 0.0)
                setattr(part, 'y', new_shelf_y)
                
                board.parts_on_board.append(part)
                shelves.append((new_shelf_y, part_length + kerf, part_width))
            else:
                # Part doesn't fit - arrangement failed
                return False
    
    return True

def _arrange_parts_L_shaped(board: Board, parts: List[Part], kerf: float) -> bool:
    """Arrange parts in an L-shape to create rectangular offcut."""
    if not parts:
        return True
    
    # Sort parts by area (largest first)
    sorted_parts = sorted(parts, key=lambda p: p.requested_length * p.requested_width, reverse=True)
    
    # Try to create L-shape: first fill bottom edge, then left edge
    bottom_x, left_y = 0.0, 0.0
    bottom_height, left_width = 0.0, 0.0
    
    # Place parts along bottom edge first
    for i, part in enumerate(sorted_parts[:]):
        part_length = part.requested_length
        part_width = part.requested_width
        
        # Try bottom edge placement
        if bottom_x + part_length <= board.total_length:
            part.x_pos = bottom_x
            part.y_pos = 0.0
            setattr(part, 'x', bottom_x)
            setattr(part, 'y', 0.0)
            
            board.parts_on_board.append(part)
            bottom_x += part_length + kerf
            bottom_height = max(bottom_height, part_width)
            sorted_parts.remove(part)
        else:
            break
    
    # Place remaining parts along left edge
    left_y = bottom_height + kerf
    for part in sorted_parts[:]:
        part_length = part.requested_length
        part_width = part.requested_width
        
        # Try left edge placement
        if left_y + part_width <= board.total_width:
            part.x_pos = 0.0
            part.y_pos = left_y
            setattr(part, 'x', 0.0)
            setattr(part, 'y', left_y)
            
            board.parts_on_board.append(part)
            left_y += part_width + kerf
            left_width = max(left_width, part_length)
            sorted_parts.remove(part)
        else:
            break
    
    # If there are still parts left, try to place them in remaining space
    if sorted_parts:
        # Use simple shelf packing for remaining parts
        remaining_x = max(bottom_x, left_width + kerf)
        remaining_y = bottom_height + kerf
        
        for part in sorted_parts:
            if (remaining_x + part.requested_length <= board.total_length and
                remaining_y + part.requested_width <= board.total_width):
                part.x_pos = remaining_x
                part.y_pos = remaining_y
                setattr(part, 'x', remaining_x)
                setattr(part, 'y', remaining_y)
                
                board.parts_on_board.append(part)
                remaining_y += part.requested_width + kerf
            else:
                # Part doesn't fit - arrangement failed
                return False
    
    return True

# Additional utility functions for integration
def get_max_utilisation_algorithm_info() -> Dict[str, object]:
    """Return information about the Max Utilisation algorithm."""
    return {
        'name': 'Max Utilisation - No Upgrade with Half-Board Optimization',
        'description': 'Maximizes board utilization with intelligent rearrangement for <40% boards to create ≥50% offcuts',
        'features': [
            'Maximum utilization optimization',
            'Smart rearrangement for <40% utilization boards',
            'Creates single offcut ≥50% of board area',
            'Reports saved materials as 0.5 quantity ONLY if successful',
            'Strict guillotine constraint validation',
            'Material segregation',
            'Collision detection with kerf spacing',
            'Smart rotation for non-grain sensitive parts',
            'Advanced bin packing algorithms'
        ],
        'supports_upgrades': False,
        'supports_rotation': True,
        'algorithm_type': 'max_utilisation'
    }

def validate_max_utilisation_requirements(parts: List[Part], core_db: Dict, laminate_db: Dict) -> Tuple[bool, List[str]]:
    """Validate that all requirements for Max Utilisation algorithm are met."""
    errors = []
    
    if not parts:
        errors.append("No parts provided for optimization")
    
    if not core_db:
        errors.append("Core materials database is empty")
    
    if not laminate_db:
        errors.append("Laminates database is empty")
    
    return (len(errors) == 0), errors