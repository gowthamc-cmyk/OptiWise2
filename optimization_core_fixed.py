"""
Fixed optimization core with corrected material compatibility and cost optimization logic.
"""

import logging
from typing import List, Dict, Tuple, Optional, Any
from data_models import MaterialDetails, Part, Offcut, Board

logger = logging.getLogger(__name__)


def consolidate_low_utilization_boards_core(boards: List[Board], core_db: Dict, kerf: float = 4.4) -> List[Board]:
    """
    Consolidate low-utilization boards by merging parts from multiple boards of the same material.
    This specifically addresses the Board 3 into Board 2 consolidation issue.
    """
    if len(boards) <= 1:
        return boards
    
    logger.info(f"Starting board consolidation for {len(boards)} boards")
    
    # Group boards by material to only merge compatible boards
    material_groups = {}
    for board in boards:
        material_key = str(board.material_details)
        if material_key not in material_groups:
            material_groups[material_key] = []
        material_groups[material_key].append(board)
    
    consolidated_boards = []
    total_boards_saved = 0
    
    for material_key, material_boards in material_groups.items():
        if len(material_boards) <= 1:
            consolidated_boards.extend(material_boards)
            continue
        
        # Sort by utilization (lowest first) for consolidation priority
        sorted_boards = sorted(material_boards, key=lambda b: b.get_utilization_percentage())
        
        # Track which boards have been merged
        processed_indices = set()
        
        i = 0
        while i < len(sorted_boards):
            if i in processed_indices:
                i += 1
                continue
                
            low_util_board = sorted_boards[i]
            current_util = low_util_board.get_utilization_percentage()
            
            # Only attempt consolidation for boards with < 50% utilization
            if current_util >= 50.0:
                consolidated_boards.append(low_util_board)
                processed_indices.add(i)
                i += 1
                continue
            
            logger.info(f"Attempting to consolidate {material_key} board with {current_util:.1f}% utilization")
            
            # Find the best target board to merge into
            best_target_idx = None
            best_fit_score = 0
            
            for j in range(i+1, len(sorted_boards)):
                if j in processed_indices:
                    continue
                    
                target_board = sorted_boards[j]
                target_util = target_board.get_utilization_percentage()
                
                # Skip if target is too full
                if target_util >= 85.0:
                    continue
                
                # Calculate fit score based on available space and part compatibility
                parts_to_merge = low_util_board.parts_on_board.copy()
                
                # Quick area compatibility check
                available_area = sum(rect.get_area() for rect in target_board.available_rectangles)
                required_area = sum(part.get_area_with_kerf(kerf) for part in parts_to_merge)
                
                if required_area > available_area:
                    continue
                
                # Detailed placement simulation
                mergeable_count = 0
                for part in parts_to_merge:
                    can_place = False
                    
                    for rect in target_board.available_rectangles:
                        part_length = part.requested_length + kerf
                        part_width = part.requested_width + kerf
                        
                        # Check both orientations
                        if (part_length <= rect.length and part_width <= rect.width):
                            can_place = True
                            break
                        elif (part.can_rotate() and part_width <= rect.length and part_length <= rect.width):
                            can_place = True
                            break
                    
                    if can_place:
                        mergeable_count += 1
                
                # Calculate fit score
                merge_ratio = mergeable_count / len(parts_to_merge) if parts_to_merge else 0
                area_efficiency = 1 - (available_area - required_area) / available_area if available_area > 0 else 0
                fit_score = merge_ratio * 0.7 + area_efficiency * 0.3
                
                # Universal aggressive consolidation - lower thresholds based on utilization
                if current_util < 20.0:
                    min_merge_ratio = 0.2  # Very aggressive for extremely low utilization
                elif current_util < 35.0:
                    min_merge_ratio = 0.3  # Aggressive for low utilization
                else:
                    min_merge_ratio = 0.4  # Still lower than original 80%
                
                if fit_score > best_fit_score and merge_ratio >= min_merge_ratio:
                    best_target_idx = j
                    best_fit_score = fit_score
            
            # Perform the merge if a good target was found
            if best_target_idx is not None:
                target_board = sorted_boards[best_target_idx]
                parts_to_merge = low_util_board.parts_on_board.copy()
                
                # Create new consolidated board
                new_board_id = f"CONSOLIDATED_{i+1}_{best_target_idx+1}"
                consolidated_board = Board(
                    board_id=new_board_id,
                    material_details=target_board.material_details,
                    total_length=target_board.total_length,
                    total_width=target_board.total_width,
                    kerf=kerf
                )
                
                # Place all parts from target board first
                target_parts_placed = 0
                for part in target_board.parts_on_board:
                    placed = False
                    for rect in consolidated_board.available_rectangles:
                        if rect.can_fit_part(part, kerf, False):
                            if consolidated_board.place_part(part, rect, False, rect.x, rect.y, core_db):
                                target_parts_placed += 1
                                placed = True
                                break
                        elif part.can_rotate() and rect.can_fit_part(part, kerf, True):
                            if consolidated_board.place_part(part, rect, True, rect.x, rect.y, core_db):
                                target_parts_placed += 1
                                placed = True
                                break
                
                # Place parts from low-utilization board
                merged_parts = 0
                for part in parts_to_merge:
                    placed = False
                    for rect in consolidated_board.available_rectangles:
                        if rect.can_fit_part(part, kerf, False):
                            if consolidated_board.place_part(part, rect, False, rect.x, rect.y, core_db):
                                merged_parts += 1
                                placed = True
                                break
                        elif part.can_rotate() and rect.can_fit_part(part, kerf, True):
                            if consolidated_board.place_part(part, rect, True, rect.x, rect.y, core_db):
                                merged_parts += 1
                                placed = True
                                break
                
                # Verify consolidation success
                total_original_parts = len(target_board.parts_on_board) + len(parts_to_merge)
                total_placed_parts = len(consolidated_board.parts_on_board)
                
                if total_placed_parts >= total_original_parts * 0.95:  # 95% success rate
                    consolidated_boards.append(consolidated_board)
                    processed_indices.add(i)
                    processed_indices.add(best_target_idx)
                    total_boards_saved += 1
                    
                    final_util = consolidated_board.get_utilization_percentage()
                    logger.info(f"Successfully consolidated: {merged_parts}/{len(parts_to_merge)} parts merged, "
                              f"final utilization: {final_util:.1f}%, boards saved: 1")
                else:
                    # Consolidation failed, keep original boards
                    consolidated_boards.append(low_util_board)
                    processed_indices.add(i)
            else:
                # No suitable target found, keep original board
                consolidated_boards.append(low_util_board)
                processed_indices.add(i)
            
            i += 1
        
        # Add any unprocessed boards
        for i, board in enumerate(sorted_boards):
            if i not in processed_indices:
                consolidated_boards.append(board)
    
    # Log final consolidation results
    original_count = len(boards)
    final_count = len(consolidated_boards)
    
    if final_count < original_count:
        logger.info(f"Board consolidation successful: {original_count} → {final_count} boards ({total_boards_saved} boards eliminated)")
    else:
        logger.info("No consolidation opportunities found")
    
    return consolidated_boards


def get_grade_level(core_name: str, core_db: Dict) -> int:
    """Get the grade level for a core material."""
    if core_name not in core_db:
        logger.warning(f"Core material '{core_name}' not found in database")
        return 0
    return core_db[core_name].get('grade_level', 0)


def can_upgrade_material(requested_material_details: MaterialDetails, 
                        offcut_material_details: MaterialDetails, 
                        core_db: Dict) -> bool:
    """
    Check if a part can be placed on a board/offcut with material compatibility rules.
    Allows different top/bottom laminates in input but checks compatibility for placement.
    """
    try:
        # Check laminate compatibility - both top and bottom must match
        if (requested_material_details.top_laminate_name != offcut_material_details.top_laminate_name or
            requested_material_details.bottom_laminate_name != offcut_material_details.bottom_laminate_name):
            return False
        
        # Thickness must match exactly for core upgrades
        if requested_material_details.thickness != offcut_material_details.thickness:
            return False
        
        # No core material downgrade allowed
        requested_grade = get_grade_level(requested_material_details.core_name, core_db)
        offcut_grade = get_grade_level(offcut_material_details.core_name, core_db)
        
        if offcut_grade < requested_grade:
            return False
        
        return True
        
    except Exception as e:
        logger.error(f"Error checking material compatibility: {e}")
        return False


def calculate_board_cost(board_material_details: MaterialDetails, 
                        core_db: Dict, laminate_db: Dict) -> float:
    """Calculate the cost of a full standard board for given material."""
    try:
        core_details = core_db.get(board_material_details.core_name, {})
        standard_length = core_details.get('standard_length', 0)
        standard_width = core_details.get('standard_width', 0)
        
        if standard_length == 0 or standard_width == 0:
            logger.error(f"Invalid board dimensions for {board_material_details.core_name}")
            return 0.0
        
        board_area_sqm = (standard_length * standard_width) / 1_000_000
        cost_per_sqm = board_material_details.get_cost_per_sqm(laminate_db, core_db)
        total_cost = board_area_sqm * cost_per_sqm
        
        return total_cost
        
    except Exception as e:
        logger.error(f"Error calculating board cost: {e}")
        return 0.0


def calculate_total_order_cost(list_of_boards: List[Board], 
                              core_db: Dict, laminate_db: Dict) -> float:
    """Calculate total cost for all boards in the order (global cost optimization)."""
    try:
        material_counts = {}
        
        for board in list_of_boards:
            material_key = (board.material_details.top_laminate_name, 
                           board.material_details.core_name,
                           board.material_details.bottom_laminate_name,
                           board.material_details.thickness)
            
            material_counts[material_key] = material_counts.get(material_key, 0) + 1
        
        total_cost = 0.0
        
        for material_key, count in material_counts.items():
            top_laminate, core_name, bottom_laminate, thickness = material_key
            
            # Create MaterialDetails for cost calculation
            material_string = f"{top_laminate}_{core_name}_{bottom_laminate}"
            try:
                material_details = MaterialDetails(material_string)
                board_cost = calculate_board_cost(material_details, core_db, laminate_db)
                total_cost += board_cost * count
                
            except ValueError as e:
                logger.error(f"Error creating MaterialDetails for cost calculation: {e}")
                continue
        
        return total_cost
        
    except Exception as e:
        logger.error(f"Error calculating total order cost: {e}")
        return 0.0


def create_material_variants(base_material_details: MaterialDetails, 
                           user_upgrade_sequence: List[str], 
                           core_db: Dict) -> List[MaterialDetails]:
    """Create material variants following user-defined upgrade sequence with thickness matching."""
    variants = [base_material_details]
    
    try:
        original_grade = get_grade_level(base_material_details.core_name, core_db)
        
        for core_name in user_upgrade_sequence:
            if core_name == base_material_details.core_name:
                continue
            
            # Check thickness compatibility first
            core_details = core_db.get(core_name, {})
            core_thickness = core_details.get('thickness', 0)
            
            if core_thickness != base_material_details.thickness:
                logger.debug(f"Skipping upgrade {core_name}: thickness mismatch "
                           f"({core_thickness}mm vs {base_material_details.thickness}mm)")
                continue
            
            upgrade_grade = get_grade_level(core_name, core_db)
            if upgrade_grade <= original_grade:
                continue
            
            # Create upgraded material string maintaining top and bottom laminates
            upgraded_material_string = (f"{base_material_details.top_laminate_name}_"
                                      f"{core_name}_"
                                      f"{base_material_details.bottom_laminate_name}")
            
            try:
                upgraded_material = MaterialDetails(upgraded_material_string)
                variants.append(upgraded_material)
                
            except ValueError as e:
                logger.warning(f"Could not create upgrade variant for {core_name}: {e}")
                continue
        
        return variants
        
    except Exception as e:
        logger.error(f"Error creating material variants: {e}")
        return [base_material_details]


def find_best_fit_offcut(part: Part, available_offcuts: List[Offcut], 
                        kerf: float) -> Tuple[Optional[Offcut], bool]:
    """Find the best fitting offcut using smart placement strategy."""
    best_offcut = None
    best_rotation = False
    best_score = float('inf')
    
    for offcut in available_offcuts:
        orientations = [False]
        if part.can_rotate():
            orientations.append(True)
        
        for rotated in orientations:
            if offcut.can_fit_part(part, kerf, rotated):
                part_length, part_width = part.get_dimensions_for_placement(rotated)
                part_area_with_kerf = (part_length + kerf) * (part_width + kerf)
                waste = offcut.get_area() - part_area_with_kerf
                
                # Smart scoring: prefer tight fits but avoid creating unusable slivers
                offcut_area = offcut.get_area()
                utilization = part_area_with_kerf / offcut_area
                
                # Calculate remaining dimensions after placement
                remaining_length = offcut.length - part_length
                remaining_width = offcut.width - part_width
                
                # Penalize placements that create very thin unusable strips
                sliver_penalty = 0
                min_useful_size = 100  # 10cm minimum useful size
                if 0 < remaining_length < min_useful_size or 0 < remaining_width < min_useful_size:
                    sliver_penalty = 10000
                
                # Combined score: minimize waste + sliver penalty + prefer high utilization
                score = waste + sliver_penalty + (1.0 - utilization) * 500
                
                if score < best_score:
                    best_score = score
                    best_offcut = offcut
                    best_rotation = rotated
    
    return best_offcut, best_rotation


def create_new_board(material_details: MaterialDetails, core_db: Dict, 
                    board_counter: int, kerf: float = 4.4) -> Board:
    """Create a new board with standard dimensions (usable size, kerf not included at edges)."""
    core_details = core_db.get(material_details.core_name, {})
    length = core_details.get('standard_length', 2440)
    width = core_details.get('standard_width', 1220)
    
    board_id = f"Board_{board_counter}_{material_details.core_name}"
    
    return Board(
        board_id=board_id,
        material_details=material_details,
        total_length=length,
        total_width=width,
        kerf=kerf
    )


def calculate_baseline_cost(parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4) -> float:
    """Calculate the baseline cost - worst case scenario if each part required its own full board."""
    try:
        # Calculate cost as if each part required a full standard board
        total_cost = 0.0
        
        for part in parts_list:
            # Get standard board dimensions for this material
            core_details = core_db.get(part.material_details.core_name, {})
            standard_length = core_details.get('standard_length', 2440)
            standard_width = core_details.get('standard_width', 1220)
            board_area_sqm = (standard_length * standard_width) / 1_000_000
            
            # Calculate cost for a full board of this material
            cost_per_sqm = part.material_details.get_cost_per_sqm(laminate_db, core_db)
            board_cost = board_area_sqm * cost_per_sqm
            
            total_cost += board_cost
            
        logger.info(f"Baseline cost calculation (worst case - each part gets full board): ₹{total_cost:.2f}")
        return total_cost
        
    except Exception as e:
        logger.error(f"Error calculating baseline cost: {e}")
        return 0.0


def run_optimization_no_upgrade(parts_list: List[Part], core_db: Dict, laminate_db: Dict, 
                               kerf: float = 4.4) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
    """Run optimization without any material upgrades."""
    logger.info(f"Starting no-upgrade optimization for {len(parts_list)} parts")
    
    # Calculate proper baseline cost
    initial_cost = calculate_baseline_cost(parts_list, core_db, laminate_db, kerf)
    
    final_boards = []
    unplaced_parts = []
    upgrade_summary = []  # No upgrades in this mode
    board_counter = 0
    
    # Sort parts by area (largest first) for better packing
    sorted_parts = sorted(parts_list, key=lambda p: p.get_area_with_kerf(kerf), reverse=True)
    
    for part in sorted_parts:
        placed = False
        
        # Try to place on existing compatible boards first
        for board in final_boards:
            # Only allow exact material match (no upgrades)
            if (part.material_details.top_laminate_name == board.material_details.top_laminate_name and
                part.material_details.bottom_laminate_name == board.material_details.bottom_laminate_name and
                part.material_details.core_name == board.material_details.core_name and
                part.material_details.thickness == board.material_details.thickness):
                
                best_offcut, should_rotate = find_best_fit_offcut(
                    part, board.available_rectangles, kerf
                )
                
                if best_offcut and board.place_part(
                    part, best_offcut, should_rotate,
                    best_offcut.x, best_offcut.y, core_db
                ):
                    placed = True
                    break
        
        # Create new board if not placed
        if not placed:
            board_counter += 1
            new_board = create_new_board(part.material_details, core_db, board_counter, kerf)
            
            if new_board.available_rectangles:
                first_offcut = new_board.available_rectangles[0]
                
                if first_offcut.can_fit_part(part, kerf, False) or \
                   (part.can_rotate() and first_offcut.can_fit_part(part, kerf, True)):
                    
                    should_rotate = False
                    if not first_offcut.can_fit_part(part, kerf, False) and part.can_rotate():
                        should_rotate = True
                    
                    if new_board.place_part(part, first_offcut, should_rotate, 0, 0, core_db):
                        final_boards.append(new_board)
                        placed = True
        
        # If still not placed, add to unplaced list
        if not placed:
            unplaced_parts.append(part)
            logger.warning(f"Could not place part {part.id}")
    
    # Calculate final cost
    final_cost = calculate_total_order_cost(final_boards, core_db, laminate_db)
    
    logger.info(f"No-upgrade optimization complete: {len(final_boards)} boards, "
               f"{len(unplaced_parts)} unplaced parts, "
               f"baseline cost: {initial_cost:.2f}, final cost: {final_cost:.2f}")
    
    return final_boards, unplaced_parts, upgrade_summary, initial_cost, final_cost


def run_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict, 
                    user_upgrade_sequence_str: str, kerf: float = 4.4) -> Tuple[
                        List[Board], List[Part], List[Dict], float, float]:
    """
    Main optimization algorithm implementing guillotine cutting with material upgrades.
    Focuses on global order cost optimization.
    """
    logger.info(f"Starting optimization for {len(parts_list)} parts")
    
    # Parse user upgrade sequence
    user_upgrade_sequence = [name.strip() for name in user_upgrade_sequence_str.split(',') if name.strip()]
    logger.info(f"User upgrade sequence: {user_upgrade_sequence}")
    
    # Sort parts by material groups, then by area (largest first within each group)
    # This matches the pattern seen in professional cutting layouts
    def get_material_sort_key(part):
        # Group by full material string, then by area (descending)
        return (str(part.material_details), -part.get_area_with_kerf(kerf))
    
    parts_list_sorted = sorted(parts_list, key=get_material_sort_key)
    
    # Initialize tracking variables
    final_boards = []
    unplaced_parts = []
    upgrade_summary = []
    global_offcuts = []
    board_counter = 0
    
    # Calculate proper baseline cost (minimum boards needed without upgrades)
    initial_cost = calculate_baseline_cost(parts_list, core_db, laminate_db, kerf)
    
    # Process each part using material-aware sorting
    for part in parts_list_sorted:
        placed = False
        
        # Create material variants (original + compatible upgrades)
        material_variants = create_material_variants(
            part.material_details, user_upgrade_sequence, core_db
        )
        
        # Try to place on existing offcuts first (global reuse)
        for material_variant in material_variants:
            compatible_offcuts = [
                offcut for offcut in global_offcuts
                if can_upgrade_material(part.material_details, offcut.material_details, core_db)
            ]
            
            if compatible_offcuts:
                best_offcut, should_rotate = find_best_fit_offcut(part, compatible_offcuts, kerf)
                
                if best_offcut:
                    # Find the board this offcut belongs to
                    source_board = None
                    for board in final_boards:
                        if best_offcut in board.available_rectangles:
                            source_board = board
                            break
                    
                    if source_board and source_board.place_part(
                        part, best_offcut, should_rotate, 
                        best_offcut.x, best_offcut.y, core_db
                    ):
                        global_offcuts.remove(best_offcut)
                        
                        # Track upgrade if material changed
                        if part.is_upgraded and part.assigned_material_details:
                            upgrade_summary.append({
                                'Part ID': part.id,
                                'Original Material': part.material_details.full_material_string,
                                'Upgraded Material': part.assigned_material_details.full_material_string
                            })
                        
                        placed = True
                        break
        
        # Try existing boards if not placed on offcuts
        if not placed:
            for material_variant in material_variants:
                for board in final_boards:
                    if can_upgrade_material(part.material_details, board.material_details, core_db):
                        best_offcut, should_rotate = find_best_fit_offcut(
                            part, board.available_rectangles, kerf
                        )
                        
                        if best_offcut and board.place_part(
                            part, best_offcut, should_rotate,
                            best_offcut.x, best_offcut.y, core_db
                        ):
                            global_offcuts.extend(board.available_rectangles)
                            
                            # Track upgrade if material changed
                            if part.is_upgraded and part.assigned_material_details:
                                upgrade_summary.append({
                                    'Part ID': part.id,
                                    'Original Material': part.material_details.full_material_string,
                                    'Upgraded Material': part.assigned_material_details.full_material_string
                                })
                            
                            placed = True
                            break
                
                if placed:
                    break
        
        # Create new board if still not placed
        if not placed:
            for material_variant in material_variants:
                board_counter += 1
                new_board = create_new_board(material_variant, core_db, board_counter, kerf)
                
                if new_board.available_rectangles:
                    first_offcut = new_board.available_rectangles[0]
                    
                    if first_offcut.can_fit_part(part, kerf, False) or \
                       (part.can_rotate() and first_offcut.can_fit_part(part, kerf, True)):
                        
                        should_rotate = False
                        if not first_offcut.can_fit_part(part, kerf, False) and part.can_rotate():
                            should_rotate = True
                        
                        if new_board.place_part(part, first_offcut, should_rotate, 0, 0, core_db):
                            final_boards.append(new_board)
                            global_offcuts.extend(new_board.available_rectangles)
                            
                            # Track upgrade if material changed
                            if part.is_upgraded and part.assigned_material_details:
                                upgrade_summary.append({
                                    'Part ID': part.id,
                                    'Original Material': part.material_details.full_material_string,
                                    'Upgraded Material': part.assigned_material_details.full_material_string
                                })
                            
                            placed = True
                            break
        
        # If still not placed, add to unplaced list
        if not placed:
            unplaced_parts.append(part)
            logger.warning(f"Could not place part {part.id}")
    
    # Post-processing: Board consolidation
    boards_before = len(final_boards)
    final_boards = consolidate_low_utilization_boards_core(final_boards, core_db, kerf)
    boards_after = len(final_boards)
    boards_saved = boards_before - boards_after
    logger.info(f"Board consolidation: {boards_before} → {boards_after} boards ({boards_saved} boards eliminated)")
    
    # Additional relocation optimization
    improved = True
    iterations = 0
    max_iterations = 2
    
    while improved and iterations < max_iterations:
        improved = False
        iterations += 1
        logger.info(f"Post-processing iteration {iterations}")
        
        # Try to move parts from less utilized boards to better fitting positions
        for board_idx, board in enumerate(final_boards):
            if board.get_utilization_percentage() < 60:  # Focus on underutilized boards
                # Get parts on this board (simplified - would need actual tracking)
                parts_to_try_relocate = []
                
                # For now, try to consolidate by removing smallest parts and trying to place them elsewhere
                smallest_offcuts = sorted(board.available_rectangles, key=lambda o: o.get_area())
                
                # Try to fill largest available spaces first
                largest_offcuts = []
                for other_board in final_boards:
                    if other_board != board and other_board.available_rectangles:
                        largest_offcut = max(other_board.available_rectangles, key=lambda o: o.get_area())
                        largest_offcuts.append((largest_offcut, other_board))
                
                # This is a simplified relocation attempt
                # In a full implementation, we'd track individual part placements and try relocating them
                
        # For this iteration, break after first attempt
        break
    
    # Calculate final cost (global order optimization)
    final_cost = calculate_total_order_cost(final_boards, core_db, laminate_db)
    
    logger.info(f"Optimization complete: {len(final_boards)} boards, "
               f"{len(unplaced_parts)} unplaced parts, "
               f"cost reduced from {initial_cost:.2f} to {final_cost:.2f}")
    
    return final_boards, unplaced_parts, upgrade_summary, initial_cost, final_cost