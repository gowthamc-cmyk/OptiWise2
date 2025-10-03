"""
Global optimization with iterative offcut reuse and cost minimization.
Implements true global offcut inventory management.
"""
import logging
from typing import List, Dict, Tuple, Optional, Set
from data_models import Part, Board, Offcut, MaterialDetails
from optimization_core_fixed import (
    run_optimization, can_upgrade_material, find_best_fit_offcut, 
    calculate_total_order_cost, get_grade_level, create_new_board
)

logger = logging.getLogger(__name__)

# Global log collector for capturing optimization messages
optimization_log_messages = []

def add_log_message(message: str):
    """Add a message to the optimization log."""
    global optimization_log_messages
    optimization_log_messages.append(message)
    logger.info(message)

class GlobalOffcutInventory:
    """Manages global offcut inventory across all boards."""
    
    def __init__(self):
        self.offcuts_by_material = {}  # Material -> List[Offcut with board reference]
        self.boards_registry = {}      # Board ID -> Board object
    
    def add_board(self, board: Board):
        """Add a board and its offcuts to the global inventory."""
        self.boards_registry[board.id] = board
        
        material_key = self._get_material_key(board.material_details)
        if material_key not in self.offcuts_by_material:
            self.offcuts_by_material[material_key] = []
        
        # Add all available rectangles from this board
        for offcut in board.available_rectangles:
            offcut.source_board_id = board.id
            self.offcuts_by_material[material_key].append(offcut)
    
    def _get_material_key(self, material_details: MaterialDetails) -> str:
        """Create a key for material grouping."""
        return f"{material_details.top_laminate_name}_{material_details.core_name}_{material_details.thickness}"
    
    def find_compatible_offcuts(self, part: Part, core_db: Dict, kerf: float) -> List[Tuple[Offcut, Board, bool]]:
        """Find all compatible offcuts for a part across all boards."""
        compatible = []
        
        # Check all material types for compatibility
        for material_key, offcuts in self.offcuts_by_material.items():
            for offcut in offcuts:
                board = self.boards_registry[offcut.source_board_id]
                
                # Check material compatibility (allow upgrades)
                if can_upgrade_material(part.material_details, offcut.material_details, core_db):
                    # Check if part fits (normal orientation)
                    if offcut.can_fit_part(part, kerf, rotated=False):
                        compatible.append((offcut, board, False))
                    
                    # Check if part fits (rotated orientation) 
                    if part.can_rotate() and offcut.can_fit_part(part, kerf, rotated=True):
                        compatible.append((offcut, board, True))
        
        # Sort by upgrade potential and space efficiency - PRIORITIZE UPGRADES
        def sort_key(item):
            offcut, board, rotated = item
            part_area = part.requested_length * part.requested_width
            offcut_area = offcut.get_area()
            
            # Get grade levels for upgrade calculation
            part_grade = get_grade_level(part.material_details.core_name, core_db)
            offcut_grade = get_grade_level(offcut.material_details.core_name, core_db)
            upgrade_potential = offcut_grade - part_grade
            
            # PRIORITIZE UPGRADES - negative penalty for upgrades means they come first
            upgrade_bonus = -1000 * upgrade_potential if upgrade_potential > 0 else 0
            
            # Prefer tight fit (less waste)
            space_efficiency = part_area / offcut_area
            
            return (upgrade_bonus, -space_efficiency)
        
        compatible.sort(key=sort_key)
        return compatible
    
    def remove_offcut(self, offcut: Offcut):
        """Remove an offcut from the global inventory."""
        material_key = self._get_material_key(offcut.material_details)
        if material_key in self.offcuts_by_material:
            if offcut in self.offcuts_by_material[material_key]:
                self.offcuts_by_material[material_key].remove(offcut)
    
    def get_total_waste_area(self) -> float:
        """Calculate total waste area across all offcuts."""
        total_area = 0
        for offcuts in self.offcuts_by_material.values():
            for offcut in offcuts:
                total_area += offcut.get_area()
        return total_area
    
    def get_offcut_summary(self) -> Dict:
        """Get summary of offcut inventory."""
        summary = {}
        for material_key, offcuts in self.offcuts_by_material.items():
            if offcuts:
                total_area = sum(offcut.get_area() for offcut in offcuts)
                summary[material_key] = {
                    'count': len(offcuts),
                    'total_area': total_area,
                    'avg_area': total_area / len(offcuts) if offcuts else 0
                }
        return summary

def relocate_part_globally(part: Part, target_offcut: Offcut, target_board: Board, 
                          source_board: Board, rotated: bool, core_db: Dict, 
                          global_inventory: GlobalOffcutInventory, kerf: float) -> bool:
    """Relocate a part using global offcut inventory."""
    try:
        # Remove part from source board
        if part in source_board.parts_on_board:
            source_board.parts_on_board.remove(part)
            
            # Place part on target board
            success = target_board.place_part(
                part, target_offcut, rotated, 
                target_offcut.x, target_offcut.y, core_db
            )
            
            if success:
                # Update global inventory
                global_inventory.remove_offcut(target_offcut)
                
                # Mark as upgraded if material changed
                if part.material_details.core_name != target_offcut.material_details.core_name:
                    part.is_upgraded = True
                
                # Refresh inventory with updated boards
                global_inventory.add_board(source_board)
                global_inventory.add_board(target_board)
                
                return True
            else:
                # Restore part if placement failed
                source_board.parts_on_board.append(part)
        
        return False
        
    except Exception as e:
        logger.error(f"Error relocating part {part.id}: {e}")
        # Restore part if there was an error
        if part not in source_board.parts_on_board:
            source_board.parts_on_board.append(part)
        return False

def run_global_optimization_iteration(boards: List[Board], unplaced_parts: List[Part], 
                                    core_db: Dict, laminate_db: Dict, kerf: float, user_upgrade_sequence: List[str]) -> Tuple[List[Board], List[Part], int, List[str]]:
    """Run one iteration of global optimization."""
    iteration_log = []
    relocations_made = 0
    
    # Create global offcut inventory
    global_inventory = GlobalOffcutInventory()
    for board in boards:
        global_inventory.add_board(board)
    
    initial_waste = global_inventory.get_total_waste_area()
    initial_cost = calculate_total_order_cost(boards, core_db, laminate_db)
    
    logger.info(f"Starting iteration - Initial waste: {initial_waste/1000000:.2f}m², Cost: ₹{initial_cost:.2f}")
    
    # Try to place unplaced parts first
    parts_placed = 0
    for part in unplaced_parts[:]:  # Copy list to modify during iteration
        compatible_offcuts = global_inventory.find_compatible_offcuts(part, core_db, kerf)
        
        if compatible_offcuts:
            offcut, board, rotated = compatible_offcuts[0]  # Best option
            
            success = board.place_part(part, offcut, rotated, offcut.x, offcut.y, core_db)
            if success:
                unplaced_parts.remove(part)
                parts_placed += 1
                
                if part.material_details.core_name != offcut.material_details.core_name:
                    part.is_upgraded = True
                
                log_entry = f"✓ Placed unplaced part {part.id} on {board.id} (upgrade: {part.material_details.core_name} → {offcut.material_details.core_name})"
                iteration_log.append(log_entry)
                logger.info(log_entry)
    
    # SYSTEMATIC UPGRADE ACROSS USER'S UPGRADE SEQUENCE
    upgrade_sequence = user_upgrade_sequence
    
    # Create grade mapping for the sequence
    grade_map = {}
    for i, core in enumerate(upgrade_sequence):
        grade_map[core] = i
    
    # Group parts and offcuts by material type
    parts_by_core = {}
    offcuts_by_core = {}
    
    for board in boards:
        core_name = board.material_details.core_name
        
        # Collect parts
        if core_name not in parts_by_core:
            parts_by_core[core_name] = []
        for part in board.parts_on_board:
            parts_by_core[core_name].append((part, board))
        
        # Collect offcuts
        if core_name not in offcuts_by_core:
            offcuts_by_core[core_name] = []
        for offcut in board.available_rectangles:
            if offcut.get_area() > 50000:  # Consider moderate-sized spaces
                offcuts_by_core[core_name].append((offcut, board))
    
    logger.info(f"Parts by core: {[(core, len(parts)) for core, parts in parts_by_core.items()]}")
    logger.info(f"Offcuts by core: {[(core, len(offcuts)) for core, offcuts in offcuts_by_core.items()]}")
    
    # Try upgrades following the sequence order
    for source_core in upgrade_sequence:
        source_grade = grade_map.get(source_core, -1)
        
        if source_core in parts_by_core:
            source_parts = parts_by_core[source_core]
            
            # Look for higher-grade targets
            for target_core in upgrade_sequence:
                target_grade = grade_map.get(target_core, -1)
                
                # Only upgrade to higher grades
                if target_grade <= source_grade:
                    continue
                
                if target_core in offcuts_by_core:
                    target_offcuts = offcuts_by_core[target_core]
                    
                    logger.info(f"Trying to upgrade {len(source_parts)} parts from {source_core} to {target_core}")
                    
                    # Try to move parts from source to target
                    for part, source_board in source_parts:
                        for offcut, target_board in target_offcuts:
                            # Skip if same board
                            if target_board.id == source_board.id:
                                continue
                            
                            # Check material compatibility (laminate and thickness must match)
                            if (part.material_details.top_laminate_name == target_board.material_details.top_laminate_name and
                                part.material_details.thickness == target_board.material_details.thickness):
                                
                                # Check if part fits (normal orientation)
                                if offcut.can_fit_part(part, kerf, rotated=False):
                                    success = relocate_part_globally(
                                        part, offcut, target_board, source_board, 
                                        False, core_db, global_inventory, kerf
                                    )
                                    
                                    if success:
                                        relocations_made += 1
                                        log_entry = f"✓ Upgraded {part.id} from {source_board.id} ({source_core}) to {target_board.id} ({target_core})"
                                        iteration_log.append(log_entry)
                                        add_log_message(log_entry)
                                        break  # Move to next part
                                
                                # Try rotated if normal didn't work
                                elif part.can_rotate() and offcut.can_fit_part(part, kerf, rotated=True):
                                    success = relocate_part_globally(
                                        part, offcut, target_board, source_board, 
                                        True, core_db, global_inventory, kerf
                                    )
                                    
                                    if success:
                                        relocations_made += 1
                                        log_entry = f"✓ Upgraded {part.id} from {source_board.id} ({source_core}) to {target_board.id} ({target_core}) - rotated"
                                        iteration_log.append(log_entry)
                                        add_log_message(log_entry)
                                        break  # Move to next part
    
    # Then try general relocations for space efficiency
    all_parts = []
    for board in boards:
        for part in board.parts_on_board:
            all_parts.append((part, board))
    
    # Sort parts by potential for improvement (larger parts first)
    all_parts.sort(key=lambda x: x[0].requested_length * x[0].requested_width, reverse=True)
    
    for part, source_board in all_parts[:30]:  # Reduced limit after MDF priority
        # Skip if this part was already processed in MDF upgrade
        if "18MDF" in source_board.material_details.core_name:
            continue
            
        # Find better placement options
        compatible_offcuts = global_inventory.find_compatible_offcuts(part, core_db, kerf)
        
        for offcut, target_board, rotated in compatible_offcuts[:3]:  # Check top 3 options
            # Skip if it's the same board (no improvement)
            if target_board.id == source_board.id:
                continue
            
            # Calculate improvement potential
            current_grade = get_grade_level(source_board.material_details.core_name, core_db)
            target_grade = get_grade_level(offcut.material_details.core_name, core_db)
            
            # Only relocate if there's an upgrade or significant space efficiency gain
            if target_grade > current_grade:
                success = relocate_part_globally(
                    part, offcut, target_board, source_board, 
                    rotated, core_db, global_inventory, kerf
                )
                
                if success:
                    relocations_made += 1
                    log_entry = f"✓ Relocated {part.id} from {source_board.id} to {target_board.id}"
                    log_entry += f" (upgrade: {source_board.material_details.core_name} → {offcut.material_details.core_name})"
                    iteration_log.append(log_entry)
                    logger.info(log_entry)
                    break  # Found a good relocation, move to next part
    
    # Remove empty boards
    remaining_boards = []
    boards_removed = 0
    
    for board in boards:
        if len(board.parts_on_board) == 0:
            boards_removed += 1
            log_entry = f"✓ Eliminated empty board: {board.id}"
            iteration_log.append(log_entry)
            logger.info(log_entry)
        else:
            remaining_boards.append(board)
    
    final_waste = global_inventory.get_total_waste_area()
    final_cost = calculate_total_order_cost(remaining_boards, core_db, laminate_db)
    
    iteration_log.append(f"Iteration complete: {parts_placed} unplaced parts placed, {relocations_made} relocations, {boards_removed} boards eliminated")
    iteration_log.append(f"Waste reduction: {(initial_waste - final_waste)/1000000:.2f}m², Cost reduction: ₹{initial_cost - final_cost:.2f}")
    
    return remaining_boards, unplaced_parts, relocations_made + parts_placed + boards_removed, iteration_log

def run_global_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict, 
                          user_upgrade_sequence_str: str, kerf: float = 4.4) -> Tuple[
    List[Board], List[Part], Dict, float, float, List[str]]:
    """
    Run global optimization with iterative offcut reuse until no further improvements.
    
    Returns:
        (final_boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, optimization_log)
    """
    global optimization_log_messages
    optimization_log_messages = []  # Reset log for new optimization
    
    add_log_message("=== Global Optimization with Iterative Offcut Reuse ===")
    add_log_message(f"Total parts to optimize: {len(parts_list)}")
    add_log_message(f"Upgrade sequence: {user_upgrade_sequence_str}")
    
    # Phase 1: Initial optimization
    add_log_message("Phase 1: Initial optimization")
    boards, unplaced_parts, upgrade_summary, initial_cost, _ = run_optimization(
        parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf
    )
    
    add_log_message(f"Phase 1 complete: {len(boards)} boards, {len(unplaced_parts)} unplaced parts")
    add_log_message(f"Initial cost: ₹{initial_cost:.2f}")
    
    # Phase 2: Iterative optimization until convergence
    logger.info("Phase 2: Iterative global optimization")
    iteration = 0
    max_iterations = 10
    
    while iteration < max_iterations:
        iteration += 1
        logger.info(f"Starting iteration {iteration}")
        
        prev_board_count = len(boards)
        prev_unplaced_count = len(unplaced_parts)
        
        # Parse upgrade sequence from string
        upgrade_sequence_list = [core.strip() for core in user_upgrade_sequence_str.split(',') if core.strip()]
        
        boards, unplaced_parts, improvements, iteration_log = run_global_optimization_iteration(
            boards, unplaced_parts, core_db, laminate_db, kerf, upgrade_sequence_list
        )
        
        # Add iteration results to main log
        add_log_message(f"--- Iteration {iteration} ---")
        for log_entry in iteration_log:
            add_log_message(log_entry)
        
        # Check for convergence
        if improvements == 0:
            add_log_message(f"Convergence reached after {iteration} iterations - no further improvements possible")
            break
        
        if len(boards) == prev_board_count and len(unplaced_parts) == prev_unplaced_count:
            add_log_message(f"No structural changes in iteration {iteration} - stopping")
            break
    
    final_cost = calculate_total_order_cost(boards, core_db, laminate_db)
    
    add_log_message("=== FINAL RESULTS ===")
    add_log_message(f"Total iterations: {iteration}")
    add_log_message(f"Final boards: {len(boards)}")
    add_log_message(f"Unplaced parts: {len(unplaced_parts)}")
    add_log_message(f"Initial cost: ₹{initial_cost:.2f}")
    add_log_message(f"Final cost: ₹{final_cost:.2f}")
    add_log_message(f"Total savings: ₹{initial_cost - final_cost:.2f}")
    
    return boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, optimization_log_messages