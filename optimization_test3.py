"""
TEST 3 Algorithm - Advanced Cutting Optimizer with Global Offcut Reuse, Best-Fit Placement, and 2-Pass Packing
Implementation of the provided advanced optimizer with global offcut sharing and low-utilization board repacking
"""

import logging
from typing import List, Dict, Tuple, Optional
from copy import deepcopy
from data_models import Part, Board, MaterialDetails

logger = logging.getLogger(__name__)

class Test3Part:
    """Internal part representation for the TEST 3 algorithm."""
    def __init__(self, part_id: str, length: float, width: float, grain_sensitive: bool = True, original_part: Optional[Part] = None):
        self.id = part_id
        self.length = length
        self.width = width
        self.area = length * width
        self.placed = False
        self.grain_sensitive = grain_sensitive
        self.original_part = original_part

    def __repr__(self):
        return f"Test3Part({self.id}, {self.length}x{self.width})"


class Test3Board:
    """Internal board representation with global offcut management."""
    def __init__(self, board_id: str, length: float, width: float):
        self.id = board_id
        self.length = length
        self.width = width
        self.parts = []  # (Test3Part, x, y, rotated)
        self.offcuts = [(0, 0, length, width)]

    def utilization(self):
        if not self.parts:
            return 0.0
        used_area = sum(p.area for p, _, _, _ in self.parts)
        return (used_area / (self.length * self.width)) * 100

    def place_part(self, part: Test3Part, kerf: float = 4.4) -> bool:
        """Place part using best-fit strategy with global offcut management."""
        best_index = None
        best_score = float('inf')
        best_rotated = False
        
        for i, (x, y, l, w) in enumerate(self.offcuts):
            # Try normal orientation
            if part.length + kerf <= l and part.width + kerf <= w:
                score = (l * w) - part.area
                if score < best_score:
                    best_index, best_score, best_rotated = i, score, False
            
            # Try rotated orientation (if part can rotate)
            elif not part.grain_sensitive and part.width + kerf <= l and part.length + kerf <= w:
                score = (l * w) - part.area
                if score < best_score:
                    best_index, best_score, best_rotated = i, score, True

        if best_index is not None:
            x, y, l, w = self.offcuts.pop(best_index)
            pl, pw = (part.width, part.length) if best_rotated else (part.length, part.width)
            pl += kerf
            pw += kerf
            self.parts.append((part, x, y, best_rotated))
            self._split_offcut((x, y, l, w), pl, pw)
            part.placed = True
            return True
        
        return False

    def try_edge_fit(self, part: Test3Part, kerf: float = 4.4) -> bool:
        """Try to place part using tight edge-fit strategy for narrow strips."""
        step_size = 10  # Define step size at function level
        
        # Look for tight strip placement at right edge (for narrow parts like strips)
        margin_x = self.length - part.length - kerf
        
        if margin_x >= 0:
            # Try vertical positions in steps to find free area
            for y_pos in range(0, int(self.width - part.width - kerf + 1), step_size):
                if self._is_area_free(margin_x, y_pos, part.length + kerf, part.width + kerf):
                    self.parts.append((part, margin_x, y_pos, False))
                    part.placed = True
                    
                    # CRITICAL FIX: Update offcuts to reflect the placed part
                    # Remove any offcuts that overlap with the placed part area
                    placed_area = (margin_x, y_pos, part.length + kerf, part.width + kerf)
                    self._remove_overlapping_offcuts(placed_area)
                    
                    logger.debug(f"Edge-fit placed {part.id} at right edge ({margin_x}, {y_pos})")
                    return True
        
        # Try rotated placement if part can rotate
        if not part.grain_sensitive:
            margin_x_rot = self.length - part.width - kerf
            if margin_x_rot >= 0:
                for y_pos in range(0, int(self.width - part.length - kerf + 1), step_size):
                    if self._is_area_free(margin_x_rot, y_pos, part.width + kerf, part.length + kerf):
                        self.parts.append((part, margin_x_rot, y_pos, True))
                        part.placed = True
                        
                        # CRITICAL FIX: Update offcuts for rotated placement
                        placed_area = (margin_x_rot, y_pos, part.width + kerf, part.length + kerf)
                        self._remove_overlapping_offcuts(placed_area)
                        
                        logger.debug(f"Edge-fit placed {part.id} rotated at right edge ({margin_x_rot}, {y_pos})")
                        return True
        
        return False

    def _remove_overlapping_offcuts(self, placed_area):
        """Remove offcuts that overlap with the placed part area."""
        px, py, pl, pw = placed_area
        
        # Filter out offcuts that overlap with the placed area
        new_offcuts = []
        for ox, oy, ol, ow in self.offcuts:
            # Check if offcut overlaps with placed area
            x_overlap = px < ox + ol and ox < px + pl
            y_overlap = py < oy + ow and oy < py + pw
            
            if not (x_overlap and y_overlap):
                # No overlap, keep this offcut
                new_offcuts.append((ox, oy, ol, ow))
            # If there's overlap, we could split the offcut, but for now just remove it
            # This is safer and prevents further overlapping issues
        
        self.offcuts = new_offcuts

    def _is_area_free(self, x: float, y: float, l: float, w: float) -> bool:
        """Check if the specified area is free from existing parts."""
        for part, px, py, rotated in self.parts:
            # Get actual part dimensions considering rotation
            if rotated:
                part_length, part_width = part.width, part.length
            else:
                part_length, part_width = part.length, part.width
            
            # Check for overlap using rectangle intersection logic
            # Two rectangles overlap if they intersect in both X and Y dimensions
            x_overlap = not (x + l <= px or px + part_length <= x)
            y_overlap = not (y + w <= py or py + part_width <= y)
            
            if x_overlap and y_overlap:
                return False  # Area is occupied
        return True  # Area is free

    def _split_offcut(self, rect: Tuple[float, float, float, float], pl: float, pw: float):
        """Split offcut rectangle into reusable pieces using proper guillotine cuts."""
        x, y, l, w = rect
        
        # Proper guillotine cutting - no overlapping offcuts
        # Right offcut: full height of original rectangle
        if l - pl > 10:
            self.offcuts.append((x + pl, y, l - pl, w))
        # Bottom offcut: only the used width to avoid overlap
        if w - pw > 10:
            self.offcuts.append((x, y + pw, pl, w - pw))

    def __repr__(self):
        return f"Test3Board({self.id}, {self.length}x{self.width})"


class GlobalOptimizer:
    """Advanced optimizer with global offcut reuse and 2-pass packing."""
    def __init__(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4):
        self.parts_list = parts_list
        self.core_db = core_db
        self.laminate_db = laminate_db
        self.kerf = kerf
        self.used_boards = []
        self.unplaced_parts = []

    def optimize(self) -> Tuple[List[Board], List[Part], Dict, float, float]:
        """Run advanced global optimization with 2-pass packing."""
        logger.info(f"Starting TEST 3 Advanced Global Optimizer with {len(self.parts_list)} parts")
        
        # Group parts by material
        material_groups = self._group_parts_by_material()
        
        # Process each material group with advanced global optimization
        for material_type, parts_group in material_groups.items():
            self._optimize_material_group_global(material_type, parts_group)
        
        # Calculate costs (simplified - no upgrades)
        initial_cost = final_cost = sum(board.material_details.get_cost_per_sqm(
            self.core_db, self.laminate_db
        ) * (board.total_length * board.total_width / 1_000_000) for board in self.used_boards)
        
        logger.info(f"TEST 3 Global Optimizer complete: {len(self.used_boards)} boards, {len(self.unplaced_parts)} unplaced")
        
        return self.used_boards, self.unplaced_parts, {}, initial_cost, final_cost

    def _group_parts_by_material(self) -> Dict[str, List[Part]]:
        """Group parts by their material specifications."""
        material_groups = {}
        for part in self.parts_list:
            material_key = part.material_details.full_material_string
            if material_key not in material_groups:
                material_groups[material_key] = []
            material_groups[material_key].append(part)
        return material_groups

    def _get_board_dimensions(self, material_type: str) -> Tuple[float, float]:
        """Extract board dimensions from core database."""
        try:
            material_details = MaterialDetails(material_type)
            core_name = material_details.core_name
        except:
            logger.warning(f"Could not parse material type: {material_type}, using default HDHMR")
            core_name = "HDHMR"
        
        if core_name in self.core_db:
            length = float(self.core_db[core_name].get('Standard Length (mm)', 2420))
            width = float(self.core_db[core_name].get('Standard Width (mm)', 1220))
        else:
            logger.warning(f"Core material {core_name} not found in database, using default dimensions")
            length, width = 2420.0, 1220.0
            
        logger.info(f"Board dimensions for {core_name}: {length}x{width}mm")
        return length, width

    def _optimize_material_group_global(self, material_type: str, parts_group: List[Part]) -> None:
        """Run advanced global optimization for a single material group."""
        board_length, board_width = self._get_board_dimensions(material_type)
        
        logger.info(f"Running advanced global optimization for {len(parts_group)} parts of material {material_type}")
        
        # Convert to Test3Part objects and sort by area (largest first)
        test_parts = []
        for part in parts_group:
            grain_sensitive = part.grains == 1  # 1 = grain sensitive, 0 = can rotate
            test_part = Test3Part(
                part_id=part.id,
                length=part.requested_length,
                width=part.requested_width,
                grain_sensitive=grain_sensitive,
                original_part=part
            )
            test_parts.append(test_part)

        # Sort parts by area (largest first) for optimal packing
        test_parts = sorted(test_parts, key=lambda p: p.area, reverse=True)
        
        boards = []
        board_dim = (board_length, board_width)
        
        # Phase 1: Enhanced placement with Last-Fit Repacking
        logger.info("Phase 1: Enhanced placement with Last-Fit Repacking for grain-sensitive parts")
        for part in test_parts:
            if part.placed:
                continue
                
            placed = False
            # Try to place on existing boards first
            for board in boards:
                if board.place_part(part, kerf=self.kerf):
                    placed = True
                    logger.debug(f"Placed {part.id} on existing {board.id}")
                    break
            
            # If not placed, try tight edge-fit strategy for narrow strips
            if not placed:
                for board in boards:
                    if board.try_edge_fit(part, kerf=self.kerf):
                        placed = True
                        logger.debug(f"Edge-fit placed {part.id} on {board.id}")
                        break
            
            # If still not placed, attempt Last-Fit Repacking on low-utilization boards
            if not placed:
                low_util_boards = sorted([b for b in boards if b.utilization() < 65.0], 
                                       key=lambda b: b.utilization())
                
                for target_board in low_util_boards:
                    # Collect all parts from the target board plus the new part
                    test_parts_for_repack = [p for p, *_ in target_board.parts] + [part]
                    
                    # Reset placement status for repacking
                    for p in test_parts_for_repack:
                        p.placed = False
                    
                    # Create a new board to test repacking
                    repack_board = Test3Board(f"{target_board.id}-repacked", 
                                            target_board.length, target_board.width)
                    
                    # Try to fit all parts (sorted by area for better packing)
                    fit_success = True
                    sorted_repack_parts = sorted(test_parts_for_repack, 
                                               key=lambda x: x.area, reverse=True)
                    
                    for p in sorted_repack_parts:
                        if not repack_board.place_part(p, kerf=self.kerf):
                            fit_success = False
                            break
                    
                    if fit_success:
                        # Successful repack - replace the old board
                        logger.info(f"Last-Fit Repacking successful: {target_board.id} -> {repack_board.id}")
                        boards.remove(target_board)
                        boards.append(repack_board)
                        placed = True
                        break
                    else:
                        # Reset placement status if repack failed
                        for p in test_parts_for_repack:
                            if p != part:  # Don't reset the new part
                                p.placed = True
            
            # Create new board if all placement attempts failed
            if not placed:
                new_board = Test3Board(f"Global-Board-{len(boards)+1}", *board_dim)
                if new_board.place_part(part, kerf=self.kerf):
                    boards.append(new_board)
                    logger.debug(f"Created new board {new_board.id} for {part.id}")
                else:
                    logger.warning(f"Could not place {part.id} even on new board")

        # Phase 2: Final consolidation pass for remaining low-utilization boards
        threshold = 65.0
        final_to_repack = [b for b in boards if b.utilization() < threshold]
        
        if final_to_repack:
            logger.info(f"Phase 2: Final consolidation of {len(final_to_repack)} remaining low-utilization boards")
            
            # Collect parts from remaining low-utilization boards
            spare_parts = []
            for board in final_to_repack:
                spare_parts.extend([p for p, *_ in board.parts])
            
            # Reset placement status
            for part in spare_parts:
                part.placed = False
            
            # Keep only high-utilization boards
            boards = [b for b in boards if b.utilization() >= threshold]
            
            # Re-place the spare parts with enhanced placement strategy
            for part in sorted(spare_parts, key=lambda p: p.area, reverse=True):
                placed = False
                
                # Try existing high-utilization boards first
                for board in boards:
                    if board.place_part(part, kerf=self.kerf):
                        placed = True
                        break
                
                # Create new board if needed
                if not placed:
                    new_board = Test3Board(f"Global-Board-{len(boards)+1}", *board_dim)
                    if new_board.place_part(part, kerf=self.kerf):
                        boards.append(new_board)
                    else:
                        logger.warning(f"Failed to repack {part.id}")
        
        # Log final optimization results
        total_utilization = sum(b.utilization() for b in boards) / len(boards) if boards else 0
        logger.info(f"Final optimization complete: {len(boards)} boards, average utilization: {total_utilization:.1f}%")

        # Convert to OptiWise format
        self._convert_test3_boards_to_optiwise(boards, material_type, "GlobalOpt")

    def _convert_test3_boards_to_optiwise(self, test_boards: List[Test3Board], 
                                        material_type: str, strategy_name: str) -> None:
        """Convert Test3Board objects back to OptiWise Board objects."""
        
        for i, test_board in enumerate(test_boards, 1):
            try:
                material_details = MaterialDetails(material_type)
            except:
                material_details = MaterialDetails("Unknown_Unknown_Unknown")
            
            # Create OptiWise board
            board = Board(
                board_id=f"{material_type}_{strategy_name}_Board_{i}",
                material_details=material_details,
                total_length=test_board.length,
                total_width=test_board.width,
                kerf=self.kerf
            )
            
            # Add placed parts to the board
            for test_part, x, y, rotated in test_board.parts:
                original_part = test_part.original_part
                original_part.x_pos = x
                original_part.y_pos = y
                original_part.rotated = rotated
                
                # Set actual dimensions to what was placed (without kerf for display)
                if rotated:
                    original_part.actual_length = test_part.width
                    original_part.actual_width = test_part.length
                else:
                    original_part.actual_length = test_part.length
                    original_part.actual_width = test_part.width
                
                board.parts_on_board.append(original_part)
            
            self.used_boards.append(board)
            
            utilization_percent = test_board.utilization()
            logger.info(f"Board {board.id}: {len(board.parts_on_board)} parts, {utilization_percent:.1f}% utilization")


def run_test3_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4) -> Tuple[List[Board], List[Part], Dict, float, float]:
    """Run the advanced global TEST 3 optimization algorithm."""
    optimizer = GlobalOptimizer(parts_list, core_db, laminate_db, kerf)
    return optimizer.optimize()