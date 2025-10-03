"""
TEST 2 Algorithm - Tight Nesting Guillotine Optimizer
Implementation of best-fit layout with minimal waste using repacking of low-utilization boards
"""

import logging
from typing import List, Dict, Tuple, Optional
from copy import deepcopy
from data_models import Part, Board, MaterialDetails

logger = logging.getLogger(__name__)

class Test2Part:
    """Internal part representation for the TEST 2 algorithm."""
    def __init__(self, part_id: str, length: float, width: float, grain_sensitive: bool = True, original_part: Optional[Part] = None):
        self.id = part_id
        self.length = length
        self.width = width
        self.area = length * width
        self.grain_sensitive = grain_sensitive
        self.original_part = original_part
        self.placed = False
        self.rotation = False

    def __repr__(self):
        return f"Test2Part({self.id}, {self.length}x{self.width})"


class Test2Board:
    """Internal board representation with tight nesting and repacking capabilities."""
    def __init__(self, board_id: str, length: float, width: float):
        self.id = board_id
        self.length = length
        self.width = width
        self.parts = []  # (Test2Part, x, y, rotated)
        self.offcuts = [(0, 0, length, width)]  # (x, y, length, width)
    
    def utilization(self):
        if not self.parts:
            return 0.0
        used_area = sum(p.area for p, _, _, _ in self.parts)
        return used_area / (self.length * self.width)

    def best_fit_position(self, part: Test2Part, kerf: float = 4.4) -> Tuple[Optional[int], Optional[bool]]:
        """Find best-fit position for a part considering kerf spacing."""
        best_index = None
        best_score = float('inf')
        best_rotated = False

        for i, (x, y, l, w) in enumerate(self.offcuts):
            # Try normal orientation
            part_length_with_kerf = part.length + kerf
            part_width_with_kerf = part.width + kerf
            
            if part_length_with_kerf <= l and part_width_with_kerf <= w:
                score = (l * w) - (part_length_with_kerf * part_width_with_kerf)
                if score < best_score:
                    best_index = i
                    best_score = score
                    best_rotated = False
            
            # Try rotated orientation (if part can rotate)
            elif not part.grain_sensitive and part_width_with_kerf <= l and part_length_with_kerf <= w:
                score = (l * w) - (part_length_with_kerf * part_width_with_kerf)
                if score < best_score:
                    best_index = i
                    best_score = score
                    best_rotated = True

        return best_index, best_rotated

    def place_part(self, part: Test2Part, kerf: float = 4.4) -> bool:
        """Place part using best-fit strategy."""
        idx, rot = self.best_fit_position(part, kerf)
        if idx is not None:
            x, y, l, w = self.offcuts.pop(idx)
            if rot:
                pl, pw = part.width + kerf, part.length + kerf
                self.parts.append((part, x, y, True))
                part.rotation = True
            else:
                pl, pw = part.length + kerf, part.width + kerf
                self.parts.append((part, x, y, False))
                part.rotation = False
            
            self._generate_offcuts(x, y, pl, pw, l, w)
            part.placed = True
            return True
        return False

    def _generate_offcuts(self, x: float, y: float, pl: float, pw: float, l: float, w: float):
        """Generate guillotine-compatible 2-offcut logic."""
        if l > pl:
            self.offcuts.append((x + pl, y, l - pl, pw))
        if w > pw:
            self.offcuts.append((x, y + pw, l, w - pw))

    def repack_with(self, parts: List[Test2Part], kerf: float = 4.4) -> Optional['Test2Board']:
        """Attempt to repack all parts plus new parts on board."""
        new_board = Test2Board(self.id + '_repack', self.length, self.width)
        all_parts = sorted(parts, key=lambda p: p.area, reverse=True)
        
        for part in all_parts:
            part.placed = False  # Reset placement status
            if not new_board.place_part(part, kerf):
                return None
        return new_board


class TightNestingOptimizer:
    """TEST 2 algorithm with tight nesting and low-utilization board repacking."""
    
    def __init__(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4):
        self.parts_list = parts_list
        self.core_db = core_db
        self.laminate_db = laminate_db
        self.kerf = kerf
        self.used_boards = []
        self.unplaced_parts = []
        
    def optimize(self) -> Tuple[List[Board], List[Part], Dict, float, float]:
        """Run the tight nesting TEST 2 optimization algorithm."""
        logger.info(f"Starting TEST 2 Tight Nesting Algorithm with {len(self.parts_list)} parts")
        
        # Group parts by material
        material_groups = self._group_parts_by_material()
        
        # Process each material group with tight nesting
        for material_type, parts_group in material_groups.items():
            self._optimize_material_group_tight_nesting(material_type, parts_group)
        
        # Calculate costs (simplified - no upgrades)
        initial_cost = final_cost = sum(board.material_details.get_cost_per_sqm(
            self.core_db, self.laminate_db
        ) * (board.total_length * board.total_width / 1_000_000) for board in self.used_boards)
        
        logger.info(f"TEST 2 Tight Nesting Algorithm complete: {len(self.used_boards)} boards, {len(self.unplaced_parts)} unplaced")
        
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
    
    def _optimize_material_group_tight_nesting(self, material_type: str, parts_group: List[Part]) -> None:
        """Optimize cutting for a single material group using tight nesting with repacking."""
        board_length, board_width = self._get_board_dimensions(material_type)
        
        logger.info(f"Running tight nesting optimization for {len(parts_group)} parts of material {material_type}")
        
        # Convert to Test2Part objects and sort by area (largest first)
        test_parts = []
        for part in parts_group:
            grain_sensitive = part.grains == 1  # 1 = grain sensitive, 0 = can rotate
            test_part = Test2Part(
                part_id=part.id,
                length=part.requested_length,
                width=part.requested_width,
                grain_sensitive=grain_sensitive,
                original_part=part
            )
            test_parts.append(test_part)
        
        # Sort parts by area (largest first) for tight nesting
        test_parts.sort(key=lambda p: p.area, reverse=True)
        
        board_dims = (board_length, board_width)
        boards = []
        board_count = 0
        
        for part in test_parts:
            if part.placed:
                continue
                
            placed = False

            # Try placing in existing boards
            for board in boards:
                if board.place_part(part, self.kerf):
                    placed = True
                    logger.debug(f"Placed {part.id} on existing {board.id}")
                    break

            # Try repacking low-utilization boards (threshold: 65%)
            if not placed:
                repack_targets = [b for b in boards if b.utilization() < 0.65]
                for target_board in repack_targets:
                    current_parts = [p for p, _, _, _ in target_board.parts]
                    if part not in current_parts:
                        current_parts.append(part)
                    
                    new_board = target_board.repack_with(current_parts, self.kerf)
                    if new_board:
                        boards[boards.index(target_board)] = new_board
                        part.placed = True
                        placed = True
                        logger.info(f"Repacked {target_board.id} to fit {part.id}")
                        break

            # Open new board if all fails
            if not placed:
                board_count += 1
                new_board = Test2Board(f"TightNest-Board-{board_count}", board_dims[0], board_dims[1])
                if new_board.place_part(part, self.kerf):
                    boards.append(new_board)
                    logger.info(f"Created new board {new_board.id} for {part.id}")
                else:
                    logger.warning(f"Could not place {part.id} even on new board")
        
        # Collect unplaced parts
        unplaced_parts = [p for p in test_parts if not p.placed]
        
        # Convert to OptiWise format
        self._convert_test2_boards_to_optiwise(boards, unplaced_parts, material_type, "TightNest")

    def _convert_test2_boards_to_optiwise(self, test_boards: List[Test2Board], unplaced_test_parts: List[Test2Part], 
                                        material_type: str, strategy_name: str) -> None:
        """Convert Test2Board objects back to OptiWise Board objects."""
        
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
                
                if rotated:
                    original_part.actual_length = test_part.width
                    original_part.actual_width = test_part.length
                else:
                    original_part.actual_length = test_part.length
                    original_part.actual_width = test_part.width
                
                board.parts_on_board.append(original_part)
            
            self.used_boards.append(board)
            
            utilization_percent = test_board.utilization() * 100
            logger.info(f"Board {board.id}: {len(board.parts_on_board)} parts, {utilization_percent:.1f}% utilization")
        
        # Add unplaced parts
        for unplaced_test_part in unplaced_test_parts:
            self.unplaced_parts.append(unplaced_test_part.original_part)


def run_test2_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4) -> Tuple[List[Board], List[Part], Dict, float, float]:
    """Run the tight nesting TEST 2 optimization algorithm."""
    optimizer = TightNestingOptimizer(parts_list, core_db, laminate_db, kerf)
    return optimizer.optimize()