"""
Enhanced TEST Algorithm - Dynamic board repacking optimization
Implementation of advanced Bottom-Left-Fill with on-the-fly board consolidation
"""

import logging
from typing import List, Dict, Tuple, Optional
from copy import deepcopy
from data_models import Part, Board, MaterialDetails

logger = logging.getLogger(__name__)

class TestPart:
    """Internal part representation for the TEST algorithm."""
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
        return f"TestPart({self.id}, {self.length}x{self.width})"


class TestBoard:
    """Internal board representation with Best-Fit placement and offcut management."""
    def __init__(self, board_id: str, length: float, width: float):
        self.id = board_id
        self.length = length
        self.width = width
        self.parts = []  # (TestPart, x, y, rotated)
        self.offcuts = [(0, 0, length, width)]  # (x, y, length, width)
    
    def utilization(self):
        if not self.parts:
            return 0.0
        used_area = sum(p.area for p, _, _, _ in self.parts)
        return used_area / (self.length * self.width)

    def place_part_best_fit(self, part: TestPart, kerf: float = 4.4) -> bool:
        """Place part using Best-Fit Decreasing strategy to minimize fragmentation."""
        best_index = None
        best_fit_score = float('inf')
        best_rotated = False

        for i, (x, y, l, w) in enumerate(self.offcuts):
            # Try normal orientation (part dimensions + kerf must fit in offcut)
            part_length_with_kerf = part.length + kerf
            part_width_with_kerf = part.width + kerf
            
            if part_length_with_kerf <= l and part_width_with_kerf <= w:
                score = (l * w) - (part_length_with_kerf * part_width_with_kerf)
                if score < best_fit_score:
                    best_index = i
                    best_fit_score = score
                    best_rotated = False
            
            # Try rotated orientation (if part can rotate)
            elif not part.grain_sensitive and part_width_with_kerf <= l and part_length_with_kerf <= w:
                score = (l * w) - (part_length_with_kerf * part_width_with_kerf)
                if score < best_fit_score:
                    best_index = i
                    best_fit_score = score
                    best_rotated = True

        if best_index is not None:
            x, y, l, w = self.offcuts.pop(best_index)
            if best_rotated:
                pl, pw = part.width + kerf, part.length + kerf
                self.parts.append((part, x, y, True))
                part.rotation = True
            else:
                pl, pw = part.length + kerf, part.width + kerf
                self.parts.append((part, x, y, False))
                part.rotation = False
            
            self._generate_two_offcuts(x, y, pl, pw, l, w)
            part.placed = True
            return True

        return False

    def _generate_two_offcuts(self, x: float, y: float, pl: float, pw: float, l: float, w: float):
        """Generate exactly two offcuts to reduce fragmentation."""
        # Horizontal then vertical cut – fewer, larger offcuts
        if l > pl:
            self.offcuts.append((x + pl, y, l - pl, pw))
        if w > pw:
            self.offcuts.append((x, y + pw, l, w - pw))

    def rearrange_with(self, part: TestPart, kerf: float = 4.4) -> Optional['TestBoard']:
        """Attempt to rearrange all parts on board plus new part."""
        candidates = [p for p, _, _, _ in self.parts] + [part]
        candidates.sort(key=lambda p: p.area, reverse=True)
        
        new_board = TestBoard(self.id + '_repack', self.length, self.width)
        for p in candidates:
            p.placed = False  # Reset placement status
            if not new_board.place_part_best_fit(p, kerf):
                return None
        return new_board


class DynamicTestOptimizer:
    """Enhanced TEST algorithm with dynamic board repacking optimization."""
    
    def __init__(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4):
        self.parts_list = parts_list
        self.core_db = core_db
        self.laminate_db = laminate_db
        self.kerf = kerf
        self.used_boards = []
        self.unplaced_parts = []
        
    def optimize(self) -> Tuple[List[Board], List[Part], Dict, float, float]:
        """Run the dynamic repacking TEST optimization algorithm."""
        logger.info(f"Starting Dynamic TEST Algorithm with {len(self.parts_list)} parts")
        
        # Group parts by material
        material_groups = self._group_parts_by_material()
        
        # Process each material group with dynamic repacking
        for material_type, parts_group in material_groups.items():
            self._optimize_material_group_dynamic(material_type, parts_group)
        
        # Calculate costs (simplified - no upgrades)
        initial_cost = final_cost = sum(board.material_details.get_cost_per_sqm(
            self.core_db, self.laminate_db
        ) * (board.total_length * board.total_width / 1_000_000) for board in self.used_boards)
        
        logger.info(f"Dynamic TEST Algorithm complete: {len(self.used_boards)} boards, {len(self.unplaced_parts)} unplaced")
        
        return self.used_boards, self.unplaced_parts, {}, initial_cost, final_cost
    
    def _group_parts_by_material(self) -> Dict[str, List[Part]]:
        """Group parts by material type."""
        groups = {}
        
        for part in self.parts_list:
            material_key = str(part.material_details)
            if material_key not in groups:
                groups[material_key] = []
            groups[material_key].append(part)
            
        return groups
    
    def _get_board_dimensions(self, material_type: str) -> Tuple[float, float]:
        """Get board dimensions for material type."""
        # Extract core material from material type string
        parts = material_type.split('_')
        core_name = parts[1] if len(parts) >= 2 else "18HDHMR"
        
        if core_name in self.core_db:
            # Try multiple possible column names for board dimensions
            core_info = self.core_db[core_name]
            
            # Standard format from user data: 2420 x 1220 mm
            length = float(core_info.get('Standard Length (mm)', 
                          core_info.get('Length (mm)',
                          core_info.get('Board Length', 2420))))
            width = float(core_info.get('Standard Width (mm)',
                         core_info.get('Width (mm)', 
                         core_info.get('Board Width', 1220))))
        else:
            # Default to user-specified board size: 2420 x 1220 mm
            length, width = 2420.0, 1220.0
            
        logger.info(f"Board dimensions for {core_name}: {length}x{width}mm")
        return length, width
    
    def _optimize_material_group_dynamic(self, material_type: str, parts_group: List[Part]) -> None:
        """Optimize cutting for a single material group using Best-Fit Decreasing with rearrangement."""
        board_length, board_width = self._get_board_dimensions(material_type)
        
        logger.info(f"Running Best-Fit optimization for {len(parts_group)} parts of material {material_type}")
        
        # Convert to TestPart objects and sort by area (largest first)
        test_parts = []
        for part in parts_group:
            grain_sensitive = part.grains == 1  # 1 = grain sensitive, 0 = can rotate
            test_part = TestPart(
                part_id=part.id,
                length=part.requested_length,
                width=part.requested_width,
                grain_sensitive=grain_sensitive,
                original_part=part
            )
            test_parts.append(test_part)
        
        # Sort parts by area (largest first) for Best-Fit Decreasing
        test_parts.sort(key=lambda p: p.area, reverse=True)
        
        board_dims = (board_length, board_width)
        boards = []
        board_count = 0
        
        for part in test_parts:
            placed = False
            
            # Try to place part on existing boards
            for board in boards:
                if board.place_part_best_fit(part, self.kerf):
                    placed = True
                    logger.debug(f"Placed {part.id} on existing {board.id}")
                    break

            if not placed:
                # Try rearranging only the least-utilized board
                if boards:
                    target_board = min(boards, key=lambda b: b.utilization())
                    new_board = target_board.rearrange_with(part, self.kerf)
                    if new_board:
                        boards[boards.index(target_board)] = new_board
                        placed = True
                        logger.info(f"Rearranged {target_board.id} to fit {part.id}")

            if not placed:
                # Create new board
                board_count += 1
                new_board = TestBoard(f"BestFit-Board-{board_count}", board_dims[0], board_dims[1])
                if new_board.place_part_best_fit(part, self.kerf):
                    boards.append(new_board)
                    logger.info(f"Created new board {new_board.id} for {part.id}")
                else:
                    logger.warning(f"Could not place {part.id} even on new board")
        
        # Collect unplaced parts
        unplaced_parts = [p for p in test_parts if not p.placed]
        
        # Convert to OptiWise format
        self._convert_test_boards_to_optiwise(boards, unplaced_parts, material_type, "BestFit")

    def _convert_test_boards_to_optiwise(self, test_boards: List[TestBoard], unplaced_test_parts: List[TestPart], 
                                       material_type: str, strategy_name: str) -> None:
        """Convert TestBoard objects back to OptiWise Board objects."""
        
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
    



def run_test_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict, kerf: float = 4.4) -> Tuple[List[Board], List[Part], Dict, float, float]:
    """Run the enhanced dynamic repacking TEST optimization algorithm."""
    optimizer = DynamicTestOptimizer(parts_list, core_db, laminate_db, kerf)
    return optimizer.optimize()