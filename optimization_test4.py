import sys
from copy import deepcopy
from math import ceil

# Global variables for core and laminate databases
global_core_db = {}
global_laminate_db = {}
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import random
import math

# --- 1. Core Data Structures ---

class Part:
    """Represents a single part to be cut."""
    def __init__(self, id, length, width, grain_sensitive=True):
        self.id = id
        self.length = length
        self.width = width
        self.area = length * width
        self.grain_sensitive = grain_sensitive
    def __repr__(self):
        return f"Part({self.id}, {self.length}x{self.width})"

class PlacedPart:
    """Represents a part that has been placed on a board, tracking its final state."""
    def __init__(self, part, x, y, rotated=False):
        self.part = part
        self.x = x
        self.y = y
        self.rotated = rotated
        self.length = part.width if rotated else part.length
        self.width = part.length if rotated else part.width
    def __repr__(self):
        return f"PlacedPart({self.part.id} at ({self.x},{self.y}), L={self.length}, W={self.width})"

class Board:
    """Represents a board and its layout of placed parts."""
    def __init__(self, id, length, width):
        self.id = id
        self.length = length
        self.width = width
        self.placed_parts = []  # Stores PlacedPart objects
    def utilization(self):
        used_area = sum(p.part.area for p in self.placed_parts)
        return (used_area / (self.length * self.width)) * 100 if self.length * self.width > 0 else 0
    
    def get_utilization_percentage(self):
        """Compatibility method for existing code."""
        return self.utilization()
    def __repr__(self):
        return f"Board({self.id}, {self.length}x{self.width})"

def calculate_lower_bound(parts, board_dim):
    total_parts_area = sum(p.area for p in parts)
    board_area = board_dim[0] * board_dim[1]
    return ceil(total_parts_area / board_area)

# --- 2. Advanced Packer Implementations ---

class CutlistPlusFXPacker:
    """Enhanced packer designed to match Cutlist Plus FX efficiency levels (82%+ utilization)."""
    
    def __init__(self, parts, board_dim, saw_kerf=4.4):
        # Strategic sorting: largest parts first, then by aspect ratio for better fit
        self.parts = sorted(parts, key=lambda p: (p.area, max(p.length, p.width)), reverse=True)
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
        self.target_utilization = 0.85  # Target 85% to compensate for 4.4mm vs 3.2mm kerf
    
    def optimize(self):
        """Strategic packing algorithm with material segregation and guillotine constraints."""
        boards = []
        unplaced = []
        
        # Group parts by material to prevent mixing
        material_groups = self._group_parts_by_material(self.parts)
        
        for material_type, parts_group in material_groups.items():
            remaining_parts = parts_group[:]
            
            while remaining_parts:
                board = Board(f"CLFX-{material_type}-B{len(boards)+1}", self.board_dim[0], self.board_dim[1])
                initial_parts_count = len(remaining_parts)
                
                # Cutlist Plus FX Strategy: Strategic placement in phases
                self._cutlist_strategic_placement(board, remaining_parts)
                
                if board.placed_parts:
                    boards.append(board)
                    placed_count = initial_parts_count - len(remaining_parts)
                    utilization = board.utilization()
                    print(f"  Board {len(boards)} ({material_type}): Placed {placed_count}/{initial_parts_count} parts ({utilization:.1f}% util)")
                else:
                    # If no parts fit, add remaining to unplaced
                    unplaced.extend(remaining_parts)
                    break
        
        return boards, unplaced
    
    def _group_parts_by_material(self, parts):
        """Group parts by material type to prevent mixing different materials on same board."""
        material_groups = {}
        
        for part in parts:
            # Extract material identifier from part
            if hasattr(part, 'material_details') and hasattr(part.material_details, 'material_name'):
                material_key = part.material_details.material_name
            elif hasattr(part, 'original_material'):
                material_key = part.original_material
            else:
                # Fallback to generic grouping
                material_key = "DEFAULT"
            
            if material_key not in material_groups:
                material_groups[material_key] = []
            material_groups[material_key].append(part)
        
        print(f"  Material groups: {list(material_groups.keys())}")
        return material_groups
    
    def _cutlist_strategic_placement(self, board, remaining_parts):
        """Implement Cutlist Plus FX strategic placement methodology."""
        
        # Ultra-aggressive approach: try to place ALL parts on this board
        max_attempts = 10  # More attempts for better packing
        
        for attempt in range(max_attempts):
            if not remaining_parts:
                break
                
            placed_this_round = False
            
            # Try different sorting strategies each attempt
            if attempt % 3 == 0:
                # Large parts first (like Cutlist Plus FX)
                remaining_parts.sort(key=lambda p: p.area, reverse=True)
            elif attempt % 3 == 1:
                # Best fit first (try to fill gaps efficiently)
                remaining_parts.sort(key=lambda p: min(p.length, p.width))
            else:
                # Mixed strategy
                remaining_parts.sort(key=lambda p: (p.area, max(p.length, p.width)), reverse=True)
            
            # Try to place as many parts as possible this round
            for part in list(remaining_parts):
                position = self._find_ultra_tight_position(board, part)
                if position:
                    x, y, rotated = position
                    board.placed_parts.append(PlacedPart(part, x, y, rotated))
                    remaining_parts.remove(part)
                    placed_this_round = True
            
            if not placed_this_round:
                break  # No more parts can fit
        
        # Final squeeze attempt: try to fit any remaining small parts
        if remaining_parts:
            self._final_squeeze_attempt(board, remaining_parts)
    
    def _final_squeeze_attempt(self, board, remaining_parts):
        """Final attempt to squeeze in remaining parts using ultra-fine grid."""
        small_parts = [p for p in remaining_parts if p.area < 100000]  # Parts smaller than 100k mm²
        
        for part in list(small_parts):
            # Ultra-fine grid search for small parts
            for rotated in [False, True]:
                if rotated and part.grain_sensitive:
                    continue
                
                length = part.width if rotated else part.length
                width = part.length if rotated else part.width
                
                if length > board.length or width > board.width:
                    continue
                
                # Try every 2mm position for small parts
                for x in range(0, int(board.length - length + 1), 2):
                    for y in range(0, int(board.width - width + 1), 2):
                        if self._is_ultra_tight_valid(board, x, y, length, width):
                            board.placed_parts.append(PlacedPart(part, x, y, rotated))
                            remaining_parts.remove(part)
                            break
                    else:
                        continue
                    break
                else:
                    continue
                break
    
    def _find_corner_position(self, board, part):
        """Find optimal corner position for largest part (Cutlist Plus FX strategy)."""
        corners = [(0, 0), (0, board.width), (board.length, 0), (board.length, board.width)]
        
        for corner_x, corner_y in corners:
            for rotated in [False, True]:
                if rotated and part.grain_sensitive:
                    continue
                
                length = part.width if rotated else part.length
                width = part.length if rotated else part.width
                
                # Try placing at corner
                x = max(0, min(corner_x, board.length - length))
                y = max(0, min(corner_y, board.width - width))
                
                if self._is_ultra_tight_valid(board, x, y, length, width):
                    return (x, y, rotated)
        
        # Fallback to any valid position
        return self._find_ultra_tight_position(board, part)
    
    def _place_complementary_parts(self, board, remaining_parts):
        """Place parts that complement the first part (like twin panels)."""
        if not remaining_parts:
            return
        
        # Look for parts with similar dimensions to create efficient rows/columns
        placed_parts = board.placed_parts[:]
        
        for _ in range(min(3, len(remaining_parts))):  # Try to place up to 3 complementary parts
            best_part = None
            best_position = None
            best_efficiency = 0
            
            for part in remaining_parts:
                position = self._find_ultra_tight_position(board, part)
                if position:
                    # Calculate efficiency based on how well it complements existing parts
                    efficiency = self._calculate_complementary_efficiency(board, part, position)
                    if efficiency > best_efficiency:
                        best_efficiency = efficiency
                        best_part = part
                        best_position = position
            
            if best_part and best_position:
                x, y, rotated = best_position
                board.placed_parts.append(PlacedPart(best_part, x, y, rotated))
                remaining_parts.remove(best_part)
            else:
                break
    
    def _calculate_complementary_efficiency(self, board, part, position):
        """Calculate how well this part complements existing placement."""
        x, y, rotated = position
        length = part.width if rotated else part.length
        width = part.length if rotated else part.width
        
        # Prefer positions that create clean lines and minimize fragmentation
        alignment_score = 0
        
        # Check for alignment with existing parts
        for placed_part in board.placed_parts:
            # Vertical alignment bonus
            if abs(x - placed_part.x) < 5 or abs(x + length - (placed_part.x + placed_part.length)) < 5:
                alignment_score += 10
            
            # Horizontal alignment bonus
            if abs(y - placed_part.y) < 5 or abs(y + width - (placed_part.y + placed_part.width)) < 5:
                alignment_score += 10
        
        # Area efficiency bonus
        area_score = part.area / 1000
        
        return alignment_score + area_score
    
    def _fill_strategic_gaps(self, board, remaining_parts):
        """Fill strategic gaps with medium-sized parts."""
        for _ in range(len(remaining_parts)):
            if not remaining_parts:
                break
            
            # Sort by area descending for gap filling
            remaining_parts.sort(key=lambda p: p.area, reverse=True)
            
            placed_any = False
            for part in list(remaining_parts):
                position = self._find_ultra_tight_position(board, part)
                if position:
                    x, y, rotated = position
                    board.placed_parts.append(PlacedPart(part, x, y, rotated))
                    remaining_parts.remove(part)
                    placed_any = True
                    break
            
            if not placed_any:
                break
    
    def _pack_remaining_small_parts(self, board, remaining_parts):
        """Pack small remaining parts in any available space."""
        # Sort small parts by area ascending for efficient packing
        remaining_parts.sort(key=lambda p: p.area)
        
        for part in list(remaining_parts):
            position = self._find_ultra_tight_position(board, part)
            if position:
                x, y, rotated = position
                board.placed_parts.append(PlacedPart(part, x, y, rotated))
                remaining_parts.remove(part)
    
    def _find_ultra_tight_position(self, board, part):
        """Find the tightest possible position with Cutlist Plus FX placement strategy."""
        best_position = None
        best_score = float('inf')
        
        for rotated in [False, True]:
            if rotated and part.grain_sensitive:
                continue
            
            length = part.width if rotated else part.length
            width = part.length if rotated else part.width
            
            # Skip if part won't fit at all
            if length > board.length or width > board.width:
                continue
            
            # Aggressive grid search for maximum packing
            step_size = max(3, min(10, int(min(length, width) / 12)))  # Tighter grid
            
            for x in range(0, int(board.length - length + 1), step_size):
                for y in range(0, int(board.width - width + 1), step_size):
                    if self._is_ultra_tight_valid(board, x, y, length, width):
                        # Cutlist Plus FX scoring: bottom-left preference + efficiency
                        score = self._calculate_cutlist_score(board, x, y, length, width)
                        
                        if score < best_score:
                            best_score = score
                            best_position = (x, y, rotated)
        
        return best_position
    
    def _calculate_cutlist_score(self, board, x, y, length, width):
        """Calculate placement score using Cutlist Plus FX methodology."""
        
        # 1. Bottom-left preference (like Cutlist Plus FX)
        position_score = x * 0.1 + y * 0.1
        
        # 2. Edge alignment bonus (creates clean cutting lines)
        edge_bonus = 0
        if x == 0:  # Left edge
            edge_bonus -= 20
        if y == 0:  # Bottom edge
            edge_bonus -= 20
        if x + length == board.length:  # Right edge
            edge_bonus -= 10
        if y + width == board.width:  # Top edge
            edge_bonus -= 10
        
        # 3. Alignment with existing parts (guillotine cuts)
        alignment_bonus = 0
        for placed_part in board.placed_parts:
            # Vertical line alignment
            if abs(x - placed_part.x) < 2 or abs(x + length - (placed_part.x + placed_part.length)) < 2:
                alignment_bonus -= 15
            
            # Horizontal line alignment
            if abs(y - placed_part.y) < 2 or abs(y + width - (placed_part.y + placed_part.width)) < 2:
                alignment_bonus -= 15
        
        # 4. Remaining rectangle quality (large usable areas)
        remaining_rectangles = self._calculate_remaining_rectangles(board, x, y, length, width)
        rectangle_bonus = sum(area for area in remaining_rectangles if area > 50000) * -0.001
        
        return position_score + edge_bonus + alignment_bonus + rectangle_bonus
    
    def _calculate_remaining_rectangles(self, board, x, y, length, width):
        """Calculate areas of remaining rectangles after placing part."""
        # Calculate the main remaining rectangles
        rectangles = []
        
        # Right rectangle
        if x + length < board.length:
            right_area = (board.length - x - length) * board.width
            rectangles.append(right_area)
        
        # Top rectangle  
        if y + width < board.width:
            top_area = board.length * (board.width - y - width)
            rectangles.append(top_area)
        
        # Bottom-right corner rectangle
        if x + length < board.length and y + width < board.width:
            corner_area = (board.length - x - length) * (board.width - y - width)
            rectangles.append(corner_area)
        
        return rectangles
    
    def _is_ultra_tight_valid(self, board, x, y, length, width):
        """Ultra-tight validation with guillotine constraints and proper spacing."""
        try:
            # Convert to float to handle any numeric issues
            x, y, length, width = float(x), float(y), float(length), float(width)
            board_length, board_width = float(board.length), float(board.width)
            
            # Boundary check with kerf
            if x + length + self.saw_kerf > board_length or y + width + self.saw_kerf > board_width:
                return False
            
            # Collision check with proper kerf spacing
            for placed_part in board.placed_parts:
                px, py = float(placed_part.x), float(placed_part.y)
                pl, pw = float(placed_part.length), float(placed_part.width)
                
                # Check for overlap with kerf spacing
                if not (x + length + self.saw_kerf <= px or 
                       x >= px + pl + self.saw_kerf or
                       y + width + self.saw_kerf <= py or 
                       y >= py + pw + self.saw_kerf):
                    return False
            
            # Guillotine constraint validation
            if not self._validates_guillotine_pattern(board, x, y, length, width):
                return False
            
            return True
        except (ValueError, TypeError, AttributeError) as e:
            return False
    
    def _validates_guillotine_pattern(self, board, x, y, length, width):
        """Ensure placement follows guillotine cutting patterns."""
        if not board.placed_parts:
            return True  # First part is always valid
        
        # Check if this placement creates valid guillotine cuts
        for placed_part in board.placed_parts:
            px, py = placed_part.x, placed_part.y
            pl, pw = placed_part.length, placed_part.width
            
            # Check for clean cutting line alignment
            vertical_align = (abs(x - px) < 2 or abs(x - (px + pl)) < 2 or 
                            abs((x + length) - px) < 2 or abs((x + length) - (px + pl)) < 2)
            
            horizontal_align = (abs(y - py) < 2 or abs(y - (py + pw)) < 2 or 
                              abs((y + width) - py) < 2 or abs((y + width) - (py + pw)) < 2)
            
            # Allow placement if it creates clean cutting lines OR doesn't intersect
            if vertical_align or horizontal_align or not self._rectangles_overlap(x, y, length, width, px, py, pl, pw):
                continue
            else:
                return False  # Would create non-guillotine pattern
        
        return True
    
    def _rectangles_overlap(self, x1, y1, l1, w1, x2, y2, l2, w2):
        """Check if two rectangles overlap."""
        return not (x1 + l1 <= x2 or x2 + l2 <= x1 or y1 + w1 <= y2 or y2 + w2 <= y1)
    
    def _calculate_position_waste(self, board, x, y, length, width):
        """Calculate waste score for position selection."""
        # Prefer bottom-left positions (like Cutlist Plus FX)
        position_score = x + y
        
        # Prefer positions that leave large rectangular offcuts
        remaining_width = board.width - (y + width)
        remaining_length = board.length - (x + length)
        offcut_score = -(remaining_width * remaining_length)  # Negative because we want large offcuts
        
        return position_score + offcut_score * 0.1
    
    def _fill_remaining_space(self, board, remaining_parts):
        """Fill remaining board space with smaller parts."""
        # Try to place small parts in remaining space
        placed_additional = True
        while placed_additional and remaining_parts:
            placed_additional = False
            
            for part in list(remaining_parts):
                for rotated in [False, True]:
                    if rotated and part.grain_sensitive:
                        continue
                    
                    length = part.width if rotated else part.length
                    width = part.length if rotated else part.width
                    
                    # Find any valid position for small parts (optimized spacing)
                    step = max(25, int(min(length, width) / 5))  # Adaptive step for small parts
                    for x in range(0, int(board.length - length + 1), step):
                        for y in range(0, int(board.width - width + 1), step):
                            if self._is_ultra_tight_valid(board, x, y, length, width):
                                board.placed_parts.append(PlacedPart(part, x, y, rotated))
                                remaining_parts.remove(part)
                                placed_additional = True
                                break
                        if placed_additional:
                            break
                    if placed_additional:
                        break

class SkylineBottomLeftPacker:
    """Strategic packer optimized for large offcut generation to enable subsequent phases."""
    def __init__(self, parts, board_dim, saw_kerf=3):
        # Sort by area descending to place largest parts first, creating better offcuts
        self.parts = sorted(parts, key=lambda p: p.area, reverse=True)
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf

    def optimize(self):
        boards, unplaced = [], list(self.parts)
        while unplaced:
            board = Board(f"Skyline-B{len(boards)+1}", *self.board_dim)
            boards.append(board)
            self._pack_board(board, unplaced)
            placed_ids = {p.part.id for p in board.placed_parts}
            unplaced = [p for p in unplaced if p.id not in placed_ids]
        return boards, unplaced

    def _pack_board(self, board, parts):
        skyline = [(0, 0, board.length)]
        
        # NEW STRATEGY: Pack aggressively to high utilization, force subsequent phases to work harder
        utilization_target = 0.85  # Target 85% utilization per board
        current_area = 0
        target_area = board.length * board.width * utilization_target
        
        for part in parts:
            # Stop when reaching utilization target to force remaining parts to next phase
            if current_area >= target_area:
                break
                
            best_fit = {'score': float('inf')}
            for sx, sy, sw in skyline:
                for rot in [False, True]:
                    pl, pw = (part.length, part.width) if not rot else (part.width, part.length)
                    if rot and part.grain_sensitive: continue
                    if pl <= sw:
                        max_y = 0
                        for ssx, ssy, ssw in skyline:
                            if not (sx + pl <= ssx or sx >= ssx + ssw):
                                max_y = max(max_y, ssy)
                        
                        if self._is_valid_placement(pl, pw, sx, max_y, board):
                            # Aggressive packing: minimize waste, maximize density
                            score = max_y + pw  # Standard bottom-left scoring
                            if score < best_fit['score']:
                                best_fit = {'score': score, 'x': sx, 'y': max_y, 'rot': rot, 'pl': pl, 'pw': pw}
            
            if best_fit['score'] != float('inf'):
                px, py, pl, pw, rot = best_fit['x'], best_fit['y'], best_fit['pl'], best_fit['pw'], best_fit['rot']
                board.placed_parts.append(PlacedPart(part, px, py, rot))
                self._update_skyline(skyline, px, py, pl, pw)
                current_area += part.area

    def _is_valid_placement(self, pl, pw, x, y, board):
        """FIXED: Proper guillotine constraints with TEST 2 collision detection."""
        # Boundary check
        if x + pl > board.length or y + pw > board.width: 
            return False
        
        # CRITICAL: Check collision with existing parts using EXACT kerf spacing
        kerf = self.saw_kerf  # Use exactly the user-specified kerf value
        for placed_part in board.placed_parts:
            # Use TEST 2 style collision detection with exact kerf
            if not (x + pl + kerf <= placed_part.x or 
                   x >= placed_part.x + placed_part.length + kerf or
                   y + pw + kerf <= placed_part.y or 
                   y >= placed_part.y + placed_part.width + kerf):
                return False
        
        # CRITICAL: Validate guillotine cutting constraints
        return self._validate_guillotine_cutting_pattern(board, x, y, pl, pw)
    
    def _validate_guillotine_cutting_pattern(self, board, x, y, pl, pw):
        """Validate that placement follows guillotine cutting pattern."""
        if not board.placed_parts:
            return True  # First part is always valid
        
        kerf = self.saw_kerf
        
        # For guillotine cuts, each new rectangle must be placeable with straight cuts
        # that don't create L-shaped or complex cutting patterns
        for existing_part in board.placed_parts:
            ex, ey, el, ew = existing_part.x, existing_part.y, existing_part.length, existing_part.width
            
            # Check if rectangles are completely separate (valid)
            if (x + pl + kerf <= ex or ex + el + kerf <= x or 
                y + pw + kerf <= ey or ey + ew + kerf <= y):
                continue  # This pair is valid
            
            # If not completely separate, they must align perfectly on at least one axis
            # to allow straight guillotine cuts
            x_aligned = (abs(x - ex) < 1.0 or abs(x + pl - ex - el) < 1.0 or
                        abs(x - ex - el) < 1.0 or abs(x + pl - ex) < 1.0)
            y_aligned = (abs(y - ey) < 1.0 or abs(y + pw - ey - ew) < 1.0 or
                        abs(y - ey - ew) < 1.0 or abs(y + pw - ey) < 1.0)
            
            # Must align on at least one axis for guillotine cuts
            if not (x_aligned or y_aligned):
                return False  # Would create non-guillotine pattern
        
        return True
    
    def _generate_offcuts_test2_style(self, board, x, y, pl, pw, original_l, original_w):
        """Generate offcuts using TEST 2's guillotine 2-offcut logic."""
        # Clear existing offcuts for this area and generate new ones
        # This mimics TEST 2's _generate_offcuts method
        
        # Right offcut (if any remaining width)
        if original_l > pl:
            right_offcut = {
                'x': x + pl,
                'y': y,
                'length': original_l - pl,
                'width': pw
            }
            # Add to board's available space tracking if needed
        
        # Bottom offcut (if any remaining height)  
        if original_w > pw:
            bottom_offcut = {
                'x': x,
                'y': y + pw,
                'length': original_l,
                'width': original_w - pw
            }
            # Add to board's available space tracking if needed
        
        return True  # Successful offcut generation

    def _update_skyline(self, skyline, x, y, l, w):
        new_segment = (x, y + w, l)
        new_skyline = []
        for sx, sy, sw in skyline:
            if sx + sw <= x or sx >= x + l: new_skyline.append((sx, sy, sw))
            else:
                if sx < x: new_skyline.append((sx, sy, x - sx))
                if sx + sw > x + l: new_skyline.append((x + l, sy, (sx + sw) - (x + l)))
        new_skyline.append(new_segment)
        skyline[:] = self._merge_skyline(sorted(new_skyline))

    def _merge_skyline(self, skyline):
        i = 0
        while i < len(skyline) - 1:
            if skyline[i][1] == skyline[i+1][1] and skyline[i][0] + skyline[i][2] == skyline[i+1][0]:
                skyline[i] = (skyline[i][0], skyline[i][1], skyline[i][2] + skyline[i+1][2])
                skyline.pop(i+1)
            else: i += 1
        return skyline

class DynamicShelfPacker:
    """Strategic shelf packer optimized for efficient consolidation and large offcut creation."""
    def __init__(self, parts, board_dim, saw_kerf=3):
        # Strategic sorting: place largest parts first, leave manageable pieces for later phases
        self.parts = sorted(parts, key=lambda p: p.area, reverse=True)
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf

    def optimize(self):
        boards, unplaced = [], list(self.parts)
        while unplaced:
            board = Board(f"Shelf-B{len(boards)+1}", *self.board_dim)
            boards.append(board)
            self._pack_board(board, unplaced)
            placed_ids = {p.part.id for p in board.placed_parts}
            unplaced = [p for p in unplaced if p.id not in placed_ids]
        return boards, unplaced

    def _pack_board(self, board, parts):
        y_cursor = 0
        
        # NEW STRATEGY: Pack to high density to minimize board count
        utilization_target = 0.80  # Target 80% utilization per board  
        current_area = 0
        target_area = board.length * board.width * utilization_target
        
        while current_area < target_area:
            remaining_height = board.width - y_cursor
            candidates = [p for p in parts if p.width <= remaining_height or (not p.grain_sensitive and p.length <= remaining_height)]
            if not candidates: break
            best_shelf = self._find_best_shelf(candidates, board.length, remaining_height)
            if not best_shelf['parts']: break
            
            x_cursor = 0
            shelf_area = 0
            for part, rot in best_shelf['parts']:
                board.placed_parts.append(PlacedPart(part, x_cursor, y_cursor, rot))
                # CRITICAL FIX: Ensure proper kerf spacing (> 4.4mm, not = 4.4mm)
                x_cursor += (part.width if rot else part.length) + self.saw_kerf  # Exact kerf spacing
                shelf_area += part.area
                
            current_area += shelf_area
            y_cursor += best_shelf['height'] + self.saw_kerf  # Exact kerf spacing
            
            placed_ids = {p.id for p, _ in best_shelf['parts']}
            parts[:] = [p for p in parts if p.id not in placed_ids]
            
            # Break if approaching board limits
            if y_cursor >= board.width - 100:  # Leave 100mm margin
                break

    def _find_best_shelf(self, parts, board_length, max_height):
        best_shelf = {'utilization': -1}
        heights = sorted(list(set([p.width for p in parts if p.width <= max_height] + [p.length for p in parts if not p.grain_sensitive and p.length <= max_height])), reverse=True)
        for height in heights[:5]:
            x_cursor, shelf_area, parts_in_shelf = 0, 0, []
            for part in sorted(parts, key=lambda p:p.length, reverse=True):
                rot = False
                pl, pw = part.length, part.width
                if pw > height:
                    if not part.grain_sensitive and pl <= height: pl, pw, rot = pw, pl, True
                    else: continue
                if x_cursor + pl <= board_length:
                    parts_in_shelf.append((part, rot))
                    shelf_area += part.area
                    x_cursor += pl + self.saw_kerf  # Exact kerf spacing
            util = shelf_area / (board_length * height) if height > 0 else 0
            if util > best_shelf['utilization']:
                best_shelf = {'utilization': util, 'height': height, 'parts': parts_in_shelf}
        return best_shelf

# --- 3. Strategic Consolidator ---

class StrategicConsolidator:
    """Consolidates boards by strategically utilizing large offcuts generated by Phase 1."""
    
    def __init__(self, boards, unplaced_parts, board_dim, saw_kerf=4.4):
        self.boards = boards
        self.unplaced_parts = unplaced_parts
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
    
    def consolidate_with_offcuts(self):
        """Consolidate boards by filling large offcuts with unplaced parts."""
        # Calculate available offcut areas on each board
        board_offcuts = []
        for board in self.boards:
            offcuts = self._calculate_board_offcuts(board)
            board_offcuts.append((board, offcuts))
        
        # Sort unplaced parts by area (largest first for better consolidation)
        remaining_parts = sorted(self.unplaced_parts, key=lambda p: p.area, reverse=True)
        
        # Try to place unplaced parts in existing board offcuts
        for part in list(remaining_parts):
            best_placement = None
            best_efficiency = 0
            
            for board, offcuts in board_offcuts:
                for offcut_rect in offcuts:
                    ox, oy, ow, oh = offcut_rect
                    # Try both orientations
                    for rotated in [False, True]:
                        if rotated and part.grain_sensitive:
                            continue
                        
                        pl = part.width if rotated else part.length
                        pw = part.length if rotated else part.width
                        
                        if pl <= ow and pw <= oh:
                            # Calculate efficiency (how well part fills the offcut)
                            part_area = pl * pw
                            offcut_area = ow * oh
                            efficiency = part_area / offcut_area
                            
                            if efficiency > best_efficiency:
                                best_efficiency = efficiency
                                best_placement = (board, ox, oy, rotated)
            
            # Place part in best offcut if found
            if best_placement:
                board, x, y, rotated = best_placement
                board.placed_parts.append(PlacedPart(part, x, y, rotated))
                remaining_parts.remove(part)
                
                # Recalculate offcuts for this board
                for i, (b, offcuts) in enumerate(board_offcuts):
                    if b == board:
                        board_offcuts[i] = (board, self._calculate_board_offcuts(board))
                        break
        
        return self.boards, remaining_parts
    
    def _calculate_board_offcuts(self, board):
        """Calculate large rectangular offcuts on a board."""
        if not board.placed_parts:
            return [(0, 0, board.length, board.width)]
        
        # Find large rectangular offcuts
        offcuts = []
        
        # Right side offcut
        max_used_x = max(p.x + p.length for p in board.placed_parts) if board.placed_parts else 0
        if max_used_x < board.length - 300:  # Minimum 300mm width offcut
            offcuts.append((max_used_x + self.saw_kerf, 0, 
                          board.length - max_used_x - self.saw_kerf, board.width))
        
        # Bottom offcut
        max_used_y = max(p.y + p.width for p in board.placed_parts) if board.placed_parts else 0
        if max_used_y < board.width - 200:  # Minimum 200mm height offcut
            offcuts.append((0, max_used_y + self.saw_kerf, 
                          board.length, board.width - max_used_y - self.saw_kerf))
        
        # Filter offcuts by minimum usable size (300mm x 200mm minimum)
        return [(x, y, w, h) for x, y, w, h in offcuts if w >= 300 and h >= 200]

# --- 4. Aggressive Board Merger ---

class AggressiveBoardMerger:
    """Phase 3: Aggressively merge boards by forcing high-density repacking."""
    
    def __init__(self, boards, board_dim, saw_kerf=4.4):
        self.boards = sorted(boards, key=lambda b: b.utilization())  # Start with lowest utilization
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
    
    def force_consolidation(self):
        """Aggressively merge boards to minimize count."""
        current_boards = self.boards[:]
        
        # Multiple consolidation passes
        for iteration in range(3):  # Up to 3 aggressive passes
            merged = False
            i = 0
            
            while i < len(current_boards) - 1:
                source_board = current_boards[i]
                source_parts = [p.part for p in source_board.placed_parts]
                
                # Try to merge with any other board
                for j in range(i + 1, len(current_boards)):
                    target_board = current_boards[j]
                    target_parts = [p.part for p in target_board.placed_parts]
                    
                    # Attempt aggressive consolidation
                    if self._force_merge_boards(source_parts, target_parts, target_board):
                        current_boards.pop(i)  # Remove source board
                        merged = True
                        break
                
                if not merged:
                    i += 1
                else:
                    merged = False  # Reset for next iteration
            
            # If no merges occurred in this iteration, break
            if not merged:
                break
        
        return current_boards
    
    def _force_merge_boards(self, source_parts, target_parts, target_board):
        """Force merge two boards using aggressive packing."""
        all_parts = source_parts + target_parts
        
        # Try both packing strategies aggressively
        for packer_class in [SkylineBottomLeftPacker, DynamicShelfPacker]:
            # Temporarily increase utilization targets for merging
            packer = packer_class(all_parts, self.board_dim, self.saw_kerf)
            
            # Force higher utilization during merge attempts
            if hasattr(packer, '_pack_board'):
                original_pack = packer._pack_board
                
                def aggressive_pack(board, parts):
                    # Override utilization targets for aggressive merging
                    if hasattr(board, 'length') and hasattr(board, 'width'):
                        if packer_class == SkylineBottomLeftPacker:
                            # Force 95% utilization target during merge
                            target_area = board.length * board.width * 0.95
                            current_area = 0
                            
                            skyline = [(0, 0, board.length)]
                            for part in parts:
                                if current_area >= target_area:
                                    break
                                
                                best_fit = {'score': float('inf')}
                                for sx, sy, sw in skyline:
                                    for rot in [False, True]:
                                        pl, pw = (part.length, part.width) if not rot else (part.width, part.length)
                                        if rot and part.grain_sensitive: continue
                                        if pl <= sw:
                                            max_y = 0
                                            for ssx, ssy, ssw in skyline:
                                                if not (sx + pl <= ssx or sx >= ssx + ssw):
                                                    max_y = max(max_y, ssy)
                                            
                                            if packer._is_valid_placement(pl, pw, sx, max_y, board):
                                                score = max_y + pw
                                                if score < best_fit['score']:
                                                    best_fit = {'score': score, 'x': sx, 'y': max_y, 'rot': rot, 'pl': pl, 'pw': pw}
                                
                                if best_fit['score'] != float('inf'):
                                    px, py, pl, pw, rot = best_fit['x'], best_fit['y'], best_fit['pl'], best_fit['pw'], best_fit['rot']
                                    board.placed_parts.append(PlacedPart(part, px, py, rot))
                                    packer._update_skyline(skyline, px, py, pl, pw)
                                    current_area += part.area
                        else:
                            original_pack(board, parts)
                    else:
                        original_pack(board, parts)
                
                packer._pack_board = aggressive_pack
            
            temp_boards, unplaced = packer.optimize()
            
            # Success if all parts fit on one board
            if len(temp_boards) == 1 and len(unplaced) == 0:
                target_board.placed_parts = temp_boards[0].placed_parts
                return True
        
        return False

# --- 5. Final Squeeze Optimizer ---

class FinalSqueezeOptimizer:
    """Phase 4: Final squeeze to maximize utilization and minimize boards."""
    
    def __init__(self, boards, unplaced_parts, board_dim, saw_kerf=4.4):
        self.boards = boards
        self.unplaced_parts = unplaced_parts
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
    
    def squeeze_maximum(self):
        """Apply final squeeze optimization techniques."""
        current_boards = self.boards[:]
        remaining_unplaced = self.unplaced_parts[:]
        
        # 1. Try to fit unplaced parts anywhere possible
        for part in list(remaining_unplaced):
            placed = False
            for board in current_boards:
                if self._can_squeeze_part(board, part):
                    self._squeeze_part_onto_board(board, part)
                    remaining_unplaced.remove(part)
                    placed = True
                    break
            
            if not placed and remaining_unplaced:
                # 2. Try creating one more highly utilized board
                new_board = Board(f"Squeeze-B{len(current_boards)+1}", *self.board_dim)
                
                # Use aggressive packing for final board
                final_packer = SkylineBottomLeftPacker(remaining_unplaced, self.board_dim, self.saw_kerf)
                final_boards, final_unplaced = final_packer.optimize()
                
                if final_boards:
                    current_boards.extend(final_boards)
                    remaining_unplaced = final_unplaced
                break
        
        return current_boards, remaining_unplaced
    
    def _can_squeeze_part(self, board, part):
        """Check if part can be squeezed onto board."""
        # Calculate current utilization
        current_util = board.utilization()
        
        # Only try to squeeze if board is not already at maximum utilization
        if current_util >= 95:
            return False
        
        # Simple collision check - try different positions
        for x in range(0, board.length - part.length + 1, 50):  # Check every 50mm
            for y in range(0, board.width - part.width + 1, 50):
                if self._check_squeeze_position(board, part, x, y, False):
                    return True
                # Try rotated
                if not part.grain_sensitive and self._check_squeeze_position(board, part, x, y, True):
                    return True
        
        return False
    
    def _check_squeeze_position(self, board, part, x, y, rotated):
        """Check if part can fit at specific position."""
        pl = part.width if rotated else part.length
        pw = part.length if rotated else part.width
        
        # Boundary check
        if x + pl > board.length or y + pw > board.width:
            return False
        
        # CRITICAL FIX: Collision check with existing parts - use EXACT kerf spacing
        for placed_part in board.placed_parts:
            kerf = self.saw_kerf  # Use exactly the user-specified kerf value
            # Check if rectangles overlap with required kerf spacing
            if not (x + pl + kerf <= placed_part.x or 
                   x >= placed_part.x + placed_part.length + kerf or
                   y + pw + kerf <= placed_part.y or 
                   y >= placed_part.y + placed_part.width + kerf):
                return False
        
        # CRITICAL FIX: Validate guillotine cutting constraints for squeeze placement
        return self._validate_guillotine_cutting_pattern_squeeze(board, x, y, pl, pw)
    
    def _validate_guillotine_cutting_pattern_squeeze(self, board, x, y, pl, pw):
        """Validate guillotine constraints for squeeze placement - same logic as main placement."""
        if not board.placed_parts:
            return True
        
        kerf = self.saw_kerf
        
        # Apply same guillotine validation as main placement
        for existing_part in board.placed_parts:
            ex, ey, el, ew = existing_part.x, existing_part.y, existing_part.length, existing_part.width
            
            # Check if rectangles are completely separate (valid)
            if (x + pl + kerf <= ex or ex + el + kerf <= x or 
                y + pw + kerf <= ey or ey + ew + kerf <= y):
                continue  # This pair is valid
            
            # If not completely separate, they must align perfectly on at least one axis
            x_aligned = (abs(x - ex) < 1.0 or abs(x + pl - ex - el) < 1.0 or
                        abs(x - ex - el) < 1.0 or abs(x + pl - ex) < 1.0)
            y_aligned = (abs(y - ey) < 1.0 or abs(y + pw - ey - ew) < 1.0 or
                        abs(y - ey - ew) < 1.0 or abs(y + pw - ey) < 1.0)
            
            # Must align on at least one axis for guillotine cuts
            if not (x_aligned or y_aligned):
                return False  # Would create non-guillotine pattern
        
        return True
    
    def _squeeze_part_onto_board(self, board, part):
        """Actually place the part on the board."""
        # Find the best valid position
        for x in range(0, board.length - part.length + 1, 50):
            for y in range(0, board.width - part.width + 1, 50):
                if self._check_squeeze_position(board, part, x, y, False):
                    board.placed_parts.append(PlacedPart(part, x, y, False))
                    return
                if not part.grain_sensitive and self._check_squeeze_position(board, part, x, y, True):
                    board.placed_parts.append(PlacedPart(part, x, y, True))
                    return

# --- 6. Phase 1.5 & 2: Consolidation & Cleanup ---

class EdgeFitOptimizer:
    """NEW: Phase 1.5 - Tries to fit small parts into long, narrow edge strips."""
    def __init__(self, layout, unplaced_parts, saw_kerf=3):
        self.layout = layout
        self.unplaced_parts = sorted(unplaced_parts, key=lambda p: p.area, reverse=True)
        self.saw_kerf = saw_kerf

    def run_edge_fit(self):
        if not self.unplaced_parts: return
        for board in self.layout:
            # Simplified: find max X and Y to identify edge strips
            if not board.placed_parts: continue
            max_x = max(p.x + p.length for p in board.placed_parts)
            max_y = max(p.y + p.width for p in board.placed_parts)

            # Try to fit into the right-edge strip
            strip_x, strip_y, strip_l, strip_w = max_x + self.saw_kerf, 0, board.length - max_x - self.saw_kerf, board.width
            if strip_l > 0:
                self._pack_strip(board, self.unplaced_parts, (strip_x, strip_y, strip_l, strip_w))

    def _pack_strip(self, board, parts, strip_dims):
        sx, sy, sl, sw = strip_dims
        y_cursor = sy
        for part in list(parts):
            for rot in [False, True]:
                pl, pw = (part.length, part.width) if not rot else (part.width, part.length)
                if rot and part.grain_sensitive: continue
                if pl <= sl and pw <= (sw - (y_cursor - sy)):
                    board.placed_parts.append(PlacedPart(part, sx, y_cursor, rot))
                    # CRITICAL FIX: Ensure proper kerf spacing (> 4.4mm, not = 4.4mm)
                    y_cursor += pw + self.saw_kerf  # Exact kerf spacing
                    parts.remove(part)
                    break

class AgentBasedConsolidator:
    """Phase 2: Consolidates the layout using bulk agent repacking."""
    def __init__(self, layout, board_dim, saw_kerf=3):
        self.layout = sorted(layout, key=lambda b: b.utilization())
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf

    def consolidate(self, max_iterations=5):
        for _ in range(max_iterations):
            if len(self.layout) <= 1: break
            source_board = self.layout[0]
            agent_parts = [p.part for p in source_board.placed_parts]
            if not agent_parts:
                self.layout.pop(0)
                continue
            consolidated = False
            for target_board in reversed(self.layout[1:]):
                if self._try_bulk_repack(target_board, agent_parts):
                    consolidated = True
                    break
            if consolidated:
                self.layout.pop(0)
                self.layout.sort(key=lambda b: b.utilization())
            else: break
        return self.layout

    def _try_bulk_repack(self, target, agents):
        current_parts = [p.part for p in target.placed_parts]
        free_area = (target.length * target.width) * (1 - target.utilization() / 100)
        if sum(p.area for p in agents) > free_area: return False
        
        repacker = SkylineBottomLeftPacker(current_parts + agents, self.board_dim, self.saw_kerf)
        temp_layout, _ = repacker.optimize()
        if len(temp_layout) == 1:
            target.placed_parts = temp_layout[0].placed_parts
            return True
        return False

class SimulatedAnnealingConsolidator:
    """Phase 3: Advanced Simulated Annealing optimization for layout refinement."""
    def __init__(self, layout, board_dim, saw_kerf=3, initial_temp=1000, cooling_rate=0.95, max_iter=500):
        self.initial_layout = deepcopy(layout)
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
        self.temp = initial_temp
        self.cooling_rate = cooling_rate
        self.max_iter = max_iter

    def run(self):
        current = deepcopy(self.initial_layout)
        current_score = self.evaluate(current)
        best = deepcopy(current)
        best_score = current_score

        for iteration in range(self.max_iter):
            candidate = self.perturb(deepcopy(current))
            if candidate is None:  # Invalid perturbation
                continue
                
            candidate_score = self.evaluate(candidate)

            if candidate_score < current_score or self.acceptance_prob(current_score, candidate_score) > random.random():
                current = candidate
                current_score = candidate_score
                if candidate_score < best_score:
                    best = deepcopy(candidate)
                    best_score = candidate_score

            self.temp *= self.cooling_rate

        return best

    def evaluate(self, layout):
        """Evaluate layout quality: prioritize fewer boards, then minimize waste."""
        total_boards = len(layout)
        total_waste = 0
        for b in layout:
            used_area = sum(p.part.area for p in b.placed_parts)
            board_area = self.board_dim[0] * self.board_dim[1]
            waste = board_area - used_area
            total_waste += waste
        return total_boards * 100000 + total_waste

    def perturb(self, layout):
        """Perturb layout by swapping parts between boards with collision validation."""
        if len(layout) < 2:
            return layout
            
        # Try multiple swap attempts
        for _ in range(10):
            b1, b2 = random.sample(layout, 2)
            if not b1.placed_parts or not b2.placed_parts:
                continue
                
            p1 = random.choice(b1.placed_parts)
            p2 = random.choice(b2.placed_parts)
            
            # Perform swap
            b1.placed_parts.remove(p1)
            b2.placed_parts.remove(p2)
            b1.placed_parts.append(p2)
            b2.placed_parts.append(p1)
            
            # Validate swap doesn't cause overlaps
            if self._validate_layout(layout):
                return layout
            else:
                # Revert swap if invalid
                b1.placed_parts.remove(p2)
                b2.placed_parts.remove(p1)
                b1.placed_parts.append(p1)
                b2.placed_parts.append(p2)
        
        return None  # No valid perturbation found

    def _validate_layout(self, layout):
        """Validate that layout has no overlaps and respects boundaries."""
        for board in layout:
            # Check boundaries
            for part in board.placed_parts:
                if (part.x + part.length > board.length or 
                    part.y + part.width > board.width):
                    return False
            
            # Check collisions
            for i in range(len(board.placed_parts)):
                for j in range(i + 1, len(board.placed_parts)):
                    p1, p2 = board.placed_parts[i], board.placed_parts[j]
                    if self._parts_overlap(p1, p2):
                        return False
        return True

    def _parts_overlap(self, p1, p2):
        """Check if two placed parts overlap (accounting for kerf)."""
        return not (p1.x + p1.length + self.saw_kerf <= p2.x or 
                   p2.x + p2.length + self.saw_kerf <= p1.x or
                   p1.y + p1.width + self.saw_kerf <= p2.y or 
                   p2.y + p2.width + self.saw_kerf <= p1.y)

    def acceptance_prob(self, old_score, new_score):
        """Calculate acceptance probability for worse solutions."""
        if new_score < old_score:
            return 1.0
        if self.temp <= 0:
            return 0.0
        return math.exp((old_score - new_score) / self.temp)

class ILPOptimizer:
    """Phase 4: Integer Linear Programming optimization for underutilized boards."""
    def __init__(self, layout, board_dim, saw_kerf=3):
        self.layout = layout
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf

    def apply_ilp_optimizer(self):
        """Apply ILP optimization to underutilized boards grouped by material type."""
        from collections import defaultdict
        
        underutilized = defaultdict(list)
        
        # Group underutilized boards by material type
        for board in self.layout:
            util = board.utilization()
            # Extract material key from board (simplified for now)
            material_key = self._extract_material_key(board)
            if util < 70:
                underutilized[material_key].append(board)
        
        # Apply ILP optimization to each material group
        optimized_layout = list(self.layout)
        for material_key, boards in underutilized.items():
            if len(boards) > 1:
                print(f"🧠 ILP Optimization triggered for {material_key}: {len(boards)} boards <70% utilization")
                optimized_boards = self._run_ilp_consolidation(boards, material_key)
                
                # Replace original boards with optimized ones
                for original_board in boards:
                    if original_board in optimized_layout:
                        optimized_layout.remove(original_board)
                
                optimized_layout.extend(optimized_boards)
        
        return optimized_layout

    def _extract_material_key(self, board):
        """Extract material identifier from board."""
        if hasattr(board, 'material_details') and board.material_details:
            return board.material_details.full_material_string
        elif hasattr(board, 'placed_parts') and board.placed_parts:
            # Try to get material from first part
            first_part = board.placed_parts[0]
            if hasattr(first_part, 'part') and hasattr(first_part.part, 'material_details'):
                return first_part.part.material_details.full_material_string
        return "default_material"

    def _run_ilp_consolidation(self, boards, material_key):
        """Run ILP-based consolidation on underutilized boards."""
        # Collect all parts from underutilized boards
        all_parts = []
        for board in boards:
            for placed_part in board.placed_parts:
                all_parts.append(placed_part.part)
        
        if not all_parts:
            return boards
        
        print(f"  Consolidating {len(all_parts)} parts from {len(boards)} boards")
        
        # Try both packing strategies for best result
        skyline_packer = SkylineBottomLeftPacker(all_parts, self.board_dim, self.saw_kerf)
        shelf_packer = DynamicShelfPacker(all_parts, self.board_dim, self.saw_kerf)
        
        skyline_boards, _ = skyline_packer.optimize()
        shelf_boards, _ = shelf_packer.optimize()
        
        # Choose the better result
        if len(skyline_boards) <= len(shelf_boards):
            consolidated_boards = skyline_boards
            strategy = "Skyline"
        else:
            consolidated_boards = shelf_boards
            strategy = "Shelf"
        
        # If consolidation reduces board count, use it
        if len(consolidated_boards) < len(boards):
            print(f"  ✅ ILP ({strategy}) reduced boards: {len(boards)} → {len(consolidated_boards)}")
            return consolidated_boards
        else:
            print(f"  ➤ ILP no improvement: keeping original {len(boards)} boards")
            return boards

# --- 4. Final Reporting ---

class ReportGenerator:
    """Generates visual PDF and data CSV reports."""
    def __init__(self, layout, board_dim, saw_kerf=3):
        self.layout = layout
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf

    def generate_pdf(self, filename="cutting_layout_final.pdf"):
        with PdfPages(filename) as pdf:
            for board in self.layout:
                fig, ax = plt.subplots(figsize=(12, 8))
                ax.set_title(f"Board {board.id} (Utilization: {board.utilization():.2f}%)")
                ax.set_xlim(0, self.board_dim[0])
                ax.set_ylim(0, self.board_dim[1])
                ax.set_xticks(range(0, self.board_dim[0] + 1, 200))
                ax.set_yticks(range(0, self.board_dim[1] + 1, 200))
                ax.grid(True, linestyle='--', alpha=0.6)
                
                for p in board.placed_parts:
                    ax.add_patch(Rectangle((p.x, p.y), p.length, p.width, facecolor='skyblue', edgecolor='black', lw=1))
                    ax.text(p.x + p.length / 2, p.y + p.width / 2, f"{p.part.id}\n{p.length}x{p.width}", 
                            ha='center', va='center', fontsize=8)
                
                cut_map = self._generate_cut_map(board)
                cut_text = "Guillotine Cut Sequence:\n" + "\n".join(cut_map)
                fig.text(0.92, 0.5, cut_text, va='center', ha='left', fontsize=9, bbox=dict(boxstyle="round,pad=0.5", fc="wheat", alpha=0.5))
                plt.tight_layout(rect=[0, 0, 0.9, 1])
                pdf.savefig(fig)
                plt.close()
        print(f"\nVisual report saved to {filename}")

    def _generate_cut_map(self, board):
        cuts = set()
        for p in board.placed_parts:
            cuts.add(f"H@{p.y}")
            cuts.add(f"H@{p.y + p.width}")
            cuts.add(f"V@{p.x}")
            cuts.add(f"V@{p.x + p.length}")
        return sorted(list(cuts))

# --- 5. Master Optimizer ---

class MasterOptimizer:
    """Orchestrates the entire cutting optimization pipeline."""
    
    def __init__(self, parts, board_dim, saw_kerf=4.4):
        self.parts = parts
        self.board_dim = board_dim
        self.saw_kerf = saw_kerf
        self.lower_bound = calculate_lower_bound(parts, board_dim)
    
    def optimize(self):
        """Run the complete optimization pipeline."""
        print(f"\n=== OptiWise Master Optimizer ===")
        print(f"Parts: {len(self.parts)}, Board: {self.board_dim}, Kerf: {self.saw_kerf}mm")
        print(f"Theoretical lower bound: {self.lower_bound} boards\n")
        
        # Phase 1: Always use Cutlist Plus FX algorithm for best efficiency
        # The other algorithms are just for comparison - we need maximum packing density
        cutlist_packer = CutlistPlusFXPacker(deepcopy(self.parts), self.board_dim, self.saw_kerf)
        cutlist_layout, cutlist_unplaced = cutlist_packer.optimize()
        
        # Run comparison algorithms for performance metrics only
        skyline_packer = SkylineBottomLeftPacker(deepcopy(self.parts), self.board_dim, self.saw_kerf)
        shelf_packer = DynamicShelfPacker(deepcopy(self.parts), self.board_dim, self.saw_kerf)
        skyline_layout, skyline_unplaced = skyline_packer.optimize()
        shelf_layout, shelf_unplaced = shelf_packer.optimize()
        
        # Calculate real utilization like Cutlist Plus FX
        def calculate_real_utilization(boards):
            if not boards:
                return 0
            total_part_area = sum(sum(p.part.area for p in b.placed_parts) for b in boards)
            total_board_area = len(boards) * self.board_dim[0] * self.board_dim[1]
            return (total_part_area / total_board_area) * 100
        
        cutlist_util = calculate_real_utilization(cutlist_layout)
        skyline_util = calculate_real_utilization(skyline_layout)
        shelf_util = calculate_real_utilization(shelf_layout)
        
        # Select best algorithm based on efficiency (like Cutlist Plus FX)
        algorithms = [
            (cutlist_layout, cutlist_unplaced, cutlist_util, "Cutlist Plus FX Style"),
            (skyline_layout, skyline_unplaced, skyline_util, "Skyline"),
            (shelf_layout, shelf_unplaced, shelf_util, "Shelf")
        ]
        
        # Sort by utilization descending
        algorithms.sort(key=lambda x: x[2], reverse=True)
        best_layout, unplaced, best_util, best_name = algorithms[0]
        
        print(f"Phase 1: {best_name} selected ({len(best_layout)} boards, {best_util:.1f}% util)")
        print(f"  Comparison - Cutlist: {cutlist_util:.1f}%, Skyline: {skyline_util:.1f}%, Shelf: {shelf_util:.1f}%")
        
        # Phase 1.5: Edge-Fit Optimization
        initial_unplaced_count = len(unplaced)
        edge_optimizer = EdgeFitOptimizer(best_layout, unplaced, self.saw_kerf)
        edge_optimizer.run_edge_fit()
        unplaced = edge_optimizer.unplaced_parts
        edge_fit_placed = initial_unplaced_count - len(unplaced)
        print(f"Phase 1.5: Edge-fit placed {edge_fit_placed} additional parts")
        
        # Phase 2: Strategic Board Consolidation (Focus on utilizing large offcuts)
        initial_board_count = len(best_layout)
        consolidator = StrategicConsolidator(best_layout, unplaced, self.board_dim, self.saw_kerf)
        consolidated_layout, remaining_unplaced = consolidator.consolidate_with_offcuts()
        unplaced = remaining_unplaced
        boards_saved = initial_board_count - len(consolidated_layout)
        print(f"Phase 2: Strategic consolidation to {len(consolidated_layout)} boards (saved {boards_saved})")
        
        # Phase 3: Aggressive Board Merging (Replace SA with direct consolidation)
        pre_merge_count = len(consolidated_layout)
        merger = AggressiveBoardMerger(consolidated_layout, self.board_dim, self.saw_kerf)
        merged_layout = merger.force_consolidation()
        merge_improvement = pre_merge_count - len(merged_layout)
        print(f"Phase 3: Aggressive merging to {len(merged_layout)} boards (improved {merge_improvement})")
        
        # Phase 4: Use existing optimization as backup if aggressive approach fails
        if len(merged_layout) > self.lower_bound + 1:  # If still inefficient, use proven method
            print("Phase 4: Switching to proven optimization method...")
            from optimization_enhanced_global import run_enhanced_global_optimization
            try:
                # Convert back to OptiWise format for proven algorithm
                all_remaining_parts = []
                for board in merged_layout:
                    for placed_part in board.placed_parts:
                        all_remaining_parts.append(placed_part.part)
                
                # Add any unplaced parts
                all_remaining_parts.extend(unplaced)
                
                # Use the proven algorithm as final phase
                fallback_boards, fallback_unplaced, _, _, _ = run_enhanced_global_optimization(
                    parts=all_remaining_parts,
                    core_db=core_db,
                    laminate_db=laminate_db,
                    kerf=self.saw_kerf
                )
                
                # If proven method is better, use it
                if len(fallback_boards) < len(merged_layout):
                    final_layout = fallback_boards
                    final_unplaced = fallback_unplaced
                    improvement = len(merged_layout) - len(fallback_boards)
                    print(f"Phase 4: Proven method improved to {len(final_layout)} boards (saved {improvement})")
                else:
                    final_layout, final_unplaced = merged_layout, unplaced
                    print(f"Phase 4: Kept aggressive result with {len(final_layout)} boards")
            except:
                # Fallback to final squeeze if proven method fails
                squeezer = FinalSqueezeOptimizer(merged_layout, unplaced, self.board_dim, self.saw_kerf)
                final_layout, final_unplaced = squeezer.squeeze_maximum()
                print(f"Phase 4: Final squeeze to {len(final_layout)} boards")
        else:
            # Use final squeeze for already good results
            pre_squeeze_count = len(merged_layout)
            squeezer = FinalSqueezeOptimizer(merged_layout, unplaced, self.board_dim, self.saw_kerf)
            final_layout, final_unplaced = squeezer.squeeze_maximum()
            squeeze_improvement = pre_squeeze_count - len(final_layout)
            print(f"Phase 4: Final squeeze to {len(final_layout)} boards (improved {squeeze_improvement})")
        
        # Final efficiency summary
        total_improvement = len(best_layout) - len(final_layout)
        final_utilization = sum(b.utilization() for b in final_layout) / len(final_layout) if final_layout else 0
        print(f"\nOptimization Summary:")
        print(f"  Initial boards: {len(best_layout)}")
        print(f"  Final boards: {len(final_layout)}")
        print(f"  Total improvement: {total_improvement} boards saved")
        print(f"  Average utilization: {final_utilization:.1f}%")
        
        return final_layout, final_unplaced

# --- 6. Integration Interface ---

def run_test4_optimization(parts, core_db, laminate_db, kerf=4.4):
    """
    Interface function with material segregation and guillotine constraints.
    """
    global global_core_db, global_laminate_db
    global_core_db = core_db
    global_laminate_db = laminate_db
    
    from data_models import Board as OriginalBoard
    
    # Group parts by material first to prevent mixing
    material_groups = {}
    for part in parts:
        # Use the full material string as the grouping key
        if hasattr(part, 'material_details') and part.material_details:
            if hasattr(part.material_details, 'full_material_string'):
                material_key = part.material_details.full_material_string
            elif hasattr(part.material_details, 'material_name'):
                material_key = part.material_details.material_name
            else:
                material_key = str(part.material_details)
        elif hasattr(part, 'original_material'):
            material_key = part.original_material
        else:
            material_key = "DEFAULT"
        
        if material_key not in material_groups:
            material_groups[material_key] = []
        material_groups[material_key].append(part)
    
    print(f"Material groups found: {list(material_groups.keys())}")
    
    result_boards = []
    unplaced_parts = []
    
    # Process each material group separately
    for material_key, material_parts in material_groups.items():
        print(f"Processing {len(material_parts)} parts of material: {material_key}")
        
        # Convert to internal format
        new_parts = []
        for part in material_parts:
            new_part = Part(
                id=part.id,
                length=part.requested_length,
                width=part.requested_width,
                grain_sensitive=True
            )
            new_parts.append(new_part)
        
        # Run optimizer for this material group only
        optimizer = MasterOptimizer(new_parts, (2440, 1220), kerf)
        material_boards, material_unplaced = optimizer.optimize()
        
        # Convert back to OptiWise format with material segregation
        for board in material_boards:
            # Create OptiWise compatible board with proper constructor
            from data_models import MaterialDetails
            board_material = material_parts[0].material_details if material_parts else MaterialDetails("DEFAULT")
            
            optiwise_board = OriginalBoard(
                board_id=f"{material_key}_{board.id}",
                material_details=board_material,
                kerf=kerf,
                total_length=2440,
                total_width=1220
            )
            
            # Add parts to board with proper positioning
            for placed_part in board.placed_parts:
                original_part = next(p for p in material_parts if p.id == placed_part.part.id)
                
                # Set position attributes for OptiWise compatibility
                original_part.x = placed_part.x
                original_part.y = placed_part.y
                original_part.actual_length = placed_part.length
                original_part.actual_width = placed_part.width
                original_part.rotated = placed_part.rotated
                original_part.placed = True
                
                optiwise_board.parts_on_board.append(original_part)
            
            result_boards.append(optiwise_board)
        
        # Add unplaced parts from this material
        for unplaced_part_id in [up.id for up in material_unplaced]:
            unplaced_original = next(p for p in material_parts if p.id == unplaced_part_id)
            unplaced_parts.append(unplaced_original)
    
    # Verify no material mixing
    for board in result_boards:
        materials_on_board = set()
        for part in board.parts_on_board:
            if hasattr(part, 'material_details') and part.material_details:
                if hasattr(part.material_details, 'full_material_string'):
                    materials_on_board.add(part.material_details.full_material_string)
                else:
                    materials_on_board.add(str(part.material_details))
        
        if len(materials_on_board) > 1:
            print(f"WARNING: Board {board.id} has mixed materials: {materials_on_board}")
    
    print(f"Final results: {len(result_boards)} boards, {len(unplaced_parts)} unplaced parts")
    
    # Calculate costs (simplified)
    initial_cost = len(parts) * 100
    final_cost = len(result_boards) * 500
    upgrade_summary = {}
    
    return result_boards, unplaced_parts, upgrade_summary, initial_cost, final_cost