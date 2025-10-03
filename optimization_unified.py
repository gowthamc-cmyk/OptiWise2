"""
Unified Optimization Engine
Integrates all optimization algorithms with intelligent algorithm selection.
"""

import logging
import time
from typing import List, Dict, Tuple, Optional
from data_models import MaterialDetails, Part, Offcut, Board

logger = logging.getLogger(__name__)


class OptimizationStrategy:
    """Enum-like class for optimization strategies."""
    FAST = "fast"
    BALANCED = "balanced"
    MAXIMUM_EFFICIENCY = "maximum_efficiency"
    MATHEMATICAL = "mathematical"
    TEST_ALGORITHM = "test_algorithm"


class UnifiedOptimizer:
    """Unified optimization engine with intelligent algorithm selection."""
    
    def __init__(self, strategy: str = OptimizationStrategy.BALANCED):
        self.strategy = strategy
        
    def optimize(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                user_upgrade_sequence_str: str, kerf: float = 4.4) -> Tuple[List[Board], List[Part], List[Dict], float, float, Dict]:
        """
        Main optimization entry point with intelligent algorithm selection.
        """
        start_time = time.time()
        logger.info(f"Starting unified optimization with strategy: {self.strategy}")
        logger.info(f"Optimizing {len(parts_list)} parts")
        
        # Select optimization approach based on problem size and strategy
        approach_info = self._select_optimization_approach(parts_list)
        
        if self.strategy == OptimizationStrategy.FAST:
            results = self._run_fast_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        elif self.strategy == OptimizationStrategy.MATHEMATICAL:
            # Force exclusive use of mathematical algorithm for all materials
            logger.info("Using pure mathematical optimization with enhanced consolidation")
            results = self._run_mathematical_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        elif self.strategy == OptimizationStrategy.MAXIMUM_EFFICIENCY:
            results = self._run_maximum_efficiency_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        elif self.strategy == OptimizationStrategy.TEST_ALGORITHM:
            results = self._run_test_algorithm_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        elif self.strategy == "no_upgrade":
            results = self._run_no_upgrade_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        else:  # BALANCED
            results = self._run_balanced_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        
        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = results
        
        # Generate optimization report
        optimization_time = time.time() - start_time
        report = self._generate_optimization_report(boards, unplaced_parts, optimization_time, approach_info)
        
        logger.info(f"Unified optimization complete in {optimization_time:.2f}s: {len(boards)} boards, {len(unplaced_parts)} unplaced")
        
        return boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, report
    
    def _select_optimization_approach(self, parts_list: List[Part]) -> Dict:
        """Intelligently select optimization approach based on problem characteristics."""
        num_parts = len(parts_list)
        
        # Analyze problem complexity
        unique_materials = len(set(str(part.material_details) for part in parts_list))
        avg_part_area = sum(part.get_area_with_kerf(4.4) for part in parts_list) / num_parts
        
        # Calculate complexity score
        complexity_score = (num_parts * 0.1) + (unique_materials * 0.5) + (avg_part_area / 100000)
        
        approach_info = {
            'num_parts': num_parts,
            'unique_materials': unique_materials,
            'avg_part_area': avg_part_area,
            'complexity_score': complexity_score,
            'recommended_algorithms': []
        }
        
        if self.strategy == OptimizationStrategy.FAST:
            approach_info['recommended_algorithms'] = ['greedy_best_fit']
        elif self.strategy == OptimizationStrategy.MATHEMATICAL:
            if num_parts <= 50:
                approach_info['recommended_algorithms'] = ['branch_and_bound', 'linear_programming', 'mixed_integer']
            else:
                approach_info['recommended_algorithms'] = ['linear_programming', 'mixed_integer', 'genetic_fallback']
        elif self.strategy == OptimizationStrategy.MAXIMUM_EFFICIENCY:
            if num_parts <= 30:
                approach_info['recommended_algorithms'] = ['genetic_algorithm', 'dynamic_programming', 'simulated_annealing']
            else:
                approach_info['recommended_algorithms'] = ['genetic_algorithm', 'bottom_left_fill', 'consolidation']
        else:  # BALANCED
            if num_parts <= 20:
                approach_info['recommended_algorithms'] = ['dynamic_programming', 'bottom_left_fill']
            elif num_parts <= 100:
                approach_info['recommended_algorithms'] = ['genetic_algorithm', 'consolidation']
            else:
                approach_info['recommended_algorithms'] = ['bottom_left_fill', 'consolidation']
        
        return approach_info
    
    def _run_fast_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                              user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run fast optimization using existing greedy algorithm."""
        from optimization_core_fixed import run_optimization
        return run_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
    
    def _run_mathematical_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                     user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run mathematical optimization using Integer Linear Programming."""
        try:
            from optimization_mathematical_fixed import run_mathematical_optimization_fixed
            
            upgrade_sequence = self._parse_upgrade_sequence(user_upgrade_sequence_str)
            
            # Calculate initial cost
            initial_cost = sum(part.material_details.get_cost_per_sqm(laminate_db, core_db) * 
                             part.get_area_with_kerf(kerf) / 1_000_000 for part in parts_list)
            
            # Run mathematical optimization that respects existing upgrade logic
            boards, unplaced_parts, summary = run_mathematical_optimization_fixed(
                parts_list, core_db, laminate_db, upgrade_sequence, kerf
            )
            
            # Calculate final cost
            final_cost = sum(board.material_details.get_cost_per_sqm(laminate_db, core_db) * 
                           (board.total_length * board.total_width) / 1_000_000 for board in boards)
            
            # Create upgrade summary in expected format
            upgrade_summary = [{
                'optimization_method': summary.get('optimization_method', 'Mathematical (Fixed)'),
                'total_boards': len(boards),
                'material_upgrades': summary.get('material_upgrades', 0),
                'average_utilization': summary.get('total_utilization', 0),
                'cost_reduction': max(0, initial_cost - final_cost),
                'efficiency_score': summary.get('total_utilization', 0)
            }]
            
            return boards, unplaced_parts, upgrade_summary, initial_cost, final_cost
            
        except Exception as e:
            logger.warning(f"Mathematical optimization failed: {e}, falling back to fast optimization")
            return self._run_fast_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
    
    def _parse_upgrade_sequence(self, user_upgrade_sequence_str: str) -> List[str]:
        """Parse comma-separated upgrade sequence string into list."""
        if not user_upgrade_sequence_str:
            return ['18MR', '18BWR']  # Default sequence
        
        # Clean and split the sequence
        sequence = [core.strip() for core in user_upgrade_sequence_str.split(',')]
        return [core for core in sequence if core]  # Remove empty strings
    
    def _run_maximum_efficiency_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                           user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run maximum efficiency optimization using all advanced algorithms."""
        try:
            from optimization_advanced_efficiency import run_advanced_optimization
            return run_advanced_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf,
                                           use_genetic=True, use_dp=True, use_consolidation=True)
        except Exception as e:
            logger.warning(f"Advanced optimization failed: {e}, falling back to balanced optimization")
            return self._run_balanced_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
    
    def _run_test_algorithm_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                        user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run pure TEST algorithm optimization with Simulated Annealing - no material upgrades."""
        try:
            from optimization_test_enhanced import run_test_algorithm_optimization
            
            # Calculate initial cost
            initial_cost = sum(part.material_details.get_cost_per_sqm(laminate_db, core_db) * 
                             part.get_area_with_kerf(kerf) / 1_000_000 for part in parts_list)
            
            # Run pure TEST algorithm (no material upgrades)
            boards, unplaced_parts, summary = run_test_algorithm_optimization(
                parts_list, core_db, laminate_db, [], kerf  # Empty upgrade sequence - pure consolidation only
            )
            
            # Calculate final cost
            final_cost = sum(board.material_details.get_cost_per_sqm(laminate_db, core_db) * 
                           (board.total_length * board.total_width) / 1_000_000 for board in boards)
            
            # Create upgrade summary in expected format
            upgrade_summary = [{
                'optimization_method': summary.get('optimization_method', 'TEST Algorithm (Pure)'),
                'total_boards': len(boards),
                'material_upgrades': 0,  # Pure TEST algorithm - no upgrades ever
                'average_utilization': summary.get('total_utilization', 0),
                'cost_reduction': max(0, initial_cost - final_cost),
                'efficiency_score': summary.get('total_utilization', 0)
            }]
            
            return boards, unplaced_parts, upgrade_summary, initial_cost, final_cost
            
        except Exception as e:
            logger.error(f"Pure TEST algorithm optimization failed: {e}")
            # Don't fall back to upgrade-based algorithms - return empty result to maintain no-upgrade promise
            return [], parts_list, [{
                'optimization_method': 'TEST Algorithm (Failed)',
                'total_boards': 0,
                'material_upgrades': 0,
                'average_utilization': 0,
                'cost_reduction': 0,
                'efficiency_score': 0
            }], 0, 0

    def _run_balanced_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                  user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run balanced optimization combining fast and efficient algorithms."""
        try:
            # Use advanced algorithms with reduced parameters for balance
            from optimization_advanced_efficiency import run_advanced_optimization
            return run_advanced_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf,
                                           use_genetic=len(parts_list) > 50, use_dp=len(parts_list) < 30, use_consolidation=True)
        except Exception as e:
            logger.warning(f"Balanced optimization failed: {e}, falling back to fast optimization")
            return self._run_fast_optimization(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
    
    def _run_no_upgrade_optimization(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                   user_upgrade_sequence_str: str, kerf: float) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run optimization without any material upgrades."""
        from optimization_core_fixed import run_optimization_no_upgrade
        return run_optimization_no_upgrade(parts_list, core_db, laminate_db, kerf)
    
    def _generate_optimization_report(self, boards: List[Board], unplaced_parts: List[Part], 
                                    optimization_time: float, approach_info: Dict) -> Dict:
        """Generate comprehensive optimization report."""
        total_board_area = sum(board.total_length * board.total_width for board in boards)
        total_used_area = sum(board.total_length * board.total_width - board.get_remaining_area() for board in boards)
        total_waste_area = sum(board.get_remaining_area() for board in boards)
        
        utilization_rates = [board.get_utilization_percentage() for board in boards]
        avg_utilization = sum(utilization_rates) / len(utilization_rates) if utilization_rates else 0
        
        # Material usage breakdown
        material_usage = {}
        for board in boards:
            material_key = str(board.material_details)
            if material_key not in material_usage:
                material_usage[material_key] = {'boards': 0, 'total_area': 0, 'used_area': 0}
            
            material_usage[material_key]['boards'] += 1
            material_usage[material_key]['total_area'] += board.total_length * board.total_width
            material_usage[material_key]['used_area'] += (board.total_length * board.total_width - board.get_remaining_area())
        
        # Offcut analysis
        large_offcuts = []
        medium_offcuts = []
        small_offcuts = []
        
        for board in boards:
            for offcut in board.available_rectangles:
                offcut_area = offcut.get_area()
                if offcut_area > 100000:  # > 100k mm²
                    large_offcuts.append(offcut_area)
                elif offcut_area > 10000:  # 10k-100k mm²
                    medium_offcuts.append(offcut_area)
                else:  # < 10k mm²
                    small_offcuts.append(offcut_area)
        
        return {
            'optimization_time': optimization_time,
            'strategy_used': self.strategy,
            'algorithms_applied': approach_info.get('recommended_algorithms', []),
            'total_boards': len(boards),
            'unplaced_parts': len(unplaced_parts),
            'total_board_area': total_board_area,
            'total_used_area': total_used_area,
            'total_waste_area': total_waste_area,
            'overall_utilization': (total_used_area / total_board_area * 100) if total_board_area > 0 else 0,
            'average_board_utilization': avg_utilization,
            'material_usage': material_usage,
            'offcut_analysis': {
                'large_offcuts': len(large_offcuts),
                'medium_offcuts': len(medium_offcuts),
                'small_offcuts': len(small_offcuts),
                'largest_offcut': max(large_offcuts) if large_offcuts else 0,
                'total_waste_percentage': (total_waste_area / total_board_area * 100) if total_board_area > 0 else 0
            },
            'efficiency_metrics': {
                'parts_per_board': len([p for board in boards for p in board.parts_on_board]) / len(boards) if boards else 0,
                'waste_per_board': total_waste_area / len(boards) if boards else 0,
                'utilization_variance': sum((u - avg_utilization)**2 for u in utilization_rates) / len(utilization_rates) if utilization_rates else 0
            }
        }


class MultiObjectiveOptimizer:
    """Multi-objective optimization balancing multiple criteria."""
    
    def __init__(self, weights: Dict[str, float] = None):
        self.weights = weights or {
            'board_count': 0.4,
            'waste_minimization': 0.3,
            'material_cost': 0.2,
            'cutting_complexity': 0.1
        }
    
    def optimize_multi_objective(self, parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                                user_upgrade_sequence_str: str, kerf: float = 4.4) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
        """Run multi-objective optimization with weighted criteria."""
        logger.info("Starting multi-objective optimization")
        
        # Generate multiple solutions with different strategies
        solutions = []
        
        strategies = [OptimizationStrategy.FAST, OptimizationStrategy.BALANCED, OptimizationStrategy.MAXIMUM_EFFICIENCY]
        
        for strategy in strategies:
            try:
                optimizer = UnifiedOptimizer(strategy)
                results = optimizer.optimize(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
                boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, report = results
                
                # Calculate multi-objective score
                score = self._calculate_multi_objective_score(boards, unplaced_parts, final_cost, report)
                
                solutions.append({
                    'strategy': strategy,
                    'results': (boards, unplaced_parts, upgrade_summary, initial_cost, final_cost),
                    'score': score,
                    'report': report
                })
                
            except Exception as e:
                logger.warning(f"Strategy {strategy} failed: {e}")
        
        if not solutions:
            # Fallback to fast optimization
            optimizer = UnifiedOptimizer(OptimizationStrategy.FAST)
            results = optimizer.optimize(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
            return results[:5]  # Return only the core results without report
        
        # Select best solution based on multi-objective score
        best_solution = min(solutions, key=lambda x: x['score'])
        logger.info(f"Best multi-objective solution: {best_solution['strategy']} (score: {best_solution['score']:.2f})")
        
        return best_solution['results']
    
    def _calculate_multi_objective_score(self, boards: List[Board], unplaced_parts: List[Part], 
                                       final_cost: float, report: Dict) -> float:
        """Calculate weighted multi-objective score (lower is better)."""
        # Board count objective (minimize)
        board_count_score = len(boards) * self.weights['board_count']
        
        # Waste minimization objective
        waste_percentage = report['offcut_analysis']['total_waste_percentage']
        waste_score = waste_percentage * self.weights['waste_minimization']
        
        # Material cost objective
        cost_score = (final_cost / 10000) * self.weights['material_cost']  # Normalize cost
        
        # Cutting complexity objective (prefer simpler cuts)
        total_parts = sum(len(board.parts_on_board) for board in boards)
        avg_parts_per_board = total_parts / len(boards) if boards else 0
        complexity_score = (10 - avg_parts_per_board) * self.weights['cutting_complexity']  # Prefer more parts per board
        
        # Unplaced parts penalty
        unplaced_penalty = len(unplaced_parts) * 100
        
        return board_count_score + waste_score + cost_score + complexity_score + unplaced_penalty


def create_optimizer(strategy: str = "balanced", multi_objective: bool = False) -> object:
    """Factory function to create appropriate optimizer."""
    if multi_objective:
        return MultiObjectiveOptimizer()
    else:
        return UnifiedOptimizer(strategy)


def run_unified_optimization(parts_list: List[Part], core_db: Dict, laminate_db: Dict,
                           user_upgrade_sequence_str: str, kerf: float = 4.4, 
                           strategy: str = "balanced", multi_objective: bool = False) -> Tuple[List[Board], List[Part], List[Dict], float, float]:
    """
    Main entry point for unified optimization system.
    
    Args:
        parts_list: List of parts to optimize
        core_db: Core materials database
        laminate_db: Laminate materials database
        user_upgrade_sequence_str: Comma-separated upgrade sequence
        kerf: Kerf width in mm
        strategy: Optimization strategy ('fast', 'balanced', 'maximum_efficiency', 'mathematical')
        multi_objective: Whether to use multi-objective optimization
    
    Returns:
        Tuple of (boards, unplaced_parts, upgrade_summary, initial_cost, final_cost)
    """
    optimizer = create_optimizer(strategy, multi_objective)
    
    if multi_objective:
        return optimizer.optimize_multi_objective(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
    else:
        results = optimizer.optimize(parts_list, core_db, laminate_db, user_upgrade_sequence_str, kerf)
        return results[:5]  # Return only core results, exclude report for compatibility