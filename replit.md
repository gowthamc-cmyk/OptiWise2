# OptiWise - Smart Beam Saw Optimization Tool

## Overview
OptiWise is a Python-based web application designed for intelligent 2D cutting optimization in furniture manufacturing using beam saws. Its primary purpose is to maximize material utilization and minimize waste and cost through advanced algorithms, including material upgrading capabilities. The project aims to provide a comprehensive, user-friendly solution for optimizing cutting layouts and generating professional reports.

## User Preferences

Preferred communication style: Simple, everyday language.

## System Architecture

### Frontend
The application features a Streamlit-based web interface (`app_complete.py`) with a clean, intuitive, step-by-step workflow. It includes tabbed interfaces for real-time optimization feedback and professional PDF cutting layouts designed with pastel colors and minimal cutting lines, matching industry standards.

### Backend
The core of OptiWise is built with Python, utilizing `pandas` for robust CSV/Excel data handling. The optimization algorithms, primarily within `optimization_core_fixed.py`, include standard greedy, global optimization with offcut pooling, and enhanced mathematical optimization with maximal rectangles and material upgrading. Comprehensive Excel reports (10 tabs) and PDF cutting layouts are generated for output.

### Data Models
Key data models include `MaterialDetails` for laminate and core material specifications, `Part` for individual cut pieces, `Board` for raw materials, and `Offcut` for tracking usable leftover material. These models support intelligent material upgrading logic and compatibility rules (e.g., MDF→HDHMR).

### Core Features
- **Material Upgrade System**: Intelligent core material upgrading based on compatibility matrices.
- **Professional Reports**: Generation of shop-floor-ready PDF cutting layouts and comprehensive 10-tab Excel reports (Optimised Cutlist, Core Material Report, Laminate Report, Edge Band Summary).
- **Edge Band Analysis**: Automated perimeter calculation grouped by top laminate finish.
- **Half-Board Optimization**: Consolidates low-utilization boards to improve material efficiency.
- **Robust Data Handling**: Comprehensive input validation for CSV files (parts list, core materials, laminates) and robust error handling.
- **Guillotine Constraints**: All optimization algorithms enforce strict guillotine cutting patterns, preventing non-manufacturable layouts.
- **Material Segregation**: Ensures each board contains only one material type for manufacturing compliance.
- **Rotation Optimization**: Supports intelligent rotation for non-grain sensitive parts to enhance material utilization.

## External Dependencies

### Python Libraries
- `streamlit`: For the web application framework.
- `pandas`: For data processing and CSV/Excel file handling.
- `matplotlib`: Used for visualization and PDF report generation.
- `openpyxl`: For processing Excel files.
- `numpy`: Used via pandas for mathematical operations.
- `scikit-learn`: For machine learning support in advanced algorithms (e.g., K-means clustering).
- `OR-Tools`: For exact trim solving in specific optimization algorithms.

### System Requirements
- Python 3.11+
- Streamlit server