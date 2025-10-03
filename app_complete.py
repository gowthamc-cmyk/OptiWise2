"""
OptiWise - Smart Beam Saw Optimization Tool (Complete Working Version)
Streamlit web application for 2D guillotine cutting optimization with material upgrading.
"""

import streamlit as st
import io
import os
import tempfile
import zipfile
from typing import Dict, List, Tuple, Optional
import logging

# Import our modules with standalone parsers
from data_models import MaterialDetails, Part, Offcut, Board
from parsers_csv_standalone import load_parts_data, load_core_materials_config, load_laminates_config, validate_data_consistency
from optimization_core_fixed import run_optimization
from optimization_global import run_global_optimization
from optimization_unified import run_unified_optimization
from simple_reports import (generate_cutting_layout_text, generate_optimized_cutlist_csv, 
                           generate_material_summary_csv, create_comprehensive_report_package)
from pdf_layout_generator import generate_cutting_layout_pdf
from utils import (setup_logging, validate_file_upload, format_currency, format_area, 
                  format_percentage, display_optimization_metrics, display_board_summary,
                  display_error_summary, get_material_options, validate_upgrade_sequence)

# Page configuration
st.set_page_config(
    page_title="OptiWise - Smart Beam Saw Optimization",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

def create_sample_data():
    """Create sample data files for testing."""
    
    # Sample cutlist data in new format
    sample_cutlist_data = """CLIENT NAME,ORDER ID / UNIQUE CODE,SL NO.,ROOM TYPE,SUB CATEGORY,TYPE,PANEL NAME,FULL NAME DESCRIPTION,QTY,GROOVE,CUT LENGTH,CUT WIDTH,FINISHED THICKNESS,MATERIAL TYPE,EB1,EB2,EB3,EB4,GRAINS,REMARKS
ABC Furniture,ABC001-P001,1,Kitchen,Base Cabinet,Panel,Base Panel 1,Kitchen Base Cabinet Door Panel,1,,600,400,18,2614 SF_18HDHMR_2614 SF,2614 SF,2614 SF,2614 SF,2614 SF,0,
ABC Furniture,ABC001-P002,2,Kitchen,Base Cabinet,Panel,Base Panel 2,Kitchen Base Cabinet Side Panel,2,,800,600,18,2614 SF_18HDHMR_2614 SF,2614 SF,2614 SF,2614 SF,2614 SF,1,Grain Direction Critical
ABC Furniture,ABC001-P003,3,Kitchen,Wall Cabinet,Panel,Wall Panel 1,Kitchen Wall Cabinet Door Panel,1,,500,350,18,5584 SGL_18HDHMR_2614 SF,5584 SGL,5584 SGL,5584 SGL,5584 SGL,0,
ABC Furniture,ABC001-P004,4,Kitchen,Wall Cabinet,Panel,Wall Panel 2,Kitchen Wall Cabinet Shelf,4,,480,300,18,5584 SGL_18HDHMR_2614 SF,5584 SGL,5584 SGL,5584 SGL,5584 SGL,0,
ABC Furniture,ABC001-P005,5,Bedroom,Wardrobe,Panel,Wardrobe Door,Bedroom Wardrobe Door Panel,2,,1800,600,18,362 SUD_18HDHMR_2614 SF,362 SUD,362 SUD,362 SUD,362 SUD,1,Large Panel - Handle with Care
ABC Furniture,ABC001-P006,6,Bedroom,Wardrobe,Panel,Wardrobe Shelf,Bedroom Wardrobe Internal Shelf,6,,1750,400,18,362 SUD_18HDHMR_2614 SF,362 SUD,362 SUD,362 SUD,362 SUD,0,"""
    
    # Sample core materials data
    sample_core_data = """Core Name,Standard Length (mm),Standard Width (mm),Thickness (mm),Price per SqM,Grade Level
18MR,2440,1220,18,850,1
18BWR,2440,1220,18,950,2
18HDHMR,2440,1220,18,1050,3"""
    
    # Sample laminates data
    sample_laminates_data = """Laminate Name,Price per SqM
SF,120
2614 SF,150"""
    
    return sample_cutlist_data, sample_core_data, sample_laminates_data

def safe_file_read(uploaded_file):
    """Safely read uploaded file content with proper encoding detection."""
    try:
        if hasattr(uploaded_file, 'read'):
            content = uploaded_file.read()
            if isinstance(content, bytes):
                for encoding in ['utf-8', 'latin-1', 'cp1252']:
                    try:
                        return content.decode(encoding)
                    except UnicodeDecodeError:
                        continue
                return content.decode('utf-8', errors='replace')
            else:
                return content
        else:
            return str(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def process_csv_data(parts_text, core_text, laminate_text):
    """Process CSV data from text input, handling both comma and tab separated formats."""
    
    def normalize_csv_format(text):
        """Convert tab-separated or inconsistent CSV to proper comma-separated format."""
        if not text:
            return text
            
        lines = text.strip().split('\n')
        normalized_lines = []
        
        for line in lines:
            if '\t' in line:
                normalized_line = line.replace('\t', ',')
            else:
                normalized_line = line
            normalized_lines.append(normalized_line)
        
        return '\n'.join(normalized_lines)
    
    try:
        parts_text = normalize_csv_format(parts_text)
        core_text = normalize_csv_format(core_text)
        laminate_text = normalize_csv_format(laminate_text)
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as parts_file:
            parts_file.write(parts_text)
            parts_file_path = parts_file.name
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as core_file:
            core_file.write(core_text)
            core_file_path = core_file.name
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as laminate_file:
            laminate_file.write(laminate_text)
            laminate_file_path = laminate_file.name
        
        # Use the updated parsers_csv module instead of standalone
        from parsers_csv import load_parts_data as load_parts_new
        from parsers_csv_standalone import load_core_materials_config, load_laminates_config
        
        parts_list = load_parts_new(parts_file_path)
        parts_errors = []  # Updated parser returns just the list
        core_db, core_errors = load_core_materials_config(core_file_path)
        laminate_db, laminate_errors = load_laminates_config(laminate_file_path)
        
        # Collect all loading errors
        all_errors = parts_errors + core_errors + laminate_errors
        if all_errors:
            return None, None, None, all_errors
        
        # Filter parts to only include those with known materials
        from parsers_csv_standalone import filter_parts_with_known_materials
        filtered_parts, skipped_parts = filter_parts_with_known_materials(parts_list, core_db, laminate_db)
        
        # Update parts_list to only include valid parts
        parts_list = filtered_parts
        
        # Create informational messages about skipped parts
        skip_messages = []
        if skipped_parts:
            skip_messages.append(f"Skipped {len(skipped_parts)} parts with unknown materials:")
            skip_messages.extend([f"  • {part}" for part in skipped_parts[:10]])  # Show first 10
            if len(skipped_parts) > 10:
                skip_messages.append(f"  • ... and {len(skipped_parts) - 10} more")
        
        # Validate remaining parts
        is_valid, error_messages = validate_data_consistency(parts_list, core_db, laminate_db)
        
        # Add skip messages to errors for display
        all_messages = skip_messages + error_messages
        
        os.unlink(parts_file_path)
        os.unlink(core_file_path)
        os.unlink(laminate_file_path)
        
        if is_valid and parts_list:  # Ensure we have parts after filtering
            return parts_list, core_db, laminate_db, True, all_messages
        else:
            return [], {}, {}, False, all_messages
            
    except Exception as e:
        return [], {}, {}, False, [f"Error processing CSV data: {str(e)}"]

def process_uploaded_files(parts_file, core_file, laminate_file):
    """Process uploaded files with improved error handling."""
    try:
        parts_content = safe_file_read(parts_file)
        core_content = safe_file_read(core_file)
        laminate_content = safe_file_read(laminate_file)
        
        if parts_content and core_content and laminate_content:
            return process_csv_data(parts_content, core_content, laminate_content)
        else:
            return [], {}, {}, False, ["Failed to read uploaded files"]
    except Exception as e:
        return [], {}, {}, False, [f"Error processing files: {e}"]

def main():
    """Main application function."""
    st.title("⚡ OptiWise - Smart Beam Saw Optimization")
    st.markdown("**Intelligent 2D Cutting Optimization with Material Upgrading**")
    
    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox(
        "Choose a page:",
        ["🏠 Home", "📊 Data Input", "⚙️ Optimization", "📋 Results", "📁 Download Files", "❓ Help"]
    )
    
    # Initialize session state
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'optimization_complete' not in st.session_state:
        st.session_state.optimization_complete = False
    
    # Route to appropriate page
    if page == "🏠 Home":
        show_home_page()
    elif page == "📊 Data Input":
        show_data_input_page()
    elif page == "⚙️ Optimization":
        show_optimization_page()
    elif page == "📋 Results":
        show_results_page()
    elif page == "📁 Download Files":
        show_download_page()
    elif page == "❓ Help":
        show_help_page()

def show_home_page():
    """Display the home page with tool overview."""
    st.header("Welcome to OptiWise")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### 🎯 What is OptiWise?
        
        OptiWise is an advanced beam saw optimization tool designed for furniture manufacturers and woodworking professionals. 
        It uses intelligent algorithms to:
        
        - **Optimize Material Usage**: Minimize waste through smart cutting layouts
        - **Reduce Costs**: Intelligent material upgrading when cost-effective
        - **Save Time**: Automated optimization with multiple algorithm options
        - **Generate Reports**: Professional cutting layouts and material reports
        
        ### 🚀 Key Features
        
        - **Multi-Algorithm Optimization**: Choose from Fast, Balanced, Maximum Efficiency, Mathematical, or TEST approaches
        - **Material Upgrading**: Automatically upgrade to higher-grade materials when it reduces waste
        - **Guillotine Cutting**: Respects real-world beam saw constraints
        - **Grain Sensitivity**: Handles parts that cannot be rotated
        - **Professional Reports**: Cutting layouts and optimization reports
        - **Advanced Algorithms**: Includes Enhanced Global Optimization and Mathematical Optimization
        """)
    
    with col2:
        st.info("""
        **Quick Start:**
        
        1. Go to 📊 **Data Input**
        2. Upload your files or use samples
        3. Navigate to ⚙️ **Optimization**
        4. Choose optimization strategy
        5. View results in 📋 **Results**
        """)
        
        # Display sample data download
        st.subheader("📥 Sample Data")
        sample_cutlist, sample_core, sample_laminates = create_sample_data()
        
        st.download_button(
            "Download Sample Cutlist",
            sample_cutlist,
            "sample_cutlist.csv",
            "text/csv"
        )
        
        st.download_button(
            "Download Sample Core Materials",
            sample_core,
            "sample_core_materials.csv",
            "text/csv"
        )
        
        st.download_button(
            "Download Sample Laminates",
            sample_laminates,
            "sample_laminates.csv",
            "text/csv"
        )

def show_data_input_page():
    """Display the data input page."""
    st.header("📊 Data Input")
    st.markdown("Upload your cutting data and material specifications")
    
    # Create tabs for different input methods
    tab1, tab2 = st.tabs(["📋 Text Input", "📎 File Upload"])
    
    with tab1:
        st.subheader("Paste Your Data")
        st.markdown("Copy and paste your data directly. Supports both comma and tab-separated formats.")
        
        sample_cutlist, sample_core, sample_laminates = create_sample_data()
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("Cutlist Data")
            parts_text = st.text_area(
                "Parts List (CSV format)",
                value="",
                height=300,
                placeholder=sample_cutlist,
                help="Format: CLIENT NAME, ORDER ID / UNIQUE CODE, SL NO., ROOM TYPE, SUB CATEGORY, TYPE, PANEL NAME, FULL NAME DESCRIPTION, QTY, GROOVE, CUT LENGTH, CUT WIDTH, FINISHED THICKNESS, MATERIAL TYPE, EB1, EB2, EB3, EB4, GRAINS, REMARKS"
            )
        
        with col2:
            st.subheader("Core Materials")
            core_text = st.text_area(
                "Core Materials (CSV format)",
                value="",
                height=300,
                placeholder=sample_core,
                help="Format: Core Name, Standard Length (mm), Standard Width (mm), Thickness (mm), Price per SqM, Grade Level"
            )
        
        with col3:
            st.subheader("Laminates")
            laminate_text = st.text_area(
                "Laminates (CSV format)",
                value="",
                height=300,
                placeholder=sample_laminates,
                help="Format: Laminate Name, Price per SqM"
            )
        
        # Load sample data button
        if st.button("📥 Load Sample Data", type="secondary"):
            st.session_state.sample_loaded = True
            st.rerun()
        
        # Check if sample data should be loaded
        if hasattr(st.session_state, 'sample_loaded') and st.session_state.sample_loaded:
            parts_text = sample_cutlist
            core_text = sample_core
            laminate_text = sample_laminates
            st.session_state.sample_loaded = False
        
        # Process data button
        if st.button("🔄 Process Data", type="primary"):
            if parts_text and core_text and laminate_text:
                st.write("Processing CSV text input...")
                # Debug the parts text format
                parts_lines = parts_text.strip().split('\n')
                if parts_lines:
                    st.write(f"Headers detected: {parts_lines[0]}")
                    if len(parts_lines) > 1:
                        st.write(f"Sample data: {parts_lines[1]}")
                
                with st.spinner("Processing data..."):
                    parts_list, core_db, laminate_db, success, error_messages = process_csv_data(
                        parts_text, core_text, laminate_text
                    )
                    
                    if success:
                        st.session_state.parts_list = parts_list
                        st.session_state.core_db = core_db
                        st.session_state.laminate_db = laminate_db
                        st.session_state.data_loaded = True
                        
                        # Quick Panel Summary
                        st.subheader("📋 Panel Summary")
                        
                        # Calculate total area needed
                        total_area_sqft = 0
                        grain_sensitive_count = 0
                        for part in parts_list:
                            try:
                                part_area_sqft = (part.requested_length * part.requested_width / 1_000_000) * 10.764
                                total_area_sqft += part_area_sqft
                                if hasattr(part, 'grains') and part.grains == 1:
                                    grain_sensitive_count += 1
                            except AttributeError:
                                continue
                        
                        # Quick metrics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Parts", len(parts_list))
                        with col2:
                            st.metric("Total Area", f"{total_area_sqft:.1f} sqft")
                        with col3:
                            st.metric("Grain Sensitive", grain_sensitive_count)
                        with col4:
                            try:
                                material_types = len(set(str(getattr(p, 'material_details', 'Unknown')) for p in parts_list))
                                st.metric("Material Types", material_types)
                            except:
                                st.metric("Material Types", "N/A")
                        
                        # Material distribution summary
                        material_counts = {}
                        for part in parts_list:
                            try:
                                material_key = str(getattr(part, 'material_details', 'Unknown'))
                                material_counts[material_key] = material_counts.get(material_key, 0) + 1
                            except:
                                continue
                        
                        with st.expander("📊 Material Distribution Details"):
                            for material, count in sorted(material_counts.items()):
                                try:
                                    material_area = sum((p.requested_length * p.requested_width / 1_000_000) * 10.764 
                                                      for p in parts_list 
                                                      if hasattr(p, 'material_details') and str(p.material_details) == material)
                                    st.write(f"• **{material}**: {count} parts ({material_area:.1f} sqft)")
                                except:
                                    st.write(f"• **{material}**: {count} parts")
                        
                        st.success(f"Successfully loaded {len(parts_list)} parts, {len(core_db)} core materials, {len(laminate_db)} laminates")
                    else:
                        st.error("Data processing failed:")
                        for error in error_messages:
                            st.error(f"• {error}")
            else:
                st.error("Please provide all three data sets.")
    
    with tab2:
        st.subheader("Upload CSV Files")
        st.markdown("Upload your CSV files directly")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            parts_file = st.file_uploader(
                "Cutlist CSV",
                type=['csv'],
                help="Upload your parts cutlist file"
            )
        
        with col2:
            core_file = st.file_uploader(
                "Core Materials CSV",
                type=['csv'],
                help="Upload your core materials file"
            )
        
        with col3:
            laminate_file = st.file_uploader(
                "Laminates CSV",
                type=['csv'],
                help="Upload your laminates file"
            )
        
        if st.button("📁 Process Uploaded Files", type="primary"):
            if parts_file and core_file and laminate_file:
                with st.spinner("Processing uploaded files..."):
                    parts_list, core_db, laminate_db, success, error_messages = process_uploaded_files(
                        parts_file, core_file, laminate_file
                    )
                    
                    if success:
                        st.session_state.parts_list = parts_list
                        st.session_state.core_db = core_db
                        st.session_state.laminate_db = laminate_db
                        st.session_state.data_loaded = True
                        
                        st.success(f"Successfully loaded {len(parts_list)} parts, {len(core_db)} core materials, {len(laminate_db)} laminates")
                        st.rerun()
                    else:
                        st.error("Data processing failed:")
                        for error in error_messages:
                            st.error(f"• {error}")
            else:
                st.error("Please upload all three required files.")
    
    # Display data preview if loaded
    if hasattr(st.session_state, 'data_loaded') and st.session_state.data_loaded:
        show_data_preview()

def show_data_preview():
    """Display preview of loaded data."""
    st.markdown("---")
    st.subheader("📋 Data Preview")
    
    parts_list = st.session_state.parts_list
    core_db = st.session_state.core_db
    laminate_db = st.session_state.laminate_db
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Parts Loaded", len(parts_list))
    with col2:
        st.metric("Core Materials", len(core_db))
    with col3:
        st.metric("Laminates", len(laminate_db))
    
    # Detailed preview in expandable sections
    with st.expander("Parts Preview"):
        if parts_list:
            preview_data = []
            for i, part in enumerate(parts_list[:10]):
                preview_data.append({
                    "ORDER ID / UNIQUE CODE": part.id,
                    "CLIENT NAME": getattr(part, 'client_name', ''),
                    "ROOM TYPE": getattr(part, 'room_type', ''),
                    "PANEL NAME": getattr(part, 'panel_name', ''),
                    "CUT LENGTH": f"{part.requested_length}mm",
                    "CUT WIDTH": f"{part.requested_width}mm",
                    "MATERIAL TYPE": str(part.material_details),
                    "GRAINS": "Yes" if part.grains == 1 else "No"
                })
            
            st.table(preview_data)
            if len(parts_list) > 10:
                st.info(f"Showing first 10 of {len(parts_list)} parts")

def show_optimization_page():
    """Display the optimization page with settings and execution."""
    st.header("⚙️ Optimization Settings")
    
    # Check if data is loaded
    if not hasattr(st.session_state, 'data_loaded') or not st.session_state.data_loaded:
        st.warning("Please load data first in the 'Data Input' section.")
        return
    
    # Order Information
    st.subheader("📋 Order Information")
    order_name = st.text_input(
        "Order Name",
        value="",
        placeholder="e.g., Kitchen Project 2025-001, Client ABC Order (leave blank to use CLIENT NAME from data)",
        help="Enter a name for this order. If left blank, CLIENT NAME from the parts data will be used. This will be used in downloaded reports and file names."
    )
    
    if order_name:
        st.session_state.order_name = order_name
    elif not hasattr(st.session_state, 'order_name'):
        # Try to get CLIENT NAME from parts data if no order name provided
        if hasattr(st.session_state, 'parts_list') and st.session_state.parts_list:
            client_name = getattr(st.session_state.parts_list[0], 'client_name', '')
            st.session_state.order_name = client_name if client_name else ""
        else:
            st.session_state.order_name = ""
    
    st.markdown("---")
    
    # Optimization settings
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Material Upgrade Settings")
        
        available_cores = get_material_options(st.session_state.core_db)
        
        st.markdown("""
        **Upgrade Sequence**: Define the order of core materials for upgrading parts when needed.
        Parts can be upgraded to higher-grade cores but never downgraded.
        """)
        
        # Display available cores with grade levels
        with st.expander("Available Core Materials (sorted by grade level)"):
            for core_name in available_cores:
                core_details = st.session_state.core_db[core_name]
                st.text(f"{core_name} (Grade {core_details['grade_level']}) - {format_currency(core_details['price_per_sqm'])}/m²")
        
        # Get selected algorithm type first
        algorithm_type = st.selectbox(
            "Optimization Strategy",
            [
                ("fast", "🔧 Upgrader - Quick greedy algorithm"),
                ("test_algorithm", "🧪 TEST - Best-Fit Decreasing with smart rearrangement (no upgrades)"),
                ("test3_algorithm", "🧪 TEST 3 - Competitive optimizer with Bottom-Left-Fill vs Shelf Packing (no upgrades)"),
                ("test4_algorithm", "🧪 TEST 4 - Master Optimizer with Skyline & Shelf Packing, Edge-Fit & Agent Consolidation (no upgrades)"),
                ("test5_algorithm", "🧪 TEST 5 - AMBP (Adaptive Multi-stage Bucket-Strip Pipeline) with guillotine constraints (no upgrades)"),
                ("test5_duplicate", "🧪 TEST 5 (duplicate) - AMBP with Off-cut Optimization - Post-optimization rearrangement for low-utilization boards (no upgrades)"),
                ("max_utilisation", "🎯 Max Utilisation - Smart Rearrangement - Rearranges parts on boards <50% utilization to create 50%+ offcuts and reports as 0.5 materials (no upgrades)"),
                ("no_upgrade", "📋 Standard - No Upgradation - Use exact material specifications")
            ],
            format_func=lambda x: x[1],
            index=0,  # Default to fast
            help="Choose optimization approach based on your needs and time constraints"
        )
        
        # Show upgrade sequence only for algorithms that use it
        show_upgrade_sequence = algorithm_type[0] not in ["no_upgrade", "test_algorithm", "test3_algorithm", "test4_algorithm", "test5_algorithm", "test5_duplicate", "max_utilisation"]
        
        if show_upgrade_sequence:
            upgrade_sequence = st.text_input(
                "Upgrade Sequence (comma-separated)",
                value=", ".join(available_cores),
                help="Enter core material names in order of preference for upgrades"
            )
            
            is_valid, error_msg, cleaned_sequence = validate_upgrade_sequence(upgrade_sequence, st.session_state.core_db)
            
            if not is_valid:
                st.error(f"Invalid upgrade sequence: {error_msg}")
            else:
                if cleaned_sequence:
                    st.success(f"Valid upgrade sequence: {cleaned_sequence}")
        else:
            upgrade_sequence = ""
            is_valid = True
            cleaned_sequence = ""
            if algorithm_type[0] == "no_upgrade":
                st.info("📋 Standard mode uses exact material specifications without upgrades.")
            elif algorithm_type[0] == "test_algorithm":
                st.info("🧪 TEST algorithm uses Best-Fit Decreasing placement with intelligent board rearrangement to minimize waste - no material upgrades.")
            elif algorithm_type[0] == "test5_duplicate":
                st.info("🧪 TEST 5 (duplicate) algorithm uses AMBP with Off-cut Optimization - rearranges parts on boards with <60% utilization to maximize largest single off-cut area - no material upgrades.")
            elif algorithm_type[0] == "max_utilisation":
                st.info("🎯 Max Utilisation algorithm maximizes board utilization with intelligent rearrangement for <40% boards to create ≥50% offcuts - only reports 0.5 quantity if successful.")
            elif algorithm_type[0] == "test3_algorithm":
                st.info("🧪 TEST 3 algorithm uses advanced global offcut reuse with Tight Edge-Fit placement for narrow strips, Last-Fit Repacking, and 2-pass packing optimization - no material upgrades.")
            elif algorithm_type[0] == "test4_algorithm":
                st.info("🧪 TEST 4 algorithm uses Master Optimizer with competitive Skyline vs Shelf packing strategies, followed by tight edge-fit placement and agent-based board consolidation - no material upgrades.")
            elif algorithm_type[0] == "fast":
                st.info("🔧 Upgrader algorithm uses quick greedy optimization for rapid results with material upgrades.")
    
    with col2:
        st.subheader("Optimization Parameters")
        
        multi_objective = st.checkbox(
            "Multi-Objective Optimization",
            value=False,
            help="Compare multiple strategies and automatically select the best overall solution"
        )
        
        kerf = st.number_input(
            "Kerf Width (mm)",
            min_value=0.0,
            max_value=10.0,
            value=4.4,
            step=0.1,
            help="Saw blade width for cutting calculations"
        )
        
        # Enhanced AMBP parameters for TEST 5, TEST 5 duplicate, and Max Utilisation
        if algorithm_type[0] in ["test5_algorithm", "test5_duplicate", "max_utilisation"]:
            st.markdown("#### AMBP Algorithm Parameters")
            
            col1, col2 = st.columns(2)
            with col1:
                bucket_slack = st.number_input(
                    "Bucket Slack",
                    min_value=0.0,
                    max_value=1.0,
                    value=0.1,
                    step=0.05,
                    help="Tolerance for width clustering (0.1 = 10%)"
                )
                
                ruin_fraction = st.number_input(
                    "Ruin Fraction",
                    min_value=0.1,
                    max_value=0.5,
                    value=0.2,
                    step=0.05,
                    help="Fraction of worst bins to ruin and recreate"
                )
            
            with col2:
                utilisation_floor = st.number_input(
                    "Utilisation Floor",
                    min_value=0.5,
                    max_value=0.95,
                    value=0.80,
                    step=0.05,
                    help="Minimum acceptable board utilisation"
                )
                
                thin_strip_width = st.number_input(
                    "Thin Strip Width (mm)",
                    min_value=50.0,
                    max_value=200.0,
                    value=130.0,
                    step=10.0,
                    help="Width threshold for thin strip classification"
                )
        else:
            # Default values for non-AMBP algorithms
            bucket_slack = 0.1
            ruin_fraction = 0.2
            utilisation_floor = 0.80
            thin_strip_width = 130.0
        
        # Display current data summary
        st.subheader("Data Summary")
        parts_count = len(st.session_state.parts_list)
        total_area = sum(p.requested_length * p.requested_width for p in st.session_state.parts_list) / 1_000_000
        
        st.metric("Total Parts", parts_count)
        st.metric("Total Area", f"{total_area:.2f} m²")
        st.metric("Materials", len(set(str(p.material_details) for p in st.session_state.parts_list)))
    
    # Run optimization
    st.markdown("---")
    
    if st.button("🚀 Run Optimization", type="primary", disabled=not is_valid):
        if is_valid:
            try:
                with st.spinner("Running optimization..."):
                    strategy_key = algorithm_type[0]
                    
                    if algorithm_type[0] == "no_upgrade":
                        final_upgrade_sequence = ""
                    else:
                        final_upgrade_sequence = cleaned_sequence
                    
                    # Run optimization based on selected algorithm
                    if algorithm_type[0] == "test_algorithm":
                        # Import and run simple TEST algorithm
                        from optimization_test_simple import run_test_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test_optimization(
                            parts_list=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "test2_algorithm":
                        # Import and run TEST 2 algorithm
                        from optimization_test2 import run_test2_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test2_optimization(
                            parts_list=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "test3_algorithm":
                        # Import and run TEST 3 algorithm
                        from optimization_test3 import run_test3_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test3_optimization(
                            parts_list=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "test4_algorithm":
                        # Import and run TEST 4 algorithm
                        from optimization_test4 import run_test4_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test4_optimization(
                            parts=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "test5_algorithm":
                        # Import and run TEST 5 algorithm with guillotine constraints and material segregation
                        from optimization_test5 import run_test5_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test5_optimization(
                            parts=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "test5_duplicate":
                        # Import and run TEST 5 (duplicate) algorithm with enhanced AMBP implementation
                        from optimization_test5_duplicate import run_test5_duplicate_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_test5_duplicate_optimization(
                            parts=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )
                    elif algorithm_type[0] == "max_utilisation":
                        # Import and run Max Utilisation algorithm 
                        from optimization_max_utilisation import run_max_utilisation_optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_max_utilisation_optimization(
                            parts=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            kerf=kerf
                        )

                    else:
                        # Run standard unified optimization
                        boards, unplaced_parts, upgrade_summary, initial_cost, final_cost = run_unified_optimization(
                            parts_list=st.session_state.parts_list,
                            core_db=st.session_state.core_db,
                            laminate_db=st.session_state.laminate_db,
                            user_upgrade_sequence_str=final_upgrade_sequence,
                            kerf=kerf,
                            strategy=strategy_key,
                            multi_objective=multi_objective
                        )
                    
                    # Store results
                    st.session_state.optimization_results = {
                        'boards': boards,
                        'unplaced_parts': unplaced_parts,
                        'upgrade_summary': upgrade_summary,
                        'initial_cost': initial_cost,
                        'final_cost': final_cost,
                        'kerf': kerf,
                        'strategy_used': strategy_key,
                        'multi_objective': multi_objective
                    }
                    st.session_state.optimization_complete = True
                    
                    if algorithm_type[0] in ["test_algorithm", "test3_algorithm", "test4_algorithm", "test5_algorithm", "test5_duplicate", "max_utilisation"]:
                        st.success(f"✅ Optimization complete using {algorithm_type[1]} (no material upgrades)!")
                    else:
                        st.success(f"✅ Optimization complete using {algorithm_type[1]}!")
                    
                    # Quick Results Summary
                    st.subheader("🎯 Quick Results Summary")
                    
                    # Calculate efficiency metrics
                    total_board_area = sum(board.total_length * board.total_width for board in boards)
                    total_board_area_sqft = total_board_area / 1_000_000 * 10.764
                    total_waste_area = sum(board.get_remaining_area() for board in boards)
                    total_waste_area_sqft = total_waste_area / 1_000_000 * 10.764
                    waste_percentage = (total_waste_area / total_board_area * 100) if total_board_area > 0 else 0
                    avg_utilization = sum(board.get_utilization_percentage() for board in boards) / len(boards) if boards else 0
                    cost_savings = initial_cost - final_cost
                    
                    # Main metrics
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("Total Boards", len(boards))
                    with col2:
                        st.metric("Board Area", f"{total_board_area_sqft:.1f} sqft")
                    with col3:
                        st.metric("Waste Area", f"{total_waste_area_sqft:.1f} sqft")
                    with col4:
                        st.metric("Utilization", f"{avg_utilization:.1f}%")
                    with col5:
                        st.metric("Cost Savings", f"₹{cost_savings:.2f}")
                    
                    # Additional info
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        if unplaced_parts:
                            st.error(f"⚠️ {len(unplaced_parts)} parts could not be placed")
                        else:
                            st.success("✅ All parts successfully placed")
                    with col2:
                        if upgrade_summary:
                            st.info(f"🔄 {len(upgrade_summary)} material upgrades applied")
                        else:
                            st.info("No material upgrades needed")
                    with col3:
                        st.info(f"💰 Final cost: ₹{final_cost:.2f}")
                    
                    # Generate reports automatically
                    with st.spinner("Generating reports..."):
                        try:
                            # Generate comprehensive report package with detailed Excel
                            reports = create_comprehensive_excel_report(
                                boards, unplaced_parts, upgrade_summary,
                                initial_cost, final_cost, st.session_state.get('order_name', ''),
                                st.session_state.get('core_db', {}),
                                st.session_state.get('laminate_db', {})
                            )
                            st.session_state.latest_reports = reports
                            st.success("Reports generated successfully!")
                        except Exception as e:
                            st.warning(f"Report generation failed: {e}")
                    
                    st.info("Go to 'Results' page to view detailed optimization results and download reports")
                    
            except Exception as e:
                st.error(f"Optimization failed: {e}")
                logging.error(f"Optimization error: {e}")

def create_comprehensive_excel_report(boards, unplaced_parts, upgrade_summary, initial_cost, final_cost, 
                                    order_name, core_db, laminate_db):
    """Create comprehensive Excel report with multiple tabs for detailed analysis."""
    reports = {}
    
    try:
        # Use the Excel generator from report_generators
        from report_generators import generate_optimized_cutlist_excel
        
        excel_content = generate_optimized_cutlist_excel(
            boards, unplaced_parts, upgrade_summary,
            initial_cost, final_cost, core_db, laminate_db,
            order_name=order_name
        )
        
        reports['optimization_report.xlsx'] = excel_content
        
        # Generate cutting layouts - PDF or text fallback
        try:
            from pdf_layout_generator import generate_cutting_layout_pdf
            pdf_content = generate_cutting_layout_pdf(boards, order_name)
            reports['cutting_layouts.pdf'] = pdf_content
            logging.info("PDF cutting layouts generated successfully")
        except Exception as e:
            logging.warning(f"PDF generation failed: {e}, using text fallback")
            try:
                reports['cutting_layout.txt'] = generate_cutting_layout_text(boards, order_name)
                logging.info("Text cutting layout generated as fallback")
            except Exception as e2:
                logging.error(f"Text fallback also failed: {e2}")
                # Create minimal text report as last resort
                minimal_text = f"OptiWise Report - {order_name}\n"
                minimal_text += f"Generated {len(boards)} optimized boards\n"
                minimal_text += f"Please check logs for detailed information."
                reports['cutting_layout.txt'] = minimal_text
        
    except Exception as e:
        logging.error(f"Excel report generation failed: {e}")
        # Ensure basic reports are always available
        try:
            reports = create_comprehensive_report_package(
                boards, unplaced_parts, upgrade_summary, 
                initial_cost, final_cost, order_name
            )
        except Exception as e2:
            logging.error(f"Fallback report generation failed: {e2}")
            # Last resort: create minimal reports
            reports = {
                'cutting_layout.txt': f"OptiWise Report - {order_name}\nGenerated {len(boards)} boards\nSee Results page for details.",
                'optimization_summary.txt': f"Optimization Summary\nBoards: {len(boards)}\nParts: {len([p for b in boards for p in b.parts_on_board])}"
            }
    
    return reports

# Old Excel functions removed - now using robust excel_generator_fixed.py

def create_simple_pdf_content(boards, order_name):
    """Create simple PDF-style content as text."""
    content = []
    content.append(f"OptiWise Cutting Layout Report - {order_name}")
    content.append("=" * 60)
    content.append("")
    
    for i, board in enumerate(boards, 1):
        content.append(f"Board {i}: {board.id}")
        content.append(f"Material: {board.material_details}")
        content.append(f"Size: {board.total_length}mm x {board.total_width}mm")
        content.append(f"Utilization: {board.get_utilization_percentage():.1f}%")
        content.append("")
        
        if board.parts_on_board:
            content.append("Parts on this board:")
            for j, part in enumerate(board.parts_on_board, 1):
                x_pos = getattr(part, 'x_pos', 0) or 0
                y_pos = getattr(part, 'y_pos', 0) or 0
                content.append(f"  {j}. {part.id} - {part.requested_length}x{part.requested_width}mm @ ({x_pos},{y_pos})")
        content.append("-" * 40)
        content.append("")
    
    return "\n".join(content)

def calculate_material_wise_summary(boards):
    """Calculate comprehensive material-wise summary with wastage."""
    material_summary = {}
    
    for board in boards:
        material_key = str(board.material_details)
        
        if material_key not in material_summary:
            material_summary[material_key] = {
                'board_count': 0,
                'total_board_area': 0,
                'utilized_area': 0,
                'wastage_area': 0,
                'parts_placed': 0,
                'boards': []
            }
        
        # Calculate areas
        board_area = board.total_length * board.total_width / 1_000_000
        utilized_area = (board_area * board.get_utilization_percentage() / 100)
        wastage_area = board_area - utilized_area
        
        # Update summary
        summary = material_summary[material_key]
        summary['board_count'] += 1
        summary['total_board_area'] += board_area
        summary['utilized_area'] += utilized_area
        summary['wastage_area'] += wastage_area
        summary['parts_placed'] += len(board.parts_on_board)
        summary['boards'].append(board)
    
    return material_summary

def generate_core_material_report_data(boards, core_db):
    """Generate core material report data for Streamlit display matching Excel format."""
    
    core_data = {}
    
    for board in boards:
        # Extract core material from board material details with enhanced extraction
        core_material = 'Unknown'
        if hasattr(board, 'material_details') and board.material_details:
            # Method 1: Direct core material attribute
            if hasattr(board.material_details, 'core_material'):
                core_material = board.material_details.core_material
            elif hasattr(board.material_details, 'core_name'):
                core_material = board.material_details.core_name
            # Method 2: Extract from full material string
            elif hasattr(board.material_details, 'full_material_string'):
                material_string = board.material_details.full_material_string
                if '_' in material_string:
                    parts = material_string.split('_')
                    for part in parts:
                        if any(core in part.upper() for core in ['MDF', 'MR', 'BWR', 'HDHMR', 'BWP']):
                            core_material = part
                            break
        
        if core_material not in core_data:
            core_data[core_material] = {
                'board_count': 0,
                'standard_area': 0,
                'utilized_area': 0,
                'total_cost': 0
            }
        
        # Calculate areas using correct board dimensions and sqft conversion
        board_length = getattr(board, 'total_length', 0)
        board_width = getattr(board, 'total_width', 0)
        standard_area_sqft = (board_length * board_width) / 92903.04  # Convert mm² to sqft
        
        utilized_area_sqft = 0
        if hasattr(board, 'parts_on_board') and board.parts_on_board:
            for part in board.parts_on_board:
                part_length = getattr(part, 'requested_length', 0)
                part_width = getattr(part, 'requested_width', 0)
                utilized_area_sqft += (part_length * part_width) / 92903.04  # Convert mm² to sqft
        
        # Get pricing - use correct database key
        unit_price_per_sqft = 0
        for core_name, core_info in core_db.items():
            if core_name == core_material or core_name in core_material:
                if isinstance(core_info, dict):
                    # Convert ₹/m² to ₹/sqft
                    unit_price_per_sqft = float(core_info.get('price_per_sqm', 0)) / 10.764
                else:
                    unit_price_per_sqft = float(core_info) / 10.764
                break
        
        core_data[core_material]['board_count'] += 1
        core_data[core_material]['standard_area'] += standard_area_sqft
        core_data[core_material]['utilized_area'] += utilized_area_sqft
        core_data[core_material]['total_cost'] += standard_area_sqft * unit_price_per_sqft
    
    # Convert to list of dictionaries for Streamlit matching Excel format
    report_data = []
    for core_material, data in core_data.items():
        wastage_area = data['standard_area'] - data['utilized_area']
        utilization_pct = (data['utilized_area'] / data['standard_area'] * 100) if data['standard_area'] > 0 else 0
        wastage_pct = 100 - utilization_pct
        
        # Get unit price for display
        unit_price_display = 0
        for core_name, core_info in core_db.items():
            if core_name == core_material or core_name in core_material:
                if isinstance(core_info, dict):
                    unit_price_display = float(core_info.get('price_per_sqm', 0)) / 10.764
                else:
                    unit_price_display = float(core_info) / 10.764
                break
        
        report_data.append({
            'Core Material': core_material,
            'Board Count': data['board_count'],
            'Standard Area (sqft)': round(data['standard_area'], 2),
            'Utilized Area (sqft)': round(data['utilized_area'], 2),
            'Wastage Area (sqft)': round(wastage_area, 2),
            'Utilization %': f"{utilization_pct:.1f}%",
            'Wastage %': f"{wastage_pct:.1f}%",
            'Unit Price (₹/sqft)': f"₹{unit_price_display:.2f}",
            'Total Cost (₹)': f"₹{data['total_cost']:.2f}"
        })
    
    return report_data if report_data else None

def generate_laminate_report_data(boards, laminate_db):
    """Generate laminate report data for Streamlit display matching Excel format."""
    
    laminate_data = {}
    
    for board in boards:
        # Extract laminate types from board material string
        laminate_types = []
        if hasattr(board, 'material_details') and board.material_details:
            if hasattr(board.material_details, 'full_material_string'):
                material_string = board.material_details.full_material_string
                if '_' in material_string:
                    parts = material_string.split('_')
                    if len(parts) >= 3:  # Format: top_core_bottom
                        top_laminate = parts[0]
                        bottom_laminate = parts[2]
                        laminate_types = [top_laminate, bottom_laminate]
            
            # Fallback to direct attributes
            if not laminate_types:
                if hasattr(board.material_details, 'top_laminate_name'):
                    top_laminate = board.material_details.top_laminate_name
                    bottom_laminate = getattr(board.material_details, 'bottom_laminate_name', top_laminate)
                    laminate_types = [top_laminate, bottom_laminate]
                elif hasattr(board.material_details, 'laminate_material'):
                    laminate_type = board.material_details.laminate_material
                    laminate_types = [laminate_type, laminate_type]  # Assume same for top/bottom
        
        # If no laminate types found, skip this board
        if not laminate_types:
            continue
            
        # Process each laminate type separately (top and bottom counted separately)
        for laminate_type in laminate_types:
            if laminate_type not in laminate_data:
                laminate_data[laminate_type] = {
                    'board_count': 0,
                    'standard_area': 0,
                    'utilized_area': 0,
                    'total_cost': 0
                }
            
            # Calculate areas using correct board dimensions and sqft conversion
            board_length = getattr(board, 'total_length', 0)
            board_width = getattr(board, 'total_width', 0)
            standard_area_sqft = (board_length * board_width) / 92903.04  # Convert mm² to sqft
            
            utilized_area_sqft = 0
            if hasattr(board, 'parts_on_board') and board.parts_on_board:
                for part in board.parts_on_board:
                    part_length = getattr(part, 'requested_length', 0)
                    part_width = getattr(part, 'requested_width', 0)
                    utilized_area_sqft += (part_length * part_width) / 92903.04  # Convert mm² to sqft
            
            # Get pricing - find best match in laminate_db and convert to ₹/sqft
            unit_price_per_sqft = 0
            for laminate_name, laminate_price in laminate_db.items():
                if laminate_name == laminate_type or laminate_name in laminate_type:
                    unit_price_per_sqft = float(laminate_price) / 10.764  # Convert ₹/m² to ₹/sqft
                    break
            
            # Each laminate face is counted separately (no doubling here as each is processed individually)
            laminate_data[laminate_type]['board_count'] += 1
            laminate_data[laminate_type]['standard_area'] += standard_area_sqft
            laminate_data[laminate_type]['utilized_area'] += utilized_area_sqft
            laminate_data[laminate_type]['total_cost'] += standard_area_sqft * unit_price_per_sqft
    
    # Convert to list of dictionaries for Streamlit matching Excel format
    report_data = []
    for laminate_type, data in laminate_data.items():
        wastage_area = data['standard_area'] - data['utilized_area']
        utilization_pct = (data['utilized_area'] / data['standard_area'] * 100) if data['standard_area'] > 0 else 0
        wastage_pct = 100 - utilization_pct
        
        # Get unit price for display
        unit_price_display = 0
        if laminate_type in laminate_db:
            unit_price_display = float(laminate_db[laminate_type]) / 10.764  # ₹/m² to ₹/sqft
        
        report_data.append({
            'Laminate Type': laminate_type,
            'Board Count': data['board_count'],
            'Standard Area (sqft)': round(data['standard_area'], 2),
            'Utilized Area (sqft)': round(data['utilized_area'], 2),
            'Wastage Area (sqft)': round(wastage_area, 2),
            'Utilization %': f"{utilization_pct:.1f}%",
            'Wastage %': f"{wastage_pct:.1f}%",
            'Unit Price (₹/sqft)': f"₹{unit_price_display:.2f}",
            'Total Cost (₹)': f"₹{data['total_cost']:.2f}"
        })
    
    return report_data if report_data else None

def generate_edge_band_report_data(boards):
    """Generate edge band summary report data for Streamlit display."""
    
    edge_band_data = {}
    
    for board in boards:
        if not hasattr(board, 'parts_on_board') or not board.parts_on_board:
            continue
            
        for part in board.parts_on_board:
            # Get top laminate from part material details
            top_laminate = 'Unknown'
            
            if hasattr(part, 'material_details') and part.material_details:
                if hasattr(part.material_details, 'top_laminate_name'):
                    top_laminate = part.material_details.top_laminate_name
                elif hasattr(part.material_details, 'laminate_material'):
                    top_laminate = part.material_details.laminate_material
                elif hasattr(part.material_details, 'top_laminate'):
                    top_laminate = part.material_details.top_laminate
            
            # Use cut length and cut width for dimensions
            part_length = getattr(part, 'requested_length', 0)
            part_width = getattr(part, 'requested_width', 0)
            
            # Get edgeband names from EB1, EB2, EB3, EB4 columns
            if hasattr(part, 'original_data') and part.original_data:
                eb1_name = str(part.original_data.get('EB1', '')).strip()  # Length edge
                eb2_name = str(part.original_data.get('EB2', '')).strip()  # Width edge  
                eb3_name = str(part.original_data.get('EB3', '')).strip()  # Length edge
                eb4_name = str(part.original_data.get('EB4', '')).strip()  # Width edge
                
                # Process each edgeband type separately
                edgeband_entries = []
                
                # EB1 (Length edge)
                if eb1_name and eb1_name != '0' and eb1_name.lower() != 'none':
                    edgeband_entries.append((eb1_name, part_length, 'length'))
                
                # EB2 (Width edge)  
                if eb2_name and eb2_name != '0' and eb2_name.lower() != 'none':
                    edgeband_entries.append((eb2_name, part_width, 'width'))
                
                # EB3 (Length edge)
                if eb3_name and eb3_name != '0' and eb3_name.lower() != 'none':
                    edgeband_entries.append((eb3_name, part_length, 'length'))
                
                # EB4 (Width edge)
                if eb4_name and eb4_name != '0' and eb4_name.lower() != 'none':
                    edgeband_entries.append((eb4_name, part_width, 'width'))
                
                # Add each edgeband to summary
                for eb_name, dimension, edge_type in edgeband_entries:
                    if eb_name not in edge_band_data:
                        edge_band_data[eb_name] = {
                            'panel_count': 0,
                            'total_length_mm': 0,
                            'total_width_mm': 0
                        }
                    
                    edge_band_data[eb_name]['panel_count'] += 1
                    if edge_type == 'length':
                        edge_band_data[eb_name]['total_length_mm'] += dimension
                    else:
                        edge_band_data[eb_name]['total_width_mm'] += dimension
            
    
    # Convert to list of dictionaries for Streamlit
    report_data = []
    for eb_name, data in edge_band_data.items():
        # Calculate total edgeband length (length + width dimensions)
        total_length_mm = data['total_length_mm'] + data['total_width_mm']
        total_length_m = total_length_mm / 1000
        
        report_data.append({
            'Edge Band Name': eb_name,
            'Panel Count': data['panel_count'],
            'Total Length (mm)': f"{total_length_mm:.0f}",
            'Total Length (m)': f"{total_length_m:.2f}"
        })
    
    # Sort by edge band name for consistent display
    report_data.sort(key=lambda x: x['Edge Band Name'])
    
    return report_data if report_data else None

def generate_material_upgrade_report_data(boards):
    """Generate material upgrade report data for Streamlit display."""
    
    upgrade_data = {}
    
    for board in boards:
        if not hasattr(board, 'parts_on_board') or not board.parts_on_board:
            continue
            
        # Get board material details
        board_core = 'Unknown'
        board_laminate = 'Unknown'
        if hasattr(board, 'material_details') and board.material_details:
            if hasattr(board.material_details, 'core_material'):
                board_core = board.material_details.core_material
            elif hasattr(board.material_details, 'core_name'):
                board_core = board.material_details.core_name
                
            if hasattr(board.material_details, 'laminate_material'):
                board_laminate = board.material_details.laminate_material
            elif hasattr(board.material_details, 'top_laminate_name'):
                board_laminate = board.material_details.top_laminate_name
        
        board_material = f"{board_laminate}_{board_core}_{board_laminate}"
        
        for part in board.parts_on_board:
            # Get original part material requirements
            original_core = 'Unknown'
            original_laminate = 'Unknown'
            
            if hasattr(part, 'material_details') and part.material_details:
                if hasattr(part.material_details, 'core_material'):
                    original_core = part.material_details.core_material
                elif hasattr(part.material_details, 'core_name'):
                    original_core = part.material_details.core_name
                    
                if hasattr(part.material_details, 'laminate_material'):
                    original_laminate = part.material_details.laminate_material
                elif hasattr(part.material_details, 'top_laminate_name'):
                    original_laminate = part.material_details.top_laminate_name
            
            original_material = f"{original_laminate}_{original_core}_{original_laminate}"
            
            # Only track actual upgrades (where materials are different)
            if original_material != board_material and original_material != 'Unknown_Unknown_Unknown':
                upgrade_key = (original_material, board_material)
                
                if upgrade_key not in upgrade_data:
                    upgrade_data[upgrade_key] = 0
                upgrade_data[upgrade_key] += 1
    
    # Convert to list of dictionaries for Streamlit
    report_data = []
    for (original, upgraded), count in upgrade_data.items():
        report_data.append({
            'Original Material': original,
            'Upgraded Material': upgraded,
            'Parts Count': count
        })
    
    return report_data if report_data else None

def show_results_page():
    """Display optimization results and generate reports."""
    st.header("📋 Optimization Results")
    
    # Check if optimization is complete
    if not hasattr(st.session_state, 'optimization_complete') or not st.session_state.optimization_complete:
        st.warning("No optimization results available. Please run optimization first.")
        return
    
    results = st.session_state.optimization_results
    boards = results['boards']
    unplaced_parts = results['unplaced_parts']
    upgrade_summary = results['upgrade_summary']
    
    # Summary metrics
    st.subheader("📊 Summary")
    
    # Calculate comprehensive metrics
    total_board_area = sum(board.total_length * board.total_width for board in boards) / 1_000_000
    total_waste_area = sum(board.get_remaining_area() for board in boards) / 1_000_000
    avg_utilization = sum(board.get_utilization_percentage() for board in boards) / len(boards) if boards else 0
    total_parts = len(st.session_state.parts_list)
    placed_parts = total_parts - len(unplaced_parts)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Boards", len(boards))
    with col2:
        st.metric("Parts Placed", f"{placed_parts}/{total_parts}")
    with col3:
        st.metric("Board Area", f"{total_board_area:.2f} m²")
    with col4:
        st.metric("Waste Area", f"{total_waste_area:.2f} m²")
    with col5:
        st.metric("Avg Utilization", f"{avg_utilization:.1f}%")
    

    # Tabbed interface for improved UI organization
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Analysis Reports", 
        "📋 Board Details",
        "🔄 Upgrades",
        "⚠️ Issues"
    ])
    
    with tab1:
        st.subheader("Core Material Analysis")
        core_report_data = generate_core_material_report_data(boards, st.session_state.core_db)
        if core_report_data:
            # Convert to dataframe-style display for better readability
            core_df_dict = {
                "Core Material": [],
                "Board Count": [],
                "Standard Area (sqft)": [],
                "Utilized Area (sqft)": [],
                "Wastage Area (sqft)": [],
                "Utilization %": [],
                "Wastage %": [],
                "Unit Price (₹/sqft)": [],
                "Total Cost (₹)": []
            }
            
            for row in core_report_data:
                core_df_dict["Core Material"].append(row.get('Core Material', ''))
                core_df_dict["Board Count"].append(row.get('Board Count', 0))
                core_df_dict["Standard Area (sqft)"].append(row.get('Standard Area (sqft)', 0))
                core_df_dict["Utilized Area (sqft)"].append(row.get('Utilized Area (sqft)', 0))
                core_df_dict["Wastage Area (sqft)"].append(row.get('Wastage Area (sqft)', 0))
                core_df_dict["Utilization %"].append(row.get('Utilization %', '0%'))
                core_df_dict["Wastage %"].append(row.get('Wastage %', '0%'))
                core_df_dict["Unit Price (₹/sqft)"].append(row.get('Unit Price (₹/sqft)', '₹0'))
                core_df_dict["Total Cost (₹)"].append(row.get('Total Cost (₹)', '₹0'))
            
            st.dataframe(core_df_dict, use_container_width=True)
        else:
            st.info("No core material data available.")
        
        st.subheader("Laminate Type Analysis")
        laminate_report_data = generate_laminate_report_data(boards, st.session_state.laminate_db)
        if laminate_report_data:
            laminate_df_dict = {
                "Laminate Type": [],
                "Board Count": [],
                "Standard Area (sqft)": [],
                "Utilized Area (sqft)": [],
                "Wastage Area (sqft)": [],
                "Utilization %": [],
                "Wastage %": [],
                "Unit Price (₹/sqft)": [],
                "Total Cost (₹)": []
            }
            
            for row in laminate_report_data:
                laminate_df_dict["Laminate Type"].append(row.get('Laminate Type', ''))
                laminate_df_dict["Board Count"].append(row.get('Board Count', 0))
                laminate_df_dict["Standard Area (sqft)"].append(row.get('Standard Area (sqft)', 0))
                laminate_df_dict["Utilized Area (sqft)"].append(row.get('Utilized Area (sqft)', 0))
                laminate_df_dict["Wastage Area (sqft)"].append(row.get('Wastage Area (sqft)', 0))
                laminate_df_dict["Utilization %"].append(row.get('Utilization %', '0%'))
                laminate_df_dict["Wastage %"].append(row.get('Wastage %', '0%'))
                laminate_df_dict["Unit Price (₹/sqft)"].append(row.get('Unit Price (₹/sqft)', '₹0'))
                laminate_df_dict["Total Cost (₹)"].append(row.get('Total Cost (₹)', '₹0'))
            
            st.dataframe(laminate_df_dict, use_container_width=True)
        else:
            st.info("No laminate data available.")
        
        st.subheader("Edge Band Summary")
        edge_band_report_data = generate_edge_band_report_data(boards)
        if edge_band_report_data:
            edge_band_df_dict = {
                "Edge Band Name": [],
                "Panel Count": [],
                "Total Length (mm)": [],
                "Total Length (m)": []
            }
            
            for row in edge_band_report_data:
                edge_band_df_dict["Edge Band Name"].append(row.get('Edge Band Name', ''))
                edge_band_df_dict["Panel Count"].append(row.get('Panel Count', 0))
                edge_band_df_dict["Total Length (mm)"].append(row.get('Total Length (mm)', '0'))
                edge_band_df_dict["Total Length (m)"].append(row.get('Total Length (m)', '0.00'))
            
            st.dataframe(edge_band_df_dict, use_container_width=True)
        else:
            st.info("No edge band data available.")
        
        # Half-board savings section in Analysis Reports
        if upgrade_summary and 'half_board_savings' in upgrade_summary and upgrade_summary['half_board_savings']:
            st.subheader("🎯 Half-Board Material Savings")
            st.info("Low-utilization boards were rearranged to create large offcuts and save materials")
            
            savings_data = {
                "Core Material": [],
                "Top Laminate": [],
                "Bottom Laminate": [],
                "Saved Quantity": [],
                "Material Signature": []
            }
            
            for saved_board in upgrade_summary['half_board_savings']:
                savings_data["Core Material"].append(saved_board.get('core_material', 'Unknown'))
                savings_data["Top Laminate"].append(saved_board.get('top_laminate', 'Unknown'))
                savings_data["Bottom Laminate"].append(saved_board.get('bottom_laminate', 'Unknown'))
                savings_data["Saved Quantity"].append(f"{saved_board.get('quantity', 0.5):.1f}")
                savings_data["Material Signature"].append(saved_board.get('material_signature', ''))
            
            st.dataframe(savings_data, use_container_width=True)
            
            total_saved_boards = len(upgrade_summary['half_board_savings'])
            total_saved_quantity = sum(saved_board.get('quantity', 0.5) for saved_board in upgrade_summary['half_board_savings'])
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Half-Boards Saved", f"{total_saved_boards}")
            with col2:
                st.metric("Total Material Saved", f"{total_saved_quantity:.1f} boards")
            
            st.success(f"Successfully rearranged {total_saved_boards} low-utilization boards and saved {total_saved_quantity:.1f} board materials!")
    
    with tab2:
        st.subheader("Board Details")
        
        # Create board details table
        board_data = {
            "Board ID": [],
            "Material": [],
            "Size (mm)": [],
            "Parts Count": [],
            "Utilization %": [],
            "Remaining Area (m²)": []
        }
        
        for i, board in enumerate(boards):
            board_data["Board ID"].append(f"Board {i+1}")
            board_data["Material"].append(str(board.material_details))
            board_data["Size (mm)"].append(f"{board.total_length}×{board.total_width}")
            board_data["Parts Count"].append(len(board.parts_on_board))
            board_data["Utilization %"].append(f"{board.get_utilization_percentage():.1f}%")
            board_data["Remaining Area (m²)"].append(f"{board.get_remaining_area()/1_000_000:.2f}")
        
        st.dataframe(board_data, use_container_width=True)
        
        # Detailed part placement for each board
        st.write("**Part Placement Details**")
        for i, board in enumerate(boards):
            with st.expander(f"Board {i+1} - {board.material_details} ({board.get_utilization_percentage():.1f}% utilized)"):
                if board.parts_on_board:
                    part_details = {
                        "ORDER ID / UNIQUE CODE": [],
                        "ROOM TYPE": [],
                        "PANEL NAME": [],
                        "CUT LENGTH × CUT WIDTH": [],
                        "Position (x,y)": [],
                        "MATERIAL TYPE": []
                    }
                    
                    for part in board.parts_on_board:
                        x_pos = getattr(part, 'x', 0)
                        y_pos = getattr(part, 'y', 0)
                        part_details["ORDER ID / UNIQUE CODE"].append(part.id)
                        part_details["ROOM TYPE"].append(getattr(part, 'room_type', ''))
                        part_details["PANEL NAME"].append(getattr(part, 'panel_name', ''))
                        part_details["CUT LENGTH × CUT WIDTH"].append(f"{part.requested_length}×{part.requested_width}")
                        part_details["Position (x,y)"].append(f"({x_pos},{y_pos})")
                        part_details["MATERIAL TYPE"].append(str(part.material_details))
                    
                    st.dataframe(part_details, use_container_width=True)
                else:
                    st.info("No parts placed on this board.")
    
    with tab3:
        st.subheader("Material Upgrades")
        
        # Material Upgrade Report
        upgrade_report_data = generate_material_upgrade_report_data(boards)
        if upgrade_report_data:
            upgrade_df_dict = {
                "Original Material": [],
                "Upgraded Material": [],
                "Parts Count": [],
                "Upgrade Type": []
            }
            
            for row in upgrade_report_data:
                upgrade_df_dict["Original Material"].append(row.get('Original Material', ''))
                upgrade_df_dict["Upgraded Material"].append(row.get('Upgraded Material', ''))
                upgrade_df_dict["Parts Count"].append(row.get('Parts Count', 0))
                upgrade_df_dict["Upgrade Type"].append(row.get('Upgrade Type', ''))
            
            st.dataframe(upgrade_df_dict, use_container_width=True)
        else:
            st.info("No material upgrades detected in this optimization.")
        
        if upgrade_summary:
            st.success(f"Applied {len(upgrade_summary)} material upgrades to improve efficiency")
        

    
    with tab4:
        st.subheader("Issues & Warnings")
        
        # Unplaced parts
        if unplaced_parts:
            st.error(f"{len(unplaced_parts)} parts could not be placed:")
            
            unplaced_data = {
                "ORDER ID / UNIQUE CODE": [],
                "CUT LENGTH × CUT WIDTH": [],
                "MATERIAL TYPE": [],
                "Issue": []
            }
            
            for part in unplaced_parts:
                unplaced_data["ORDER ID / UNIQUE CODE"].append(part.id)
                unplaced_data["CUT LENGTH × CUT WIDTH"].append(f"{part.requested_length}×{part.requested_width}")
                unplaced_data["MATERIAL TYPE"].append(str(part.material_details))
                unplaced_data["Issue"].append("Could not fit on any board")
            
            st.dataframe(unplaced_data, use_container_width=True)
            st.info("Consider using a higher-grade material or splitting large parts to resolve placement issues.")
        else:
            st.success("All parts were successfully placed!")
    
    # Cost analysis
    st.subheader("💰 Cost Analysis")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Initial Cost", f"₹{results['initial_cost']:.2f}")
    with col2:
        st.metric("Final Cost", f"₹{results['final_cost']:.2f}")
    with col3:
        cost_savings = results['initial_cost'] - results['final_cost']
        st.metric("Cost Savings", f"₹{cost_savings:.2f}", delta=f"{cost_savings:.2f}")
    
    # Download reports section
    if hasattr(st.session_state, 'latest_reports') and st.session_state.latest_reports:
        st.subheader("📥 Download Reports")
        
        reports = st.session_state.latest_reports
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if 'cutting_layouts.pdf' in reports:
                st.download_button(
                    "📄 Download Cutting Layouts (PDF)",
                    reports['cutting_layouts.pdf'],
                    f"{st.session_state.get('order_name', 'OptiWise')}_cutting_layouts.pdf",
                    "application/pdf"
                )
            elif 'cutting_layout.txt' in reports:
                st.download_button(
                    "📄 Download Cutting Layout (Text)",
                    reports['cutting_layout.txt'],
                    f"{st.session_state.get('order_name', 'OptiWise')}_cutting_layout.txt",
                    "text/plain"
                )
        
        with col2:
            if 'optimization_report.xlsx' in reports:
                st.download_button(
                    "📊 Download Detailed Report (Excel)",
                    reports['optimization_report.xlsx'],
                    f"{st.session_state.get('order_name', 'OptiWise')}_detailed_report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif 'optimization_report.csv' in reports:
                st.download_button(
                    "📊 Download Optimization Report (CSV)",
                    reports['optimization_report.csv'],
                    "optimization_report.csv",
                    "text/csv"
                )
        
        with col3:
            if 'material_summary.csv' in reports:
                st.download_button(
                    "📋 Download Material Summary (CSV)",
                    reports['material_summary.csv'],
                    "material_summary.csv",
                    "text/csv"
                )
        
        # Additional reports if available
        if 'upgrade_summary.csv' in reports:
            st.download_button(
                "🔄 Download Upgrade Summary (CSV)",
                reports['upgrade_summary.csv'],
                "upgrade_summary.csv",
                "text/csv"
            )
        
        st.info("Reports include detailed cutting layouts, part positions, material utilization, and cost analysis in professional formats.")

def create_project_zip():
    """Create a ZIP file containing all OptiWise project files."""
    zip_buffer = io.BytesIO()
    
    # List of files to include
    files_to_include = [
        'app_complete.py',
        'data_models.py',
        'optimization_core_fixed.py',
        'optimization_global.py',
        'optimization_unified.py',
        'parsers_csv_standalone.py',
        'utils.py',
        'pyproject.toml'
    ]
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_name in files_to_include:
            if os.path.exists(file_name):
                zip_file.write(file_name, file_name)
        
        # Add sample data
        sample_cutlist, sample_core, sample_laminates = create_sample_data()
        zip_file.writestr('sample_cutlist.csv', sample_cutlist)
        zip_file.writestr('sample_core_materials.csv', sample_core)
        zip_file.writestr('sample_laminates.csv', sample_laminates)
        
        # Add startup script
        startup_script = """#!/bin/bash
# OptiWise Startup Script
echo "Starting OptiWise..."
streamlit run app_complete.py --server.port 5000
"""
        zip_file.writestr('start_optiwise.sh', startup_script)
    
    zip_buffer.seek(0)
    return zip_buffer.read()

def show_download_page():
    """Display download page for all OptiWise project files."""
    st.header("📁 Download OptiWise Files")
    st.markdown("Download all the files you need to run OptiWise on your own system")
    
    # System requirements
    st.subheader("📋 System Requirements")
    st.markdown("""
    **Python Requirements:**
    - Python 3.11 or higher
    - pip package manager
    
    **Required Libraries:**
    - streamlit
    
    **Installation Command:**
    ```bash
    pip install streamlit
    ```
    """)
    
    # Create and download project package
    if st.button("📦 Create Complete Project Package"):
        with st.spinner("Creating project package..."):
            try:
                zip_data = create_project_zip()
                
                st.download_button(
                    label="📦 Download Complete OptiWise Package",
                    data=zip_data,
                    file_name="OptiWise_Complete_Package.zip",
                    mime="application/zip"
                )
                
                st.success("✅ Project package created successfully!")
                
                # Instructions
                st.subheader("📖 Installation Instructions")
                st.markdown("""
                **After downloading:**
                1. Extract the ZIP file to your desired location
                2. Open terminal/command prompt in the extracted folder
                3. Install required libraries: `pip install streamlit`
                4. Run OptiWise: `streamlit run app_complete.py --server.port 5000`
                5. Open your browser to `http://localhost:5000`
                """)
                
            except Exception as e:
                st.error(f"Failed to create package: {e}")
    
    # Individual file downloads
    st.subheader("📄 Sample Files")
    
    sample_cutlist, sample_core, sample_laminates = create_sample_data()
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            "Sample Cutlist CSV",
            sample_cutlist,
            "sample_cutlist.csv",
            "text/csv"
        )
    
    with col2:
        st.download_button(
            "Sample Core Materials CSV",
            sample_core,
            "sample_core_materials.csv",
            "text/csv"
        )
    
    with col3:
        st.download_button(
            "Sample Laminates CSV",
            sample_laminates,
            "sample_laminates.csv",
            "text/csv"
        )

def show_help_page():
    """Display help and documentation."""
    st.header("❓ Help & Documentation")
    
    st.markdown("""
    ## 🚀 Getting Started
    
    ### 1. Data Preparation
    Prepare three CSV files with your cutting data:
    - **Parts List**: Details of all parts to be cut
    - **Core Materials**: Available board materials and specifications
    - **Laminates**: Laminate pricing information
    
    ### 2. Data Input Methods
    - **Text Input**: Copy and paste data directly
    - **File Upload**: Upload CSV files
    
    ### 3. Optimization Strategies
    - **Fast**: Quick greedy algorithm for immediate results
    - **Balanced**: Good compromise between speed and efficiency
    - **Maximum Efficiency**: Best material utilization
    - **Mathematical**: Exact algorithms for optimal solutions
    - **Standard**: No material upgrades, exact specifications
    
    ### 4. Material Upgrading
    OptiWise can automatically upgrade parts to higher-grade materials when it:
    - Reduces total waste
    - Improves board utilization
    - Maintains cost effectiveness
    
    ## 📋 File Formats
    
    ### Parts List CSV Format
    ```
    Part ID,Length (mm),Width (mm),Quantity,Material,Grain Sensitive
    PART001,600,400,1,SF-18MR-SF,0
    PART002,800,300,1,2614 SF-18MR-2614 SF,1
    ```
    
    ### Core Materials CSV Format
    ```
    Core Name,Standard Length (mm),Standard Width (mm),Thickness (mm),Price per SqM,Grade Level
    18MR,2440,1220,18,850,1
    18BWR,2440,1220,18,950,2
    ```
    
    ### Laminates CSV Format
    ```
    Laminate Name,Price per SqM
    SF,120
    2614 SF,150
    ```
    
    ## 🔧 Troubleshooting
    
    ### Common Issues:
    
    **Data Loading Fails:**
    - Check CSV format and column headers
    - Ensure material names are consistent across all files
    - Verify numeric values are properly formatted
    
    **Optimization Errors:**
    - Confirm all materials in parts list exist in materials database
    - Check that board dimensions can accommodate largest parts
    - Verify upgrade sequence contains valid material names
    
    **Poor Results:**
    - Try different optimization strategies
    - Adjust material upgrade sequence
    - Consider allowing part rotation (set Grain Sensitive = 0)
    - Check if parts are too large for available boards
    """)

if __name__ == "__main__":
    setup_logging()
    main()