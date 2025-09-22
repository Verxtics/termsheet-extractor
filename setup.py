#!/usr/bin/env python3
"""
Setup script for Multi-Issuer Termsheet Extractor
Run this once to set up your master Excel file and test the system
"""

import os
import sys
from config_helper import create_master_template, validate_master_file

def setup_extractor():
    """Initialize the termsheet extractor system"""
    
    print("🚀 Setting up Multi-Issuer Termsheet Extractor...")
    print("=" * 50)
    
    # Step 1: Create master template
    print("\n1️⃣ Creating master Excel template...")
    master_file = "master_termsheets.xlsx"
    
    if os.path.exists(master_file):
        response = input(f"❓ {master_file} already exists. Overwrite? (y/N): ")
        if response.lower() != 'y':
            print("✅ Using existing master file")
        else:
            create_master_template(master_file)
    else:
        create_master_template(master_file)
    
    # Step 2: Validate master file
    print("\n2️⃣ Validating master file structure...")
    if validate_master_file(master_file):
        print("✅ Master file is ready!")
    else:
        print("❌ Master file validation failed")
        return False
    
    # Step 3: Check dependencies
    print("\n3️⃣ Checking dependencies...")
    try:
        import streamlit
        import pandas
        import pdfplumber
        import openpyxl
        print("✅ All required packages installed")
    except ImportError as e:
        print(f"❌ Missing package: {e}")
        print("💡 Run: pip install -r requirements_improved.txt")
        return False
    
    # Step 4: Instructions
    print("\n🎉 Setup Complete!")
    print("=" * 50)
    print("\n📋 Next Steps:")
    print("1. Run the extractor: streamlit run improved_extractor.py")
    print("2. Upload your PDF termsheets")
    print("3. Data will be appended to master_termsheets.xlsx")
    print("\n📁 Files created:")
    print(f"   • {master_file} - Your master data workbook")
    print("   • improved_extractor.py - Main application")
    print("   • config_helper.py - Configuration utilities")
    print("\n🔧 To add custom issuers:")
    print("   • Edit the issuer_patterns in improved_extractor.py")
    print("   • Use config_helper.py for pattern examples")
    
    return True

if __name__ == "__main__":
    success = setup_extractor()
    
    if success:
        print("\n✅ Ready to extract termsheets!")
        
        # Ask if user wants to run the app
        run_app = input("\n❓ Start the extractor app now? (y/N): ")
        if run_app.lower() == 'y':
            os.system("streamlit run improved_extractor.py")
    else:
        print("\n❌ Setup failed. Please check the errors above.")
        sys.exit(1)
