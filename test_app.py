#!/usr/bin/env python3
"""
Test script for webapp/app.py
Tests Excel file processing and conversion logic
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def test_excel_validation():
    """Test Excel file validation"""
    print("=== Testing Excel Validation ===")
    
    # Test with sample Excel file
    sample_excel = os.path.join('..', '0. Hoso.xlsx')
    if not os.path.exists(sample_excel):
        print("❌ Sample Excel file not found")
        return False
    
    try:
        from app import core_excel_to_word
        
        # Create temp output directory
        with tempfile.TemporaryDirectory() as temp_dir:
            print(f"Testing with Excel file: {sample_excel}")
            print(f"Output directory: {temp_dir}")
            
            success, result = core_excel_to_word(sample_excel, temp_dir)
            
            if success:
                print(f"✅ Conversion successful!")
                print(f"Created {len(result)} files:")
                for file_path in result:
                    print(f"  - {file_path}")
                return True
            else:
                print(f"❌ Conversion failed: {result}")
                return False
                
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_excel_reader():
    """Test ExcelReader functionality"""
    print("\n=== Testing ExcelReader ===")
    
    sample_excel = os.path.join('..', '0. Hoso.xlsx')
    if not os.path.exists(sample_excel):
        print("❌ Sample Excel file not found")
        return False
    
    try:
        from excel_reader import ExcelReader
        
        excel_reader = ExcelReader(sample_excel)
        hoso_codes = excel_reader.get_all_hoso_codes()
        
        print(f"✅ ExcelReader initialized successfully")
        print(f"Found {len(hoso_codes)} hồ sơ codes: {hoso_codes}")
        
        if hoso_codes:
            # Test context creation for first hoso
            first_hoso = hoso_codes[0]
            context = excel_reader.create_complete_context(first_hoso)
            print(f"✅ Context created for {first_hoso}")
            print(f"Context keys: {list(context.keys())}")
            return True
        else:
            print("❌ No hồ sơ codes found")
            return False
            
    except Exception as e:
        print(f"❌ ExcelReader test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_templates():
    """Test template files"""
    print("\n=== Testing Templates ===")
    
    template_folder = os.path.join('..', 'templates')
    if not os.path.exists(template_folder):
        print("❌ Template folder not found")
        return False
    
    from pathlib import Path
    templates = list(Path(template_folder).glob("*.docx"))
    templates = [t for t in templates if not t.name.startswith('~$') and t.name != '9. Phieu YCTN TNN.docx']
    
    print(f"✅ Found {len(templates)} templates:")
    for template in templates:
        print(f"  - {template.name}")
    
    return len(templates) > 0

def main():
    """Run all tests"""
    print("🧪 Starting WebApp Tests...")
    
    tests = [
        test_templates,
        test_excel_reader,
        test_excel_validation
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"❌ Test {test.__name__} crashed: {e}")
    
    print(f"\n📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! WebApp is ready.")
        return True
    else:
        print("⚠️  Some tests failed. Please check the issues above.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
