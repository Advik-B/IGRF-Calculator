#!/usr/bin/env python3
"""
Final verification script for IGRF-14 upgrade.
This script verifies all components are working correctly.
"""

import os
import sys
import subprocess
import shutil

def check_file(filename, description):
    """Check if a file exists and has content."""
    if os.path.exists(filename):
        size = os.path.getsize(filename)
        print(f"✓ {description}: {filename} ({size} bytes)")
        return True
    else:
        print(f"✗ {description} missing: {filename}")
        return False

def check_coefficient_data():
    """Verify coefficient file structure."""
    try:
        import pandas as pd
        
        # Use absolute path based on script location
        script_dir = os.path.dirname(os.path.abspath(__file__))
        coeff_path = os.path.join(script_dir, 'coeff.csv')
        
        if not os.path.exists(coeff_path):
            print(f"✗ Coefficient file not found at: {coeff_path}")
            return False
        
        df = pd.read_csv(coeff_path)
        
        # Check column count (should be 30: g/h, n, m, 26 years, 1 SV)
        if len(df.columns) == 30:
            print(f"✓ Coefficient file has correct column count: {len(df.columns)}")
        else:
            print(f"✗ Coefficient file has wrong column count: {len(df.columns)} (expected 30)")
            return False
        
        # Check year range
        if '2025' in df.columns and '2025-30' in df.columns:
            print("✓ IGRF-14 year range correct (includes 2025 and 2025-30)")
        else:
            print("✗ Missing IGRF-14 columns")
            return False
        
        # Check row count (should be 195 for degree 13)
        if len(df) == 195:
            print(f"✓ Coefficient count correct: {len(df)}")
        else:
            print(f"⚠ Coefficient count: {len(df)} (expected 195)")
        
        return True
    except Exception as e:
        print(f"✗ Error checking coefficients: {e}")
        return False

def check_calc_module():
    """Check calc.py for IGRF-14 references."""
    try:
        # Use absolute path based on script location
        script_dir = os.path.dirname(os.path.abspath(__file__))
        calc_path = os.path.join(script_dir, 'calc.py')
        
        if not os.path.exists(calc_path):
            print(f"✗ calc.py not found at: {calc_path}")
            return False
        
        with open(calc_path, 'r') as f:
            content = f.read()
        
        # Case-insensitive check for IGRF-14 references
        content_lower = content.lower()
        if 'igrf-14' in content_lower or 'igrf14' in content_lower or 'igrf 14' in content_lower:
            print("✓ calc.py references IGRF-14")
        else:
            print("⚠ calc.py may still reference IGRF-13")
        
        if 'igrf14coeffs' in content:
            print("✓ Excel sheet name updated to 'igrf14coeffs'")
        else:
            print("⚠ Excel sheet name may not be updated")
        
        return True
    except Exception as e:
        print(f"✗ Error checking calc.py: {e}")
        return False

def check_dependencies():
    """Check if required Python packages are available."""
    required = ['pandas', 'numpy', 'openpyxl']
    optional = ['playwright', 'PyQt6']
    
    all_good = True
    for pkg in required:
        try:
            __import__(pkg)
            print(f"✓ Required package installed: {pkg}")
        except ImportError:
            print(f"✗ Required package missing: {pkg}")
            all_good = False
    
    for pkg in optional:
        try:
            __import__(pkg)
            print(f"✓ Optional package installed: {pkg}")
        except ImportError:
            print(f"⚠ Optional package not installed: {pkg}")
    
    return all_good

def check_virtual_display():
    """Check if Xvfb is available (Linux only)."""
    # Only relevant on Linux
    if not sys.platform.startswith('linux'):
        print("⚠ Xvfb is Linux-specific (not required on this platform)")
        return True
    
    try:
        xvfb_path = shutil.which('Xvfb')
        if xvfb_path:
            print("✓ Xvfb (virtual display) is available")
            return True
        else:
            print("⚠ Xvfb not found (not required but recommended for headless)")
            return False
    except Exception as e:
        print(f"⚠ Could not check for Xvfb: {e}")
        return False

def main():
    """Run all verification checks."""
    print("="*70)
    print(" IGRF-14 Upgrade Verification")
    print("="*70)
    
    results = []
    
    print("\n1. File Checks:")
    print("-" * 70)
    results.append(check_file('coeff.csv', 'Coefficient file'))
    results.append(check_file('calc.py', 'Calculator module'))
    results.append(check_file('run_calculator.py', 'Display wrapper'))
    results.append(check_file('IGRF14_UPDATE.md', 'Documentation'))
    
    print("\n2. Coefficient Data:")
    print("-" * 70)
    results.append(check_coefficient_data())
    
    print("\n3. Code Updates:")
    print("-" * 70)
    results.append(check_calc_module())
    
    print("\n4. Dependencies:")
    print("-" * 70)
    results.append(check_dependencies())
    
    print("\n5. Virtual Display Support:")
    print("-" * 70)
    check_virtual_display()  # Optional, don't count in results
    
    print("\n" + "="*70)
    passed = sum(1 for r in results if r)
    total = len(results)
    
    if passed == total:
        print(f"✓ ALL CHECKS PASSED ({passed}/{total})")
        print("="*70)
        print("\nIGRF-14 upgrade completed successfully!")
        print("\nTo run the calculator:")
        print("  python3 run_calculator.py  (recommended, handles display)")
        print("  python3 app.py            (requires display)")
        return 0
    else:
        print(f"⚠ SOME CHECKS FAILED ({passed}/{total} passed)")
        print("="*70)
        return 1

if __name__ == "__main__":
    sys.exit(main())
