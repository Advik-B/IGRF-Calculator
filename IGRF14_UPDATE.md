# IGRF-14 Update

This document describes the changes made to upgrade the IGRF Calculator from IGRF-13 to IGRF-14.

## Changes Made

### 1. Downloaded IGRF-14 Coefficients
- Used Playwright to browse the official IAGA website (Kyoto University)
- Downloaded the official IGRF-14 coefficient file (igrf14coeffs.xlsx)
- Downloaded from: https://wdc.kugi.kyoto-u.ac.jp/igrf/

### 2. Converted Coefficients to CSV
- Converted the Excel file to CSV format (coeff.csv)
- Updated coefficient data from IGRF-13 (1900-2020, SV 2020-25) to IGRF-14 (1900-2025, SV 2025-30)
- Maintained compatibility with existing code structure

### 3. Updated Code References
- Updated `calc.py`:
  - Changed docstring from "IGRF13" to "IGRF-14"
  - Updated sheet name from 'IGRF13coeffs' to 'igrf14coeffs'
  - Updated expected year range in comments (2020 → 2025)
- Code structure remains the same, ensuring backward compatibility

### 4. Virtual Display Support
- Added `run_calculator.py` wrapper script for headless environments
- Automatically detects if display is available
- Starts Xvfb virtual display if needed
- Falls back to Qt offscreen platform if Xvfb unavailable

### 5. Testing
- Created test suite (`test_igrf14.py`) to verify IGRF-14 calculations
- All tests passing with reasonable values
- Verified secular variation is working correctly

## IGRF-14 vs IGRF-13

Key differences:
- **Time Coverage**: IGRF-14 extends to 2025 (vs 2020 for IGRF-13)
- **Secular Variation**: Now covers 2025-2030 (vs 2020-2025)
- **Accuracy**: Updated with latest geomagnetic measurements
- **Additional Data**: One more epoch (2025) included

## Files Changed

### Modified Files:
- `calc.py` - Updated IGRF version references
- `coeff.csv` - Replaced with IGRF-14 coefficients
- `.gitignore` - Added temporary files to ignore

### New Files:
- `run_calculator.py` - Wrapper for virtual display support
- Helper scripts (not committed):
  - `download_igrf14.py` - Playwright script to download coefficients
  - `convert_igrf14_to_csv.py` - Excel to CSV converter
  - `test_igrf14.py` - Test suite

## Running the Application

### With Display:
```bash
python3 app.py
```

### Without Display (Headless):
```bash
# Option 1: Use wrapper script (recommended)
python3 run_calculator.py

# Option 2: Manual Xvfb
Xvfb :99 -screen 0 1024x768x24 &
export DISPLAY=:99
python3 app.py

# Option 3: Use Qt offscreen
export QT_QPA_PLATFORM=offscreen
python3 app.py
```

## Dependencies

Required packages:
- pandas
- numpy
- openpyxl
- PyQt6
- pyexcel
- pyexcel-xls
- xlrd

For Playwright (development only, not needed for running the app):
- playwright

For virtual display (optional, for headless environments):
- xvfb (system package: `sudo apt-get install xvfb`)
- libegl1 (system package)
- libgl1 (system package)
- libxcb-cursor0 (system package)

## Verification

To verify the IGRF-14 implementation is working:
```bash
python3 test_igrf14.py
```

Expected output:
- All tests pass
- Values for Boulder, CO at 2025: TI ≈ 51,300 nT, DEC ≈ 7.9°, DIP ≈ 66°
- Values change over time due to secular variation

## References

- IGRF-14 Official Page: https://www.ncei.noaa.gov/products/international-geomagnetic-reference-field
- IAGA Download Site: https://wdc.kugi.kyoto-u.ac.jp/igrf/
- GitHub Repository: https://github.com/IAGA-VMOD/IGRF14eval
