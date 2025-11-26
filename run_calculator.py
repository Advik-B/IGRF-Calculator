#!/usr/bin/env python3
"""
IGRF Calculator - Wrapper script to handle display setup.

This script ensures that Qt can run properly, either by:
1. Using the existing display (if available)
2. Starting a virtual display (Xvfb) if running headless
3. Using Qt's offscreen platform as a fallback

For IGRF-14, this application uses coefficients downloaded from the official
IAGA website using Playwright.
"""

import os
import sys
import subprocess
import time

def check_display():
    """Check if a display is available."""
    display = os.environ.get('DISPLAY')
    if display:
        # Try to connect to the display
        try:
            # Check if we can actually use the display
            result = subprocess.run(['xdpyinfo'], 
                                  capture_output=True, 
                                  timeout=2)
            return result.returncode == 0
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return False
    return False

def start_xvfb():
    """Start Xvfb virtual display."""
    try:
        # Check if Xvfb is available
        result = subprocess.run(['which', 'Xvfb'], 
                              capture_output=True, 
                              timeout=2)
        if result.returncode != 0:
            print("Xvfb not found. Install with: sudo apt-get install xvfb")
            return None
        
        # Start Xvfb on display :99
        display_num = ':99'
        print(f"Starting Xvfb on display {display_num}...")
        
        xvfb_proc = subprocess.Popen([
            'Xvfb', display_num,
            '-screen', '0', '1024x768x24',
            '-nolisten', 'tcp',
            '-nolisten', 'unix'
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # Give it time to start
        time.sleep(2)
        
        # Check if it's still running
        if xvfb_proc.poll() is None:
            os.environ['DISPLAY'] = display_num
            print(f"✓ Xvfb started successfully on {display_num}")
            return xvfb_proc
        else:
            print("✗ Xvfb failed to start")
            return None
            
    except Exception as e:
        print(f"Error starting Xvfb: {e}")
        return None

def main():
    """Main entry point."""
    print("="*60)
    print("IGRF-14 Calculator")
    print("="*60)
    
    xvfb_proc = None
    
    # Check if we have a display
    if check_display():
        print("✓ Display detected")
    else:
        print("⚠ No display detected, attempting to use virtual display...")
        xvfb_proc = start_xvfb()
        
        if xvfb_proc is None:
            print("⚠ Virtual display not available, using Qt offscreen platform")
            os.environ['QT_QPA_PLATFORM'] = 'offscreen'
    
    # Import and run the main application
    try:
        print("\nStarting IGRF Calculator application...")
        import app
        exit_code = app.main()
        
    except KeyboardInterrupt:
        print("\nApplication interrupted by user")
        exit_code = 130
        
    except Exception as e:
        print(f"\nError running application: {e}")
        import traceback
        traceback.print_exc()
        exit_code = 1
        
    finally:
        # Clean up Xvfb if we started it
        if xvfb_proc is not None:
            print("\nStopping virtual display...")
            xvfb_proc.terminate()
            try:
                xvfb_proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                xvfb_proc.kill()
    
    return exit_code

if __name__ == "__main__":
    sys.exit(main())
