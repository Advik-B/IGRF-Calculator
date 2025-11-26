import sys
import os
import shutil
import subprocess
from setuptools import setup, Command
from setuptools.command.build import build

# --- Configuration ---
APP_NAME = "IGRF Calculator"
MAIN_SCRIPT = "app.py"  # **REPLACE THIS if your script name is different**
VERSION = "1.0.0"
AUTHOR = "Advik" # Replace with your name/org
DESCRIPTION = "An application to calculate IGRF values from input data."

# Files/folders to bundle with the application (relative paths)
# Format: "source_path{os.pathsep}destination_in_bundle"
# '.' means the root directory inside the bundle.
DATA_FILES = [
    f"coeff.csv{os.pathsep}.",
    f"icon.ico{os.pathsep}.",
]

# Optional: Specify an icon (provide path to .ico on Windows, .icns on macOS)
ICON_FILE = "icon.ico" # Replace with your icon file or set to None

# PyInstaller options
PYINSTALLER_OPTIONS = [
    '--noconfirm',         # Don't ask for confirmation before overwriting
    '--clean',             # Clean PyInstaller cache and remove temporary files before building
    '--windowed',          # Build a windowed application (no console) - use '--console' for debugging
    # '--onefile',           # Uncomment for a single executable file (slower startup)
    # If not using --onefile, output is a folder in 'dist'
]

# Hidden imports that PyInstaller might miss
# Add modules here if you get ImportError at runtime
HIDDEN_IMPORTS = [
    'PyQt6.sip',
    'PyQt6.QtGui',
    'PyQt6.QtCore',
    'PyQt6.QtWidgets',
    'pandas._libs.tslibs.timedeltas', # Common pandas hidden imports
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.timestamps',
    'numpy', # Explicitly include numpy might help sometimes
    'openpyxl', # Needed for .xlsx saving
    'pyexcel',
    'pyexcel_xls'
    # Add more based on runtime errors if needed
]
# --- End Configuration ---


# Helper function to get platform-specific path separator
def get_path_sep():
    return ';' if sys.platform == 'win32' else ':'

# Custom Command to run PyInstaller
class BuildExeCommand(Command):
    """A custom command to run PyInstaller"""
    description = 'Build the application using PyInstaller'
    user_options = [
        ('onefile=', None, 'Build as a one-file executable (True/False)'),
        ('console=', None, 'Build as a console application (True/False)'),
    ]

    def initialize_options(self):
        """Set default values for options."""
        self.onefile = 'False' # Default to one-folder
        self.console = 'False' # Default to windowed

    def finalize_options(self):
        """Post-process options."""
        # Ensure boolean interpretation
        self.onefile = str(self.onefile).lower() in ('true', '1', 'yes')
        self.console = str(self.console).lower() in ('true', '1', 'yes')

    def run(self):
        """Run command."""
        # Check if main script exists
        if not os.path.exists(MAIN_SCRIPT):
            print(f"Error: Main script '{MAIN_SCRIPT}' not found!", file=sys.stderr)
            sys.exit(1)

        # Clean previous builds
        print("--- Cleaning previous build directories (dist/, build/) ---")
        if os.path.isdir('dist'):
            shutil.rmtree('dist')
        if os.path.isdir('build'):
            shutil.rmtree('build')
        spec_file = f"{os.path.splitext(MAIN_SCRIPT)[0]}.spec"
        if os.path.exists(spec_file):
             os.remove(spec_file)


        # Construct the PyInstaller command
        command = ['pyinstaller'] + PYINSTALLER_OPTIONS

        if self.onefile:
            command.append('--onefile')
        if self.console:
            # Remove --windowed if console is requested
            if '--windowed' in command: command.remove('--windowed')
            command.append('--console')
        elif '--windowed' not in command: # Ensure --windowed if not console
             command.append('--windowed')
        
        # *** Add the --name option here ***
        command.extend(['--name', APP_NAME])
        print(f"--- Setting output name to: {APP_NAME} ---")

        # Add data files
        path_sep = get_path_sep()
        for data_file in DATA_FILES:
            # Ensure source file exists
            source_path = data_file.split(path_sep)[0]
            if not os.path.exists(source_path):
                print(f"Warning: Data file '{source_path}' not found, skipping.", file=sys.stderr)
                continue
            command.extend(['--add-data', data_file])

        # Add hidden imports
        for hidden in HIDDEN_IMPORTS:
            command.extend(['--hidden-import', hidden])

        # Add icon if specified and exists
        if ICON_FILE and os.path.exists(ICON_FILE):
            command.extend(['--icon', ICON_FILE])
        elif ICON_FILE:
             print(f"Warning: Icon file '{ICON_FILE}' not found, skipping icon.", file=sys.stderr)


        # Add the main script
        command.append(MAIN_SCRIPT)

        print(f"--- Running PyInstaller ---")
        print(f"Command: {' '.join(command)}")
        try:
            subprocess.check_call(command)
            print("--- PyInstaller build completed successfully! ---")
            # Determine output path using APP_NAME
            output_dir = os.path.join('dist', APP_NAME) if not self.onefile else 'dist'
            exe_name = f"{APP_NAME}.exe" if sys.platform == 'win32' else APP_NAME
            output_path = os.path.join(output_dir, exe_name)
            print(f"--- Executable created: {os.path.abspath(output_path)} ---")

        except subprocess.CalledProcessError as e:
            print(f"--- PyInstaller build failed! Error: {e} ---", file=sys.stderr);
            sys.exit(1)
        except FileNotFoundError:
            print("Error: 'pyinstaller' command not found. Install PyInstaller and ensure it's in PATH.",
                  file=sys.stderr)
            sys.exit(1)

with open("requirements.txt") as f:
    req = f.readlines()

# Standard setup() function
setup(
    name=APP_NAME,
    version=VERSION,
    author=AUTHOR,
    description=DESCRIPTION,
    py_modules=[os.path.splitext(MAIN_SCRIPT)[0]],
    # List runtime dependencies here for informational purposes
    # or if you intend to distribute as a package too.
    install_requires=req,
    # Define the custom command
    cmdclass={
        'build_exe': BuildExeCommand,
    },
    # Optional: You might need to include your package if it's structured
    # packages=find_packages(),
    include_package_data=True,
)