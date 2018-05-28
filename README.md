# allure-docx
docx report generation based on allure-generated result files using `python-docx` library:
https://python-docx.readthedocs.io/en/latest/

Currently tested only with Python 3.6

## Limitations
Check the issues on the repository to see the current limitations.

## Questions

Feel free to open an issue ticket with any problems or questions.

## Installation

### Requirements
This project uses `cairosvg` package, which in turn needs Cairo binary libraries installed in a place in PATH.

For windows, cairo precompiled binary dlls can be found here:
https://github.com/preshing/cairo-windows/releases

### Downloading and using the binaries directly
We publish windows standalone executable files. With them you can use it without having to install anything else (no python, etc..).

You can download them at: https://github.com/typhoon-hil/allure-docx/releases

Then you can use the executable directly (possibly adding it a folder added to PATH).

### Installing with Python and pip
You need a installed python interpreter.

Install directly from this git repository using with `pip install git+https://github.com/typhoon-hil/allure-docx.git`. Then if your python Scripts folder is in the PATH, you can run directly from command line as `allure-docx`.

## Usage
Check usage by running `allure-docx --help`

You can generate the docx file by running `allure-docx ALLUREDIR filename.docx`

The generated docx contain a Table of Contents that needs to be manually updated after generation. Generating PDFs (see below) will automatically update the TOC though.

### PDF
The `--pdf` option will search for either `OfficeToPDF` or `LibreOfficeToPDF` application in the PATH to generate the PDF.

On Windows, PDFs can be generated from generated docx files using OfficeToPDF application. MS Word needs to be installed.

https://github.com/cognidox/OfficeToPDF/releases

On Windows and Linux, PDFs can be generated using `LibreOfficeToPDF` application. LibreOffice should be installed.

https://github.com/typhoon-hil/LibreOfficeToPDF/releases

### Custom Title and Logo
You can use the `--title` option to customize the docx report title.
 
If you want to remove the title altogether (e.g. your logo already has the company title), you can set `--title=""`.

You can also use the `--logo` option with a path to a custom image to add the test title and the `--logo-height` option to adjust the height size of the logo image (in centimeters).

Example invocation:

`allure-docx --pdf --title="My Company" --logo=C:\mycompanylogo.png --logo-height=2 allure allure.docx`

## Building a standalone executable
We use PyInstaller to create standalone executables. If you want to build an executable yourself, follow these steps:
- Create a new virtual environment with the proper python version (tested using python 3, 32 or 64 bit so far)
- Install using ONLY pip needed packages defined in `setup.py`. This prevents your executable to become too large with unnecessary dependencies
- Delete any previous `dist` folder from a previous PyInstaller run
- Run the `build_exe.cmd` command to run PyInstaller and create a single file executable.


