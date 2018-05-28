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

### Installing with Python and pip
You need a installed python interpreter.

Clone the repository and install with `pip install .`. Then if your python Scripts folder is in the PATH, you can run directly from command line as `allure-docx`.

## Usage
Check usage by running `allure-docx --help`

You can generate the docx file by running `allure-docx ALLUREDIR filename.docx`

The generated docx contain a Table of Contents that needs to be manually updated after generation. Generating PDFs (see below) will automatically update the TOC though.

### PDF
The `--pdf` option will search for either `OfficeToPDF` or `LibreOfficeToPDF` application in the PATH to generate the PDF.

On Windows, PDFs can be generated from generated docx files using OfficeToPDF application. MS Word needs to be installed.

https://github.com/cognidox/OfficeToPDF

On Windows and Linux, PDFs can be generated using `LibreOfficeToPDF` application. LibreOffice should be installed.

https://github.com/typhoon-hil/LibreOfficeToPDF

### Custom Title and Logo
You can use the `--title` option to customize the docx report title.
 
If you want to remove the title altogether (e.g. your logo already has the company title), you can set `--title=""`.

You can also use the `--logo` option with a path to a custom image to add the test title and the `--logo-height` option to adjust the height size of the logo image (in centimeters).

Example invocation:

`allure-docx --pdf --title="My Company" --logo=C:\mycompanylogo.png --logo-height=2 allure allure.docx`




