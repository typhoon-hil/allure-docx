# allure-docx
docx report generation based on allure-generated result files using `python-docx` library:
https://python-docx.readthedocs.io/en/latest/

Currently tested only with Python 3.6

## Installation
Clone the repository and install with `pip install .`. Then if your python Scripts folder is in the path, you can run directly from command line as `allure-docx`.

## Usage
Check usage by running `allure-docx --help`

The generated docx contain a Table of Contents that needs to be manually updated after generation. Generating PDFs (see below) will automatically update the TOC though.

### PDF
The `--pdf` option will search for either `OfficeToPDF` or `LibreOfficeToPDF` application in the path to generate the PDF.

On Windows, PDFs can be generated from generated docx files using OfficeToPDF application. MS Word needs to be installed.

https://github.com/cognidox/OfficeToPDF

On Windows and Linux, PDFs can be generated using `LibreOfficeToPDF` application. LibreOffice should be installed.

https://github.com/typhoon-hil/LibreOfficeToPDF


## TODOs:
- This report does not takes all the fields in allure data model into account. Missing `descriptionHtml`, `links`, `stage`, `labels`, `statusDetails` specifics (`flaky`, `known`). This however could be very easily added in the script. Contributors are more then welcome to help making the report more complete.
- Does not support a test result folder with old runs for the same tests. It shows them all as individual test cases.
- This package doesn't have any tests yet.
- Only images are attached to the docx file. For other types of attachment, only name is shown.

