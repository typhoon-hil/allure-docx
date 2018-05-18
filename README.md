# allure-docx
docx report generation based on allure-generated result files using `python-docx` library:
https://python-docx.readthedocs.io/en/latest/

## Installation
Install with `pip install allure-docx`. Then if your python Scripts folder is in the path, you can run directly from command line as `allure-docx`.

## Usage
Check usage by running `allure-docx --help`

The generated docx contain a Table of Contents that needs to be manually updated after generation. Generating PDFs (see below) will automatically update the TOC though.

## ToDos:
- This report does not takes all the fields in allure data model into account. Missing `descriptionHtml`, `links`, `stage`, `labels`, `statusDetails` specifics (`flaky`, `known`). This however could be very easily added in the script. Contributors are more then welcome to help making the report more complete.
- Does not support a test result folder with old runs for the same tests.

## PDF
On Windows, PDFs can be generated from generated docx files using OfficeToPdf application, which should be in the path. MS Word needs to be installed as well.
https://github.com/cognidox/OfficeToPDF

The `--pdf` option will search for OfficeToPdf application in the path to generate the PDF.
