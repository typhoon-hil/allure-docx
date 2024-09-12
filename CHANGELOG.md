# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [0.4.1] - 2024-09-12
### Changed
- Changed template to include new styles: `JSON Attachment` for JSON attachments in code style and `Normal table` for table attachments.
- Added tests for both type of attachments to suite.

## [0.4.0] - 2023-01-19
### Changed
- Configuration is now controlled with .ini files
- Layout and template changed
- Added new information that can be added to the report
- Removed `LibreOfficeToPDF`, instead use `soffice`
- Removed `OfficeToPdf`, instead use `docx2pdf`
- Pie chart replaced by ``matplotlib.pyploy.pie``, without necessity of ``cairo.dll`` and ``cairosvg`` (Python package).

## [0.3.2] - 2018-10-31
### Changed
- Now generating 32-bit executable.

## [0.3.1] - 2018-10-30
### Fixed
- Problems loading Cairo DLL.

## [0.3.0] - 2018-07-08
### Added 
- Cairo DLL is now bundled inside the executable.
### Fixed
- Issues with `statusDetails` without `message` or `trace` attributes (issue #16).

## [0.2.0] - 2018-06-13
### Added
- Command line options to control level of details in the generated docx report (issue #14).

## [0.1.3] - 2018-06-13
### Fixed
- Gracefully treating cases where no test results are found (avoids `ZeroDivisionError`, Issue #15).

## [0.1.2] - 2018-05-30
### Fixed
- Fixed problem when tests doesn't have `start` or `stop` attributes.
  - Also selecting newest test based on json file modification date.

## [0.1.1] - 2018-05-29
### Changed
- Using `shutil.which` to find `OfficeToPDF` or `LibreOfficeToPDF`, as this function will also return batch files.

