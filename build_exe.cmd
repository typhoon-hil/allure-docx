@echo off
setlocal

rem // Following lines are just to add the relative libs (cairo dlls) folder to path
set REL_PATH=libs
set ABS_PATH=
rem // Save current directory and change to target directory
pushd %REL_PATH%
rem // Save value of CD variable (current directory)
set ABS_PATH=%CD%
rem // Restore original directory
popd

set PATH=%ABS_PATH%;%PATH%

pyinstaller -F --clean --add-data=src\allure_docx\template.docx;. --additional-hooks-dir=. --log-level=DEBUG -n allure-docx src\allure_docx\commandline.py > output.txt 2>&1 | type output.txt
