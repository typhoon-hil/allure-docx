pyinstaller -F --clean --add-data=libs\cairo.dll;cairo --add-data=src\allure_docx\template.docx;. --additional-hooks-dir=. --log-level=DEBUG -n allure-docx src\allure_docx\commandline.py
