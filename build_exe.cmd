pyinstaller -F --clean --add-data=src\allure_docx\template.docx;. --additional-hooks-dir=. --log-level=DEBUG -n allure-docx src\allure_docx\commandline.py
