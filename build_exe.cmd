rem pyinstaller -F --clean --add-data=src\allure_docx\template.docx;. --add-data=venv\Lib\site-packages\pygal\css;pygal\css --log-level=DEBUG -n allure-docx src\allure_docx\commandline.py

pyinstaller -F --clean --add-data=src\allure_docx\template.docx;. --additional-hooks-dir=. --log-level=DEBUG -n allure-docx src\allure_docx\commandline.py
