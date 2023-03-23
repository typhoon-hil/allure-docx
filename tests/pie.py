from allure_docx import create_piechart


data = {
        "broken": 1,
        "failed": 2,
        "skipped": 3,
        "passed": 4,
}


create_piechart(data, "C:\\Users\\victo\\Desktop\\piechart.png")