import subprocess
import os


def pytest_addoption(parser):
    parser.getgroup("reporting").addoption("--allure-docx",
                                           action="store_true",
                                           help="Create a docx file from allure results")

    parser.getgroup("reporting").addoption("--allure-docx-pdf",
                                           action="store_true",
                                           help="Create a docx file from allure results and generate a pdf")


report_dir = None
allure_docx = None
allure_pdf = None


def pytest_configure(config):
    global report_dir, allure_docx, allure_pdf

    report_dir = config.option.allure_report_dir
    allure_docx = config.option.allure_docx
    allure_pdf = config.option.allure_docx_pdf


def pytest_unconfigure():
    if report_dir is not None:
        output_file = os.path.basename(os.path.normpath(report_dir))+".docx"
        if allure_docx or allure_pdf:
            args = '--pdf' if allure_pdf else ''
            subprocess.check_output("allure-docx {} {} {}".format(args, report_dir, output_file), shell=True)
