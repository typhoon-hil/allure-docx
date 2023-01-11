import os
import shutil

import pytest

from allure_docx import commandline
from click.testing import CliRunner

file_dir = os.path.dirname(os.path.realpath(__file__))


def test_create_from_commandline():
    os.makedirs(os.path.join(file_dir, "build"), exist_ok=True)
    runner = CliRunner()
    result = runner.invoke(commandline.main, [
        os.path.join(file_dir, "allure-results"),
        os.path.join(file_dir, "build/report.docx"),
        "--config", os.path.join(file_dir, "custom.ini")
    ])

    assert result.exit_code == 0

@pytest.fixture(autouse=True)
def test_remove_build():
    yield
    build_dir = os.path.join(file_dir, "build")
    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir)
