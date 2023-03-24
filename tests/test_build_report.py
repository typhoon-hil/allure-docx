import os
import shutil
import pytest

from allure_docx import commandline
from allure_docx import ReportConfig
from click.testing import CliRunner
from allure_docx import ConfigTags

file_dir = os.path.dirname(os.path.realpath(__file__))

def test_create_from_commandline():
    os.makedirs(os.path.join(file_dir, "build"), exist_ok=True)
    runner = CliRunner()
    result = runner.invoke(commandline.main, [
        os.path.join(file_dir, "allure-results"),
        os.path.join(file_dir, "build/report.docx"),
        "--config_file", os.path.join(file_dir, "custom.ini")
    ])

    if result.exit_code != 0:
        raise result.exception

    result = runner.invoke(commandline.main, [
        os.path.join(file_dir, "allure-results"),
        os.path.join(file_dir, "build/report.docx"),
        "--config_tag", "no_trace"
    ])

    if result.exit_code != 0:
        raise result.exception

def test_config():
    config = ReportConfig()
    assert "description" in config["info"]["failed"]
    assert "trace" in config["info"]["failed"]

    config = ReportConfig(tag=ConfigTags.NO_TRACE)
    assert "trace" not in config["info"]["failed"]

    config = ReportConfig(config_file = os.path.join(file_dir, "custom.ini"))
    assert "trace" not in config["info"]["failed"]
    assert "setup" not in config["info"]["failed"]
    assert "teardown" not in config["info"]["failed"]
    assert config["cover"]["company"] == "Test company"

@pytest.fixture(autouse=True)
def test_remove_build():
    yield
    build_dir = os.path.join(file_dir, "build")
    if os.path.isdir(build_dir):
        shutil.rmtree(build_dir)
