from subprocess import check_output
from . import process
import os
import click

project_dir = os.path.dirname(os.path.realpath(__file__))

@click.command()
@click.argument('alluredir')
@click.argument('output')
@click.option('--template', default=None, help='Path (absolute or relative) to a custom docx template file with styles')
@click.option('--pdf', is_flag=True, help='(Windows Only) Also generate a pdf file from created docx using OfficeToPdf (needs MS Word installed)')
def main(alluredir, output, template, pdf):
    """alluredir: Path (relative or absolute) to alluredir folder with test results

    output: Path (relative or absolute) with filename for the generated docx file"""
    cwd = os.getcwd()
    if not os.path.isabs(alluredir):
        alluredir = os.path.join(cwd, alluredir)
    if not os.path.isabs(output):
        output = os.path.join(cwd, output)
    if template is None:
        template = os.path.join(project_dir, "template.docx")
    else:
        if not os.path.isabs(template):
            template = os.path.join(cwd, template)
    process.run(alluredir, template, output)
    if pdf:
        filepath, ext = os.path.splitext(output)
        output_pdf = filepath+".pdf"
        print(check_output("OfficeToPDF /bookmarks /print {} {}".format(output, output_pdf), shell=True).decode())
