import subprocess
from . import process
import os
import click

bin_dir = os.path.dirname(os.path.realpath(__file__))

@click.command()
@click.argument('--alluredir', help='Path (relative or absolute) to alluredir folder with test results')
@click.argument('--output', help='Path (relative or absolute) with filename for the generated docx file')
@click.option('--template', default=None, help='Path to a custom docx template file with styles')
@click.option('--pdf', is_flag=True, help='Also generate a pdf file from created docx (needs MS Word installed)')
def main(alluredir, output, template, pdf):
    cwd = os.getcwd()
    if not os.path.isabs(alluredir):
        alluredir = os.path.join(cwd, alluredir)
    if not os.path.isabs(output):
        output = os.path.join(cwd, output)
    if template is None:
        # internal template
        pass
    else:
        if not os.path.isabs(template):
            template = os.path.join(cwd, template)
    print("Paths:\n{}\n{}\n{}".format(alluredir, output, template))
    process.run(alluredir, template, output)
    if pdf:
        #generate pdf
        #chdir(alluredir)
        #print(check_output("OfficeToPDF.exe /bookmarks /print demo.docx demo.pdf", shell=True).decode())
        pass
