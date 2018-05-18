import json
from os import listdir, chdir
from os.path import join, isfile
from time import ctime
from datetime import timedelta
from matplotlib import pyplot as plt
import matplotlib as mpl
import os
from subprocess import check_output

from docx.shared import Mm
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


INDENT = 6


def _format_argval(argval):
    """Remove newlines and limit max length

    From Allure-pytest logger (formats argument in the CLI live logs).
    Consider using the same function."""
    MAX_ARG_LENGTH = 100
    argval = argval.replace("\n"," ")
    if len(argval) > MAX_ARG_LENGTH:
        argval = argval[:3]+" ... "+argval[-MAX_ARG_LENGTH:]
    return argval


def build_data(alluredir):

    def _process_steps(session, node):
        if "steps" in node:
            for step in node['steps']:
                if session['start'] is None or step['start'] < session['start']:
                    session['start'] = step['start']
                if session['stop'] is None or step['stop'] > session['stop']:
                    session['stop'] = step['stop']
                _process_steps(session, step)

    json_results = [f for f in listdir(alluredir) if isfile(join(alluredir, f)) and "result" in f]
    json_containers = [f for f in listdir(alluredir) if isfile(join(alluredir, f)) and "container" in f]

    session = {"start": None,
               "stop": None,
               "results": {
                   "broken": 0,
                   "failed": 0,
                   "skipped": 0,
                   "passed": 0,
               },
               "results_relative": {
                   "broken": 0,
                   "failed": 0,
                   "skipped": 0,
                   "passed": 0,
               },
               "total": 0
               }

    data_containers = []
    for file in json_containers:
        with open(join(alluredir, file), encoding="utf-8") as f:
            container = json.load(f)
            if 'befores' in container:
                for before in container['befores']:
                    _process_steps(session, before)
            if 'afters' in container:
                for after in container['afters']:
                    _process_steps(session, after)
            data_containers.append(container)

    data_results = []
    for file in json_results:
        with open(join(alluredir, file), encoding="utf-8") as f:
            result = json.load(f)
            _process_steps(session, result)
            session['total'] += 1
            session['results'][result['status']] += 1
            result["parents"] = []
            for container in data_containers:
                if "children" not in container:
                    continue
                if result["uuid"] in container["children"]:
                    result["parents"].append(container)
            data_results.append(result)

    def getsortingkey(d):
        classification = {"broken": 0,
                          "failed": 1,
                          "skipped": 2,
                          "passed": 3}
        return "{}-{}".format(classification[d["status"]], d["name"])

    sorted_results = sorted(data_results, key=getsortingkey)

    session['duration'] = str(timedelta(seconds=(session['stop']-session['start'])/1000.0))
    session['start'] = ctime(session['start']/1000.0)
    session['stop'] = ctime(session['stop']/1000.0)

    for item in session['results']:
        session['results_relative'][item] = "{:.2f}%".format(100*session['results'][item]/session['total'])

    return sorted_results, session


def create_piechart(session):
    mpl.rcParams['font.size'] = 17.0
    explode = (0.05, 0.05, 0.05, 0.05)
    fig1, ax1 = plt.subplots()
    ax1.pie(session['results'].values(), explode=explode, labels=session['results'].keys(), autopct=lambda x:"{:.2f}%".format(x), colors=["y", "r", "grey", "g"], radius=1.5)
    ax1.axis('equal')
    plt.savefig(session['piechart_source'], bbox_inches="tight")


def create_docx(sorted_results, session, template_path, output_path):

    def create_TOC(document):
        paragraph = document.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = 'TOC \\o "1-1" \\h \\z'   # change 1-3 depending on heading levels you need

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "Right-click to update field."
        fldChar2.append(fldChar3)

        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

    def print_steps(document, parent_step, indent=0):
        indent_str = indent*INDENT*" "
        for step in parent_step:
            if step['status'] in ["failed", "broken"]:
                stepstyle = "Step Failed"
            else:
                stepstyle = "Step"
            document.add_paragraph("{}> {}".format(indent_str,step['name']), style=stepstyle)
            if 'parameters' in step:
                for p in step['parameters']:
                    paragraph = document.add_paragraph("{}    ".format(indent_str), style='Step Param Parag')
                    paragraph.add_run("{} = {}".format(p['name'], _format_argval(p['value'])), style='Step Param')
            if 'statusDetails' in step:
                document.add_paragraph(step['statusDetails']['message'], style=stepstyle)
                table = document.add_table(rows=1, cols=1, style="Trace table")
                hdr_cells = table.rows[0].cells
                hdr_cells[0].add_paragraph(step['statusDetails']['trace']+'\n', style='Code')
            if 'attachments' in step:
                for attachment in step['attachments']:
                    document.add_paragraph("{} [Attachment] {}".format(indent_str, attachment['name']), style="Step")
                    document.add_picture(os.path.join(alluredir, attachment['source']), width=Mm(100))
                    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if 'steps' in step:
                print_steps(step['steps'], document, indent+1)

    document = Document(template_path)

    document.add_heading('TyphoonTest', 0)
    document.add_paragraph('Test Report', style='Subtitle')

    document.add_paragraph('Test Session Summary', style='Alternative Heading 1')
    table = document.add_table(rows=1, cols=2)

    summary_cell = table.rows[0].cells[0]
    summary_cell.add_paragraph('Start: {}\nEnd: {}\nDuration: {}'.format(session['start'], session['stop'], session['duration']))

    results_strs = []
    for item in session['results']:
        results_strs.append("{}: {} ({})".format(item, session['results'][item], session['results_relative'][item]))
    summary_cell.add_paragraph("\n".join(results_strs))

    piechart_cell = table.rows[0].cells[1]
    paragraph = piechart_cell.paragraphs[0]
    run = paragraph.add_run()
    run.add_picture(session['piechart_source'], width=Mm(75))

    document.add_paragraph('Test Results', style="Alternative Heading 1")
    create_TOC(document)

    for test in sorted_results:
        document.add_page_break()
        document.add_heading('{}-{}'.format(test['name'], test['status']), level=1)

        if 'description' in test:
            document.add_paragraph(test['description'])

        if 'parameters' in test:
            document.add_heading('Parameters', level=2)
            for p in test['parameters']:
                document.add_paragraph("{}: {}".format(p['name'], p['value']), style='Step')

        if 'statusDetails' in test:
            document.add_heading('Details', level=2)
            if test['status'] in ["failed", "broken"]:
                style = "Normal Failed"
            else:
                style = None
            document.add_paragraph(test['statusDetails']['message'], style=style)
            table = document.add_table(rows=1, cols=1, style="Trace table")
            hdr_cells = table.rows[0].cells
            hdr_cells[0].add_paragraph(test['statusDetails']['trace']+'\n', style='Code')

        document.add_heading('Test Setup', level=2)
        for parent in test['parents']:
            if 'befores' in parent:
                for before in parent['befores']:
                    document.add_paragraph('[Fixture] {}'.format(before['name']), style="Step")
                    if 'steps' in before:
                        print_steps(before['steps'], 1)

        document.add_heading('Test Body', level=2)
        if 'steps' in test:
            print_steps(test['steps'])

        document.add_heading('Test Teardown', level=2)
        for parent in test['parents']:
            if 'afters' in parent:
                for after in parent['afters']:
                    document.add_paragraph('[Fixture] {}'.format(after['name']), style="Step")
                    if 'steps' in after:
                        print_steps(after['steps'], 1)

    document.save(output_path)


def run(alluredir, template_path, output_filename):
    results, session = build_data(alluredir)
    session['piechart_source'] = os.path.join(alluredir, 'pie.png')
    create_piechart(session)
    create_docx(results, session, template_path, output_filename)



