from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.shared import RGBColor
from polarion.polarion import Polarion
import oxml_helpers
import argparse
import json
import re
import progressbar


def findWorkitemInDoc(doc):
    workitems_in_doc = {}
    doc_elm = doc._element

    content_list = doc_elm.xpath('.//w:sdt')
    print('Finding workitems in document')
    for content in progressbar.progressbar(content_list, redirect_stdout=True):
        if content.hasTag('workItem'):
            workitem = content.getContent()
            if workitem.hasField('id'):
                id = workitem.getField('id').all_text
                workitems_in_doc[id] = None
    return workitems_in_doc

def getTestRuns(polarion_config):
    test_runs = {}
    try:
        pol = Polarion(polarion_config['url'], polarion_config['username'], polarion_config['password'])
    except Exception as e:
        print(f"Connection to polarion ({polarion_config['url']} {polarion_config['username']}) failed with the following error: {e}")

    try:
        project = pol.getProject(polarion_config['project'])
    except Exception as e:
        print(f"Opening the project {polarion_config[' project ']} failed with the following error: {e}")

    print(f'Loading test runs from Polarion')
    for run in progressbar.progressbar(polarion_config['test_runs'], redirect_stdout=True):
        try:
            test_runs[run] = project.getTestRun(run)
        except Exception as e:
            print(f'Opening test run {run} failed with the following error: {e}')

    return test_runs

def matchResultsToDoc(workitems, test_runs):
    print('Mathing test run results to document')
    for test_run in progressbar.progressbar(test_runs, redirect_stdout=True):
        for record in test_runs[test_run].records:
            if record.getTestCaseName() in workitems:
                workitems[record.getTestCaseName()] = record
            else:
                print(f'{record.getTestCaseName()} was in test run {test_run} but not in the document')
    for workitem in workitems:
        if workitems[workitem] == None:
            print(f'{workitem} not found in test runs')

def cleanhtml(raw_html):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', raw_html)
    return cleantext

def buildResultString(paragraph, result, config):
    result_string = config['result_string']
    if result.executed == None:
        result_string = result_string.replace('{result}', '-')
        result_string = result_string.replace('{result_color}', '-')
        result_string = result_string.replace('{executed}', '-')
        result_string = result_string.replace('{user}', '-')
        result_string = result_string.replace('{comment}', '-')
    else:
        result_string = result_string.replace('{result}', result.result.id)
        # result.executedByURI
        result_string = result_string.replace('{executed}', result.executed.strftime(config['date_format'])) 
        result_string = result_string.replace('{user}', '-')
        if result.comment == None:
            result_string = result_string.replace('{comment}', '-')
        else:
            # result.comment.content can contain HTML, make sure to clean it before putting in doc
            comment = cleanhtml(result.comment.content)
            result_string = result_string.replace('{comment}', comment)

    if '{result_color}' in result_string:
        result_color = result.result.id.upper()

        parts = result_string.split('{result_color}')
        paragraph.add_run(parts[0])
        result_run = paragraph.add_run(result_color)
        paragraph.add_run(parts[1])

        if 'result_name_color' in config:
            if result.result.id in config['result_name_color']:
                config_color = config['result_name_color'][result.result.id]
                result_run.font.color.rgb = RGBColor(config_color[0], config_color[1], config_color[2])
        result_run.font.bold = True
        pass
    else:

        paragraph.add_run(f'\n{result_string}')
            


def fillDocWithResults(doc, workitems, config):
    doc_elm = doc._element

    content_list = doc_elm.xpath('.//w:sdt')
    print('Filling out result in document')
    for content in progressbar.progressbar(content_list, redirect_stdout=True):
        if content.hasTag('workItem'):
            workitem = content.getContent()
            if workitem.hasField('id'):
                id = workitem.getField('id').all_text

                result = workitems[id]
                if result != None:
                    par_idx = 0
                    if config['result_position'] < len(workitem.p_lst):
                        par_idx = config['result_position']
                    else:
                        print(f'Position failure for {id}')

                    par = Paragraph(workitem.p_lst[par_idx], document._body)

                    buildResultString(par, result, config)
                    
                                           


                    

progressbar.streams.wrap_stderr()

parser = argparse.ArgumentParser(description='Add test result to polarion exported document')
parser.add_argument('-c', '--config', type=str, default='config.json', help='json configuration file')
args = parser.parse_args()

with open(args.config) as f:
  config = json.load(f)

document = Document(config['input'])

workitems = findWorkitemInDoc(document)
test_runs = getTestRuns(config['polarion'])
matchResultsToDoc(workitems, test_runs)
fillDocWithResults(document, workitems, config)

document.save(config['output'])

