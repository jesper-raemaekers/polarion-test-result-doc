from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from polarion.polarion import Polarion
import oxml_helpers
import argparse
import json

def findWorkitemInDoc(doc):
    workitems_in_doc = {}
    doc_elm = doc._element

    content_list = doc_elm.xpath('.//w:sdt')
    for content in content_list:
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
        
    for run in polarion_config['test_runs']:
        try:
            test_runs[run] = project.getTestRun(run)
        except Exception as e:
            print(f'Opening test run {run} failed with the following error: {e}')

    return test_runs

def matchResultsToDoc(workitems, test_runs):
    for test_run in test_runs:
        for record in test_runs[test_run].records:
            if record.getTestCaseName() in workitems:
                workitems[record.getTestCaseName()] = record
            else:
                print(f'{record.getTestCaseName()} was in test run {test_run} but not in the document')
    for workitem in workitems:
        if workitems[workitem] == None:
            print(f'{workitem} not found in test runs')

           
def fillDocWithResults(doc, workitems, config):
    doc_elm = doc._element

    content_list = doc_elm.xpath('.//w:sdt')
    for content in content_list:
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

                    result_string = config['result_string']
                    
                    if result.executed == None:
                        result_string = result_string.replace('{result}', 'No')
                        result_string = result_string.replace('{executed}', 'No date')
                    else:
                        result_string = result_string.replace('{result}', result.result.id)
                        result_string = result_string.replace('{executed}', result.executed.strftime(config['date_format']))                        


                    par.add_run(f'\n{result_string}')


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

