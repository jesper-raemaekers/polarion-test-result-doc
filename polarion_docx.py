from docx import Document
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table

from polarion.polarion import Polarion
import oxml_helpers
import argparse
import json

import logging
import progressbar
import copy
from polarion_helpers import getTestRuns
from document_helpers import findWorkitemInDoc, matchResultsToDoc, extendPolarionTables, fillPolarionTables, fillDocWithResults

#setup progress hars
progressbar.streams.wrap_stderr()

#setup logging
logging.basicConfig(filename='log.log', level=logging.INFO, filemode='w',
                    format='[%(asctime)s] {%(pathname)s:%(lineno)d} %(levelname)s - %(message)s')


parser = argparse.ArgumentParser(description='Add test result to polarion exported document')
parser.add_argument('-c', '--config', type=str, default='config.json', help='json configuration file')
args = parser.parse_args()

logging.info(f'Starting with config file "{args.config}"')

config = None
try:
    with open(args.config) as f:
        config = json.load(f)

except Exception as e:
    logging.critical(f'Config filed not opened, exception: {e}')

if config != None:
    logging.info(f"Opening source document {config['input']}")
    document = Document(config['input'])

    
    workitems = findWorkitemInDoc(document)
        
    test_runs = getTestRuns(config['polarion'])


    matchResultsToDoc(workitems, test_runs)
    extendPolarionTables(document)
    fillPolarionTables(document, workitems, config)
    fillDocWithResults(document, workitems, config)

    document.save(config['output'])

