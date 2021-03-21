
from polarion.polarion import Polarion
import progressbar
import logging

def getTestRuns(polarion_config):
    test_runs = {}
    try:
        pol = Polarion(polarion_config['url'], polarion_config['username'], polarion_config['password'])
    except Exception as e:
        logging.error(f"Connection to polarion ({polarion_config['url']} {polarion_config['username']}) failed with the following error: {e}")
        print(f"Connection to polarion ({polarion_config['url']} {polarion_config['username']}) failed with the following error: {e}")

    try:
        project = pol.getProject(polarion_config['project'])
    except Exception as e:
        logging.error(f"Opening the project {polarion_config[' project ']} failed with the following error: {e}")
        print(f"Opening the project {polarion_config[' project ']} failed with the following error: {e}")

    print(f'Loading test runs from Polarion')
    records_sum = 0
    for run in progressbar.progressbar(polarion_config['test_runs'], redirect_stdout=True):
        try:
            test_runs[run] = project.getTestRun(run)
            records_sum += len(test_runs[run].records)
        except Exception as e:
            logging.error(f'Opening test run {run} failed with the following error: {e}')
            print(f'Opening test run {run} failed with the following error: {e}')

    logging.info(f"Found {records_sum} test records in polarion test runs: {polarion_config['test_runs']}")
    return test_runs



