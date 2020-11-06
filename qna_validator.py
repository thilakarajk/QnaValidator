import os
import re
import sys
import traceback
from time import sleep

import click

import openpyxl
import pandas as pd
import requests
from fuzzywuzzy import fuzz
from pandas.errors import ParserError
from progress.bar import ChargingBar
from requests.compat import urljoin

result_file_path = os.path.join(os.path.abspath(os.getcwd()), 'result.xlsx')

def prepare_query_request(hostname, authorization_key, knowledge_base, question):
    if hostname.endswith('/'):
        hostname = hostname[:-1]
    url = f'{hostname}/knowledgebases/{knowledge_base}/generateAnswer'
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'EndpointKey {authorization_key}'
    }
    query = {
        'question': question
    }
    return url, headers, query

def color_negative_red(val):
    color = 'red' if isinstance(val,str) and r"\n" in val else 'white'
    return 'background-color:%s' % color

@click.command()
@click.argument('filepath', type=click.Path(exists=True))
@click.option('--hostname', prompt='Hostname without trailing slash', help='Hostname of your QnA maker')
@click.option('--authkey', prompt='Authorization Key', help='Authorization key for Knowledge base')
@click.option('--knowledgebase', prompt='Knowledge Base', required=True)
@click.option('--confidencescore', prompt='Confidence Score')
def process(filepath, hostname, authkey, knowledgebase, confidencescore=75):
    try:
        if os.path.exists(result_file_path):
            result = input(
                f'Would you like to delete {result_file_path} (y/n): ')
            if result.lower() in ['y', 'yes', 'yeah', 'yup']:
                try:
                    os.remove(result_file_path)
                except:
                    print(
                        "Couldn't delete this file. delete manually and run this script")
                    sys.exit(1)
            else:
                print('Delete this file manually and run this script')
                sys.exit(0)
        try:
            data = pd.read_excel(filepath, header=0)
        except (ParserError, OSError, PermissionError) as e:
            if isinstance(e, PermissionError):
                print(f"Please close {filepath} and run this script.")
                sys.exit(1)
            if isinstance(e, ParserError):
                data = pd.read_csv(filepath, header=0)
        if os.sys.platform == 'win32':
            os.system('cls')
        else:
            os.system('clear')
        bar = ChargingBar('Processing',max=data.shape[0]-1)
        if fuzz.ratio('question', data.columns[0]) >= 80:
            question = data.columns[0]
            answer = data.columns[1]
        else:
            question = data.columns[1]
            answer = data.columns[2]
        data['Returned_response'] = None
        data['Confidence_score'] = None
        data['Pass/Fail'] = None
        for index, row in data.iterrows():
            url, headers, query = prepare_query_request(
                hostname, authkey, knowledgebase, row[question])
            response = requests.post(url=url, json=query, headers=headers)
            sleep(1)
            response = response.json()['answers']
            if response:
                data.loc[index, 'Returned_response'] = response[0]['answer']
                data.loc[index, 'Confidence_score'] = response[0]['score']
                pattern = re.compile(r'\s+', flags=re.MULTILINE)
                if response[0]['score'] != 0 and fuzz.ratio(re.sub(pattern, " ", row[answer].lower()),
                                                            re.sub(pattern, " ", response[0]['answer'].lower())) >= int(confidencescore):
                    data.loc[index, 'Pass/Fail'] = 'PASS'
                else:
                    data.loc[index, 'Pass/Fail'] = 'FAIL'
            bar.next()
        print("\nTest complete. Writing results in results.xlsx")
        data.style.applymap(color_negative_red).to_excel(result_file_path, engine='openpyxl',index=False)
    except Exception as e:
        print('=============================================')
        print(traceback.format_exc())
        print('=============================================')
        bar.finish()
        sys.exit(1)


if __name__ == "__main__":
    process()
