"""
Copyright (c) 2018 Cisco and/or its affiliates.

This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.0 (the "License"). You may obtain a copy of the
License at

               https://developer.cisco.com/docs/licenses

All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

__author__ = "Chris McHenry"
__copyright__ = "Copyright (c) 2018 Cisco and/or its affiliates."
__license__ = "Cisco Sample Code License, Version 1.0"

import json
import argparse
import xlsxwriter
from tetpyclient import RestClient
import requests.packages.urllib3
from terminaltables import AsciiTable
import csv

API_ENDPOINT="{{TETRATION URL}}"
API_CREDS="{{PATH TO TETRATION API CREDS JSON}}"

def selectTetrationApps(endpoint,credentials):

    restclient = RestClient(endpoint,
                            credentials_file=credentials,
                            verify=False)

    requests.packages.urllib3.disable_warnings()
    resp = restclient.get('/openapi/v1/applications')

    if not resp:
        sys.exit("No data returned for Tetration Apps! HTTP {}".format(resp.status_code))

    app_table = []
    app_table.append(['Number','Name','Author','Primary'])
    for i,app in enumerate(resp.json()):
        app_table.append([i+1,app['name'],app['author'],app['primary']])
    print(AsciiTable(app_table).table)
    choice = raw_input('\nSelect Tetration App: ')

    choice = choice.split(',')
    appIDs = []
    for app in choice:
        if '-' in app:
            for app in range(int(app.split('-')[0])-1,int(app.split('-')[1])):
                appIDs.append(resp.json()[int(app)-1]['id'])
        else:
            appIDs.append(resp.json()[int(app)-1]['id'])
    return appIDs

def filterToString(invfilter):
    if 'filters' in invfilter.keys():
        query=[]
        for x in invfilter['filters']:
            if 'filters' in x.keys():
                query.append(filterToString(x))
            elif 'filter' in x.keys():
                query.append(x['type'] + filterToString(x['filter']))
            else:
                query.append(x['field'].replace('user_','*')+ ' '+ x['type'] + ' '+ str(x['value']))
        operator = ' '+invfilter['type']+' '
        return '('+operator.join(query)+')'
    else:
        return invfilter['field']+ ' '+ invfilter['type'] + ' '+ str(invfilter['value'])

def main():
    """
    Main execution routine
    """
    parser = argparse.ArgumentParser(description='Tetration Policy to XLS')
    parser.add_argument('--maxlogfiles', type=int, default=10, help='Maximum number of log files (default is 10)')
    parser.add_argument('--debug', nargs='?',
                        choices=['verbose', 'warnings', 'critical'],
                        const='critical',
                        help='Enable debug messages.')
    parser.add_argument('--config', default=None, help='Configuration file')
    args = parser.parse_args()
    apps = []
    if args.config is None:
        print '%% No configuration file given - connecting directly to Tetration'
        try:
            restclient = RestClient(API_ENDPOINT,credentials_file=API_CREDS,verify=False)
            appIDs = selectTetrationApps(endpoint=API_ENDPOINT,credentials=API_CREDS)
            for appID in appIDs:
                print('Downloading app details for '+appID)
                apps.append(restclient.get('/openapi/v1/applications/%s/details'%appID).json())
        except:
            print('Error connecting to Tetration')
    else:
        # Load in the configuration
        try:
            with open(args.config) as config_file:
                apps.append(json.load(config_file))
        except IOError:
            print '%% Could not load configuration file'
            return
        except ValueError:
            print 'Could not load improperly formatted configuration file'
            return

    # Load in the IANA Protocols
    protocols = {}
    try:
        with open('protocol-numbers-1.csv') as protocol_file:
            reader = csv.DictReader(protocol_file)
            for row in reader:
                protocols[row['Decimal']]=row
    except IOError:
        print '%% Could not load protocols file'
        return
    except ValueError:
        print 'Could not load improperly formatted protocols file'
        return

    for app in apps:
        workbook = xlsxwriter.Workbook('./'+app['name']+'.xlsx')
        bold = workbook.add_format({'bold': True})

        if 'clusters' in app.keys():
            worksheet = workbook.add_worksheet(name='App Servers')
            worksheet.set_row(0, None, bold)
            worksheet.write_row(0,0,['Hostname','IP','Cluster Membership'])
            i=1
            clusters = app['clusters']
            for cluster in clusters:
                hosts = []
                for node in cluster['nodes']:
                    hosts.append(node['name'])
                    worksheet.write_row(i,0,[node['name'],node['ip'],cluster['name']])
                    i+=1
            worksheet.set_column(0, 0, 30)
            worksheet.set_column(1, 1, 15)

        if 'inventory_filters' in app.keys():
            i=1
            worksheet = workbook.add_worksheet(name='External Groups')
            worksheet.set_row(0, None, bold)
            worksheet.write_row(0,0,['Inventory Filter Name','Filter Definition'])
            worksheet.set_column(0, 0, 30)

            filters = app['inventory_filters']
            for invfilter in filters:
                worksheet.write_row(i,0,[invfilter['name'],filterToString(invfilter['query'])])
                i+=1

        if 'default_policies' in app.keys():
            i=1
            worksheet = workbook.add_worksheet(name='Policies')
            worksheet.set_row(0, None, bold)
            worksheet.write_row(0,0,['Consumer Group','Provider Group','Services'])
            worksheet.set_column(0, 0, 30)
            worksheet.set_column(1, 1, 30)

            policies = app['default_policies']
            for policy in policies:
                pols = {}
                for rule in policy['l4_params']:
                    if 'port' in rule:
                        if rule['port'][0] == rule['port'][1]:
                            port = str(rule['port'][0])
                        else:
                            port = str(rule['port'][0]) + '-' + str(rule['port'][1])
                    else:
                        port = None

                    if port == None:
                        pols[protocols[str(rule['proto'])]['Keyword']] = []
                    elif protocols[str(rule['proto'])]['Keyword'] in pols.keys():
                        pols[protocols[str(rule['proto'])]['Keyword']].append(port)
                    else:
                        pols[protocols[str(rule['proto'])]['Keyword']] = [port]

                policy_list = []
                for key, val in pols.iteritems():
                    print(key,val)
                    if len(val)>0:
                        policy_list.append('{}={}'.format(key,', '.join(val)))
                    else:
                        policy_list.append(key)
                worksheet.write_row(i,0,[policy["consumer_filter_name"],policy["provider_filter_name"],'; '.join(policy_list)])
                i+=1
        
        workbook.close()

if __name__ == '__main__':
    main()
