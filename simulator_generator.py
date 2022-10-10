import json
from base64 import b64encode
import argparse
import logging
import logging.handlers
import re
import sys
import csv
import traceback
import os
import time
import pandas as pd
import numpy as np  
import requests

import xlsxwriter
import xml.etree.ElementTree as ET
from utils.utils import RestUtils, AIPRestAPI, LogUtils, ObjectViolationMetric, RulePatternDetails, FileUtils, StringUtils, Metric, Contribution, Violation, ViolationFilter,\
    SnapshotFilter
from utils import excel_format

'''
 Author : MMR & TGU
 March 2020
'''
########################################################################

# Total Quality Index,Security,Efficiency,Robustness,Transferability,Changeability,Coding Best Practices/Programming Practices,Documentation,Architectural Design,Green IT
#      ISO-5055-Index, ISO-5055-Maintainability, ISO-5055-Performance-Efficiency, ISO-5055-Performance-Efficiency, ISO-5055-Security
bcids = ["60017","60016","60014","60013","60012","60011","66031","66032","66033","20140522",
         "1061000","1061001", "1061002", "1061003", "1061004"
         ]
broundgrades = False
########################################################################


########################################################################
def init_parse_argument():
    # get arguments
    parser = argparse.ArgumentParser(add_help=False)
    requiredNamed = parser.add_argument_group('Required named arguments')
    requiredNamed.add_argument('-restapiurl', required=True, dest='restapiurl', help='Rest API URL using format https://demo-eu.castsoftware.com/CAST-RESTAPI or http://demo-eu.castsoftware.com/Engineering')
    requiredNamed.add_argument('-edurl', required=False, dest='edurl', help='Engineering dashboard URL using format http://demo-eu.castsoftware.com/Engineering, if empty will be same as restapiurl')
    requiredNamed.add_argument('-user', required=False, dest='user', help='Username')    
    requiredNamed.add_argument('-password', required=False, dest='password', help='Password')
    requiredNamed.add_argument('-apikey', required=False, dest='apikey', help='Api key')
    requiredNamed.add_argument('-log', required=True, dest='log', help='log file')
    requiredNamed.add_argument('-of', required=True, dest='outputfolder', help='output folder')    
    requiredNamed.add_argument('-effortcsvfilepath', required=False, dest='effortcsvfilepath', help='Inputs quality rules effort csv file path (default=CAST_QualityRulesEffort.csv)')    
    requiredNamed.add_argument('-aggregationmode', required=False, dest='aggregationmode', help='Aggregation mode FullApplication/ByNumberOfArtifacts (default=FullApplication)')    
    requiredNamed.add_argument('-loadmodules', required=False, dest='loadmodules', help='Load the modules true/false default=false')
    requiredNamed.add_argument('-loadviolations', required=False, dest='loadviolations', help='Load the violations true/false default=false')
    requiredNamed.add_argument('-qridfilter', required=False, dest='qridfilter', help='For violations filtering, violation quality rule id regexp filter')
    requiredNamed.add_argument('-qrnamefilter', required=False, dest='qrnamefilter', help='For violations filtering, violation quality rule name regexp filter')
    requiredNamed.add_argument('-criticalrulesonlyfilter', required=False, dest='criticalrulesonlyfilter', help='For violations filtering, violation quality rules filter (True|False)')
    requiredNamed.add_argument('-businesscriterionfilter', required=False, dest='businesscriterionfilter', help='For violations filtering, business criterion filter : 60016,60012, ...)')
    requiredNamed.add_argument('-technofilter', required=False, dest='technofilter', help='For violations filtering, violation quality rule technology filter (JEE, SQL, HTML5, Cobol...)')    
    
    requiredNamed.add_argument('-applicationfilter', required=False, dest='applicationfilter', help='Application name regexp filter')
    requiredNamed.add_argument('-loglevel', required=False, dest='loglevel', help='Log level (INFO|DEBUG) default = INFO')
    requiredNamed.add_argument('-nbrows', required=False, dest='nbrows', help='max number of rows extracted from the rest API, default = 1000000000')
    requiredNamed.add_argument('-extensioninstallationfolder', required=False, dest='extensioninstallationfolder', help='extension installation folder')

    return parser
########################################################################

def get_excelfilepath(outputfolder, appName):
    fpath = ''
    if outputfolder != None:
        fpath = outputfolder
    fpath += appName + "_simulation.xlsx"
    return fpath 

########################################################################

def checkoutputfilelocked(logger, filepath):
    if FileUtils.is_file_locked_with_retries(logger, filepath):
        LogUtils.logerror(logger, 'File is locked. Aborting', True)
        return True
    return False

########################################################################
if __name__ == '__main__':

    global logger
    # load the data or just generate an empty excel file
    loaddata = True
    # load only 10 metrics
    loadonlyXmetrics = True    
    # round the grades or not


    parser = init_parse_argument()
    args = parser.parse_args()
    restapiurl = args.restapiurl
    if restapiurl != None and restapiurl[-1:] == '/':
        # remove the trailing / 
        restapiurl = restapiurl[:-1] 
    edurl = StringUtils.remove_trailing_suffix(restapiurl, 'rest') 
    # the engineering dashboard url can be different from the rest api url, but if not specified we will take the same value are rest api url
    if args.edurl != None:
        edurl = args.edurl
    user = 'N/A'
    if args.user != None: 
        user = args.user 
    password = 'N/A'
    if args.password != None: 
        password = args.password    
    apikey = 'N/A'
    if args.apikey != None: 
        apikey = args.apikey    
    log = args.log
    extensioninstallationfolder = "."
    if args.extensioninstallationfolder != None:
        extensioninstallationfolder = args.extensioninstallationfolder
    # add trailing / if not exist 
    if extensioninstallationfolder[-1:] != '/' and extensioninstallationfolder[-1:] != '\\' :
        extensioninstallationfolder += '/'
    
    outputfolder = args.outputfolder
    if not outputfolder.endswith('/') and not outputfolder.endswith('\\'):
        outputfolder += '/'
    
    aggregationmode = 'FullApplication'
    if args.aggregationmode != None and aggregationmode in ('FullApplication','ByNumberOfArtifacts'):
        aggregationmode = args.aggregationmode 
    
    effortcsvfilepath = "CAST_QualityRulesEffort.csv"
    if args.effortcsvfilepath != None:
        effortcsvfilepath = args.effortcsvfilepath 

    loadmodules = False
    if args.loadmodules != None and args.loadmodules in ('True','true'):
        loadmodules = True    

    loadviolations = False
    if args.loadviolations != None and args.loadviolations in ('True','true'):
        loadviolations = True                                  
    
    qridfilter = args.qridfilter
    qrnamefilter = args.qrnamefilter
    criticalrulesonlyfilter = False
    if args.criticalrulesonlyfilter != None and (args.criticalrulesonlyfilter == 'True' or args.criticalrulesonlyfilter == 'true'):
        criticalrulesonlyfilter = True
    businesscriterionfilter = args.businesscriterionfilter
    technofilter = args.technofilter

    # new params
    applicationfilter = args.applicationfilter
    loglevel = "INFO"
    if args.loglevel != None and (args.loglevel == 'INFO' or args.loglevel == 'DEBUG'):
        loglevel = args.loglevel
    csvfile = False
    nbrows = 1000000000
    if args.nbrows != None and type(nbrows) == int: 
        nbrows=args.nbrows

    ###########################################################################

    # setup logging
    logger = logging.getLogger(__name__)
    handler = logging.FileHandler(log, mode="w")
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    if loglevel == 'INFO':
        logger.setLevel(logging.INFO)
    elif loglevel == 'DEBUG':
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)

    try:
        # Version
        script_version = 'Unknown'
        try:
            pluginfile = extensioninstallationfolder + 'plugin.nuspec'
            LogUtils.logdebug(logger,pluginfile,True)
            tree = ET.parse(pluginfile)
            root = tree.getroot()
            namespace = "{http://schemas.microsoft.com/packaging/2011/08/nuspec.xsd}"
            for versiontag in root.findall('{0}metadata/{0}version'.format(namespace)):
                script_version = versiontag.text
        except:
            None 
        
        # log params
        logger.info('********************************************')
        LogUtils.loginfo(logger,'version='+script_version,True)
        logger.info('python version='+sys.version)
        logger.info('****************** params ******************')
        LogUtils.loginfo(logger, 'restapiurl='+restapiurl, True)
        LogUtils.loginfo(logger,'user='+str(user),True)
        if password == None or password == "N/A":
            logger.info('password=' + password)
        else: 
            logger.info('password=*******')
        if apikey == None or apikey== "N/A":
            logger.info('apikey=N/A')
        else:
            logger.info('apikey=*******') 
        LogUtils.loginfo(logger,'log file='+log,True)
        LogUtils.loginfo(logger, 'applicationfilter='+str(applicationfilter), True)
        LogUtils.loginfo(logger, 'output folder='+str(outputfolder), True)
        LogUtils.loginfo(logger,'loadmodules='+str(loadmodules), True)
        LogUtils.loginfo(logger, 'aggregationmode='+str(aggregationmode), True)
        logger.info('effortcsvfilepath='+str(effortcsvfilepath))
        logger.info('loadviolations='+str(loadviolations))
        logger.info('****************** params ******************')
        logger.info('edurl='+str(edurl))        
        logger.info('qridfilter='+str(qridfilter))
        logger.info('qrnamefilter='+str(qrnamefilter))
        logger.info('criticalrulesonlyfilter='+str(criticalrulesonlyfilter))
        logger.info('businesscriterionfilter='+str(businesscriterionfilter))
        logger.info('technofilter='+str(technofilter))
        logger.info('extensioninstallationfolder='+extensioninstallationfolder)
        logger.info('log level='+loglevel)
        logger.info('nbrows='+str(nbrows)) 
        logger.info('********************************************')
        
        if loadmodules and loadviolations:
            message = 'Aborting : please select one of the 2 options : loadmodules or loadviolations !'
            LogUtils.logerror(logger, message, True)
            raise ValueError(message) 
        
        LogUtils.loginfo(logger, 'Initialization', True) 
        rest_utils = RestUtils(logger, restapiurl, RestUtils.CLIENT_REQUESTS, user, password, apikey, uselocalcache=None, cachefolder=None, extensionid='com.castsoftware.uc.simulatorgenerator')
        rest_utils.open_session()
        rest_service_aip = AIPRestAPI(rest_utils) 
        
        # Few checks on the server 
        server = rest_service_aip.get_server()
        if server != None: logger.info('server version=%s, memory (free)=%s' % (str(server.version), str(server.freememory)))
        
        # retrieve the domains & the applications in those domains 
        # retrieve the domains & the applications in those domains 
        listdomains = rest_service_aip.get_domains()
        
        #json_domains = rest_service_aip.get_domains_json()
        bAEDdomainFound = False
        for it_domain in listdomains:
            if not it_domain.isAAD():
                bAEDdomainFound = True
                
        idomain = 0            
        for domain in listdomains:
            idomain += 1
            LogUtils.loginfo(logger, "Domain " + domain.name + " | progress:" + str(idomain) + "/" + str(len(listdomains)), True)
 
            # only engineering domains, or AAD domain only in case there is no engineering domain, we prefer to have engineering domains containing of action plan summary
            if domain.name == 'AAD' and bAEDdomainFound:
                logger.info("  Skipping domain " + domain.name + ", because we process in priority Engineering domains")
                continue
                
            if domain.name != 'AAD' or not bAEDdomainFound:
                listapplications = rest_service_aip.get_applications(domain)
                for objapp in listapplications:
                    if applicationfilter != None and not re.match(applicationfilter, objapp.name):
                        logger.info('Skipping application : ' + objapp.name)
                        continue                
                    elif applicationfilter == None or applicationfilter == '' or re.match(applicationfilter, objapp.name):
                        LogUtils.loginfo(logger, "Processing application " + objapp.name, True)
                        # testing if csv file can be written
                        fpath = get_excelfilepath(outputfolder, objapp.name)
                        # if the output file is locked we move to next application
                        if checkoutputfilelocked(logger, fpath):
                            continue

                        listbusinesscriteria = []
                        dicremediationabacus = {}
                        # applications health factors for last snapshot
                        if loaddata:
                            logger.info('Loading the remediation effort from file ' + effortcsvfilepath)
                            csvdelimiter = ";"
                            csvquotechar='"'
                            if not os.path.exists(effortcsvfilepath):
                                logger.warning('File ' + effortcsvfilepath + ' do not exist ! Remediation efforts will not be loaded.')
                            else:
                                with open(effortcsvfilepath, newline='') as infile:
                                    reader = csv.reader(infile,delimiter=csvdelimiter,quotechar=csvquotechar)
                                    for row in reader:
                                        effortqrname = ''
                                        try:
                                            # remove unicode characters
                                            effortqrname = StringUtils.remove_unicode_characters(row[1])
                                            
                                        except UnicodeDecodeError: 
                                            logger.error('Non UTF-8 character in the row [' + str(row) + '] of the csv file')
                                            effortqrname =  'Non UTF-8 quality rule name'

                                        dicremediationabacus.update({row[0]:{"id":row[0],"name":effortqrname,"uniteffortinhours":row[2]}})
                        # snapshot list
                        logger.info('Loading the application snapshot')
                        listsnapshots = rest_service_aip.get_application_snapshots(domain.name, objapp.id)
                        
                        for objsnapshot in listsnapshots:
                            logger.info("    Snapshot " + objsnapshot.href + '#' + objsnapshot.snapshotid)
                            listbccontributions = []
                            listtccontributions = []
                            dictapsummary = {}
                            dictaptriggers = {}
                            tqiqm = {}
                            dictmetrics = {}
                            dicttechnicalcriteria = {}                            
                            
                            try:
                                if not loaddata:
                                    logger.info("NOT Extracting the snapshot quality model")                                           
                                else:
                                    logger.info("Extracting the snapshot quality model")
                                    tqiqm = rest_service_aip.get_snapshot_tqi_quality_model(domain.name, objsnapshot.snapshotid)
                            except:
                                LogUtils.logwarning(logger, 'Not able to extract the TQI quality model ***',True)

                            dict_modulesweight = None
                            if aggregationmode != None and aggregationmode != 'FullApplication':
                                try:
                                    if not loaddata:
                                        logger.info("NOT Extracting the modules weight")                                           
                                    else:
                                        logger.info("Extracting the modules weight")
                                        if aggregationmode == 'ByNumberOfArtifacts':
                                            snapshotsfilter = SnapshotFilter(None, objsnapshot.snapshotid)
                                            dict_modulesweight = rest_service_aip.get_nb_artifacts_dict(domain.name, snapshotsfilter, "$all")
                                except:
                                    LogUtils.logwarning(logger, 'Not able to extract the  modules weight ***',True)
                                    
                            try:
                                if loaddata:
                                    logger.info("Extracting the action plan summary")
                                    dictapsummary = rest_service_aip.get_actionplan_summary(domain.name, objapp.id, objsnapshot.snapshotid)
                            except:
                                LogUtils.logwarning(logger, 'Not able to extract the action plan summary ***',True)                                        
                            try:                                     
                                if loaddata:
                                    logger.info("Extracting the action plan triggers")
                                    dictaptriggers = rest_service_aip.get_actionplan_triggers(domain.name, objapp.id, objsnapshot.snapshotid)
                            except:
                                LogUtils.logwarning(logger, 'Not able to extract the action plan triggers ***',True)
                            
                            if loaddata:
                                logger.info('Extracting the quality results per application and per module (quality-rules, quality-measures, quality-distributions, technical-criteria, business-criteria) ')
                                strmodules_filter = None
                                if loadmodules: 
                                    strmodules_filter = "$all" 
                                dictmetrics, dicttechnicalcriteria, listbusinesscriteria, dictmodules = rest_service_aip.get_qualitymetrics_results(domain.name, objapp.id, objsnapshot.snapshotid, tqiqm=tqiqm, criticalonly=False, modules=strmodules_filter, aggregationmode=aggregationmode, nbrows=nbrows)                                
                                
                                logger.info('Extracting the technical criteria contributors at application level')
                                #TODO:Keep this section to collect the information for the old versions of Rest API, because in the latest versions select=aggregators in the previous REST call is already providing this 
                                for item_tc in dicttechnicalcriteria:
                                    tciterator = dicttechnicalcriteria[item_tc]
                                    for item in rest_service_aip.get_metric_contributions(domain.name, tciterator.id, objsnapshot.snapshotid):
                                        listtccontributions.append(item)

                                logger.info('Extracting the business criteria contributors at application level')
                                for bcid in bcids:
                                    for item in rest_service_aip.get_metric_contributions(domain.name, bcid, objsnapshot.snapshotid):
                                        listbccontributions.append(item)
                                    index = 0 
                                    for cont in listbccontributions:
                                        # if no results for this technical criteria, we remove it from the list
                                        if dicttechnicalcriteria.get(cont.metricid) == None:
                                            del listbccontributions[index]
                                        index += 1
                            
                            json_violations = None
                            listviolations = []
                            if loaddata and loadviolations: 
                                LogUtils.loginfo(logger,'Extracting violations',True)
                                LogUtils.loginfo(logger,'Loading violations from the REST API',True)
                                violationfilter = ViolationFilter(criticalrulesonlyfilter, businesscriterionfilter, technofilter, None, qridfilter, qrnamefilter, nbrows)
                                listviolations = rest_service_aip.get_snapshot_violations(domain.name, objapp.id, objsnapshot.snapshotid, edurl,objsnapshot.href, tqiqm, listtccontributions, violationfilter)                                            
                            # generate output file                                    
                            fpath = get_excelfilepath(outputfolder, objapp.name)
                            LogUtils.loginfo(logger,"Generating xlsx file " + fpath,True)
                            
                            excel_format.generate_excelfile(logger, fpath, objapp.name, objsnapshot.version, objsnapshot.isodate, loadviolations, listbusinesscriteria, dicttechnicalcriteria, listbccontributions, listtccontributions, dictmetrics, dictapsummary, dicremediationabacus, listviolations, broundgrades, dictaptriggers, dictmodules, dict_modulesweight, aggregationmode)

                            # keep only last snapshot
                            break
                                        
    except: # catch *all* exceptions
        tb = traceback.format_exc()
        #e = sys.exc_info()[0]
        LogUtils.logerror(logger, '  Error during the processing %s' % tb,True)

    LogUtils.loginfo(logger,'Done !',True)