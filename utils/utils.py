import requests
from base64 import b64encode
import re
import os
import sys
import time

'''
Created on 13 avr. 2020

@author: MMR
'''

####################################################################################

class StringUtils:
    @staticmethod 
    def NonetoEmptyString(obj):
        if obj == None or obj == 'None':
            return ''
        return obj

    ########################################################################
    # workaround to remove the unicode characters before sending them to the CSV/Excel file
    # and avoid the below error
    #UnicodeEncodeError: 'charmap' codec can't encode character '\x82' in position 105: character maps to <undefined>    
    @staticmethod    
    def remove_unicode_characters(astr):
        return astr.encode('ascii', 'ignore').decode("utf-8")


####################################################################################


class FileUtils:
    
    """Checks if a file is locked by opening it in append mode.
    If no exception thrown, then the file is not locked.
    """
    @staticmethod
    def is_file_locked_with_retries(logger, filepath):
        filelocked = False
        icount = 0
        while icount < 10 and FileUtils.is_file_locked(filepath):
            icount += 1
            filelocked = True
            LogUtils.logwarning(logger,'File %s is locked. Please unlock it ! Waiting 5 seconds before retrying (try %s/10) ' % (filepath, str(icount)),True)
            time.sleep(5)
        if not FileUtils.is_file_locked(filepath):
            filelocked = False
        return filelocked        
    
    """Checks if a file is locked by opening it in append mode.
    If no exception thrown, then the file is not locked.
    """
    @staticmethod
    def is_file_locked(filepath):
    
        locked = False
        file_object = None
        if os.path.exists(filepath):
            try:
                #print ("Trying to open %s." % filepath)
                buffer_size = 8
                # Opening file in append mode and read the first 8 characters.
                file_object = open(filepath, 'a', buffer_size)
                if file_object:
                    #print ("%s is not locked." % filepath)
                    locked = False
            
            except IOError:
                e = sys.exc_info()[0]
                #print ("File is locked (unable to open in append mode). %s." % e)
                locked = True
            finally:
                if file_object:
                    file_object.close()
                    #print ("%s closed." % filepath)
        #else:
        #    print "%s not found." % filepath
        return locked


####################################################################################

class RestUtils:
    
    def __init__(self, logger, url, user, password, apikey = None): 
        self.session_cookie = None
        self.session = None
        self.logger = logger
        self.url = url
        self.user = user
        self.password = password
        self.apikey = apikey
    
    ####################################################################################################
    
    def get_response_cookies(self, response):
        return response.headers._store 

    ####################################################################################################
        
    # get the response session cookie containing JSESSION
    def get_session_cookie(self, response):
        session_cookie = None
        response_cookies = self.get_response_cookies(response)
        if response_cookies != None:
            sc = response_cookies.get('set-cookie')
            if sc != None and sc[1]  != None:
                session_cookie = sc[1]        
        return session_cookie
    
    ####################################################################################################
    # retrieve the connection
    def open_session(self):
        self.logger.info('Opening session to ' + self.url)
        response = None
        request_headers = {}
        request_headers.update({"User-Agent": "XY"})        
        request_headers.update({'X-Client':'com.castsoftware.uc.violationextraction'})
                
        try:
            self.session = requests.Session()
            if self.user != None and self.password != None and self.user != 'N/A' and self.user != 'N/A':
                self.logger.info ('Using user and password')
                #we need to base 64 encode it 
                #and then decode it to acsii as python 3 stores it as a byte string
                #userAndPass = b64encode(user_password).decode("ascii")
                auth = str.encode("%s:%s" % (self.user, self.password))
                #user_and_pass = b64encode(auth).decode("ascii")
                user_and_pass = b64encode(auth).decode("iso-8859-1")
                request_headers.update({'Authorization':'Basic %s' %  user_and_pass})
            # else if the api key is provided
            elif self.apikey != None and self.apikey != 'N/A':
                self.logger.info ('Using api key')
                # API key configured in the WAR
                request_headers.update({'X-API-KEY':self.apikey})
                # we are provide a user name hardcoded' 
                request_headers.update({'X-API-USER':'admin_apikey'})            
            
            response = self.session.get(self.url, headers=request_headers)
            
        except:
            self.logger.error ('Error connecting to ' + self.url)
            self.logger.error ('URL is not reachable. Please check your connection (web application down, VPN not active, ...)')
            raise SystemExit
        #finally:
        #    self.logger.info ('Headers = ' + str(response.headers))
            
        if response.status_code != 200:
            # This means something went wrong.
            self.logger.error ('Error connecting to ' + self.url)
            self.logger.error ('Status code = ' + str(response.status_code))
            self.logger.error ('Headers = ' + str(response.headers))
            self.logger.error ('Please check the URL, user and password or api key')
            raise SystemExit
        else: 
            self.logger.info ('Successfully connected to  : ' + self.url)    
    
    ####################################################################################################

    def get_default_http_headers(self):
        # User agent & Name of the client added in the header (for the audit trail)
        default_headers = {"User-Agent": "XY", "X-Client":"com.castsoftware.uc.simulatorgenerator"} 
        return default_headers

    ####################################################################################################
    def execute_request(self, requesttype, request, contenttype='application/json', inputjson=None):
        if self.session == None:
            self.open_session()
            
        if request == None:
            request_text = self.url
        else:
            request_text = self.url + "/rest/" + request
        self.logger.debug('Sending ' + requesttype + ' ' + request_text + ' with contenttype=' + contenttype)   
        
        request_headers = {}
        request_headers.update(self.get_default_http_headers())
        request_headers.update({'accept' : contenttype})
    
        # send the request
        if 'GET' == requesttype:
            response = self.session.get(request_text,headers=request_headers)
        elif 'POST' == requesttype:
            response = self.session.post(request_text,inputjson,headers=request_headers)
        elif 'PUT' == requesttype:
            response = self.session.post(request_text,inputjson,headers=request_headers)        
        elif 'DELETE' == requesttype:
            response = self.session.post(request_text,inputjson,headers=request_headers)    
        else:
            LogUtils.logerror(self.logger,'Invalid HTTP request type' + requesttype)
        
        output = None
        if response != None: 
            # Error
            if response.status_code != 200:
                LogUtils.logerror(self.logger,'HTTPS request failed ' + str(response.status_code) + ' :' + request_text,True)
                return None
            else:
                # get the session cookie containing JSESSION
                # look for the Set-Cookie in response headers, to inject it for future requests
                session_cookie = self.get_session_cookie(response)
                if session_cookie != None:
                    # copy the session cookie
                    self.session_cookie = session_cookie
                    #print('3='+session_cookie)
                
                if contenttype == 'application/json':
                    output = response.json()
                else:
                    output = response.text
    
        return output 
    
    ####################################################################################################
    
    def execute_request_get(self, request, contenttype='application/json'):
        return self.execute_request('GET', request, contenttype)
    
    ####################################################################################################
    
    def execute_request_post(self, request, contenttype='application/json', inputjson=None):
        return self.execute_request('POST', request, contenttype)
    
    ####################################################################################################
    
    def execute_request_put(self, request, contenttype='application/json', inputjson=None):
        return self.execute_request('PUT', request, contenttype)
    
    ####################################################################################################
    
    def execute_request_delete(self, request, contenttype='application/json', inputjson=None):
        return self.execute_request('DELETE', request, contenttype)

########################################################################

# CAST AIP Dahshboard REST API 
class AIPRestAPI:
    FILTER_ALL = "$all"
    FILTER_SNAPSHOTS_ALL = FILTER_ALL
    FILTER_SNAPSHOTS_LAST = "-1"
    FILTER_SNAPSHOTS_LAST_TWO = "-2"
    
    ########################################################################
    
    def __init__(self, restutils): 
        self.restutils = restutils

        
    ########################################################################

    def get_server(self):
        request = "server"
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_domains(self):
        request = ""
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_applications(self, domain):
        request = domain + "/applications"
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_transactions_per_business_criterion(self, domain, applicationid, snapshotid, bcid, nbrows):
        self.restutils.logger.info("Extracting the transactions for business criterion " + bcid)
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/transactions/" + bcid
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)    
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_application_snapshots(self, domain, applicationid):
        request = domain + "/applications/" + applicationid + "/snapshots" 
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_total_number_violations(self, domain, applicationid, snapshotid):
        self.restutils.logger.info("Extracting the number of violations")
        request = domain + "/results?sizing-measures=67011,67211&application=" + applicationid + "&snapshot=" + snapshotid
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_qualitydistribution_details(self, domain, applicationid, snapshotid, metricid, category, nbrows):
        request = domain + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + '/components/' + str(metricid) + '/'
        request += str(category)+'?business-criterion=60017&startRow=1&nbRows=' + str(nbrows)
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_dict_cyclomaticcomplexity_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        #Very High Complexity Artifacts
        categories = [1,2,3,4]
        labels = {1:'Very High Complexity Artifacts',2:'High Complexity Artifacts',3:'Moderate Complexity Artifacts',4:'Low Complexity Artifacts'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_CYCLOMATIC_COMPLEXITY, cat, nbrows)
            icount = 0
            if json != None:
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Cyclomatic complexity distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict

    ########################################################################
    def get_dict_costcomplexity_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        categories = [1,2,3,4]
        labels = {1:'Very High Complexity',2:'High Complexity',3:'Moderate Complexity',4:'Low Complexity'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_COST_COMPLEXITY, cat, nbrows)
            icount = 0
            if json != None:
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Cost complexity distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict

    ########################################################################
    def get_dict_fanout_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        categories = [1,2,3,4]
        labels = {1:'Very High Fan-Out classes',2:'High Fan-Out classes',3:'Moderate Fan-Out classes',4:'Low Fan-Out classes'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_FAN_OUT, cat, nbrows)
            icount = 0
            if json != None:            
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Fan-Out classes distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict

    ########################################################################
    def get_dict_fanin_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        #Very High Fan-In classes
        categories = [1,2,3,4]
        labels = {1:'Very High Fan-In classes',2:'High Fan-In classes',3:'Moderate Fan-In classes',4:'Low Fan-In classes'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_FAN_IN, cat, nbrows)
            icount = 0
            if json != None:            
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Fan-In classes distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict

    ########################################################################
    def get_dict_coupling_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        #Very High Coupling Artifacts
        categories = [1,2,3,4]
        labels = {1:'Very High Coupling Artifacts',2:'High Coupling Artifacts',3:'Average Coupling Artifacts',4:'Low Coupling Artifacts'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_COUPLING, cat, nbrows)
            icount = 0
            if json != None:                 
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Coupling distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict

    ########################################################################
    def get_dict_size_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        categories = [1,2,3,4]
        # Very Large Size Artifacts
        labels = {1:'Very Large Size Artifacts',2:'Large Size Artifacts',3:'Average Size Artifacts',4:'Small Size Artifacts'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_SIZE, cat, nbrows)
            icount = 0
            if json != None:
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'Size distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict
    ########################################################################
    def get_dict_SQLcomplexity_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        categories = [1,2,3,4]
        # Very Large Size Artifacts
        labels = {1:'Very High SQL Complexity Artifacts',2:'High SQL Complexity Artifacts',3:'Moderate SQL Complexity Artifacts',4:'Low Complexity Artifacts'}
        for cat in categories:
            json = self.get_qualitydistribution_details(domain, applicationid, snapshotid, Metric.DIST_SQL_COMPLEXITY, cat, nbrows)
            icount = 0
            if json != None:            
                for it in json:
                    icount += 1
                    dict.update({it['href']:labels.get(cat)})
                LogUtils.loginfo(self.restutils.logger, 'SQL complexity distribution cat ' + str(cat) + ' : ' + str(icount), False)
        return dict
    ########################################################################
    def get_distributions_details(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        # cyclomatic complexity dist.
        dict.update({Metric.DIST_CYCLOMATIC_COMPLEXITY:self.get_dict_cyclomaticcomplexity_distribution(domain, applicationid, snapshotid, nbrows)})
        # cost complexity dist.
        dict.update({Metric.DIST_COST_COMPLEXITY:self.get_dict_costcomplexity_distribution(domain, applicationid, snapshotid, nbrows)})        
        # fan in dist.
        dict.update({Metric.DIST_FAN_IN:self.get_dict_fanin_distribution(domain, applicationid, snapshotid, nbrows)})
        # fan out dist.
        dict.update({Metric.DIST_FAN_OUT:self.get_dict_fanout_distribution(domain, applicationid, snapshotid, nbrows)})
        # size dist.
        dict.update({Metric.DIST_SIZE:self.get_dict_size_distribution(domain, applicationid, snapshotid, nbrows)})
        # coupling dist.
        dict.update({Metric.DIST_COUPLING:self.get_dict_coupling_distribution(domain, applicationid, snapshotid, nbrows)})
        # SQL dist.
        dict.update({Metric.DIST_SQL_COMPLEXITY:self.get_dict_SQLcomplexity_distribution(domain, applicationid, snapshotid, nbrows)})        
        
        return dict

    ########################################################################
    def get_snapshot_violations(self, domain, applicationid, snapshotid, criticalonly, violationStatus, businesscriterionfilter, technoFilter, nbrows):
        self.restutils.logger.info("Extracting the snapshot violations")
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/violations'
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)
        if criticalonly != None and criticalonly:         
            request += '&rule-pattern=critical-rules'
        if violationStatus != None:
            request += '&status=' + violationStatus
        if businesscriterionfilter == None:
            businesscriterionfilter = "60017"
        if businesscriterionfilter != None:
            strbusinesscriterionfilter = str(businesscriterionfilter)        
            # we can have multiple value separated with a comma
            if ',' not in strbusinesscriterionfilter:
                request += '&business-criterion=' + strbusinesscriterionfilter
            request += '&rule-pattern=('
            for item in strbusinesscriterionfilter.split(sep=','):
                request += 'cc:' + item + ','
                if criticalonly == None or not criticalonly:   
                    request += 'nc:' + item + ','
            request = request[:-1]
            request += ')'
            
        if technoFilter != None:
            request += '&technologies=' + technoFilter
            
        return self.restutils.execute_request_get(request)
        
    ########################################################################
    def get_tqi_transactions_violations(self, domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows):    
        request = domain + "/transactions/" + transactionid + "/snapshots/" + snapshotid + '/violations'
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)
        if criticalonly != None and criticalonly:         
            request += '&rule-pattern=critical-rules'
        if violationStatus != None:
            request += '&status=' + violationStatus
        
        businesscriterionfilter = "60017"
        if businesscriterionfilter != None:
            strbusinesscriterionfilter = str(businesscriterionfilter)        
            # we can have multiple value separated with a comma
            if ',' not in strbusinesscriterionfilter:
                request += '&business-criterion=' + strbusinesscriterionfilter
            request += '&rule-pattern=('
            for item in strbusinesscriterionfilter.split(sep=','):
                request += 'cc:' + item + ','
                if criticalonly == None or not criticalonly:   
                    request += 'nc:' + item + ','
            request = request[:-1]
            request += ')'
            
        if technoFilter != None:
            request += '&technologies=' + technoFilter
            
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def create_scheduledexclusions(self, domain, applicationid, snapshotid, json_exclusions_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/exclusions/requests'
        return self.restutils.execute_request_post(request, 'application/json',json_exclusions_to_create)
    
    ########################################################################
    def create_actionplans(self, domain, applicationid, snapshotid, json_actionplans_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/action-plan/issues'
        return self.restutils.execute_request_post(request, 'application/json',json_actionplans_to_create)

    ########################################################################
    def get_rule_pattern(self, rulepatternHref):
        self.restutils.logger.debug("Extracting the rule pattern details")   
        request = rulepatternHref
        json_rulepattern =  self.restutils.execute_request_get(request)
        obj = None
        if json_rulepattern != None:
            obj = RulePatternDetails()    
            try:
                obj.associatedValueName = json_rulepattern['associatedValueName']
            except KeyError:
                None
            try:
                qslist = json_rulepattern['qualityStandards']
                for qs in qslist:
                    obj.listQualityStandard.append(qs['standard']+"/"+qs['id'])
            except KeyError:
                None
        return obj 

    ########################################################################
    def get_actionplan_summary(self, domain, applicationid, snapshotid):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/action-plan/summary"
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_qualitymetrics_results(self, domain, applicationid, criticalonly, nbrows):
        LogUtils.loginfo(self.restutils.logger,'Extracting the quality metrics results',True)
        request = domain + "/applications/" + applicationid + "/results?quality-indicators"
        request += '=(cc:60017'
        if criticalonly == None or not criticalonly:   
            request += ',nc:60017'
        request += ')&select=(evolutionSummary,violationRatio)'
        # last snapshot only
        request += '&snapshots=-1'
        request += '&startRow=1'
        request += '&nbRows=' + str(nbrows)
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_all_snapshots(self, domain):
        request = domain + "/results?snapshots=" + AIPRestAPI.FILTER_SNAPSHOTS_ALL
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_loc(self, domain, snapshotsfilter=None):
        return self.get_sizing_measures_by_id(domain, Metric.TS_LOC, snapshotsfilter)
  
    ########################################################################
    def get_afp(self, domain, snapshotsfilter=None):
        return self.get_sizing_measures_by_id(domain, Metric.FS_AFP, snapshotsfilter)
  
    ########################################################################
    def get_tqi(self, domain, snapshotsfilter=None):
        return self.get_quality_indicators_by_id(domain, Metric.BC_TQI, snapshotsfilter)  
    
    ########################################################################
    def get_sizing_measures_by_id(self, domain, metricids, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?sizing-measures=(" + str(metricids) + ")&snapshots="+snapshotsfilter
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_quality_distribution_by_id_as_json(self, domain, metricids, snapshotsfilter=None):
        json_cost_complexity = self.get_quality_indicators_by_id(domain, metricids, snapshotsfilter)
        json_snapshots = {}
        if json_cost_complexity != None:
            for metric in json_cost_complexity:
                snapshothref = metric['applicationSnapshot']['href']
                if json_snapshots.get(snapshothref) == None:
                    json_snapshots[snapshothref] = {}
                for appresults in metric['applicationResults']:
                    json_distr = {}
                    key = appresults['reference']['key']
                    categories = None
                    try:
                        categories = appresults['result']['categories']
                    except KeyError:
                        None
                    if categories != None:
                        icount = 0
                        NbVeryHigh = 0.0
                        NbHigh = 0.0
                        NbAverage = 0.0
                        NbLow = 0.0
                        for c in categories:
                            icount += 1
                            if icount == 1: very_high = c
                            elif icount == 2: high = c
                            elif icount == 3: average = c
                            elif icount == 4: low = c
                        try: NbVeryHigh=very_high.get('value') 
                        except: None                                 
                        try: NbHigh=high.get('value')
                        except: None
                        try: NbAverage=average.get('value')
                        except: None
                        try: NbLow=low.get('value')
                        except: None
                        total = NbVeryHigh+NbHigh+NbAverage+NbLow
                        PercentVeryHigh=0.0
                        PercentHigh=0.0
                        PercentAverage=0.0
                        PercentLow=0.0                          
                        if total > 0:
                            PercentVeryHigh = NbVeryHigh / total
                            PercentHigh = NbHigh / total
                            PercentAverage = NbAverage / total
                            PercentLow = NbLow / total
                        json_distr = {key+"_NbVeryHigh":NbVeryHigh,key+"_NbHigh":NbHigh,key+"_NbAverage":NbAverage,
                                                            key+"_NbLow":NbLow,key+"_PercentVeryHigh":PercentVeryHigh,key+"_PercentHigh":PercentHigh,
                                                            key+"_PercentAverage":PercentAverage,key+"_PercentLow":PercentLow}       
                        json_snapshots.get(snapshothref).update(json_distr) 
        return json_snapshots

    ########################################################################
    def get_sizing_measures(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?sizing-measures=(technical-size-measures,technical-debt-statistics,run-time-statistics,functional-weight-measures)&snapshots="+snapshotsfilter
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_quality_indicators(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?quality-indicators=(quality-rules,business-criteria,technical-criteria,quality-distributions,quality-measures)&select=(evolutionSummary,violationRatio,aggregators,categories)&snapshots="+snapshotsfilter
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_quality_indicators_by_id(self, domain, metricids, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?quality-indicators=(" + str(metricids) + ")&select=(evolutionSummary,violationRatio,aggregators,categories)&snapshots="+snapshotsfilter
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_metric_contributions(self, domain, metricid, snapshotid):
        request = domain + "/quality-indicators/" + str(metricid) + "/snapshots/" + snapshotid 
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_qualityrules_thresholds(self, domain, snapshotid, qrid):
        #LogUtils.loginfo(self.restutils.logger,'Extracting the quality rules thresholds',True)
        request = domain + "/quality-indicators/" + str(qrid) + "/snapshots/"+ snapshotid
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_businesscriteria_grades(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST        
        request = domain + "/results?quality-indicators=(60017,60016,60014,60013,60011,60012,66031,66032,66033)&applications=($all)&snapshots=" + snapshotsfilter 
        return self.restutils.execute_request_get(request)

    ########################################################################
    def get_snapshot_tqi_quality_model (self, domain, snapshotid):
        self.restutils.logger.info("Extracting the snapshot quality model")   
        request = domain + "/quality-indicators/60017/snapshots/" + snapshotid + "/base-quality-indicators" 
        return self.restutils.execute_request_get(request)
    ########################################################################
    def get_snapshot_bc_tc_mapping(self, domain, snapshotid, bcid):
        self.restutils.logger.info("Extracting the snapshot business criterion " + bcid +  " => technical criteria mapping")  
        request = domain + "/quality-indicators/" + str(bcid) + "/snapshots/" + snapshotid
        return self.restutils.execute_request_get(request)
    
    ########################################################################
    def get_components_pri (self, domain, applicationid, snapshotid, bcid,nbrows):
        self.restutils.logger.info("Extracting the components PRI for business criterion " + bcid)  
        request = domain + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + '/components/' + str(bcid)
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)
        return self.restutils.execute_request_get(request)
         
    ########################################################################
    def get_sourcecode(self, sourceCodesHref):
        return self.restutils.execute_request_get(sourceCodesHref)     
    
    ########################################################################
    def get_sourcecode_file(self, filehref, srcstartline, srcendline):
        strstartendlineparams = ''
        if srcstartline != None and srcstartline >= 0 and srcendline != None and srcendline >= 0:
            strstartendlineparams = '?start-line='+str(srcstartline)+'&end-line='+str(srcendline)
        return self.restutils.execute_request_get(filehref+strstartendlineparams, 'text/plain')    
    
    ########################################################################
    def get_objectviolation_metrics(self, objHref):
        self.restutils.logger.debug("Extracting the component metrics")    
        request = objHref
        json_component = self.restutils.execute_request_get(request)
        obj = None
        if json_component != None:
            obj = ObjectViolationMetric()
            try:
                obj.componentType = json_component['type']['label']
            except KeyError:
                None                              
            try:
                obj.codeLines = json_component['codeLines']
                if obj.codeLines == None :  obj.codeLines = 0
            except KeyError:
                obj.codeLines = 0
            try:
                obj.commentedCodeLines = json_component['commentedCodeLines']
                if obj.commentedCodeLines == None :  obj.commentedCodeLines = 0
            except KeyError:
                obj.commentedCodeLines = 0                                          
            try:
                obj.commentLines = json_component['commentLines']
                if obj.commentLines == None :  obj.commentLines = 0
            except KeyError:
                obj.commentLines = 0
                  
            try:
                obj.fanIn = json_component['fanIn']
            except KeyError:
                obj.fanIn = 0
            try:
                obj.fanOut = json_component['fanOut']
            except KeyError:
                obj.fanOut = 0 
            try:
                obj.cyclomaticComplexity = json_component['cyclomaticComplexity']
            except KeyError:
                obj.cyclomaticComplexity ='Not available'   
            #Incorrect ratio, recomputing manually
            #try:
            #    obj.ratioCommentLinesCodeLines = json_component['ratioCommentLinesCodeLines']
            #except KeyError:
            #    obj.ratioCommentLinesCodeLines = None
    
            obj.ratioCommentLinesCodeLines = None
            if obj.codeLines != None and obj.commentLines != None and obj.codeLines != 'Not available' and obj.commentLines != 'Not available' and (obj.codeLines + obj.commentLines != 0) :
                obj.ratioCommentLinesCodeLines = obj.commentLines / (obj.codeLines + obj.commentLines) 
            else:
                obj.ratioCommentLinesCodeLines = 0
            try:
                obj.halsteadProgramLength = json_component['halsteadProgramLength']
            except KeyError:
                obj.halsteadProgramLength = 0
            try:
                obj.halsteadProgramVocabulary = json_component['halsteadProgramVocabulary']
            except KeyError:
                obj.halsteadProgramVocabulary = 0
            try:
                obj.halsteadVolume = json_component['halsteadVolume']
            except KeyError:
                obj.halsteadVolume = 0 
            try:
                obj.distinctOperators = json_component['distinctOperators']
            except KeyError:
                obj.distinctOperators = 0 
            try:
                obj.distinctOperands = json_component['distinctOperands']
            except KeyError:
                obj.distinctOperands = 0                                             
            try:
                obj.integrationComplexity = json_component['integrationComplexity']
            except KeyError:
                obj.integrationComplexity = 0
            try:
                obj.criticalViolations = json_component['criticalViolations']
            except KeyError:
                obj.criticalViolations = 'Not available'
    
        return obj
        
    
    ########################################################################
    def get_objectviolation_findings(self, objHref, qrid):
        self.restutils.logger.debug("Extracting the component findings")    
        request = objHref + '/findings/' + qrid 
        return self.restutils.execute_request_get(request)
         
    ########################################################################
    # extract the transactions TRI & violations component list per business criteria
    def init_transactions (self, domain, applicationid, snapshotid, criticalonly, violationStatus, technoFilter,nbrows):
        #Security,Efficiency,Robustness,TQI
        bcids = ["60017","60016","60014","60013"]
        transaclist = {}
        for bcid in bcids: 
            json_transactions = self.get_transactions_per_business_criterion(domain, applicationid, snapshotid, bcid, nbrows)
            if json_transactions != None:
                transaclist[bcid] = []
                icount = 0
                for trans in json_transactions:
                    icount += 1
                    tri = None
                    shortname = 'Undefined'
                    name = 'Undefined'
                    transactionHref = 'Undefined'
                    transactionid = -1 
                    
                    try:
                        name =  trans['name']
                    except KeyError:
                        None  
                    try:
                        tri =  trans['transactionRiskIndex']
                    except KeyError:
                        None
                    try:
                        transactionHref =  trans['href']
                    except KeyError:
                        None
                    rexuri = "/transactions/([0-9]+)/"
                    m0 = re.search(rexuri, transactionHref)
                    if m0: transactionid = m0.group(1)
                    try:
                        shortName =  trans['shortName']
                    except KeyError:
                        None 
                        
                    mytransac = {
                        "name":name,
                        "shortName":shortName,
                        "href":transactionHref,
                        "business criteria id":bcid,
                        "transactionRiskIndex":tri,
                        "componentsWithViolations":[]
                    }
                    
                    # for TQI only we retrieve the list of components in violations on that transaction
                    # for the other we need only the transaction TRI
                    json_tran_violations = None
                    # look for the transaction violation only for the TQI, for the other HF take the violation already extracted for the TQI  
                    if bcid == "60017":
                        self.restutils.logger.info("Extracting the violations for transaction " + transactionid + ' (' + str(icount) + '/' + str(len(json_transactions)) + ')')
                        json_tran_violations = self.get_tqi_transactions_violations(domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows)                  
                        if json_tran_violations != None:
                            for tran_viol in json_tran_violations:
                                mytransac.get("componentsWithViolations").append(tran_viol['component']['href'])
                                #print(shortName + "=>" + tran_viol['component']['href'])
                    else:
                        if transaclist["60017"] != None:
                            for t in transaclist["60017"]:
                                if mytransac['href'] == t['href']:
                                    mytransac.update({"componentsWithViolations":t.get("componentsWithViolations")})
                                    break
                    transaclist[bcid].append(mytransac)
        return transaclist
    
    ########################################################################
    def initialize_components_pri (self, domain, applicationid, snapshotid,bcids,nbrows):
        comppridict = {}
        for bcid in bcids:
            comppridict.update({bcid:{}})
            json_snapshot_components_pri = self.get_components_pri(domain, applicationid, snapshotid, bcid,nbrows)
            if json_snapshot_components_pri != None:
                for val in json_snapshot_components_pri:
                    compid = None
                    try:
                        treenodehref = val['treeNodes']['href']
                    except KeyError:
                        self.restutils.logger.error('KeyError treenodehref ' + str(val))
                    if treenodehref != None:
                        rexuri = "/components/([0-9]+)/"
                        m0 = re.search(rexuri, treenodehref)
                        if m0: compid = m0.group(1)
                        pri = val['propagationRiskIndex']                                                 
                        if treenodehref != None and pri != None: 
                            comppridict.get(bcid).update({compid:pri})
                            #if (bcid == 60016 or bcid == "60016"):
                                #print(str(compid))
            json_snapshot_components_pri = None
        return comppridict
    
    
    ########################################################################
    def initialize_bc_tch_mapping(self, domain, applicationid, snapshotid, bcids):
        outputtcids = {}
        for bcid in bcids:
            outputtcid = []
            json = self.get_snapshot_bc_tc_mapping(domain, snapshotid, bcid)
            if json != None:
                if json != None:
                    for val in json['gradeContributors']:
                        outputtcid.append(val['key'])
            outputtcids.update({bcid:outputtcid}) 
            json = None
        return outputtcids    
    
########################################################################

class ObjectViolationMetric:
    componentType = '<Not extracted>'
    criticalViolations = '<Not extracted>'
    cyclomaticComplexity = '<Not extracted>'
    codeLines = '<Not extracted>'
    commentLines = '<Not extracted>'
    ratioCommentLinesCodeLines = '<Not extracted>'
    commentedCodeLines = None
    fanIn = None
    fanOut = None
    halsteadProgramLength = None
    halsteadProgramVocabulary = None
    halsteadVolume = None
    distinctOperators = None
    distinctOperands = None
    integrationComplexity = None
    criticalViolations = None

########################################################################

class RulePatternDetails:
    def __init__(self):
        self.associatedValueName = ''
        self.listQualityStandard = []

    def get_quality_standards(self):
        strqualitystandards = ''
        #print(len(self.listQualityStandard))
        for qs in self.listQualityStandard:
            strqualitystandards += qs + ","
        if strqualitystandards != '': strqualitystandards = strqualitystandards[:-1]
        return strqualitystandards

########################################################################

# metric class
class Metric:
    # Business criteria metrics
    BC_TQI = 60017
    
    # Technical sizes metrics
    TS_LOC="10151"
    
    # Functional size metrics
    FS_AFP="10202"
    
    # Distributions metrics
    DIST_CYCLOMATIC_COMPLEXITY = "65501"
    DIST_COST_COMPLEXITY = "67001"
    DIST_FAN_OUT = "66020"
    DIST_FAN_IN = "66021"
    DIST_COUPLING = "65350"
    DIST_SIZE = "65105"
    DIST_SQL_COMPLEXITY = "65801"
    DIST_METRICS = [DIST_CYCLOMATIC_COMPLEXITY,DIST_COST_COMPLEXITY,DIST_FAN_OUT,DIST_FAN_IN,DIST_COUPLING,DIST_SIZE,DIST_SIZE]
    
    id = None
    name = None
    type = None
    critical = None
    grade = None
    failedchecks = None
    successfulchecks = None
    totalchecks = None
    ratio = None
    threshold1 = None
    threshold2 = None
    threshold3 = None
    threshold4 = None
    addedviolations = None
    removedviolations = None
    applicationName = None

    @staticmethod
    def get_distributionsmetrics():
        metrics = ''
        for m in Metric.DIST_METRICS:
            metrics += m + ','
        return metrics[:-1]

########################################################################

    
# contribution class (technical criteria contributions to business criteria, or quality metrics to technical criteria) 
class Contribution:
    parentmetricid = None
    parentmetricname = None
    metricid = None
    metricname = None
    weight = None
    critical = None
  
########################################################################
    
# violation class
class Violation:
    id = None
    qrid = None
    qrname = None
    critical = None
    componentid = None
    componentShortName = None
    componentNameLocation = None
    hasActionPlan = False
    actionplanstatus = ''
    actionplantag = ''
    actionplancomment = ''
    hasExclusionRequest = False
    url = None
    violationstatus = None
    componentstatus = None


########################################################################
# snapshot class
class Snapshot:
    def __init__(self, href=None, domain=None, applicationid=None, applicationname=None, snapshotid=None, isodate=None, version=None):
        self.href = href
        self.domain = domain
        self.applicationid = applicationid
        self.snapshotid = snapshotid
        self.isodate = isodate
        self.version = version
        self.applicationname = applicationname

    def load(self, json):
        if json != None:
            self.version = json['version']
            self.href = json['applicationSnapshot']['href']
            self.applicationname = json['applicationSnapshot']['name']
            self.isodate = json['date']['isoDate']
            self.applicationid = -1
            self.snapshotid = -1
            rexappsnapid = "([A-Z0-9_]+)/applications/([0-9]+)/snapshots/([0-9]+)"
            m0 = re.search(rexappsnapid, self.href)
            if m0: 
                self.domain = m0.group(1)
                self.applicationid = m0.group(2)
                self.snapshotid = m0.group(3)

########################################################################
   
# Logging utils
class LogUtils:

    @staticmethod
    def loginfo(logger, msg, tosysout = False):
        logger.info(msg)
        if tosysout:
            print(msg)

    @staticmethod
    def logwarning(logger, msg, tosysout = False):
        logger.warning(msg)
        if tosysout:
            print("#### " + msg)

    @staticmethod
    def logerror(logger, msg, tosysout = False):
        logger.error(msg)
        if tosysout:
            print("#### " + msg)

#########################################################################