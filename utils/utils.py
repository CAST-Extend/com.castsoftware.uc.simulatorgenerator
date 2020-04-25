import requests
from base64 import b64encode
import re
import os
import sys

'''
Created on 13 avr. 2020

@author: MMR
'''

# Cookie used for the jsessionid cookie to make sure the session persists
setcookie = None


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



class FileUtils:
    ########################################################################
    
    """Checks if a file is locked by opening it in append mode.
    If no exception thrown, then the file is not locked.
    """
    @staticmethod
    def is_file_locked(filepath):
    
        locked = None
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


####################################################################################################

# retrieve the connection depending on 
def open_connection(logger, url, user, pwd):
    logger.info('Opening connection to ' + url)
    try:
        resp = requests.get(url, headers={"User-Agent": "XY"}, auth=(user, pwd))
    except:
        logger.error ('Error connecting to ' + url)
        logger.error ('URL is not reachable. Please check your connection (web application down, VPN not active, ...)')
        raise SystemExit
    
    if resp.status_code != 200:
        # This means something went wrong.
        logger.error ('Error connecting to ' + url)
        logger.error ('Status code = ' + str(resp.status_code))
        logger.error ('Headers = ' + str(resp.headers))
        logger.error ('Please check the URL, user and password ')
        raise SystemExit
    else: 
        logger.error ('Successfully connected to  : ' + url)    
    
    global setcookie
    setcookie = None

####################################################################################################
# retrieve the connection depending on 
def close_connection(logger):
    global setcookie
    setcookie = None

####################################################################################################

def execute_request(logger, requesttype, url, request, user, password, apikey, contenttype='application/json', inputjson=None):
    global setcookie
    
    request_headers = {}
    request_text = url + "/rest/" + request
    logger.debug('Sending ' + requesttype + ' ' + request_text + ' with contenttype=' + contenttype)   

    # if the user and password are provided, we take them first
    if user != None and password != None and user != 'N/A' and user != 'N/A':
        #we need to base 64 encode it 
        #and then decode it to acsii as python 3 stores it as a byte string
        #userAndPass = b64encode(user_password).decode("ascii")
        auth = str.encode("%s:%s" % (user, password))
        #user_and_pass = b64encode(auth).decode("ascii")
        user_and_pass = b64encode(auth).decode("iso-8859-1")
        request_headers.update({'Authorization':'Basic %s' %  user_and_pass})
    # else if the api key is provided
    elif apikey != None and apikey != 'N/A':
        print (apikey)
        # API key configured in the WAR
        request_headers.update({'X-API-KEY:':apikey})
        # we are provide a user name hardcoded' 
        request_headers.update({'X-API-USER:':'admin_apikey'})
        
    # Name of the client added in the header (for the audit trail)
    request_headers.update({'X-Client':'com.castsoftware.uc.violationextraction'})
    request_headers.update({'accept' : contenttype})
    
    # if the session JSESSIONID is already defined we inject the cookie to reuse previous session
    if setcookie != None:
        request_headers.update({'Set-Cookie':setcookie})

    # send the request
    if 'GET' == requesttype:
        response = requests.get(request_text,headers=request_headers,auth=(user, password))
    elif 'POST' == requesttype:
        response = requests.post(request_text,inputjson,headers=request_headers,auth=(user, password))
    elif 'PUT' == requesttype:
        response = requests.post(request_text,inputjson,headers=request_headers,auth=(user, password))        
    elif 'DELETE' == requesttype:
        response = requests.post(request_text,inputjson,headers=request_headers,auth=(user, password))    
    else:
        LogUtils.logerror(logger,'Invalid HTTP request type' + requesttype)
    
    output = None
    if response != None: 
        # Error
        if response.status_code != 200:
            LogUtils.logerror(logger,'HTTPS request failed ' + str(response.status_code) + ' :' + request_text,True)
            return None
        else:
            # look for the Set-Cookie in response headers, to inject it for future requests
            if setcookie == None: 
                sc = response.headers._store.get('set-cookie')
                if sc != None and sc[1]  != None:
                    setcookie = sc[1]
            if contenttype == 'application/json':
                output = response.json()
            else:
                output = response.text

    return output 

####################################################################################################

def execute_request_get(logger, url, request, user, password, apikey, contenttype='application/json'):
    return execute_request(logger, 'GET', url, request, user, password, apikey, contenttype)

####################################################################################################

def execute_request_post(logger, url, request, user, password, apikey, contenttype='application/json', inputjson=None):
    return execute_request(logger, 'POST', url, request, user, password, apikey, contenttype)

####################################################################################################

def execute_request_put(logger, url, request, user, password, apikey, contenttype='application/json', inputjson=None):
    return execute_request(logger, 'PUT', url, request, user, password, apikey, contenttype)

####################################################################################################

def execute_request_delete(logger, url, request, user, password, apikey, contenttype='application/json', inputjson=None):
    return execute_request(logger, 'DELETE', url, request, user, password, apikey, contenttype)

####################################################################################################

'''
# Using http.client, not working with Python 3.6+ 
def execute_request(logger, connection, requesttype, request, warname, user, password, apikey, inputjson, contenttype='application/json'):
    global setcookie
    
    request_headers = {}
    json_data = None

    request_text = "/"
    if warname != None:
        request_text +=  warname +"/"
    request_text += "rest/" + request
    logger.debug('Sending ' + requesttype + ' ' + request_text + ' with contenttype=' + contenttype)   

    # if the user and password are provided, we take them first
    if user != None and password != None and user != 'N/A' and user != 'N/A':
        #we need to base 64 encode it 
        #and then decode it to acsii as python 3 stores it as a byte string
        #userAndPass = b64encode(user_password).decode("ascii")
        auth = str.encode("%s:%s" % (user, password))
        #user_and_pass = b64encode(auth).decode("ascii")
        user_and_pass = b64encode(auth).decode("iso-8859-1")
        request_headers.update({'Authorization':'Basic %s' %  user_and_pass})
    # else if the api key is provided
    elif apikey != None and apikey != 'N/A':
        #print (apikey)
        # API key configured in the WAR
        request_headers.update({'X-API-KEY:':apikey})
        # we are provide a user name hardcoded' 
        request_headers.update({'X-API-USER:':'admin_apikey'})
        
    # Name of the client added in the header (for the audit trail)
    request_headers.update({'X-Client':'com.castsoftware.uc.violationextraction'})
    
    # GET request
    if requesttype == 'GET':
        request_headers.update({'accept' : contenttype})
    # POST/PUT/DELETE request, we are provining a json file
    else:
        request_headers.update({'Content-type' : contenttype})
        json_data = json.dumps(inputjson)
    
    # if the session JSESSIONID is already defined we inject the cookie to reuse previous session
    if setcookie != None:
        request_headers.update({'Set-Cookie':setcookie})

    # sent the request
    connection.request(requesttype, request_text, json_data, headers=request_headers)        
     
    #get the response back
    response = connection.getresponse()
    #logger.debug('     response status ' + str(response.status) + ' ' + str(response.reason))    
    
    # Error 
    if  response.status != 200:
        logerror(logger,'HTTPS request failed ' + str(response.status) + ' ' + str(response.reason) + ':' + request_text,True)
        return None
    
    # look for the Set-Cookie in response headers, to inject it for future requests
    if setcookie == None: 
        for h1 in response.headers._headers:
            if h1 != None and h1[0] == 'Set-Cookie':
                setcookie = h1[1]
                break
    
    #send back the date
    encoding = response.info().get_content_charset('iso-8859-1')
    responseread_decoded = response.read().decode(encoding)
    
    if contenttype=='application/json':
        output = json.loads(responseread_decoded)
    else:
        output = responseread_decoded
    
    return output
'''

########################################################################

# CAST AIP Dashboard 
class AIPRestAPI:

    # Returns a json
    @staticmethod
    def get_server(logger, url, user, password, apikey):
        request = "server"
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def get_domains(logger, url, user, password, apikey):
        request = ""
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def get_applications(logger, url, user, password, apikey, domain):
        request = domain + "/applications"
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def get_transactions_per_business_criterion(logger, url, user, password, apikey, domain, applicationid, snapshotid, bcid, nbrows):
        logger.info("Extracting the transactions for business criterion " + bcid)
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/transactions/" + bcid
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)    
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod    
    def get_application_snapshots(logger, url, user, password, apikey, domain, applicationid):
        request = domain + "/applications/" + applicationid + "/snapshots" 
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def get_total_number_violations(logger, url, user, password, apikey, domain, applicationid,snapshotid):
        logger.info("Extracting the number of violations")
        request = domain + "/results?sizing-measures=67011,67211&application=" + applicationid + "&snapshot=" + snapshotid
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    #get_snapshot_violations(logger, url, user, password, apikey, domain, applicationid, snapshotid, criticalrulesonlyfilter, None, businesscriterionfilter, technofilter, nbrows)                                            
    @staticmethod
    def get_snapshot_violations(logger, url, user, password, apikey, domain, applicationid, snapshotid, criticalonly, violationStatus, businesscriterionfilter, technoFilter, nbrows):
        logger.info("Extracting the snapshot violations")
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
            
        return execute_request_get(logger, url, request, user, password, apikey)
        
    ########################################################################
    @staticmethod
    def get_tqi_transactions_violations(logger, url, user, password, apikey, domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows):    
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
            
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def create_scheduledexclusions(logger, url, domain, user, password, apikey, applicationid, snapshotid, json_exclusions_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/exclusions/requests'
        return execute_request_post(logger, url, request, user, password, apikey,'application/json',json_exclusions_to_create)
    
    ########################################################################
    @staticmethod
    def create_actionplans(logger, url, domain, user, password, apikey, applicationid, snapshotid, json_actionplans_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/action-plan/issues'
        return execute_request_post(logger, url, request, user, password, apikey,'application/json',json_actionplans_to_create)

    ########################################################################
    
    @staticmethod
    def get_rule_pattern(logger, url, user, password, apikey, rulepatternHref):
        logger.debug("Extracting the rule pattern details")   
        request = rulepatternHref
        json_rulepattern = execute_request_get(logger, url, request, user, password, apikey)
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
    @staticmethod
    def get_actionplan_summary(logger, url, user, password, apikey, domain, applicationid, snapshotid):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/action-plan/summary"
        return execute_request_get(logger, url, request, user, password, apikey)

    ########################################################################
    @staticmethod
    def get_qualitymetrics_results(logger, url, user, password, apikey, domain, applicationid, criticalonly, nbrows):
        request = domain + "/applications/" + applicationid + "/results?quality-indicators"
        request += '=(cc:60017'
        if criticalonly == None or not criticalonly:   
            request += ',nc:60017'
        request += ')&select=(evolutionSummary,violationRatio)'
        # last snapshot only
        request += '&snapshots=-1'
        request += '&startRow=1'
        request += '&nbRows=' + str(nbrows)
        return execute_request_get(logger, url, request, user, password, apikey)

    ########################################################################
    @staticmethod
    def get_metric_contributions(logger, url, user, password, apikey, domain, metricid, snapshotid):
        request = domain + "/quality-indicators/" + metricid + "/snapshots/" + snapshotid 
        return execute_request_get(logger, url, request, user, password, apikey)

    ########################################################################
    @staticmethod
    def get_qualityrules_thresholds(logger, url, user, password, apikey, domain, snapshotid, qrid):
        request = domain + "/quality-indicators/" + str(qrid) + "/snapshots/"+ snapshotid
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod
    def get_businesscriteria_grades(logger, url, user, password, apikey, domain):
        request = domain + "/results?quality-indicators=(60017,60016,60014,60013,60011,60012,66031,66032,66033)&applications=($all)&snapshots=(-1)" 
        return execute_request_get(logger, url, request, user, password, apikey)

    ########################################################################
    @staticmethod
    def get_snapshot_tqi_quality_model (logger, url, user, password, apikey, domain, snapshotid):
        logger.info("Extracting the snapshot quality model")   
        request = domain + "/quality-indicators/60017/snapshots/" + snapshotid + "/base-quality-indicators" 
        return execute_request_get(logger, url, request, user, password, apikey)
    ########################################################################
    @staticmethod
    def get_snapshot_bc_tc_mapping(logger, url, user, password, apikey, domain, snapshotid, bcid):
        logger.info("Extracting the snapshot business criterion " + bcid +  " => technical criteria mapping")  
        request = domain + "/quality-indicators/" + str(bcid) + "/snapshots/" + snapshotid
        return execute_request_get(logger, url, request, user, password, apikey)
    
    ########################################################################
    @staticmethod 
    def get_components_pri (logger, url, user, password, apikey, domain, applicationid, snapshotid, bcid,nbrows):
        logger.info("Extracting the components PRI for business criterion " + bcid)  
        request = domain + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + '/components/' + str(bcid)
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)
        return execute_request_get(logger, url, request, user, password, apikey)
         
    ########################################################################
    @staticmethod     
    def get_sourcecode(logger, url, sourceCodesHref, warname, user, password, apikey):
        return execute_request_get(logger, url, sourceCodesHref, user, password, apikey)     
    
    ########################################################################
    @staticmethod
    def get_sourcecode_file(logger, url, filehref, srcstartline, srcendline, user, password, apikey):
        strstartendlineparams = ''
        if srcstartline != None and srcstartline >= 0 and srcendline != None and srcendline >= 0:
            strstartendlineparams = '?start-line='+str(srcstartline)+'&end-line='+str(srcendline)
        return execute_request_get(logger, url, filehref+strstartendlineparams, user, password, apikey,'text/plain')    
    
    ########################################################################
    @staticmethod
    def get_objectviolation_metrics(logger, url, user, password, apikey, objHref):
        logger.debug("Extracting the component metrics")    
        request = objHref
        json_component = execute_request_get(logger, url, request, user, password, apikey)
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
    @staticmethod
    def get_objectviolation_findings(logger, url, user, password, apikey, objHref, qrid):
        logger.debug("Extracting the component findings")    
        request = objHref + '/findings/' + qrid 
        return execute_request_get(logger, url, request, user, password, apikey)
         
    ########################################################################
    # extract the transactions TRI & violations component list per business criteria
    @staticmethod
    def init_transactions (logger, url, usr, pwd, apikey,domain, applicationid, snapshotid, criticalonly, violationStatus, technoFilter,nbrows):
        #Security,Efficiency,Robustness,TQI
        bcids = ["60017","60016","60014","60013"]
        transaclist = {}
        for bcid in bcids: 
            json_transactions = AIPRestAPI.get_transactions_per_business_criterion(logger, url, usr, pwd, apikey,domain, applicationid, snapshotid, bcid, nbrows)
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
                        logger.info("Extracting the violations for transaction " + transactionid + ' (' + str(icount) + '/' + str(len(json_transactions)) + ')')
                        json_tran_violations = AIPRestAPI.get_tqi_transactions_violations(logger, url, usr, pwd, apikey,domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows)                  
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
    @staticmethod
    def initialize_components_pri (logger, url, user, password, apikey, domain, applicationid, snapshotid,bcids,nbrows):
        comppridict = {}
        for bcid in bcids:
            comppridict.update({bcid:{}})
            json_snapshot_components_pri = AIPRestAPI.get_components_pri(logger, url, user, password, apikey, domain, applicationid, snapshotid, bcid,nbrows)
            if json_snapshot_components_pri != None:
                for val in json_snapshot_components_pri:
                    compid = None
                    try:
                        treenodehref = val['treeNodes']['href']
                    except KeyError:
                        logger.error('KeyError treenodehref ' + str(val))
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
    @staticmethod 
    def initialize_bc_tch_mapping(logger, url, user, password, apikey, domain, applicationid, snapshotid, bcids):
        outputtcids = {}
        for bcid in bcids:
            outputtcid = []
            json = AIPRestAPI.get_snapshot_bc_tc_mapping(logger, url, user, password, apikey, domain, snapshotid, bcid)
            if json != None:
                if json != None:
                    for val in json['gradeContributors']:
                        outputtcid.append(val['key'])
            outputtcids.update({bcid:outputtcid}) 
            json = None
        return outputtcids    
    
########################################################################

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
    
# contribution class (technical criteria contributions to business criteria, or quality metrics to technical criteria) 
class Contribution:
    parentmetricid = None
    parentmetricname = None
    metricid = None
    metricname = None
    weight = None
    critical = None
    
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

####################################################################################################