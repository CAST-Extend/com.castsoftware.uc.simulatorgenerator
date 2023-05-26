import requests
from base64 import b64encode
import re
import os
import sys
import time
import subprocess
import json
#import psycopg2
import traceback

'''
Created on 13 avr. 2020

@author: MMR
'''

####################################################################################

class Filter:
    def __init__(self):
        None
        
class ViolationFilter(Filter):
    def __init__(self, criticalrulesonlyfilter, businesscriterionfilter, technofilter, violationstatusfilter, qridfilter, qrnamefilter, nbrowsfilter):
        self.criticalrulesonly = criticalrulesonlyfilter
        self.businesscriterion = businesscriterionfilter
        self.techno = technofilter
        self.violationstatus = violationstatusfilter
        self.nbrows = nbrowsfilter
        self.qrid = qridfilter
        self.qrname = qrnamefilter

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

    ########################################################################
    @staticmethod    
    def remove_semicolumn(astr):
        return astr.replace(';', '')

    ########################################################################
    @staticmethod
    def remove_trailing_suffix (mystr, suffix='rest'):
        # remove trailing /
        while mystr.endswith('/'):
            mystr = mystr[:-1]
        if mystr.endswith(suffix):
            return (mystr[:len(mystr)-len(suffix)-1])        
        else:
            return mystr

######################################################################################################################

class DateUtils:
    @staticmethod 
    # Format a timestamp date into a string
    def get_formatted_dateandtime(mydate):
        formatteddate = str(mydate.year) + "_"
        if mydate.month < 10:
            formatteddate += "0"
        formatteddate += str(mydate.month) + "_"
        if mydate.day < 10:
            formatteddate += "0"
        formatteddate += str(mydate.day)
        
        formatteddate += "_" 
        if mydate.hour < 10:
            formatteddate += "0"    
        formatteddate += str(mydate.hour)
        if mydate.minute < 10:
            formatteddate += "0"    
        formatteddate += str(mydate.minute)    
        if mydate.second < 10:
            formatteddate += "0"    
        formatteddate += str(mydate.second)    
        
        return formatteddate       

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
    CLIENT_CURL = 'curl'
    CLIENT_REQUESTS = 'requests'
    
    USERAGENT = 'XY'
   
    def __init__(self, logger, url, restclient, user=None, password = None, apikey = None, uselocalcache=False, cachefolder=None, extensionid='Community Extension'): 
        self.session_cookie = None
        self.session = None
        self.restclient = restclient
        self.logger = logger
        self.url = url
        self.extensionid = extensionid
        self.user = user
        self.password = password
        self.apikey = apikey
        self.uselocalcache = uselocalcache
        self.cachefolder = cachefolder
        self.cachesubfolder = None 
    
    ####################################################################################################
    
    def get_json(self, request, apikey=None, cachefilename=None):
        json_filepath = self.get_cachefilepath(cachefilename)
        # create parent folder if required 
        if json_filepath != None and not os.path.exists(os.path.dirname(json_filepath)):
            os.path.dirname(json_filepath)
        # run rest command only if file do not exist or we don't use the local cache and force the data to be loaded again
        if self.uselocalcache and json_filepath != None and os.path.isfile(json_filepath):
            try:
                with open(json_filepath, 'r', encoding='utf-8') as json_file:
                #with open(json_filepath, 'r') as json_file:
                    return json.load(json_file)
            except UnicodeDecodeError:
                LogUtils.logwarning(self.logger, 'Unicode decode error in json file %s: Skipping' % json_filepath, True)                
            except json.decoder.JSONDecodeError:
                LogUtils.logwarning(self.logger, 'Invalid json file %s: Skipping' % json_filepath, True)  
        else:
            if self.restclient == 'curl':
                return self.execute_curl(request, apikey, json_filepath)
            elif self.restclient == 'requests':
                return self.execute_requests(request)    
    
    ####################################################################################################
    
    def modify_with_json(self, requesttype, request, apikey, inputjson, cachefilename=None, contenttype='text/plain'):
        if type(inputjson) == 'str':
            inputjsonstr = inputjson
        else:
            inputjsonstr = json.dumps(inputjson)  
        json_filepath = self.get_cachefilepath(cachefilename)
        # create parent folder if required 
        if not os.path.exists(os.path.dirname(json_filepath)):
            os.path.dirname(json_filepath)
        if self.restclient == 'curl':
            return self.execute_curl(request, apikey, json_filepath, requesttype, 'application/json', inputjsonstr)
        elif self.restclient == 'requests':
            return self.execute_requests(request, requesttype, 'application/json', inputjsonstr, contenttype)      
    
    ####################################################################################################
    """
    def put_json(self, request, apikey, inputjson, cachefilename=None):
        if type(inputjson) == 'str':
            inputjsonstr = inputjson
        else:
            inputjsonstr = json.dumps(inputjson)  
        json_filepath = self.get_cachefilepath(cachefilename)
        # create parent folder if required 
        if not os.path.exists(os.path.dirname(json_filepath)):
            os.path.dirname(json_filepath)
        if self.restclient == 'curl':
            return self.execute_curl(request, apikey, json_filepath, 'PUT', 'application/json', inputjsonstr)
        elif self.restclient == 'requests':
            return self.execute_requests(request, 'PUT', 'application/json', inputjsonstr, 'text/plain')        
    """
    ####################################################################################################
    def get_response_cookies(self, response):
        return response.headers._store 

    ####################################################################################################
    def get_cachefolderpath(self):
        folder = None
        if self.cachefolder != None:
            folder = self.cachefolder
            if self.cachesubfolder != None:
                folder += "\\" + self.cachesubfolder 
        return folder
    ####################################################################################################
    def get_cachefilepath(self, cachefilename):
        if cachefilename == None:
            return None
        folder = os.path.dirname(self.get_cachefolderpath() + '\\' + cachefilename)
        if folder != None and cachefilename != None:
            return "%s\%s" % (folder, cachefilename)
        else:
            return None

    ####################################################################################################
    def execute_curl(self, request, apikey, cachefilepath, requesttype='GET', accept='application/json', inputjsonstr=None):
        json_output = None
        request_text = self.url + request
        
        strcmd = 'curl %s' % request_text
        strcmd += ' -X %s' % requesttype
        strcmd += ' -H "Accept: ' + accept + '"'
        strcmd += ' -H "User-Agent: '+ RestUtils.USERAGENT + '"'
        strcmd += ' -H "X-Client: ' + self.extensionid + '"'
        if self.apikey != None and self.apikey != 'N/A':
            strcmd += ' -H "X-API-KEY: ' + self.apikey+ '"'
        if self.user != None and self.password != None and self.user != 'N/A' and self.password != 'N/A':
            strcmd += ' -u ' + self.user + ':' + self.password            
        strcmd += ' -H "Connection: keep-alive"'
        if requesttype != 'GET':
            strcmd += ' -H  "Content-Type: application/json"' 
            strcmd += ' -d "' + inputjsonstr.replace('"','\\"') + '"'
        strcmd += ' -o "%s"' % cachefilepath

        if not os.path.exists(os.path.dirname(cachefilepath)):
            os.makedirs(os.path.dirname(cachefilepath))

        LogUtils.logdebug(self.logger,"curl running: " + strcmd, True)
        status, curl_output = subprocess.getstatusoutput(strcmd)
        if status != 0:
            # error
            LogUtils.logerror(self.logger,"Error running %s - curl status %s" % (request_text, str(status)), True)
            LogUtils.logerror(self.logger,"curl output %s" % curl_output)
            raise SystemError                

        # if no error send back a json string containing data from the cache file just loaded
        try:
            with open(cachefilepath, 'r') as json_file:
                json_output = json.load(json_file)
        except UnicodeDecodeError:
            LogUtils.logwarning(self.logger, 'Unicode decode error in json file %s: Skipping' % cachefilepath, True)
        except json.decoder.JSONDecodeError:
            LogUtils.logwarning(self.logger, 'Invalid json file %s: Skipping' % cachefilepath, True)   

        return json_output
    
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
    def open_session(self, resturi=''):
        
        if self.restclient == 'curl':
            # Nothing to do for curl
            None
        elif self.restclient == 'requests':
            uri = self.url + '/' +  resturi
            self.logger.info('Opening session to ' + uri)
            response = None
            request_headers = {}
            #request_headers.update(self.get_default_http_headers())        
            request_headers.update({'accept':'application/json'})        
            try:
                self.session = requests.session()
                if self.user != None and self.password != None and self.user != 'N/A' and self.password != 'N/A':
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
                    # API key configured in the Health / Engineering / REST-API WAR
                    request_headers.update({'X-API-KEY':self.apikey})
                    # we are provide a user name hardcoded' 
                    #request_headers.update({'X-API-USER':'admin_apikey'})            
                    # API key configured in ExtenNG
                    #request_headers.update({'x-nuget-apikey':self.apikey})
                
                self.logger.info ('request headers = ' + str(request_headers))
                
                response = self.session.get(uri, headers=request_headers)
                
            except:
                self.logger.error ('Error connecting to ' + uri)
                self.logger.error ('URL is not reachable. Please check your connection (web application down, VPN not active, ...)')
                raise SystemExit
            #finally:
            #    self.logger.info ('Headers = ' + str(response.headers))
                
            if response.status_code != 200:
                # This means something went wrong.
                self.logger.error ('Error connecting to ' + uri)
                self.logger.info ('Status code = ' + str(response.status_code))
                self.logger.info ('response headers = ' + str(response.headers))
                self.logger.info ('Please check the URL, user and password or api key')
                raise SystemExit
            else: 
                self.logger.info ('Successfully connected to  : ' + self.url)    
    
    ####################################################################################################

    def get_default_http_headers(self):
        # User agent & Name of the client added in the header (for the audit trail)
        default_headers = {"User-Agent": "XY", "X-Client": self.extensionid} 
        return default_headers

    ####################################################################################################
    def execute_requests(self, request, requesttype='GET', accept='application/json', inputjsonstr=None, contenttype='application/json'):
        if self.session == None:
            self.session = self.open_session()
            
        if request == None:
            request_text = self.url
        else:
            request_text = self.url
            if not request_text.endswith('/'): request_text += '/'
            request_text += request
        
        request_headers = {}
        request_headers.update(self.get_default_http_headers())
        request_headers.update({'accept' : accept})
        try:
            request_headers.update({'X-XSRF-TOKEN': self.session.cookies['XSRF-TOKEN']})
        except KeyError:
            None
        request_headers.update({'Content-Type': 'application/json'})
    
        LogUtils.logdebug(self.logger,'Sending ' + requesttype + ' ' + request_text + ' with contenttype=' + contenttype + ' json=' + str(inputjsonstr), False)
        #LogUtils.logdebug(self.logger,'  Request headers=' + json.dumps(request_headers) , False)
    
        # send the request
        if 'GET' == requesttype:
            response = self.session.get(request_text,headers=request_headers, verify=False)
        elif 'POST' == requesttype:
            response = self.session.post(request_text,inputjsonstr,headers=request_headers)
        elif 'PUT' == requesttype:
            response = self.session.put(request_text,inputjsonstr,headers=request_headers)        
        elif 'DELETE' == requesttype:
            response = self.session.delete(request_text,inputjsonstr,headers=request_headers)    
        else:
            LogUtils.logerror(self.logger,'Invalid HTTP request type' + requesttype)
        
        output = None
        if response != None:
            #LogUtils.logdebug(self.logger,'  HTTP code=%s headers=%s'% (str(response.status_code), json.dumps(response.headers._store)), False)
            
            # Error
            if response.status_code not in (200, 201):
                LogUtils.logerror(self.logger,'HTTP(S) request failed ' + str(response.status_code) + ' :' + request_text,True)
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
    
    def execute_requests_get(self, request, accept='application/json', content_type='application/json'):
        return self.execute_requests(request, 'GET', accept, None, content_type)
    
    ####################################################################################################
    
    def execute_requests_post(self, request, accept='application/json', inputjson=None, contenttype='application/json'):
        return self.execute_requests(request, 'POST', inputjson, accept, contenttype)
    
    ####################################################################################################
    
    def execute_requests_put(self, request, accept='application/json', inputjson=None, contenttype='application/json'):
        return self.execute_requests(request, 'PUT', accept, inputjson, contenttype)
    
    ####################################################################################################
    
    def execute_requests_delete(self, request, accept='application/json', inputjson=None, contenttype='application/json'):
        return self.execute_requests(request, 'DELETE', accept, inputjson, contenttype)


########################################################################


########################################################################
# snapshot class
class Domain:
    def __init__(self, href=None, name=None, version=None, schema=None):
        self.href = href
        self.name = name
        self.version = version
        self.schema = schema
    
    @staticmethod
    def load(json):
        if json != None:
            d = Domain()
            d.href = json['href']
            d.name = json['name']
            d.version = json['version']
            d.schema = json['schema']
            return d
        else: return None
    
    def isAAD(self):
        return self.name != None and self.name == 'AAD'
    
    @staticmethod
    def loadlist(json):
        domainlist = []
        if json != None:
            for item in json:
                domainlist.append(Domain.load(item)) 
        return domainlist

########################################################################

# Application
class Application:
    def __init__(self):
        self.href = None
        self.name = None
        self.id = None
        self.schema_central = None
        self.schema_local = None
        self.schema_mgnt = None
    
    @staticmethod   
    def load(json):
        if json != None:   
            x = Application()
            x.href = json['href']
            x.name = json['name']
            x.id = AIPRestAPI.get_href_id(x.href)
            return x  
        else: return None
        
    @staticmethod
    def loadlist(json):
        applicationlist = []
        if json != None:
            for item in json:
                applicationlist.append(Application.load(item)) 
        return applicationlist        
        
########################################################################

# server status
class Server:
    def __init__(self):
        self.version = None
        self.status = None
        self.freememory = None
    
    @staticmethod   
    def load(json):
        if json != None:   
            x = Server()
            #servversion2digits = servversion[-4:] 
            #if float(servversion2digits) <= 1.13 : 
            #    None
            x.version = json['version']
            x.status = json['status']
            x.freememory = json['memory']['freeMemory']
            return x                
        else: return None
        
########################################################################

# snapshot filter class
class SnapshotFilter:
    def __init__(self, snapshot_index, snapshot_ids):
        self.snapshot_index = snapshot_index
        self.snapshot_ids = snapshot_ids

########################################################################

# snapshot class
class Snapshot:
    def __init__(self, href=None, domainname=None, applicationid=None, applicationname=None, snapshotid=None, isodate=None, version=None):
        self.href = href
        self.domainname = domainname
        self.applicationid = applicationid
        self.snapshotid = snapshotid
        self.isodate = isodate
        self.version = version
        self.versionname = None
        self.applicationname = applicationname
        self.time = None
        self.number = None
        self.technologies = None
        self.modules = None
        self.last = None
        self.beforelast = None
        self.first = None

    def get_technologies_as_string(self):
        strtechnologies = ''
        if self.technologies != None:
            for t in self.technologies:
                strtechnologies += t + ','
            if ',' in strtechnologies:
                strtechnologies = strtechnologies[:-1]
        return strtechnologies

    @staticmethod
    def load(json_snapshot, last, beforelast, first):
        x = Snapshot()
        if json_snapshot != None:
            x.version = json_snapshot['annotation']['version']
            x.versionname = json_snapshot['annotation']['name']
            x.href = json_snapshot['href']
            x.applicationname = json_snapshot['name']
            x.isodate = json_snapshot['annotation']['date']['isoDate']
            x.time = json_snapshot['annotation']['date']['time']
            x.number = json_snapshot['number']
            x.last = last
            x.beforelast = beforelast
            x.first = first
            try:
                x.technologies = json_snapshot['technologies']
            except KeyError:
                None

            x.applicationid = -1
            x.snapshotid = -1
            """rexappsnapid = "([-A-Z0-9_]+)/applications/([0-9]+)/snapshots/([0-9]+)"
            m0 = re.search(rexappsnapid, x.href)
            if m0: 
                x.domainname = m0.group(1)
                x.applicationid = m0.group(2)
                x.snapshotid = m0.group(3)
            """
            rex = "/snapshots/([0-9]+)"
            m0 = re.search(rex, x.href)
            if m0: 
                x.snapshotid = m0.group(1)
            rex = "/applications/([0-9]+)/"
            m0 = re.search(rex, x.href)
            if m0: 
                x.applicationid = m0.group(1)      
            rex = "(.*)/applications"
            m0 = re.search(rex, x.href)
            if m0: 
                x.domainname = m0.group(1)      
        return x
    @staticmethod
    def loadlist(json_snapshots):
        snapshotlist = []
        if json_snapshots != None:
            icount = 0
            for json_snapshot in json_snapshots:
                icount += 1
                snapshotlist.append(Snapshot.load(json_snapshot, icount==1, icount==2, icount==len(json_snapshots))) 
        return snapshotlist  

########################################################################
   
# Module class
class Module:
    def __init__(self):
        self.href = None
        self.domainname = None
        self.snapshotid = None
        self.moduleid = None
        self.modulename = None
        self.technologies = None

    def get_technologies_as_string(self):
        strtechnologies = ''
        if self.technologies != None:
            for t in self.technologies:
                strtechnologies += t + ','
            if ',' in strtechnologies:
                strtechnologies = strtechnologies[:-1]
        return strtechnologies

    @staticmethod
    def load(json_module):
        x = Module()
        if json_module != None:
            x.href = json_module['href']
            x.modulename = json_module['name']
            try:
                x.technologies = json_module['technologies']
            except KeyError:
                None
            x.moduleid = -1
            x.snapshotid = -1
            rexappsnapid = "([A-Z0-9_]+)/modules/([0-9]+)/snapshots/([0-9]+)"
            m0 = re.search(rexappsnapid, x.href)
            if m0: 
                x.domainname = m0.group(1)
                x.moduleid = m0.group(2)
                x.snapshotid = m0.group(3)   
        return x
    @staticmethod
    def loadlist(json_modules):
        listmodules = []
        if json_modules != None:
            for json in json_modules:
                listmodules.append(Module.load(json)) 
        return listmodules 
  
########################################################################
  

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
    # extract the last element that is the id
    @staticmethod
    def get_href_id(href, separator='/'):
        if href == None or separator == None:
            return None
        href_id = ""
        hrefsplit = href.split('/')
        for elem in hrefsplit:
            # the last element is the id
            href_id = elem    
        return href_id      
        
    ########################################################################
    # get mngt or schema name from the central schema name
    # assumption for the naming convention all the triplet schemas have same prefix + have default suffixes (_mngt, _central, _local)  
    @staticmethod
    def get_schema_name(centralschemaname, suffix='mngt'):
        if centralschemaname == None or suffix == None:
            return None
        schema = ''
        schema_split = centralschemaname.split("_")
        icount=0
        for sc in schema_split:
            icount+=1
            if (icount < len(schema_split)):
                schema += sc
            if (icount < len(schema_split) - 1):
                schema += "_"
        schema += "_" + suffix
        return schema        

    ########################################################################
    # Extract the packages list (platform and extension) installed & referenced in the mngt schema
    '''def get_mngt_schema_packages(self, domainname, appname, mngt_schema, host="localhost", port="2282", database="postgres", user="operator", password="CastAIP"):
        listpackages = []
        json_packages = []
        conn = None
        cur = None
        try:
            if not os.path.exists(self.restutils.get_cachefolderpath()): 
                os.makedirs(self.restutils.get_cachefolderpath())
            cachefilepath = self.restutils.get_cachefolderpath() + '\\packages_' + domainname + "_" + appname + ".json"
            if self.restutils.uselocalcache and cachefilepath != None and os.path.isfile(cachefilepath):
                with open(cachefilepath, 'r') as json_file:
                    json_packages = json.load(json_file)
            else:
                # load data from DB and save json data to disk
                conn = psycopg2.connect(host=host, port = port, database=database, user=user, password=password)
                cur = conn.cursor()
                sql = "SELECT package_name,version FROM " + mngt_schema + ".sys_package_version where package_name like '%CORE_PMC%' or package_name like '%com%' order by 1 desc"
                self.restutils.logger.debug("sql="+sql)
                #minjus_genesis_82_mngt.sys_package_version where package_name like '%CORE_PMC%' or package_name like '%com%' order by 1 desc"""
                cur.execute(sql) 
                for package_name, version in cur.fetchall():
                    if package_name == 'CORE_PMC':
                        package_type = "platform"
                    else:
                        package_type = "extension"
                        if package_name[0:1] == '/':
                            package_name = package_name[1:]
                    json_packages.append(
                            {
                                "package_type": package_type, 
                                "package_mngt_id": package_name, 
                                "package_mngt_version": version}
                            )
                # create cache file
                with open(cachefilepath, 'w') as json_file:
                    json.dump(json_packages, json_file)

            if json_packages != None:
                listpackages = PackageMngt.loadlist_from_mngt(json_packages)

        except:
            tb = traceback.format_exc()
            LogUtils.logerror(self.restutils.logger, "Error extracting the versions from postgresql %s" % tb, True)
        finally:
            if cur is not None:
                cur.close()
            if conn is not None:
                conn.close()
    '''   

        
    ########################################################################

    def get_server_json(self):
        request = "server"
        return self.restutils.get_json(request)
    
    def get_server(self):
        return Server.load(self.get_server_json())

    ########################################################################
    def get_domains_json(self):
        request = ""
        return self.restutils.execute_requests_get(request)

    def get_domains(self):
        return Domain.loadlist(self.get_domains_json())
        
    ########################################################################
    def get_applications_json(self, domain):
        request = domain + "/applications"
        return self.restutils.execute_requests_get(request)
    
    def get_applications(self, domain):
        applicationlist = Application.loadlist(self.get_applications_json(domain.name))
        for app in applicationlist:
            if domain.schema != None and "_central" in domain.schema:
                app.schema_central = domain
                app.schema_mngt = AIPRestAPI.get_schema_name(domain.schema, "mngt")
                app.schema_local = AIPRestAPI.get_schema_name(domain.schema, "local")
        return applicationlist
    
    ########################################################################
    def get_transactions_per_business_criterion(self, domainname, applicationid, snapshotid, bcid, nbrows):
        self.restutils.logger.info("Extracting the transactions for business criterion " + bcid)
        request = domainname + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/transactions/" + bcid
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)    
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def get_application_snapshots_json(self, domainname, applicationid):
        request = domainname + "/applications/" + applicationid + "/snapshots" 
        return self.restutils.execute_requests_get(request)
    
    def get_application_snapshot_modules_json(self, domainname, applicationid, snapshotid):
        request = domainname + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + "/modules" 
        return self.restutils.execute_requests_get(request)    
    
    def get_application_snapshots(self, domainname, applicationid):
        snapshotlist = Snapshot.loadlist(self.get_application_snapshots_json(domainname, applicationid))
        for it in snapshotlist:
            modulelist = Module.loadlist(self.get_application_snapshot_modules_json(domainname, applicationid, it.snapshotid))
            it.modules = modulelist 
        return snapshotlist 
    
    ########################################################################
    def get_total_number_violations_json(self, domain, applicationid, snapshotid):
        self.restutils.logger.info("Extracting the number of violations")
        request = domain + "/results?sizing-measures=67011,67211&application=" + applicationid + "&snapshot=" + snapshotid
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def get_qualitydistribution_details_json(self, domain, applicationid, snapshotid, metricid, category, nbrows):
        request = domain + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + '/components/' + str(metricid) + '/'
        request += str(category)+'?business-criterion=60017&startRow=1&nbRows=' + str(nbrows)
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def get_dict_cyclomaticcomplexity_distribution(self, domain, applicationid, snapshotid, nbrows):
        dict = {}
        #Very High Complexity Artifacts
        categories = [1,2,3,4]
        labels = {1:'Very High Complexity Artifacts',2:'High Complexity Artifacts',3:'Moderate Complexity Artifacts',4:'Low Complexity Artifacts'}
        for cat in categories:
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_CYCLOMATIC_COMPLEXITY, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_COST_COMPLEXITY, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_FAN_OUT, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_FAN_IN, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_COUPLING, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_SIZE, cat, nbrows)
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
            json = self.get_qualitydistribution_details_json(domain, applicationid, snapshotid, Metric.DIST_SQL_COMPLEXITY, cat, nbrows)
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
    def get_snapshot_violations_json(self, domainname, applicationid, snapshotid, violationfilter):
        self.restutils.logger.info("Extracting the snapshot violations")
        request = domainname + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/violations'
        request += '?startRow=1'
        request += '&nbRows=' + str(violationfilter.nbrows)
        if violationfilter.criticalrulesonly != None and violationfilter.criticalrulesonly:         
            request += '&rule-pattern=critical-rules'
        if violationfilter.violationstatus != None:
            request += '&status=' + violationfilter.violationstatus
        if violationfilter.businesscriterion == None:
            violationfilter.businesscriterion = "60017"
        if violationfilter.businesscriterion != None:
            strbusinesscriterionfilter = str(violationfilter.businesscriterion)        
            # we can have multiple value separated with a comma
            if ',' not in strbusinesscriterionfilter:
                request += '&business-criterion=' + strbusinesscriterionfilter
            request += '&rule-pattern=('
            for item in strbusinesscriterionfilter.split(sep=','):
                request += 'cc:' + item + ','
                if violationfilter.criticalrulesonly == None or not violationfilter.criticalrulesonly:   
                    request += 'nc:' + item + ','
            request = request[:-1]
            request += ')'
            
        if violationfilter.techno != None:
            request += '&technologies=' + violationfilter.techno
            
        return self.restutils.execute_requests_get(request)
        
        
    def get_snapshot_violations(self, domainname, applicationid, snapshotid, edurl, snapshothref, tqiqm, listtccontributions, violationfilter):
        listviolations = []
        json_violations = self.get_snapshot_violations_json(domainname, applicationid, snapshotid, violationfilter)                                            
        if json_violations != None:
            iCouterRestAPIViolations = 0
            for violation in json_violations:
                objviol = Violation()
                iCouterRestAPIViolations += 1
                currentviolurl = ''
                violations_size = len(json_violations)
                imetricprogress = int(100 * (iCouterRestAPIViolations / violations_size))
                if iCouterRestAPIViolations==1 or iCouterRestAPIViolations==violations_size or iCouterRestAPIViolations%3000 == 0:
                    LogUtils.loginfo(self.restutils.logger,"processing violation " + str(iCouterRestAPIViolations) + "/" + str(violations_size)  + ' (' + str(imetricprogress) + '%)',True)
                try:
                    objviol.qrname = violation['rulePattern']['name']
                except KeyError:
                    qrname = None    
                       
                try:                                    
                    qrrulepatternhref = violation['rulePattern']['href']
                except KeyError:
                    qrrulepatternhref = None
                                                                
                qrrulepatternsplit = qrrulepatternhref.split('/')
                for elem in qrrulepatternsplit:
                    # the last element is the id
                    objviol.qrid = elem                                            
                
                # critical contribution
                objviol.qrcritical = '<Not extracted>'
                try:
                    qrdetails = tqiqm[objviol.qrid]
                    if tqiqm != None and qrdetails != None and qrdetails.get("critical") != None:
                        objviol.critical = str(qrdetails.get("critical"))
                except KeyError:
                    LogUtils.logwarning(self.restutils.logger, 'Could not find the critical contribution for %s'% str(objviol.qrid), True)
                    
                # filter on quality rule id or name, if the filter match
                if violationfilter.qrid != None and not re.match(violationfilter.qrid, str(objviol.qrid)):
                    continue
                if violationfilter.qrname != None and not re.match(violationfilter.qrname, qrname):
                    continue
                actionPlan = violation['remedialAction']
                try:               
                    objviol.hasActionPlan = actionPlan != None
                except KeyError:
                    self.restutils.logger.warning('Not able to extract the action plan')
                if objviol.hasActionPlan:
                    try:               
                        objviol.actionplanstatus = actionPlan['status']
                        objviol.actionplantag = actionPlan['tag']
                        objviol.actionplancomment = actionPlan['comment']
                    except KeyError:
                        self.restutils.logger.warning('Not able to extract the action plan details')
                try:                                    
                    objviol.hasExclusionRequest = violation['exclusionRequest'] != None
                except KeyError:
                    self.restutils.logger.warning('Not able to extract the exclusion request')
                # filter the violations already in the exclusion list 
                try:                                    
                    objviol.violationstatus = violation['diagnosis']['status']
                except KeyError:
                    self.restutils.logger.warning('Not able to extract the violation status')
                try:
                    componentHref = violation['component']['href']
                except KeyError:
                    componentHref = None

                objviol.componentid = ''
                rexcompid = "/components/([0-9]+)/snapshots/"
                m0 = re.search(rexcompid, componentHref)
                if m0: 
                    objviol.componentid = m0.group(1)
                if qrrulepatternhref != None and componentHref != None:
                    objviol.id = qrrulepatternhref+'#'+componentHref
                try:
                    objviol.componentShortName = violation['component']['shortName']
                except KeyError:
                    self.restutils.logger.warning('Not able to extract the componentShortName')                                     
                try:
                    objviol.componentNameLocation = violation['component']['name']
                except KeyError:
                    self.restutils.logger.warning('Not able to extract the componentNameLocation')
                # filter on component name location
                try:
                    objviol.componentstatus = violation['component']['status']
                except KeyError:
                    componentStatus = None                                            
                try:
                    findingsHref = violation['diagnosis']['findings']['href']
                except KeyError:
                    findingsHref = None                                            
                try:
                    componentTreeNodeHref = violation['component']['treeNodes']['href']
                except KeyError:
                    componentTreeNodeHref = None                                        
                try:
                    sourceCodesHref = violation['component']['sourceCodes']['href']
                except KeyError:
                    sourceCodesHref = None
                
                try:
                    propagationRiskIndex = violation['component']['propagationRiskIndex']
                except KeyError:
                    propagationRiskIndex = None                                            
        
                firsttechnicalcriterionid = '#N/A#'
                for tcc in listtccontributions:
                    if tcc.metricid ==  objviol.qrid:
                        firsttechnicalcriterionid = tcc.parentmetricid
                        break 
                currentviolfullurl = edurl + '/engineering/index.html#' + snapshothref
                currentviolfullurl += '/business/60017/qualityInvestigation/0/60017/' 
                currentviolfullurl += firsttechnicalcriterionid + '/' + objviol.qrid + '/' + objviol.componentid
                objviol.url = currentviolfullurl
        
                listviolations.append(objviol)        
        return listviolations
    ########################################################################
    def get_tqi_transactions_violations_json(self, domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows):    
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
            
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def create_scheduledexclusions_json(self, domain, applicationid, snapshotid, json_exclusions_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/exclusions/requests'
        return self.restutils.execute_request_post(request, 'application/json',json_exclusions_to_create)
    
    ########################################################################
    def create_actionplans_json(self, domain, applicationid, snapshotid, json_actionplans_to_create):
        request = domain + "/applications/" + applicationid + "/snapshots/" + snapshotid + '/action-plan/issues'
        return self.restutils.execute_request_post(request, 'application/json',json_actionplans_to_create)

    ########################################################################
    def get_rule_pattern(self, rulepatternHref):
        self.restutils.logger.debug("Extracting the rule pattern details")   
        request = rulepatternHref
        json_rulepattern =  self.restutils.execute_requests_get(request)
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

    # extract the last element that is the id
    def get_hrefid(self, href, separator='/'):
        if href == None or separator == None:
            return None
        _id = ""
        hrefsplit = href.split('/')
        for elem in hrefsplit:
            # the last element is the id
            _id = elem    
        return _id


    ########################################################################
    # Action plan summary
    def get_actionplan_summary_json(self, domainname, applicationid, snapshotid):
        request = domainname + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/action-plan/summary"
        return self.restutils.execute_requests_get(request)

    def get_actionplan_summary(self,domainname,applicationid, snapshotid):
        dictapsummary = {}
        json_apsummary = self.get_actionplan_summary_json(domainname, applicationid, snapshotid)
        if json_apsummary != None:
            for qrap in json_apsummary:
                qrhref = qrap['rulePattern']['href']
                qrid = self.get_hrefid(qrhref)
                addedissues = 0
                pendingissues = 0
                addedissues  = qrap['addedIssues']
                pendingissues  = qrap['pendingIssues']
                numberofactions = addedissues + pendingissues
                dictapsummary.update({qrid:numberofactions})
        return dictapsummary

    ########################################################################
    # Educate feature
    def get_actionplan_triggers_json(self, domainname, applicationid, snapshotid):
        request = domainname + "/applications/" + applicationid + "/snapshots/" + snapshotid + "/action-plan/triggers"
        return self.restutils.execute_requests_get(request)

    def get_actionplan_triggers(self, domainname, applicationid, snapshotid):
        dict_aptriggers = {}
        json_aptriggers = self.get_actionplan_triggers_json(domainname, applicationid, snapshotid)
        if json_aptriggers != None:
            for qrap in json_aptriggers:
                qrhref = qrap['rulePattern']['href']
                qrid = self.get_hrefid(qrhref)
                active = qrap['active']
                dict_aptriggers.update({qrid:active})
        return dict_aptriggers
    ########################################################################
    def get_qualitymetrics_results_json(self, domainname, applicationid, snapshotfilter, snapshotids, criticalonly,  modules=None, nbrows=10000000):
        LogUtils.loginfo(self.restutils.logger,'Extracting the quality metrics results',True)
        request = domainname + "/applications/" + applicationid + "/results?quality-indicators"
        request += '=(business-criteria,cc:60017'
        if criticalonly == None or not criticalonly:   
            request += ',nc:60017'
        request += ')&select=(evolutionSummary,violationRatio,aggregators)'
        strsnapshotfilter = ''
        if snapshotfilter != None:
            strsnapshotfilter = "&snapshots=" + snapshotfilter
        elif snapshotids != None:
            strsnapshotfilter = "&snapshot-ids=" + snapshotids
        else:
            strsnapshotfilter = '&snapshots=-1'
        request += strsnapshotfilter
        if modules != None:
            request += '&modules=' + modules
        request += '&startRow=1'
        request += '&nbRows=' + str(nbrows)
        return self.restutils.execute_requests_get(request)

    def get_qualitymetrics_results_by_snapshotids_json(self, domainname, applicationid, snapshotids, criticalonly, modules, nbrows):
        return self.get_qualitymetrics_results_json(domainname, applicationid, None, snapshotids, criticalonly, modules, nbrows)

    def get_qualitymetrics_results_allsnapshots_json(self, domainname, applicationid, snapshotfilter, criticalonly, modules, nbrows):
        return self.get_qualitymetrics_results_json(domainname, applicationid, snapshotfilter, None, criticalonly, modules, nbrows)

    def get_metric_from_json(self, json_metric, parent_metric=None):
        if json_metric == None:
            return None
    
        metric = Metric()
        if parent_metric == None:
            
            try:
                metric.type = json_metric['type']
                #if metric.type != "quality-rules":
                #    print("not a qr")
            except KeyError:
                None
            try:
                metric.id = json_metric['reference']['key']
            except KeyError:
                None
            try:
                metric.name = json_metric['reference']['name']
            except KeyError:
                None                                                    
            try:
                metric.critical = json_metric['reference']['critical']
            except KeyError:
                None
        else:
            metric.type = parent_metric.type
            metric.id = parent_metric.id
            metric.name = parent_metric.name
            metric.critical = parent_metric.critical 
        """
        self.restutils.logger.debug("D01 " +  metric.name + " " + str(json_metric))
        self.restutils.logger.debug("D02 " +  str(json_metric['result']))
        self.restutils.logger.debug("D03 " +  str(json_metric['result']['grade']))
        if json_metric == None or json_metric['result']== None or json_metric['result']['grade'] == None:
            None
        """
        hasresult = False
        try:
            hasresult = json_metric['result'] != None
        except:
            None

        if hasresult:        
            try:
                metric.grade = json_metric['result']['grade']
            except:
                self.restutils.logger.warning("Metric %s has an empty grade " % str(metric.name))
                # if there is no grade for the modules, we skip
                # we don't skip for the application metric, even if it's not normal but we might have a grade for the module and we want to process in this case
                if parent_metric != None:
                    return None
                    
            try:
                metric.failedchecks = json_metric['result']['violationRatio']['failedChecks']
            except KeyError:
                None                                                          
            try:
                metric.successfulchecks = json_metric['result']['violationRatio']['successfulChecks']
            except KeyError:
                None                                                             
            try:
                metric.totalchecks = json_metric['result']['violationRatio']['totalChecks']
            except KeyError:
                None                                                         
            try:
                metric.ratio = json_metric['result']['violationRatio']['ratio']
            except KeyError:
                None
                
            if  metric.ratio == None:
                self.restutils.logger.warning("Metric %s has an empty compliance ratio " % str(metric.name))                                                           
            try:
                metric.addedviolations = json_metric['result']['evolutionSummary']['addedViolations']                                              
            except KeyError:
                None   
            try:
                metric.removedviolations = json_metric['result']['evolutionSummary']['removedViolations']
            except KeyError:
                None      
              
        return metric

    def get_qualitymetrics_results(self, domainname, applicationid, snapshotid, tqiqm, criticalonly, modules=None, aggregationmode='FullApplication', nbrows=10000000):
        dictmetrics = {}
        dicttechnicalcriteria = {}
        listbc = []

        dictmodules = None
        if modules != None:
            dictmodules = {}
        
        json_qr_results = self.get_qualitymetrics_results_by_snapshotids_json(domainname, applicationid, snapshotid, criticalonly, modules, nbrows)
        for res in json_qr_results:
            iCount = 0
            lastProgressReported = None
            for res_app in res['applicationResults']:
                iCount += 1
                b_has_grade = False
                metricssize = len(res['applicationResults'])
                imetricprogress = int(100 * (iCount / metricssize))
                if imetricprogress in (9,19,29,39,49,59,69,79,89,99) : 
                    if lastProgressReported == None or lastProgressReported != imetricprogress:
                        LogUtils.loginfo(self.restutils.logger,  ' ' + str(imetricprogress+1) + '% of the metrics processed',True)
                        lastProgressReported = imetricprogress
                # parse the json
                metric = self.get_metric_from_json(res_app)
                # skip the metrics that have no grade
                if metric == None:
                    continue
                metric.applicationName = res['application']['name'] 
                
                if metric.type in ("quality-measures","quality-distributions","quality-rules"):
                    if (metric.grade == None): 
                        b_has_grade = False
                    else:
                        b_has_grade = True
                    #else:
                    if 1==1:
                        if metric.type == "quality-rules":
                            try:
                                metric.threshold1 = tqiqm[''+metric.id].get("threshold1")
                                metric.threshold2 = tqiqm[''+metric.id].get("threshold2")
                                metric.threshold3 = tqiqm[''+metric.id].get("threshold3")
                                metric.threshold4 = tqiqm[''+metric.id].get("threshold4")
                            except KeyError:
                                None
                            
                            json_thresholds = None
                            # loading from another place when tresholds are empty
                            if not metric.threshold1 or not metric.threshold2 or not metric.threshold3 or not metric.threshold4:
                                #LogUtils.loginfo(logger,'Extracting the quality rules thresholds',True)
                                json_thresholds = self.get_qualityrules_thresholds_json(domainname, snapshotid, metric.id)   
                            if json_thresholds != None and json_thresholds.get('thresholds') != None and json_thresholds['thresholds'] != None:
                                icount = 0
                                for thres in json_thresholds['thresholds']:
                                    icount += 1
                                    if icount == 1: metric.threshold1=thres
                                    if icount == 2: metric.threshold2=thres
                                    if icount == 3: metric.threshold3=thres
                                    if icount == 4: metric.threshold4=thres
                        
                elif metric.type == "technical-criteria":
                    #print('tc grade=' + str(metric.grade) + str(type(metric.grade)))
                    if (metric.grade == None): 
                        LogUtils.logwarning(self.restutils.logger, "Technical criterion has no grade, removing it from the list : " + metric.name)
                    else:
                        dicttechnicalcriteria[metric.id] = metric
                elif metric.type == 'business-criteria':
                    if (metric.grade == None): 
                        self.restutils.logger.warning("Business criterions has no grade, removing it from the list : " + metric.name)
                    else:
                        listbc.append(metric)
                
                b_one_module_has_grade = False
                if modules:
                    try:
                        for res_mod in res_app['moduleResults']:
                            
                            mod_name = res_mod['moduleSnapshot']['name'] 
                            if not dictmodules.get(mod_name):
                                dictmodules[mod_name] = {}
                            metric_module = self.get_metric_from_json(res_mod, metric)
                            # skip the metrics that have no grade
                            if metric_module == None:
                                continue
                            if metric_module.grade != None:
                                b_one_module_has_grade = True
                            metric_module.applicationName = res['application']['name']
                            metric_module.threshold1 = metric.threshold1
                            metric_module.threshold2 = metric.threshold2
                            metric_module.threshold3 = metric.threshold3
                            metric_module.threshold4 = metric.threshold4
                            dictmodules[mod_name][metric_module] = metric_module
                    except KeyError:
                        None
                # we add the metric only if it has a grade at application level or has grade at least at module level
                if metric.type in ("quality-measures","quality-distributions","quality-rules"):
                    if b_has_grade or b_one_module_has_grade:
                        dictmetrics[metric.id] = metric  
                    if not b_has_grade or (modules != None and not b_one_module_has_grade):
                        LogUtils.logwarning(self.restutils.logger, "Metric %s" % metric.name, True)
                    if not b_has_grade:
                        LogUtils.logwarning(self.restutils.logger, "    has no grade at application level", True)
                    if modules != None and not b_one_module_has_grade:
                        LogUtils.logwarning(self.restutils.logger, "    has no grade at module level", True)
                    if not b_has_grade and not b_one_module_has_grade:
                        LogUtils.logwarning(self.restutils.logger, "    skipping metric", True)
                           
        return dictmetrics, dicttechnicalcriteria, listbc, dictmodules

    ########################################################################
    def get_all_snapshots_json(self, domain):
        request = domain + "/results?snapshots=" + AIPRestAPI.FILTER_SNAPSHOTS_ALL
        return self.restutils.execute_requests_get(request)

    ########################################################################
    def get_loc_json(self, domain, snapshotsfilter=None):
        return self.get_sizing_measures_by_id_json(domain, Metric.TS_LOC, snapshotsfilter)
  
    ########################################################################
    def get_nb_artifacts_json(self, domain, snapshotsfilter=None, modules=None):
        return self.get_sizing_measures_by_id_json(domain, Metric.TS_NB_ARTIFACTS, snapshotsfilter, modules)
  
    ########################################################################
    def get_nb_artifacts_dict(self, domain, snapshotsfilter=None, modules=None):
        dict_nb_art = None
        json_nb_art = self.get_sizing_measures_by_id_json(domain, Metric.TS_NB_ARTIFACTS, snapshotsfilter, modules)
        if json_nb_art != None:
            dict_nb_art = {}
            try:
                for res_mod1 in json_nb_art:
                    for res_mod2 in res_mod1['applicationResults']:
                        for res_mod3 in res_mod2['moduleResults']:
                            module_name = res_mod3['moduleSnapshot']['name']
                            module_weight = res_mod3['result']['value']
                            dict_nb_art[module_name] = module_weight
                    break
            except KeyError:
                None
        return dict_nb_art   
  
    ########################################################################
    def get_afp_json(self, domain, snapshotsfilter=None):
        return self.get_sizing_measures_by_id_json(domain, Metric.FS_AFP, snapshotsfilter)
  
    ########################################################################
    def get_tqi_json(self, domain, snapshotsfilter=None):
        return self.get_quality_indicators_by_id_json(domain, Metric.BC_TQI, snapshotsfilter)  
    
    ########################################################################
    def get_sizing_measures_by_id_json(self, domain, metricids, snapshotsfilter=None, modules=None):
        snapshot_index = None 
        snapshot_ids = None
        if snapshotsfilter != None:
            if snapshotsfilter.snapshot_index != None:
                snapshot_index = snapshotsfilter.snapshot_index
            if snapshotsfilter.snapshot_ids != None:
                snapshot_ids = snapshotsfilter.snapshot_ids
        if snapshot_index == None and snapshot_ids == None:
            snapshot_index == AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?sizing-measures=(" + str(metricids) + ")"
        if snapshot_index != None:
            request += "&snapshots=" + snapshot_index
        if snapshot_ids != None:
            request += "&snapshot-ids=" + snapshot_ids            
        if modules != None:
            request += "&modules=" + modules
        return self.restutils.execute_requests_get(request)

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
    def get_sizing_measures_json(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?sizing-measures=(technical-size-measures,technical-debt-statistics,run-time-statistics,functional-weight-measures)&snapshots="+snapshotsfilter
        return self.restutils.execute_requests_get(request)

    ########################################################################
    def get_quality_indicators_json(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?quality-indicators=(quality-rules,business-criteria,technical-criteria,quality-distributions,quality-measures)&select=(evolutionSummary,violationRatio,aggregators,categories)&snapshots="+snapshotsfilter
        return self.restutils.execute_requests_get(request)



    ########################################################################
    def get_quality_indicators_by_id_json(self, domain, metricids, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST
        request = domain + "/results?quality-indicators=(" + str(metricids) + ")&select=(evolutionSummary,violationRatio,aggregators,categories)&snapshots="+snapshotsfilter
        return self.restutils.execute_requests_get(request)

    ########################################################################
    def get_metric_contributions_json(self, domain, metricid, snapshotid):  
        request = domain + "/quality-indicators/" + str(metricid) + "/snapshots/" + snapshotid 
        return self.restutils.execute_requests_get(request)

    def get_metric_contributions(self, domain, metricid, snapshotid):
        return Contribution.loadlist(self.get_metric_contributions_json(domain, metricid, snapshotid))

    #################################Violation#######################################
    def get_qualityrules_thresholds_json(self, domain, snapshotid, qrid):
        #LogUtils.loginfo(self.restutils.logger,'Extracting the quality rules thresholds',True)
        request = domain + "/quality-indicators/" + str(qrid) + "/snapshots/"+ snapshotid
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def get_businesscriteria_grades_json(self, domain, snapshotsfilter=None):
        if snapshotsfilter == None:
            snapshotsfilter = AIPRestAPI.FILTER_SNAPSHOTS_LAST        
        request = domain + "/results?quality-indicators=(60017,60016,60014,60013,60011,60012,66031,66032,66033,60013,20140522)&applications=($all)&snapshots=" + snapshotsfilter 
        return self.restutils.execute_requests_get(request)

    ########################################################################
    def get_snapshot_tqi_quality_model_json (self, domainname, snapshotid):
        self.restutils.logger.info("Extracting the snapshot quality model")   
        request = domainname + "/quality-indicators/60017/snapshots/" + snapshotid + "/base-quality-indicators" 
        return self.restutils.execute_requests_get(request)

    def get_snapshot_quality_model_qualitymetrics_json (self, domainname, snapshotid):
        self.restutils.logger.info("Extracting the snapshot quality model - quality-rules")   
        request = domainname + "/configuration/snapshots/" + snapshotid + "/quality-rules" 
        return self.restutils.execute_requests_get(request)    
    
    def get_snapshot_quality_model_qualitydistributions_json (self, domainname, snapshotid):
        self.restutils.logger.info("Extracting the snapshot quality model - quality-distributions")   
        request = domainname + "/configuration/snapshots/" + snapshotid + "/quality-distributions" 
        return self.restutils.execute_requests_get(request)    

    def get_snapshot_quality_model_qualitymeasures_json (self, domainname, snapshotid):
        self.restutils.logger.info("Extracting the snapshot quality model - quality-measures")   
        request = domainname + "/configuration/snapshots/" + snapshotid + "/quality-measures" 
        return self.restutils.execute_requests_get(request)    

    ########################################################################
    def get_snapshot_tqi_quality_model (self, domainname, snapshotid):
        tqiqm = {}
        ''' 
        5 metrics are missing here (quality-measures and quality-distributions), because they don't contribute to the Total quality index :
        67001 : Cost Complexity distribution
        67030 : Distribution of defects to critical diagnostic-based metrics per cost complexity
        67020 : Distribution of violations to critical diagnostic-based metrics per cost complexity
        62003 : SEI Maintainability Index 3
        62004 : SEI Maintainability Index 4
        '''
        json_snapshot_quality_model = self.get_snapshot_tqi_quality_model_json(domainname, snapshotid)
        json_snapshot_qualitymetrics = self.get_snapshot_quality_model_qualitymetrics_json(domainname, snapshotid)
        json_snapshot_qualitydistributions = self.get_snapshot_quality_model_qualitydistributions_json(domainname, snapshotid)
        json_snapshot_qualitymeasures = self.get_snapshot_quality_model_qualitymeasures_json(domainname, snapshotid)
        listqualitymetrics = []
        listqualitydistributions = []
        listqualitymeasures = []
        for qmitem in json_snapshot_qualitymetrics: listqualitymetrics.append(qmitem['key'])
        for qmitem in json_snapshot_qualitydistributions: listqualitydistributions.append(qmitem['key'])
        for qmitem in json_snapshot_qualitymeasures: listqualitymeasures.append(qmitem['key'])        
        
        if json_snapshot_quality_model != None:
            for qmitem in json_snapshot_quality_model:
                maxWeight = -1
                qrid = qmitem['key']
                qrcompoundWeight = qmitem['compoundedWeight'] 
                qrcompoundWeightFormula = qmitem['compoundedWeightFormula']
                regexp = "\([0-9]+x([0-9]+)\)"                                            
                for m in re.finditer(regexp, qrcompoundWeightFormula):
                    if m.group(1) != None and int(m.group(1)) > int(maxWeight): 
                        maxWeight = int(m.group(1))                                            
                qrname = qmitem['name']
                qrcritical = qmitem['critical']
                qrtype=None
                threshold1 = None
                threshold2 = None 
                threshold3 = None
                threshold4 = None
                if qrid in listqualitymetrics: 
                    qrtype = 'quality-rules'
                    try:
                        if qmitem['thresholds'] != None:
                            icount = 0
                            for thres in qmitem['thresholds']:
                                icount += 1
                                if icount == 1: threshold1=thres
                                if icount == 2: threshold2=thres
                                if icount == 3: threshold3=thres
                                if icount == 4: threshold4=thres                                                
                    except KeyError:
                        None
                elif qrid in listqualitydistributions: qrtype = 'quality-distributions'
                elif qrid in listqualitymeasures: qrtype = 'quality-measures'
                    
                tqiqm[qrid] = {"critical":qrcritical,"type": qrtype, "hasresults":False, "name":qrname, "tc":{},"maxWeight":maxWeight,"compoundedWeight":qrcompoundWeight,"compoundedWeightFormula":qrcompoundWeightFormula}
                if qrid in listqualitymetrics: 
                    # adding the thresholds
                    tqiqm.get(qrid).update({"threshold1":threshold1, "threshold2:":threshold2,"threshold3":threshold3, "threshold4" : threshold4}) 
                    #tqiqm[qrid] = {"critical":qrcrt,"tc":{},"maxWeight":maxWeight,"compoundedWeight":qrcompoundWeight,"compoundedWeightFormula":qrcompoundWeightFormula, 
                    #""               "threshold1":threshold1, "threshold2:":threshold2,"threshold3":threshold3, "threshold4" : threshold4}  
                #contains the technical criteria (might be several) for each rule, we keep the fist one
                for tccont in qmitem['compoundedWeightTerms']:
                    _term = tccont['term'] 
                    #tqiqm[qrid] = {tccont['technicalCriterion']['key']: tccont['technicalCriterion']['name']} 
                    #TODO: add qrcompoundWeight and/or qrcompoundWeightFormula
                    tqiqm.get(qrid).get("tc").update({tccont['technicalCriterion']['key']: tccont['technicalCriterion']['name']})
        return tqiqm
    ########################################################################
    def get_snapshot_bc_tc_mapping_json(self, domain, snapshotid, bcid):
        self.restutils.logger.info("Extracting the snapshot business criterion " + bcid +  " => technical criteria mapping")  
        request = domain + "/quality-indicators/" + str(bcid) + "/snapshots/" + snapshotid
        return self.restutils.execute_requests_get(request)
    
    ########################################################################
    def get_components_pri_json (self, domain, applicationid, snapshotid, bcid,nbrows):
        self.restutils.logger.info("Extracting the components PRI for business criterion " + bcid)  
        request = domain + "/applications/" + str(applicationid) + "/snapshots/" + str(snapshotid) + '/components/' + str(bcid)
        request += '?startRow=1'
        request += '&nbRows=' + str(nbrows)
        return self.restutils.execute_requests_get(request)
         
    ########################################################################
    def get_sourcecode_json(self, sourceCodesHref):
        return self.restutils.execute_requests_get(sourceCodesHref)     
    
    ########################################################################
    def get_sourcecode_file_json(self, filehref, srcstartline, srcendline):
        strstartendlineparams = ''
        if srcstartline != None and srcstartline >= 0 and srcendline != None and srcendline >= 0:
            strstartendlineparams = '?start-line='+str(srcstartline)+'&end-line='+str(srcendline)
        return self.restutils.execute_requests_get(filehref+strstartendlineparams, 'text/plain', 'text/plain')    
    
    ########################################################################
    def get_objectviolation_metrics(self, objHref):
        self.restutils.logger.debug("Extracting the component metrics")    
        request = objHref
        json_component = self.restutils.execute_requests_get(request)
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
    def get_objectviolation_findings_json(self, objHref, qrid):
        self.restutils.logger.debug("Extracting the component findings")    
        request = objHref + '/findings/' + qrid 
        return self.restutils.execute_requests_get(request)
         
    ########################################################################
    # extract the transactions TRI & violations component list per business criteria
    def init_transactions (self, domain, applicationid, snapshotid, criticalonly, violationStatus, technoFilter,nbrows):
        #TQI,Security,Efficiency,Robustness
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
                        json_tran_violations = self.get_tqi_transactions_violations_json(domain, snapshotid, transactionid, criticalonly, violationStatus, technoFilter,nbrows)                  
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
            json_snapshot_components_pri = self.get_components_pri_json(domain, applicationid, snapshotid, bcid,nbrows)
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
            json = self.get_snapshot_bc_tc_mapping_json(domain, snapshotid, bcid)
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
    TS_NB_ARTIFACTS="10152"
    
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
  
    @staticmethod
    def loadlist(json_contributions):
        listcontributions = []
        if json_contributions != None:
            for json in json_contributions['gradeContributors']:
                listcontributions.append(Contribution.load(json, json_contributions['name'], json_contributions['key'] )) 
        return listcontributions   

    @staticmethod
    def load(json_contribution, parentmetricname, parentmetricid):
        x = Contribution()
        if json_contribution != None:
            x.parentmetricname = parentmetricname
            x.parentmetricid = parentmetricid
            x.metricname = json_contribution['name']
            x.metricid = json_contribution['key']
            x.critical = json_contribution['critical']
            x.weight = json_contribution['weight']      
        return x
  
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
   
# Logging utils
class LogUtils:

    @staticmethod
    def logdebug(logger, msg, tosysout = False):
        logger.debug(msg)
        if tosysout:
            print(msg)

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