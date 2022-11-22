from utils.utils import LogUtils, StringUtils
import xlsxwriter
import pandas as pd
import numpy as np
from io import StringIO
import csv

broundgrades = None

type_quality_rules = "quality-rules"
type_quality_distributions = "quality-distributions"
type_quality_measuree = "quality-measures"
type_technical_criteria = "technical-criteria"
type_business_criteria = "business-criteria"


# excel format class
class ExcelFormat:
    const_float_format = '{:,.2f}'
    const_format_percentage = '0.00%'
    const_float_with_2decimals = '#,##0.00'
    const_float_with_1decimal = '#,##0.0'
    const_format_align_left = 'left'
    const_format_int = '### ### ### ### ### ##0'

    #const_color_tabs_interactive = 'blue'
    
    
    const_color_green = '#C6EFCE'
    const_color_red = '#FFC7CE'
    const_color_light_grey = '#B2B2B2'
    const_color_light_blue = '#5DD5F1'
    const_color_light_rose = '#FFCCCC'
    # header color= green
    const_color_header_columns = '#D7E4BC'
    
    # Tab names
    const_TAB_README = 'README'
    const_TAB_APP_BC_GRADES = 'BC Grades'
    const_TAB_APP_TC_GRADES = 'TC Grades'
    const_TAB_APP_RULES_GRADES = 'Rules Grades'
    const_TAB_APP_VIOLATIONS = 'Violations_2'
    const_TAB_APP_BC_CONTRIBUTIONS = 'BC contributions'    
    const_TAB_APP_TC_CONTRIBUTIONS = 'TC contributions'
    const_TAB_REMEDIATION_EFFORT = 'Remediation effort'
    
    const_TAB_MOD_BC_GRADES = 'Modules BC Grades'
    const_TAB_MOD_TC_GRADES = 'Modules TC Grades'
    const_TAB_MOD_RULES_GRADES = 'Modules Rules Grades'
    const_TAB_MOD_BC_CONTRIBUTIONS = 'Modules BC contributions'    
    const_TAB_MOD_TC_CONTRIBUTIONS = 'Modules TC contributions'
    const_TAB_MOD_WEIGHT = 'Modules weight'
    
    format_percentage = None
    format_int_thousands = None
    format_align_left = None
    
    format_green_percentage = None
    format_red_percentage = None
    format_grey_float_1decimal = None
    format_green_int = None
    format_red_int = None
    format_green_int = None


tab_app_bc_grade        = ExcelFormat.const_TAB_APP_BC_GRADES
tab_app_tc_grade        = ExcelFormat.const_TAB_APP_TC_GRADES
tab_app_rule_grade      = ExcelFormat.const_TAB_APP_RULES_GRADES
tab_app_bc_cont         = ExcelFormat.const_TAB_APP_BC_CONTRIBUTIONS
tab_app_tc_cont         = ExcelFormat.const_TAB_APP_TC_CONTRIBUTIONS
tab_app_viol            = ExcelFormat.const_TAB_APP_VIOLATIONS
tab_mod_bc_grade        = ExcelFormat.const_TAB_MOD_BC_GRADES
tab_mod_tc_grade        = ExcelFormat.const_TAB_MOD_TC_GRADES
tab_mod_rule_grade      = ExcelFormat.const_TAB_MOD_RULES_GRADES
tab_mod_bc_cont         = ExcelFormat.const_TAB_MOD_BC_CONTRIBUTIONS
tab_mod_tc_cont         = ExcelFormat.const_TAB_MOD_TC_CONTRIBUTIONS
tab_mod_weigth          = ExcelFormat.const_TAB_MOD_WEIGHT
tab_remd                = ExcelFormat.const_TAB_REMEDIATION_EFFORT

###############################################################################

def get_df_readme(logger, loadviolations):
    df_readme = None
    #Readme Page content
    str_readme_content =  "Tab;Content;Comment\n"
    str_readme_content += ExcelFormat.const_TAB_README + ";Read me;\n"
    str_readme_content += ExcelFormat.const_TAB_APP_BC_GRADES + ";Business Criteria current grade and simulation grade;Use this sheet to see the global impact on application grades and total estimated effort\n"
    str_readme_content += ExcelFormat.const_TAB_APP_TC_GRADES + ";Technical criteria current grade and simulation grade;\n"
    str_readme_content += ExcelFormat.const_TAB_APP_RULES_GRADES +";Quality Rules, Distributions and Measures grades and simulation;Use this sheet to change the number of violations for action and see the impact on rules grades and estimated effort\n"
    if loadviolations:
        str_readme_content += ExcelFormat.const_TAB_APP_VIOLATIONS + ";Violations list;Use this sheet to select your violations for action\n"
    str_readme_content += ExcelFormat.const_TAB_APP_BC_CONTRIBUTIONS + ";Business Criteria contributors (Technical criteria);\n"
    str_readme_content += ExcelFormat.const_TAB_APP_TC_CONTRIBUTIONS + ";Technical Criteria contributors (Quality metrics);\n"
    str_readme_content += ExcelFormat.const_TAB_REMEDIATION_EFFORT + ";Quality rules unit remediation effort;Use to sheet to set or modify the unit remediation effort per quality rule\n"
    
    try: 
        df_readme = pd.read_csv(StringIO(StringUtils.remove_unicode_characters(str_readme_content)), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : readme',True)    
    return df_readme
###############################################################################

bc_grades_commonheader = "Business criterion;Metric Id;Grade;Simulation grade;Lowest critical grade;Weighted average of Technical criteria; Delta"
app_bc_grades_header = "Application name;" + bc_grades_commonheader
mod_bc_grades_header = "Module Name;" + bc_grades_commonheader
###############################################################################

def get_grade_for_display(broundgrades, grade):
    if grade == None:
        return ""
    return str(round_grades(broundgrades, grade))
    
###############################################################################

def get_df_app_bc_grades(logger, appName, snapshotversion, snapshotdate, listbusinesscriteria, loadviolations):
    df_app_bc_grades = None
    str_df_app_bc_grades = app_bc_grades_header
    str_df_app_bc_grades += "\n"
    for bc in listbusinesscriteria:
        if bc.applicationName == appName and not bc.name in ('SEI Maintainability'):#, 'Green IT'):
            str_df_app_bc_grades += appName + ";" + bc.name + ";" + bc.id + ";" + get_grade_for_display(broundgrades, bc.grade) + ";;;;"
            str_df_app_bc_grades += '\n'
    
    
    emptyline = ";;;;;;;\n"
    # Summary
    str_df_app_bc_grades += emptyline+emptyline+emptyline
    str_df_app_bc_grades += ";Application name;" + appName + "\n"
    str_df_app_bc_grades += ";Version;" + snapshotversion + "\n"
    str_df_app_bc_grades += ";Date;" + snapshotdate + "\n"    
    str_df_app_bc_grades += ';Number of violations for action\n'
    str_df_app_bc_grades += ';Number of quality rules for action\n'
    str_df_app_bc_grades += ';Estimated effort (man.days)\n'

    if loadviolations:
        str_df_app_bc_grades += '\n'
        str_df_app_bc_grades += ';Number of action plans added\n'
        str_df_app_bc_grades += ';Number of action plans removed\n'
        #TODO: identify the action plan modified
        str_df_app_bc_grades += ';Number of action plans modified; <Not available>\n'
        str_df_app_bc_grades += ';JSON violations added\n'
        str_df_app_bc_grades += ';JSON violations removed\n'
        str_df_app_bc_grades += ';JSON violations modified\n'
    try: 
        str_df_app_bc_grades = StringUtils.remove_unicode_characters(str_df_app_bc_grades)
        df_app_bc_grades = pd.read_csv(StringIO(str_df_app_bc_grades), sep=";")
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_app_bc_grades %s ' % str_df_app_bc_grades,True)
    return df_app_bc_grades

###############################################################################

tc_grades_commonheader = "Technical criterion name;Metric Id;Grade;Simulation grade;Lowest critical grade;Weighted average of quality rules;Delta grade (%)"
app_tc_grades_header = "Key;Application name;" + tc_grades_commonheader 
mod_tc_grades_header = "Key;Module name;" + tc_grades_commonheader
###############################################################################

def get_df_app_tc_grades(logger, appName, dictechnicalcriteria):
    df_app_tc_grades = None
    str_df_tc_grades = app_tc_grades_header + '\n'
    
    for tc in dictechnicalcriteria:
        #print('tc grade 2=' + str(tc.grade) + str(type(tc.grade)))
        otc = dictechnicalcriteria[tc]
        str_line = ';' + appName + ';' + otc.name + ';' + str(otc.id) + ';'+ get_grade_for_display(broundgrades, otc.grade) + ';;;;' + '\n'
        str_df_tc_grades += str_line
    try: 
        str_df_tc_grades = StringUtils.remove_unicode_characters(str_df_tc_grades)
        df_app_tc_grades = pd.read_csv(StringIO(str_df_tc_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_app_tc_grades %s ' % str_df_tc_grades,True)
    return df_app_tc_grades

###############################################################################
def get_df_mod_tc_grades(logger, dict_modules, dict_modulesweight = None):
    df_mod_tc_grades = None
    str_df_mod_tc_grades = mod_tc_grades_header
    if dict_modulesweight != None:
        str_df_mod_tc_grades += ';Module weight;Module weighted grade'
    str_df_mod_tc_grades += '\n'

    for module_name in dict_modules:
        module_metrics = dict_modules[module_name]
        for otc in module_metrics:
            if otc.type not in (type_technical_criteria):
                continue
            str_line = ';' + module_name + ';' + otc.name + ';' + str(otc.id) + ';'+ get_grade_for_display(broundgrades, otc.grade) + ';;;;'
            if dict_modulesweight != None:
                str_line += ';;'  
            str_line +=  '\n'
            str_df_mod_tc_grades += str_line    
    try: 
        str_df_tc_grades = StringUtils.remove_unicode_characters(str_df_mod_tc_grades)
        df_mod_tc_grades = pd.read_csv(StringIO(str_df_tc_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : str_df_mod_tc_grades %s ' % str_df_mod_tc_grades,True)
    return df_mod_tc_grades

###############################################################################
def get_df_mod_bc_grades(logger, dict_modules, dict_modulesweight=None):
    df_mod_bc_grades = None
    str_df_mod_bc_grades = mod_bc_grades_header
    if dict_modulesweight != None:
        str_df_mod_bc_grades += ';Module weight;Module weighted grade'
    str_df_mod_bc_grades += '\n'
    
    for module_name in dict_modules:
        module_metrics = dict_modules[module_name]
        for obc in module_metrics:
            if obc.type not in (type_business_criteria):
                continue
            if obc.name in ( 'SEI Maintainability'):#, 'Green IT'):
                continue
            str_line = ''
            str_line += module_name + ';' + obc.name + ';' + str(obc.id) + ';'+ get_grade_for_display(broundgrades, obc.grade) + ';;;;'
            if dict_modulesweight != None:
                str_line += ';;'            
            str_line += '\n'
            str_df_mod_bc_grades += str_line    
    try: 
        str_df_mod_bc_grades = StringUtils.remove_unicode_characters(str_df_mod_bc_grades)
        df_mod_bc_grades = pd.read_csv(StringIO(str_df_mod_bc_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : str_df_mod_grades %s ' % str_df_mod_bc_grades,True)
    return df_mod_bc_grades
###############################################################################

rule_grades_commonheader = "Snapshot Date;Snapshot version;Metric Name;Metric Id;Metric Type;Critical;Grade;Simulation grade;Grade Delta;Grade Delta (%);Nb of violations;Nb violations for action;Remaining violations;Unit effort (man.hours);Total effort (man.hours);Total effort (man.days);Total Checks;Compliance ratio;New compliance ratio;Thres.1;Thres.2;Thres.3;Thres.4;Educate (Mark for ...);Violations extracted;Grade improvement priority;Nb violations to fix on this rule before rule grade improvement;Nb rules to fix before TC grade improvement;Nb violations  to fix on the TC  rules before TC grade improvement;Technical criteria;Business criteria"
app_rule_grades_header = "Key;Application name;" + rule_grades_commonheader
mod_rule_grades_header = "Key;Module name;" + rule_grades_commonheader

###############################################################################
def get_df_app_rules_grades(logger, appName, snapshotdate, snapshotversion, dictmetrics, listtccontributions, listbccontributions, dictapsummary=None, dictaptriggers=None, aggregationmode='FullApplication'):
    df_app_rules_grades = None
    str_df_rules_grades = app_rule_grades_header + '\n'
    
    for metric in dictmetrics:
        oqr = dictmetrics[metric]
        if oqr.type not in (type_quality_rules,type_quality_distributions,type_quality_measuree):
            continue
        str_line = ';'
        str_line += appName
        str_line, skip_line  = get_def_rule_grade_line(logger, str_line, snapshotdate, snapshotversion, oqr, listtccontributions, listbccontributions, dictapsummary, dictaptriggers, aggregationmode)
        str_line += '\n'
        if not skip_line:
            str_df_rules_grades += str_line
    #logger.debug(str_df_rules_grades)
    try: 
        str_df_rules_grades = StringUtils.remove_unicode_characters(str_df_rules_grades)
        df_app_rules_grades = pd.read_csv(StringIO(str_df_rules_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_app_rules_grades %s ' % str_df_rules_grades,True)
    return df_app_rules_grades
###############################################################################
def get_def_rule_grade_line(logger, str_line, snapshotdate, snapshotversion, oqr, listtccontributions, listbccontributions, dictapsummary = None, dictaptriggers = None, aggregationmode='FullApplication'):
        skip_line = False
        str_line += ";" + str(snapshotdate) 
        str_line += ";" + str(snapshotversion) 
        str_line += ";" +  str(oqr.name)
        str_line += ";" + str(oqr.id) 
        
        #TODO : formatting problem to fix ? 
        #str_line += ";" + str(oqr.type)
        str_line += ";" + str(oqr.type) 
        
        str_line += ";" + str(oqr.critical)
        #print("qr grade=" + get_grade_for_display(broundgrades, oqr.grade)) 
        if oqr.grade == None and aggregationmode=='FullApplication':
            skip_line = True
        str_line += ";" + get_grade_for_display(broundgrades, oqr.grade) 
        
        #simulation grade, grade delta%, grade delta%
        str_line += ';;;' 
        #failed checks
        str_line += ';'
        if oqr.failedchecks != None: str_line += str(oqr.failedchecks)
        #number of actions
        str_line += ';'
        if dictapsummary != None and dictapsummary.get(oqr.id) != None and oqr.type == type_quality_rules:
            str_line += str(dictapsummary.get(oqr.id)) 
        #remaining violations
        str_line += ';'
        #unit effort mh, total effort mh, total effort md
        str_line += ';;;'
        #total checks 
        str_line += ';'
        if oqr.totalchecks != None: 
            str_line += str(oqr.totalchecks) 
        #compliance ratio
        str_line += ';'
        if oqr.ratio != None:
            #print("compl ratio grade=" + str(oqr.ratio))
            str_line += str(oqr.ratio)
        #new compliance ratio
        str_line += ';'
        #4 thresholds 
        if oqr.type == type_quality_rules:
            str_line += ';'+str(oqr.threshold1)+';'+str(oqr.threshold2)+';'+str(oqr.threshold3)+';' + str(oqr.threshold4)
        else:
            str_line += ';;;;'
        #Educate, if applicable
        str_line += ';'
        if dictaptriggers != None and dictaptriggers.get(oqr.id) != None and oqr.type == type_quality_rules:
            if dictaptriggers.get(oqr.id):
                str_line += 'Action'
            else:
                str_line += 'Continuous improvement'
        #new compliance ratio
        str_line += ';'
        #grade improvement priority
        str_line += ';'
        #Nb violations to fix before grade improvement
        str_line += ';'
        #Nb rules to fix before TC grade improvement
        str_line += ';'
        #Nb violations  to fix on the TC  rules before TC grade improvement
        str_line += ';'
        #Technical criteria
        str_line += ';'
        str_tech_criteria = ''
        str_bus_criteria = ''
        for tcc_contrib in listtccontributions:
            if tcc_contrib.metricid == oqr.id:
                str_tech_criteria += tcc_contrib.parentmetricname + ','
                for bcc_contrib in listbccontributions:
                    if bcc_contrib.metricid == tcc_contrib.parentmetricid:
                        str_bus_criteria += bcc_contrib.parentmetricname + ','
        if str_tech_criteria.endswith(','):
            str_tech_criteria = str_tech_criteria[:-1]
        str_line += str_tech_criteria
        #Business criteria
        str_line += ';'
        if str_bus_criteria.endswith(','):
            str_bus_criteria = str_bus_criteria[:-1]        
        str_line += str_bus_criteria        
        return str_line, skip_line

###############################################################################

def get_df_mod_rules_grades(logger, snapshotdate, snapshotversion, dict_modules, listtccontributions, listbccontributions, dictapsummary = None, dictaptriggers = None, dict_modulesweight=None, aggregationmode='FullApplication'):
    df_mod_rules_grades = None
    
    str_df_mod_rules_grades = mod_rule_grades_header
    if dict_modulesweight != None:
        str_df_mod_rules_grades += ';Module weight;Module weighted grade'
    str_df_mod_rules_grades += '\n'    
    
    for module_name in dict_modules:
        module_metrics = dict_modules[module_name]
        for oqr in module_metrics:
            if oqr.type not in (type_quality_rules,type_quality_distributions,type_quality_measuree):
                continue
            
            str_line = ';'
            str_line += module_name
            str_line, skip_line = get_def_rule_grade_line(logger, str_line, snapshotdate, snapshotversion, oqr, listtccontributions, listbccontributions, dictapsummary, dictaptriggers, aggregationmode)
            if dict_modulesweight != None:
                str_line += ';;'            
            str_line += '\n'
            if not skip_line:
                str_df_mod_rules_grades += str_line
    #logger.debug(str_df_mod_rules_grades)
    try: 
        str_df_mod_rules_grades = StringUtils.remove_unicode_characters(str_df_mod_rules_grades)
        df_mod_rules_grades = pd.read_csv(StringIO(str_df_mod_rules_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_mod_rules_grades %s ' % str_df_mod_rules_grades,True)
    return df_mod_rules_grades

###############################################################################

def get_df_mod_bc_contribution(logger, dict_modules, listbccontributions):
    df_mod_bc_cont = None
    str_df_mod_bc_cont = mod_bc_contribution_header
    for module_name in dict_modules:
        module_metrics = dict_modules[module_name]
        for bcc in listbccontributions:
            parent_found = False
            child_found = False
            for met in module_metrics:
                if met.type == type_business_criteria and met.id == bcc.parentmetricid:
                    parent_found = True
                if met.type == type_technical_criteria and met.id == bcc.metricid:
                    child_found = True
                if parent_found and child_found:
                    break     
            if parent_found and child_found:
                str_df_mod_bc_cont += module_name + ';' + bcc.parentmetricname + ';' + bcc.parentmetricid + ';' + bcc.metricname + ';' + bcc.metricid
                str_df_mod_bc_cont += ';' + str(bcc.weight) + ';' + str(bcc.critical) + ';;'
                str_df_mod_bc_cont += '\n'
    #logger.debug(str_df)
    try: 
        str_df_mod_bc_cont = StringUtils.remove_unicode_characters(str_df_mod_bc_cont)
        df_mod_bc_cont = pd.read_csv(StringIO(str_df_mod_bc_cont), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_mod_bc_contributions %s ' % str_df_mod_bc_cont,True)
    return df_mod_bc_cont

###############################################################################

def get_df_mod_tc_contribution(logger, dict_modules, listtccontributions):
    df_mod_tc_cont = None
    str_df_mod_tc_cont = get_tc_contribution_header('Module')
    
    for module_name in dict_modules:
        module_metrics = dict_modules[module_name]
        for tcc in listtccontributions:
            parent_found = False
            child_found = False
            for met in module_metrics:
                if met.type == type_technical_criteria and met.id == tcc.parentmetricid:
                    parent_found = True
                if met.type in (type_quality_rules,type_quality_measuree,type_quality_distributions) and met.id == tcc.metricid:
                    child_found = True
                if parent_found and child_found:
                    break     
            if parent_found and child_found:
                str_df_mod_tc_cont += ';' + module_name + ';' + tcc.parentmetricname + ';' + tcc.parentmetricid + ';' + tcc.metricname + ';' + tcc.metricid
                str_df_mod_tc_cont += ';' + str(tcc.weight) + ';' + str(tcc.critical) + ';;'
                str_df_mod_tc_cont  += ';;;;;;;;;;;;;;'
                str_df_mod_tc_cont += '\n'
    #logger.debug(str_df)
    try: 
        str_df_mod_tc_cont = StringUtils.remove_unicode_characters(str_df_mod_tc_cont)
        df_mod_tc_cont = pd.read_csv(StringIO(str_df_mod_tc_cont), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_mod_tc_contributions %s ' % str_df_mod_tc_cont,True)
    return df_mod_tc_cont

###############################################################################

def get_df_mod_weight(logger, dict_modulesweight):
    df_mod_weight = None
    str_df_mod_weight = 'Module;Nb artifacts\n'
    for mod in dict_modulesweight:
        str_df_mod_weight += str(mod) + ';' + str(dict_modulesweight[mod]) + '\n'
    #logger.debug(str_df)
    try: 
        str_df_mod_weight = StringUtils.remove_unicode_characters(str_df_mod_weight)
        df_mod_weight = pd.read_csv(StringIO(str_df_mod_weight), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : str_df_mod_weight %s ' % str_df_mod_weight,True)
    return df_mod_weight

###############################################################################
def get_df_app_violations(logger, loadviolations, listviolations):
    df_app_violations = None
    # Data for the Violations Tab
    listmetricsinviolations = []
    if loadviolations:
        LogUtils.loginfo(logger,'Loading violations for excel reporting',True)
        str_df_violations = 'Quality rule name;Quality rule Id;Critical;Component name location;Selected for action;In action plan;Action plan status;Action plan tag;Action plan comment;Has exclusion request'
        str_df_violations += ';Violation status;Component status;URL;Nb actions;Nb actions added;Nb actions removed;JSON actions added;JSON actions modified;JSON actions removed;Violation id\n'
        for objviol in listviolations:
            str_df_violations += str(objviol.qrname) + ';' + str(objviol.qrid) + ';' + str(objviol.critical) + ';' + str(objviol.componentNameLocation) + ';'+ str(objviol.hasActionPlan) + ';' + str(objviol.hasActionPlan) + ';' + str(objviol.actionplanstatus) + ';' + str(objviol.actionplantag) + ';' + str(objviol.actionplancomment) 
            str_df_violations +=  ';'+ str(objviol.hasExclusionRequest) + ';'+ str(objviol.violationstatus) + ';'+ str(objviol.componentstatus)  + ';'+ str(objviol.url) + ';;;;;;;'+ str(objviol.id) 
            str_df_violations += '\n'
            listmetricsinviolations.append(str(objviol.qrid))
        try: 
            str_df_violations = StringUtils.remove_unicode_characters(str_df_violations)
            df_app_violations = pd.read_csv(StringIO(str_df_violations), sep=";",engine='python',quoting=csv.QUOTE_NONE) 
        except: 
            LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_app_violations %s ' % str_df_violations,True)
    return df_app_violations, listmetricsinviolations

###############################################################################
bc_contribution_commonheader = 'Business criterion name;Business criterion Id;Technical criterion name;Technical criterion Id;Weight;Critical;Simulation grade;Weighted grade\n'
app_bc_contribution_header = "Application name;" + bc_contribution_commonheader
mod_bc_contribution_header = "Module name;" + bc_contribution_commonheader
###############################################################################


def get_df_app_bc_contribution(logger, appName, listbccontributions, dictechnicalcriteria):
    df_app_bc_contribution = None
    str_df_bc_contribution = app_bc_contribution_header
    for bcc in listbccontributions:
        hasresults = False
        for tc in dictechnicalcriteria:
            if tc == bcc.metricid: hasresults = True 
        # keep only the technical criteria that have results
        if not hasresults:
            continue        
        str_df_bc_contribution += appName + ';' + bcc.parentmetricname + ';' + bcc.parentmetricid + ';' + bcc.metricname + ';' + bcc.metricid
        str_df_bc_contribution += ';' + str(bcc.weight) + ';' + str(bcc.critical) + ';;'
        str_df_bc_contribution += '\n'
    #logger.debug(str_df_bc_contribution)
    try: 
        str_df_bc_contribution = StringUtils.remove_unicode_characters(str_df_bc_contribution)
        df_app_bc_contribution = pd.read_csv(StringIO(str_df_bc_contribution), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_app_bc_contribution %s ' % str_df_bc_contribution,True)
    return df_app_bc_contribution

###############################################################################

def get_tc_contribution_header(level='Application'):
    header = ''
    if level == 'Application':
        header = 'Key;Application name'
    elif level == 'Module':
        header = 'Key;Module name'
    header += ';Technical criterion name;Technical criterion Id;Metric name;Metric Id;Weight;Critical;Grade simulation;Weighted grade;Grade improvement priority;Grade improvement opportunity;TC simulation grade;TC weight;TC Lowest critical grade;TC Weighted average of quality rules;Simulation grade for improvement;Weighted grade for improvement;Delta weighted grade for improvement;TC simulation grade from improvement;TC Lowest critical grade for improvement;Nb rules to fix before TC grade improvement;Nb violations  to fix on several rules before TC grade improvement;Nb violations;Remaining violations\n'
    return header

###############################################################################

def get_df_app_tc_contribution(logger, appName, listtccontributions, dictmetrics, aggregationmode):
    df_app_tc_contribution = None
    str_df_tc_contribution = get_tc_contribution_header()
    # for each contribution TC/QR 
    for tcc in listtccontributions:
        #print(str(tcc.metricid))
        QRhasresults = False
        for met in dictmetrics:
            if str(met) == str(tcc.metricid): 
                try:
                    if aggregationmode != 'FullApplication':
                        QRhasresults = True
                    elif aggregationmode == 'FullApplication' and dictmetrics[met].grade != None:
                        QRhasresults = True
                except KeyError:
                    None
                break
        # keep only the quality metrics that have metrics that have results 
        if QRhasresults:
            #print (tcc.metricid)
            str_df_tc_contribution += ';' + appName + ';' + tcc.parentmetricname + ';' + tcc.parentmetricid + ';' + tcc.metricname + ';' + tcc.metricid
            str_df_tc_contribution += ';' + str(tcc.weight) + ';' + str(tcc.critical) + ';;'
            str_df_tc_contribution  += ';;;;;;;;;;;;;;'
            str_df_tc_contribution += '\n'
    try: 
        str_df_tc_contribution = StringUtils.remove_unicode_characters(str_df_tc_contribution)
        df_app_tc_contribution = pd.read_csv(StringIO(str_df_tc_contribution), sep=";",quoting=csv.QUOTE_NONE)
    #,engine='python'
    except: 
        LogUtils.logerror(logger, 'csv.Error: unexpected end of data : df_app_tc_contribution %s ' % str_df_tc_contribution, True)
    return df_app_tc_contribution

###############################################################################

def get_df_remediationeffort(logger, dictmetrics, dicremediationabacus):

    df_remediationeffort = None
    str_df_remediationeffort = 'Quality rule id;Quality rule name;Remediation effort (minutes)\n'
    for qr in dictmetrics:
        oqr = dictmetrics[qr]
        # we are looking only at quality rules here, not distributions or measures 
        if oqr.type == type_quality_rules:
            str_df_remediationeffort += str(oqr.id) + ';' + str(oqr.name) 
            # if the quality rule is not 
            if dicremediationabacus.get(oqr.id) != None and dicremediationabacus.get(oqr.id).get('uniteffortinhours'):
                #print (str(oqr.id) + ' in the abacus')
                str_df_remediationeffort += ';' + str(dicremediationabacus.get(oqr.id).get('uniteffortinhours'))
            else:
                #print (str(oqr.id) + ' not in the abacus => N/A')
                str_df_remediationeffort += ';#N/A'
            str_df_remediationeffort += '\n'
    try: 
        str_df_remediationeffort = StringUtils.remove_unicode_characters(str_df_remediationeffort)
        df_remediationeffort = pd.read_csv(StringIO(str_df_remediationeffort), sep=";",engine='python',quoting=csv.QUOTE_NONE) 
    except: 
        LogUtils.logerror(logger, 'csv.Error: unexpected end of data : df_remediationeffort %s ' % str_df_remediationeffort,True)
    return df_remediationeffort


###############################################################################

def generate_excelfile(logger, filepath, appName, snapshotversion, snapshotdate, loadviolations, listbusinesscriteria, dictechnicalcriteria, listbccontributions, listtccontributions, dictmetrics, dictapsummary, dicremediationabacus, listviolations, broundgrades, dictaptriggers, dictmodules=None,dict_modulesweight=None, aggregationmode='FullApplication'):
    format = ExcelFormat()
    pd.options.display.float_format = format.const_float_format.format
    
    logger.info("Loading data in Excel")
    #Readme Page content
    df_readme = get_df_readme(logger, loadviolations)
    #######################################################################################################################
    # Data for the application BC Grades Tab
    df_app_bc_grades = get_df_app_bc_grades(logger, appName, snapshotversion, snapshotdate, listbusinesscriteria, loadviolations)
    # Data for the application TC Grades Tab
    df_app_tc_grades = get_df_app_tc_grades(logger, appName, dictechnicalcriteria)
    # Data for the application Rules Grades Tab
    df_app_rules_grades = get_df_app_rules_grades(logger, appName, snapshotdate, snapshotversion, dictmetrics, listtccontributions, listbccontributions, dictapsummary, dictaptriggers, aggregationmode)
    # List of application violations
    df_app_violations, listmetricsinviolations = get_df_app_violations(logger, loadviolations, listviolations)
    # Data for the application BC Contributions Tab
    df_app_bc_contribution = get_df_app_bc_contribution(logger, appName, listbccontributions, dictechnicalcriteria)
    # Data for the application TC Contributions Tab
    df_app_tc_contribution = get_df_app_tc_contribution(logger, appName, listtccontributions, dictmetrics,aggregationmode)
    #######################################################################################################################
    # Data for the module BC Grades Tab
    if dictmodules:
        df_mod_bc_grades = get_df_mod_bc_grades(logger, dictmodules, dict_modulesweight)
        # Data for the module TC Grades Tab
        df_mod_tc_grades = get_df_mod_tc_grades(logger, dictmodules, dict_modulesweight)
        # Data for the modules Rules Grades Tab
        df_mod_rules_grades = get_df_mod_rules_grades(logger, snapshotdate, snapshotversion, dictmodules, listtccontributions, listbccontributions, None, None, dict_modulesweight, aggregationmode)  
        # Data for the cBC Contributions Tab
        df_mod_bc_contribution = get_df_mod_bc_contribution(logger, dictmodules, listbccontributions)
        # Data for the mod TC Contributions Tab
        df_mod_tc_contribution = get_df_mod_tc_contribution(logger, dictmodules, listtccontributions)
        if dict_modulesweight != None:
            df_mod_weight = get_df_mod_weight(logger, dict_modulesweight)
    #######################################################################################################################
    # Data for the Remediation Tab
    df_remediationeffort = get_df_remediationeffort(logger, dictmetrics, dicremediationabacus)
        
    ###############################################################################
    logger.info("Writing data in Excel")
    #file = open(filepath, 'w')
    with pd.ExcelWriter(filepath,engine='xlsxwriter') as writer:
        df_readme.to_excel(writer, sheet_name=format.const_TAB_README, index=False)
        
        df_app_bc_grades.to_excel(writer, sheet_name=format.const_TAB_APP_BC_GRADES, index=False)
        df_app_tc_grades.to_excel(writer, sheet_name=format.const_TAB_APP_TC_GRADES, index=False)
        df_app_rules_grades.to_excel(writer, sheet_name=format.const_TAB_APP_RULES_GRADES, index=False)
        if loadviolations:
            df_app_violations.to_excel(writer, sheet_name=format.const_TAB_APP_VIOLATIONS, index=False)        
        df_app_bc_contribution.to_excel(writer, sheet_name=format.const_TAB_APP_BC_CONTRIBUTIONS, index=False) 
        df_app_tc_contribution.to_excel(writer, sheet_name=format.const_TAB_APP_TC_CONTRIBUTIONS, index=False)
        
        if dictmodules:
            df_mod_bc_grades.to_excel(writer, sheet_name=format.const_TAB_MOD_BC_GRADES, index=False)
            df_mod_tc_grades.to_excel(writer, sheet_name=format.const_TAB_MOD_TC_GRADES, index=False)
            df_mod_rules_grades.to_excel(writer, sheet_name=format.const_TAB_MOD_RULES_GRADES, index=False)
            df_mod_bc_contribution.to_excel(writer, sheet_name=format.const_TAB_MOD_BC_CONTRIBUTIONS, index=False) 
            df_mod_tc_contribution.to_excel(writer, sheet_name=format.const_TAB_MOD_TC_CONTRIBUTIONS, index=False)        
            if dict_modulesweight != None:
                df_mod_weight.to_excel(writer, sheet_name=format.const_TAB_MOD_WEIGHT, index=False)        
        df_remediationeffort.to_excel(writer, sheet_name=format.const_TAB_REMEDIATION_EFFORT, index=False) 

        workbook = writer.book
        
        #define the number format 
        format.format_percentage = workbook.add_format({'num_format': format.const_format_percentage})
        format.format_int_thousands = workbook.add_format({'num_format': format.const_format_int})
        format.format_float_with_2decimals = workbook.add_format({'num_format': format.const_float_with_2decimals})
        #define the colors
        format.format_green_percentage= workbook.add_format({'bg_color': format.const_color_green,'num_format': format.const_format_percentage})
        format.format_red_percentage = workbook.add_format({'bg_color': format.const_color_red,'num_format': format.const_format_percentage})
        format.format_grey_float_1decimal = workbook.add_format({'bg_color': format.const_color_light_grey, 'num_format': format.const_float_with_1decimal})
        format.format_green_int = workbook.add_format({'bg_color': format.const_color_green,'num_format': format.const_format_int})
        format.format_red_int = workbook.add_format({'bg_color': format.const_color_red,'num_format': format.const_format_int})

        format.format_align_left = workbook.add_format({'align': format.const_format_align_left})
    
        worksheet = writer.sheets[format.const_TAB_README]
        format_table_readme(workbook,worksheet,df_readme,format)      
    
        ################################################################################################################
        # Application
        
        worksheet = writer.sheets[format.const_TAB_APP_BC_GRADES]
        format_table_bc_grades(workbook,worksheet,df_app_bc_grades,format,loadviolations, 'Application', dict_modulesweight)   
    
        worksheet = writer.sheets[format.const_TAB_APP_TC_GRADES]
        format_table_tc_grades(workbook,worksheet,df_app_tc_grades,format, 'Application', dict_modulesweight)  
    
        worksheet = writer.sheets[format.const_TAB_APP_RULES_GRADES]
        format_table_rules_grades(workbook,worksheet,df_app_rules_grades,format,'Application',loadviolations, listmetricsinviolations, dictmodules != None, dict_modulesweight)  
    
        if loadviolations:
            worksheet = writer.sheets[format.const_TAB_APP_VIOLATIONS]
            format_table_violations(workbook,worksheet,df_app_violations,format)      
    
        worksheet = writer.sheets[format.const_TAB_APP_BC_CONTRIBUTIONS]
        format_table_bc_contribution(workbook,worksheet,df_app_bc_contribution,format, 'Application')     
        
        worksheet = writer.sheets[format.const_TAB_APP_TC_CONTRIBUTIONS]
        format_table_tc_contribution(workbook,worksheet,df_app_tc_contribution,format, 'Application')  
        
        ################################################################################################################
        # Modules
        if dictmodules:
            worksheet = writer.sheets[format.const_TAB_MOD_BC_GRADES]
            format_table_bc_grades(workbook,worksheet,df_mod_bc_grades,format,False,'Module', dict_modulesweight)
            
            worksheet = writer.sheets[format.const_TAB_MOD_TC_GRADES]
            format_table_tc_grades(workbook,worksheet,df_mod_tc_grades,format,'Module', dict_modulesweight)
            
            worksheet = writer.sheets[format.const_TAB_MOD_RULES_GRADES]
            format_table_rules_grades(workbook,worksheet,df_mod_rules_grades,format,'Module', None, None, None, dict_modulesweight)
            
            worksheet = writer.sheets[format.const_TAB_MOD_BC_CONTRIBUTIONS]
            format_table_bc_contribution(workbook,worksheet,df_mod_bc_contribution,format,'Module')
            
            worksheet = writer.sheets[format.const_TAB_MOD_TC_CONTRIBUTIONS]
            format_table_tc_contribution(workbook,worksheet,df_mod_tc_contribution,format,'Module')
            
            if dict_modulesweight != None:
                worksheet = writer.sheets[format.const_TAB_MOD_WEIGHT]
                format_table_mod_weight(workbook,worksheet,df_mod_weight,format)
            
        worksheet = writer.sheets[format.const_TAB_REMEDIATION_EFFORT]        
        format_table_remediation_effort(workbook,worksheet,df_remediationeffort,format)  
        
        worksheet = writer.sheets[format.const_TAB_APP_BC_GRADES]
        worksheet.activate()
        
        writer.save()
    
        LogUtils.loginfo(logger, 'File ' + filepath + ' generated', True)

def format_table_readme(workbook,worksheet,table,format):
    worksheet.set_tab_color(format.const_color_light_grey)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    
    
    worksheet.set_column('A:A', 25, None) # Page 
    worksheet.set_column('B:B', 60, None) # Content  
    worksheet.set_column('C:C', 110, None) # Comments  

    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)

########################################################################

def round_grades(broundgrades, formula):
    if broundgrades:
        # if we round we do it with 2 decimals
        return excel_round(formula, 2)
    else:
        return formula

########################################################################

def excel_round(formula,decimals):
    return 'ROUND(' + formula[1:] +  ',' + decimals + ')'

########################################################################

def format_table_bc_grades(workbook,worksheet,table,format,loadviolations, level='Application', dict_modulesweight=None):
    if level == 'Application':
        worksheet.set_tab_color(format.const_color_light_blue)
        tab_bc_cont = tab_app_bc_cont
        tab_rule_grade = tab_app_rule_grade
    elif level == 'Module':
        worksheet.set_tab_color(format.const_color_light_rose)
        tab_bc_cont = tab_mod_bc_cont
        tab_rule_grade = tab_mod_rule_grade

    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
   
    # the last 6 lines don't have this formula    
    offset = 1
    if not loadviolations:
        nb_rows = len(table.index.values)+1
        if level == 'Application':
            nb_rows = nb_rows - 9
    else: 
        nb_rows = len(table.index.values)+1 - 15
    
    col_to_format = 'H'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start, col_to_format, nb_rows)
    #print("range {}".format(range_to_format))
    
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '>','value': 0.001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '<','value': -0.001, 'format':   format.format_red_percentage})

    worksheet.set_column('A:A', 20, None) # Application column
    worksheet.set_column('B:B', 32, None) # BC name
    worksheet.set_column('C:C', 7.5, None) # Metric Id
    worksheet.set_column('D:D', 11, format.format_float_with_2decimals) # Grade 
    worksheet.set_column('E:E', 11, format.format_float_with_2decimals) # Simulated grade 
    # group and hide columns lowest critical grade and weighted average
    worksheet.set_column('F:F', 15, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) #  
    worksheet.set_column('G:G', 20, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) #  
    # group and hide columns lowest critical grade and weighted average
    #worksheet.set_column('F:F', None, None, {'level': 1, 'collapsed': True})
    #worksheet.set_column('G:G', None, None, {'level': 1, 'collapsed': True})    
    worksheet.set_column('H:H', 11, format.format_percentage) # delta %
 
    if level == 'Module' and dict_modulesweight != None:
        worksheet.set_column('I:I', 11, format.format_float_with_2decimals) # Module weight
        worksheet.set_column('J:J', 15, format.format_float_with_2decimals) # Module weighted grade 
        last_column = 'J'
    else:
        last_column = 'H'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))     
    
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        # simulation grade
        
        if level == 'Module' or (level == 'Application' and dict_modulesweight == None):
            # aggregation = full application
            worksheet.write_formula(row_num, 5-1, round_grades(broundgrades,'=IF(F%d=0,G%d,MIN(F%d,G%d))') % (row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        elif level == 'Application' and dict_modulesweight != None:
            # aggregation = module weighted average
            worksheet.write_formula(row_num, 5-1, round_grades(broundgrades,"=_xlfn.SUMIFS('%s'!J:J,'%s'!B:B,B%s)/_xlfn.SUMIFS('%s'!I:I,'%s'!B:B,B%s)") % (tab_mod_bc_grade,tab_mod_bc_grade,row_num + 1,tab_mod_bc_grade,tab_mod_bc_grade,row_num + 1))

        # lowest critical
        worksheet.write_formula(row_num, 6-1,  round_grades(broundgrades,"=_xlfn.MINIFS('%s'!H:H,'%s'!C:C,C%d,'%s'!G:G,TRUE,'%s'!A:A,A%s)") % (tab_bc_cont, tab_bc_cont, row_num + 1, tab_bc_cont, tab_bc_cont, row_num + 1))
        # weighted average
        #worksheet.write_formula(row_num, 7-1, round_grades(broundgrades,"=SUMIF('%s'!C:C,C%d,'%s'!I:I)/SUMIF('%s'!C:C,C%d,'%s'!F:F)") % (tab_bc_cont, row_num + 1, tab_bc_cont, tab_bc_cont, row_num + 1, tab_bc_cont))
        
        formula = "=SUMIFS('%s'!I:I,'%s'!C:C,C%s,'%s'!A:A,A%s)/SUMIFS('%s'!F:F,'%s'!C:C,C%s,'%s'!A:A,A%s)" % (tab_bc_cont, tab_bc_cont,  row_num + 1, tab_bc_cont,  row_num + 1, tab_bc_cont, tab_bc_cont,  row_num + 1, tab_bc_cont,  row_num + 1)
        worksheet.write_formula(row_num, 7-1, formula)
        #=SUMIFS('Modules BC contributions'!I:I;'Modules BC contributions'!C:C;C2;'Modules BC contributions'!A:A;A2)/SUMIFS('Modules BC contributions'!F:F;'Modules BC contributions'!C:C;C2;'Modules BC contributions'!A:A;A2)

        #=SUMIFS('Modules BC contributions'!I:I;'Modules BC contributions'!C:C;C2;'Modules BC contributions'!A:A;A2)/SUMIFS('Modules BC contributions'!F:F;'Modules BC contributions'!C:C;C2;'Modules BC contributions'!A:A;A2)
        # Delta %
        worksheet.write_formula(row_num, 8-1, '=($E%d-$D%d)/$D%d' % (row_num + 1, row_num + 1, row_num + 1), format.format_percentage)
    
        if level == 'Module' and dict_modulesweight != None:
            worksheet.write_formula(row_num, 9-1, "=VLOOKUP(A%s,'%s'!A:B,2,FALSE)" % (row_num + 1, tab_mod_weigth), format.format_float_with_2decimals)
            worksheet.write_formula(row_num, 10-1, "=I%s*E%s" % (row_num + 1, row_num + 1), format.format_float_with_2decimals)    
    
    
    if level == 'Application':
        # 3 empty line + 3 lines for application name, snapshot version and date
        row_to_format_for_summary = nb_rows + 6        #number of violations
        worksheet.write_formula(row_to_format_for_summary, 3-1, "=SUM('%s'!N:N)" % tab_rule_grade)
        #number of quality rules for action
        worksheet.write_formula(row_to_format_for_summary+1, 3-1, "=COUNTIF('%s'!N:N,\">0\")" % tab_rule_grade)
        #estimated effort m.d
        worksheet.write_formula(row_to_format_for_summary+2, 3-1, "=SUM('%s'!R:R)" % tab_rule_grade)
        #if loadviolations:
            #Number of action plans added
            #="[" & CONCATENATE(TRANSPOSE(A1:A5)&",") &"]"
            #worksheet.write_formula(row_to_format_for_summary+4, 3-1, "=CONCATENATE(TRANSPOSE(A1:A5)&"" "")")
            #Number of action plans removed
            #worksheet.write_formula(row_to_format_for_summary+5, 3-1, "=SUM(Violations!P:P)")        
            #JSON added
            #worksheet.write_formula(row_to_format_for_summary+6, 3-1, "=SUM(Violations!P:P)")        
            #JSON removed
            #JSON modified

    
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)


########################################################################
  
def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string  
   
########################################################################
  
# extract the last element that is the id
def get_hrefid(href, separator='/'):
    if href == None or separator == None:
        return None
    _id = ""
    hrefsplit = href.split('/')
    for elem in hrefsplit:
        # the last element is the id
        _id = elem    
    return _id  
    
########################################################################    
    
def format_table_tc_grades(workbook,worksheet,table,format, level='Application', dict_modulesweight=None):
    if level == 'Application':
        worksheet.set_tab_color(format.const_color_light_blue)
        tab_tc_cont = tab_app_tc_cont
    elif level == 'Module':
        worksheet.set_tab_color(format.const_color_light_rose)
        tab_tc_cont = tab_mod_tc_cont

    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)

    nb_rows = len(table.index.values)+1

    #define the range to be formated in excel format
    col_to_format = 'I'    
    start = col_to_format + '2'
    range_to_format = "{}:{}{}".format(start, col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '>','value': 0.001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '<','value': -0.001, 'format':   format.format_red_percentage})
    
    worksheet.set_column('A:A', 10, None, {'level': 1, 'hidden': True}) #  Key        
    worksheet.set_column('B:B', 20, None) #  App name
    worksheet.set_column('C:C', 60, None) #  TC name
    worksheet.set_column('D:D', 8, format.format_align_left) # Id
    worksheet.set_column('E:E', 8, format.format_float_with_2decimals) # Grade
    worksheet.set_column('F:F', 10, format.format_float_with_2decimals) # Simulation grade
    # group and hide columns lowest critical grade and weighted average
    worksheet.set_column('G:G', 13, format.format_float_with_2decimals, {'level': 2, 'hidden': True}) # 
    worksheet.set_column('H:H', 19, format.format_float_with_2decimals, {'level': 2, 'hidden': True}) # 
    worksheet.set_column('I:I', 12, format.format_percentage) # 
 
 
    if level == 'Module' and dict_modulesweight != None:
        worksheet.set_column('J:J', 11, format.format_float_with_2decimals) # Module weight
        worksheet.set_column('K:K', 15, format.format_float_with_2decimals) # Module weighted grade 
        last_column = 'K'
    else:
        last_column = 'I'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))   
 
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        worksheet.write_formula(row_num, 1-1, "B%s&D%s" % (row_num + 1, row_num + 1))
        
        #simulation grade
        #worksheet.write_formula(row_num, 6-1, round_grades(broundgrades,"=IF(G%d=0,H%d,MIN(G%d,H%d))") % (row_num + 1, row_num + 1, row_num + 1, row_num + 1), format.format_float_with_2decimals)
        if level == 'Module' or (level == 'Application' and dict_modulesweight == None):
            # aggregation = full application
            worksheet.write_formula(row_num, 6-1, round_grades(broundgrades,"=IF(G%d=0,H%d,MIN(G%d,H%d))") % (row_num + 1, row_num + 1, row_num + 1, row_num + 1), format.format_float_with_2decimals)
        elif level == 'Application' and dict_modulesweight != None:
            # aggregation = module weighted average
            worksheet.write_formula(row_num, 6-1, round_grades(broundgrades,"=_xlfn.SUMIFS('%s'!K:K,'%s'!C:C,C%s)/_xlfn.SUMIFS('%s'!J:J,'%s'!C:C,C%s)") % (tab_mod_tc_grade,tab_mod_tc_grade,row_num + 1,tab_mod_tc_grade,tab_mod_tc_grade,row_num + 1))
        
        #lowest critical rule grade
        worksheet.write_formula(row_num, 7-1, round_grades(broundgrades,"=_xlfn.MINIFS('%s'!I:I,'%s'!D:D,D%d,'%s'!H:H,TRUE,'%s'!B:B,B%s)") % (tab_tc_cont, tab_tc_cont, row_num + 1, tab_tc_cont, tab_tc_cont, row_num + 1), format.format_float_with_2decimals)
        #weighted av
        #worksheet.write_formula(row_num, 7-1, round_grades(broundgrades,"=SUMIF('%s'!C:C, C%d,'%s'!I:I)/SUMIF('%s'!C:C, C%d,'%s'!F:F)") % (tab_tc_cont, row_num + 1, tab_tc_cont, tab_tc_cont, row_num + 1, tab_tc_cont), format.format_float_with_2decimals)
        formula="=SUMIFS('%s'!J:J,'%s'!D:D, D%s,'%s'!B:B,B%s)/SUMIFS('%s'!G:G,'%s'!D:D, D%s,'%s'!B:B,B%s)" % (tab_tc_cont, tab_tc_cont,row_num + 1, tab_tc_cont,row_num + 1,tab_tc_cont, tab_tc_cont,row_num + 1, tab_tc_cont,row_num + 1)
        
        worksheet.write_formula(row_num, 8-1, round_grades(broundgrades,formula))
        #delta %
        worksheet.write_formula(row_num, 9-1, "=($F%d-$E%d)/$E%d" % (row_num + 1, row_num + 1, row_num + 1), format.format_percentage)
        
        if level == 'Module' and dict_modulesweight != None:
            worksheet.write_formula(row_num, 10-1, "=VLOOKUP(B%s,'%s'!A:B,2,FALSE)" % (row_num + 1, tab_mod_weigth), format.format_float_with_2decimals)
            worksheet.write_formula(row_num, 11-1, "=J%s*F%s" % (row_num + 1, row_num + 1), format.format_float_with_2decimals)            
        
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format) 
 

 
########################################################################

def format_table_rules_grades(workbook,worksheet, table, format, level='Application', loadviolations=False, listmetricsinviolations = None, loadmodules=False, dict_modulesweight=None):
    if level == 'Application':
        worksheet.set_tab_color(format.const_color_light_blue)
        tab_tc_cont = tab_app_tc_cont
        tab_viol = tab_app_viol
    elif level == 'Module':
        worksheet.set_tab_color(format.const_color_light_rose)
        tab_tc_cont = tab_mod_tc_cont
        #tab_viol = tab_app_viol
        
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1})
    worksheet.set_zoom(85)
    worksheet.freeze_panes(1, 0)  # Freeze the first row.    
    nb_rows = len(table.index.values)+1
    
    # conditional formating for the Grade delta column (red and green)
    col_to_format = 'L'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '>', 'value': 0.001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '<', 'value': -0.001, 'format':   format.format_red_percentage})   

    # conditional formating for the number of violations for action
    col_to_format = 'N'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '>', 'value': 0.001, 'format':   format.format_green_int})
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '<', 'value': -0.001, 'format':   format.format_red_int})   

    # conditional formating for the unit effort column
    col_to_format = 'P'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})
    # conditional formating for the total effort column in hours
    col_to_format = 'Q'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})
    # conditional formating for the total effort column in days
    col_to_format = 'R'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})

    # conditional formating for the Grade improvement priority
    col_to_format = 'AB'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':'3_color_scale', 'min_value': 1, 'max_value': 4})

    col_group_app_or_mod = None
    if level == 'Application':
        col_group_app_or_mod = {'level': 2, 'hidden': True}
    
    worksheet.set_column('A:A', 10, None,                     {'level': 1, 'hidden': True}) #  Key
    worksheet.set_column('B:B', 25, None,                     col_group_app_or_mod) #  Application / module name
    worksheet.set_column('C:C', 12, format.format_align_left, {'level': 2, 'hidden': True}) # Application column 
    worksheet.set_column('D:D', 10, format.format_align_left, {'level': 2, 'hidden': True}) # Snapshot date
    worksheet.set_column('E:E', 60, None) # Metric name 
    worksheet.set_column('F:F', 8, None) # metric id
    worksheet.set_column('G:G', 18, None) #  
    worksheet.set_column('H:H', 6.5, None) #  
    worksheet.set_column('I:I', 9, format.format_float_with_2decimals) #   
    worksheet.set_column('J:J', 8, format.format_float_with_2decimals) #    
    worksheet.set_column('K:K', 8, format.format_float_with_2decimals)
    worksheet.set_column('L:L', 8, format.format_percentage) # % Grade delta
    
    worksheet.set_column('M:M', 10, None) # Nb violations
    worksheet.set_column('N:N', 11, None) # Nb violations for action
    worksheet.set_column('O:O', 10, None) # Remaining vioaltions
    worksheet.set_column('P:P', 11, format.format_float_with_2decimals, {'level': 3, 'hidden': False}) # Unit effort
    worksheet.set_column('Q:Q', 11, format.format_float_with_2decimals, {'level': 3, 'hidden': False}) # Total effort mh
    worksheet.set_column('R:R', 11, format.format_float_with_2decimals, {'level': 3, 'hidden': False}) # Total effort md
    worksheet.set_column('S:S', 11, format.format_int_thousands) # total checks
    worksheet.set_column('T:T', 11, format.format_percentage) # % compliance ratio
    worksheet.set_column('U:U', 11, format.format_percentage) # % new compliance ratio
    worksheet.set_column('V:V', 6.5, None, {'level': 4, 'hidden': True}) # Thres 1   
    worksheet.set_column('W:W', 6.5, None, {'level': 4, 'hidden': True}) #
    worksheet.set_column('X:X', 6.5, None, {'level': 4, 'hidden': True}) #
    worksheet.set_column('Y:Y', 6.5, None, {'level': 4, 'hidden': True}) # Thres 4
    
    worksheet.set_column('Z:Z', 11, None, {'level': 4, 'hidden': True}) # Educate   
    worksheet.set_column('AA:AA', 11, None, {'level': 4, 'hidden': True}) # violations extracted ?
    worksheet.set_column('AB:AB', 13, None) # Grade improvement priority
    worksheet.set_column('AC:AC', 13, None) # Nb violations to fix on this rule before rule grade improvement
    worksheet.set_column('AD:AD', 13, None) # Nb rules to fix before TC grade improvement
    
    worksheet.set_column('AE:AE', 13, None) # Nb violations  to fix on the TC  rules before TC grade improvement
    worksheet.set_column('AF:AF', 50, None) # Technical criterion
    worksheet.set_column('AG:AG', 60, None) # Business criteria
    
    if level == 'Module' and dict_modulesweight != None:
        worksheet.set_column('AH:AH', 11, format.format_float_with_2decimals) # Module weight
        worksheet.set_column('AI:AI', 15, format.format_float_with_2decimals) # Module weighted grade 
        last_column = 'AI'
    else:
        last_column='AG'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))

    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        metrictype = str(table.loc[row_num-1, 'Metric Type'])
        metricid = str(table.loc[row_num-1, 'Metric Id'])
        
        worksheet.write_formula(row_num, 1-1, '=$B%d&$F%d' % (row_num + 1, row_num + 1))
        
        # formulas applicable only for quality-rules, not for quality-measures and quality-distributions 
        if metrictype == type_quality_rules:
            #simulation grade
            #formula = round_grades(broundgrades,'=IF(U%s=0,I%s,IF(U%s<=V%s/100,1,IF(U%s<W%s/100,(U%s*100-V%s)/(W%s-V%s)+1,IF(U%s<X%s/100,(U%s*100-W%s)/(X%s-W%s)+2,IF(U%s<Y%s/100,(U%s*100-X%s)/(Y%s-X%s)+3,4)))))') % (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
            if level == 'Module' or (level == 'Application' and dict_modulesweight == None):
                # aggregation = full application
                formula = round_grades(broundgrades,'=IF(U%s=0,I%s,IF(U%s<V%s/100,1,IF(U%s<W%s/100,(U%s*100-V%s)/(W%s-V%s)+1,IF(U%s<X%s/100,(U%s*100-W%s)/(X%s-W%s)+2,IF(U%s<Y%s/100,(U%s*100-X%s)/(Y%s-X%s)+3,4)))))') % (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
            elif level == 'Application' and dict_modulesweight != None:
                # aggregation = module weighted average
                formula = round_grades(broundgrades,"=_xlfn.SUMIFS('%s'!AI:AI,'%s'!E:E,E%s)/_xlfn.SUMIFS('%s'!AH:AH,'%s'!E:E,E%s)") % (tab_mod_rule_grade,tab_mod_rule_grade,row_num + 1,tab_mod_rule_grade,tab_mod_rule_grade,row_num + 1)
            worksheet.write_formula(row_num, 10-1, formula)
            
            #grade delta
            worksheet.write_formula(row_num, 11-1, '=$J%d-$I%d' % (row_num + 1, row_num + 1))
            #grade delta %
            worksheet.write_formula(row_num, 12-1, '=$K%d/$I%d' % (row_num + 1, row_num + 1))
            
            #Nb violations for action, with a formula only if the violations are loaded
            # also only if we find the metric id in the list of violations
            if loadviolations:
                if listmetricsinviolations != None and len(listmetricsinviolations) > 0 and metricid in listmetricsinviolations:
                    formula = "=SUMIF(%s!B:B, E%d,%s!N:N)"% (tab_viol, row_num + 1, tab_viol)
                    #print(formula)
                    worksheet.write_formula(row_num, 14-1, formula)
            if loadmodules:
                # overwrites the value coming for the action plan 
                formula = "=SUMIFS('%s'!N:N,'%s'!F:F,F%s)" % (tab_mod_rule_grade, tab_mod_rule_grade, row_num + 1)
                worksheet.write_formula(row_num, 14-1, formula)
            #remaining violations
            worksheet.write_formula(row_num, 15-1, '=$M%d-$N%d' % (row_num + 1, row_num + 1))
            #unit effort
            worksheet.write_formula(row_num, 16-1, "=(VLOOKUP(F%d,'%s'!A:C,3,FALSE))/60" % (row_num + 1, tab_remd))        
            #total effort (mh)
            worksheet.write_formula(row_num, 17-1, "=P%d*N%d" % (row_num + 1, row_num + 1))
            #total effort (md)
            worksheet.write_formula(row_num, 18-1, "=Q%d/8" % (row_num + 1))
            #new compliance ratio
            worksheet.write_formula(row_num, 21-1, '=($S%d-$O%d)/$S%d' % (row_num + 1, row_num + 1, row_num + 1))
            #Violations extracted ? Present in violations tab
            if loadviolations:
                #formula = '=IF(NOT(ISNA(VLOOKUP($E%d,Violations!B:B,1,FALSE))),TRUE,FALSE)' % (row_num + 1)
                formula = '=IF(NOT(ISNA(VLOOKUP($F%d,%s!B:B,1,FALSE))),TRUE,FALSE)' % (row_num + 1, tab_viol)
                #print(formula)
                worksheet.write_formula(row_num, 27-1, formula)
            else:
                worksheet.write_formula(row_num, 27-1, "=FALSE")
            #Grade improvement opportunity
            #worksheet.write_formula(row_num, 28-1, "=VLOOKUP(F%s,'%s'!F:P,6,FALSE)" % (row_num + 1, tab_tc_cont))
            worksheet.write_formula(row_num, 28-1, "=VLOOKUP(A%s,'%s'!A:P,11,FALSE)" % (row_num + 1, tab_tc_cont))
            
            #Nb violations to fix before grade improvement
            worksheet.write_formula(row_num, 29-1, '=IF(J%s<4,IF(U%s<(V%s/100),O%s-(S%s-(ROUNDUP((V%s/100)*S%s,0))),1),"")' % (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1))
            #Nb rules to fix before TC grade improvement
            worksheet.write_formula(row_num, 30-1, "=VLOOKUP(A%s,'%s'!A:Y,22,FALSE)" % (row_num + 1, tab_tc_cont))
            #Nb violations  to fix on the TC  rules before TC grade improvement
            worksheet.write_formula(row_num, 31-1, "=VLOOKUP(A%s,'%s'!A:Y,23,FALSE)" % (row_num + 1, tab_tc_cont))
            
            if level == 'Module' and dict_modulesweight != None:
                worksheet.write_formula(row_num, 34-1, "=VLOOKUP(B%s,'%s'!A:B,2,FALSE)" % (row_num + 1, tab_mod_weigth), format.format_float_with_2decimals)
                worksheet.write_formula(row_num, 35-1, "=AH%s*J%s" % (row_num + 1, row_num + 1), format.format_float_with_2decimals)                 
            
        else:
            # simulation grade = grade
            worksheet.write_formula(row_num, 10-1, '=$I%d' % (row_num + 1))
            # grade delta
            worksheet.write_formula(row_num, 11-1, '=0')
            # grade delta % 
            worksheet.write_formula(row_num, 12-1, '=0')

        # Write the column headers with the defined format.
        for col_num, value in enumerate(table.columns.values):
            worksheet.write(0, col_num, value, header_format)


########################################################################


def format_table_violations(workbook,worksheet,table,format, level='Application'):
    worksheet.set_tab_color(format.const_color_light_blue)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1})
    worksheet.set_zoom(78)
    worksheet.freeze_panes(1, 0)  # Freeze the first row. 
    nb_rows = len(table.index.values)+1
    
    worksheet.set_column('A:A', 80, format.format_align_left) #  QR Name
    worksheet.set_column('B:B', 9, None) #  QR Id 
    worksheet.set_column('C:C', 10, format.format_align_left) # Critical
    worksheet.set_column('D:D', 65, None) # Fullname
    worksheet.set_column('E:E', 10, None) # Action plan
    worksheet.set_column('F:F', 10, None) # For action
    worksheet.set_column('G:G', 10, None) # AP status
    worksheet.set_column('H:H', 11, None) # AP tag 
    
    worksheet.set_column('I:I', 11, None)    
    worksheet.set_column('J:J', 11, None) # Excl request
    worksheet.set_column('K:K', 11, None) # Comp status     
    worksheet.set_column('L:L', 11, None) # Viol status
    worksheet.set_column('M:M', 11, None) # URL

    # group and hide the context
    #worksheet.set_column('N:T', None, None, {'level': 1, 'collapsed': True})
    worksheet.set_column('N:N', 8, None) # Nb actions 
    worksheet.set_column('O:O', 8, None) # Nb actions added
    worksheet.set_column('P:P', 8, None) # Nb actions removed
    worksheet.set_column('Q:Q', 8, None) # JSON actions added
    worksheet.set_column('R:R', 8, None) # JSON actions modified
    worksheet.set_column('S:S', 8, None) # JSON actions removed    
    worksheet.set_column('T:T', 60, None) # Violation id
    last_column = 'T'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))
    
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)   
   
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):    
        None
        #Nb actions 
        worksheet.write_formula(row_num, 14-1, '=IF(E%d,1,"")' % (row_num + 1))
        #Nb actions added 
        worksheet.write_formula(row_num, 15-1, '=IF(AND(E%d,NOT(F%d)),1,"")' % (row_num + 1,row_num + 1))        
        #for Nb action removed
        worksheet.write_formula(row_num, 16-1, '=IF(AND(NOT(E%d),F%d),1,"")' % (row_num + 1,row_num + 1))
        #        
        #JSON actions added
        formula_added = '=IF(O%d=1,"{""component"": {""href"":"""&MID($T%d,SEARCH("#",$T%d)+1,LEN($T%d))&""" },""rulePattern"": { ""href"":"""&MID($T%d,1,SEARCH("#",$T%d)-1)&""" },""remedialAction"": {""comment"": """&IF($I%d<>"",$I%d,"For action")&""", ""tag"": """&IF($H%d<>"",$H%d,"Moderate")&""" }}","")'% (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
        worksheet.write_formula(row_num, 17-1, formula_added)
        #{"component": {"href":"DOMAIN04/components/227700/snapshots/18" },"rulePattern": { "href":"DOMAIN04/rule-patterns/1600110" },"remedialAction": {"comment": "For action", "tag": "extreme" }}
        
        #JSON actions modified
        formula_modified = '=IF(AND(OR(O%d=0,O%d=""),OR(P%d=0,P%d="")),"{""component"": {""href"":"""&MID($T%d,SEARCH("#",$T%d)+1,LEN($T%d))&""" },""rulePattern"": { ""href"":"""&MID($T%d,1,SEARCH("#",$T%d)-1)&""" },""remedialAction"": {""comment"": """&IF($I%d<>"",$I%d,"For action")&""", ""tag"": """&IF($H%d<>"",$H%d,"Moderate")&""" }}","")'% (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
        worksheet.write_formula(row_num, 18-1, formula_modified)        
        
        #JSON actions removed
        formula_removed = '=IF(P%d=1,"{""component"": {""href"":"""&MID($T%d,SEARCH("#",$T%d)+1,LEN($T%d))&""" },""rulePattern"": { ""href"":"""&MID($T%d,1,SEARCH("#",$T%d)-1)&""" }}","")'% (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
        worksheet.write_formula(row_num, 19-1 , formula_removed)        
        
        # Data validation        
        worksheet.data_validation('E' + str(nb_rows+1), {'validate': 'list', 'source': ['TRUE', 'FALSE']})   

     
########################################################################

def format_table_bc_contribution(workbook,worksheet,table, format, level='Application'):
    if level == 'Application':
        worksheet.set_tab_color(format.const_color_light_blue)
        tab_tc_grade = tab_app_tc_grade
    elif level == 'Module':
        worksheet.set_tab_color(format.const_color_light_rose)
        tab_tc_grade = tab_mod_tc_grade
    
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1})
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
    nb_rows = len(table.index.values)+1
    
    worksheet.set_column('A:A', 25, None) # App name
    worksheet.set_column('B:B', 25, None) #  
    worksheet.set_column('C:C', 13, format.format_align_left) # Application column 
    worksheet.set_column('D:D', 60, format.format_align_left) # BC column 
    worksheet.set_column('E:E', 13, None) # Metric Id column 
    worksheet.set_column('F:E', 9, None) # HF column 
    worksheet.set_column('G:G', 9, None) # HF column 
    worksheet.set_column('H:H', 13, format.format_float_with_2decimals) # HF column 
    worksheet.set_column('I:I', 13, format.format_float_with_2decimals) # HF column  
    last_column = 'I'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))


    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        worksheet.write_formula(row_num, 8 - 1, "=VLOOKUP(A%s&E%d,'%s'!A:F,6,FALSE)" % (row_num + 1, row_num + 1, tab_tc_grade))
        worksheet.write_formula(row_num, 9 - 1, '=$H%d*$F%d' % (row_num + 1, row_num + 1))

    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)

########################################################################

def format_table_tc_contribution(workbook,worksheet,table,format, level='Application'):
    if level == 'Application':
        worksheet.set_tab_color(format.const_color_light_blue)
        tab_rule_grade = tab_app_rule_grade
        tab_tc_grade = tab_app_tc_grade 

    elif level == 'Module':
        worksheet.set_tab_color(format.const_color_light_rose)
        tab_rule_grade = tab_mod_rule_grade
        tab_tc_grade = tab_mod_tc_grade 
    
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
    
    nb_rows = len(table.index.values)+1
    
    # conditional formating for the Grade improvement priority
    col_to_format = 'K'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':'3_color_scale', 'min_value': 1, 'max_value': 4})
    
    worksheet.set_column('A:A', 10, None, {'level': 1, 'hidden': True}) # Key
    
    worksheet.set_column('B:B', 25, None) # App name
    
    worksheet.set_column('C:C', 60, None) #  criterion name
    worksheet.set_column('D:D', 13, format.format_align_left)  
    worksheet.set_column('E:E', 70, format.format_align_left)  
    worksheet.set_column('F:F', 8, None) # 
    worksheet.set_column('G:G', 9, None) #  
    worksheet.set_column('H:H', 9, None) #  
    worksheet.set_column('I:I', 13, format.format_float_with_2decimals) #  grade simulation
    worksheet.set_column('J:J', 13, format.format_float_with_2decimals) #  weighted grade
    
    worksheet.set_column('K:K', 13, None) #  Grade imp priority
    worksheet.set_column('L:L', 13, None) #  Grade imp opp
    worksheet.set_column('M:M', 11, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #  TC simu grad
    worksheet.set_column('N:N', 9,  None                                    , {'level': 2, 'hidden': True}) #  TC weight
    worksheet.set_column('O:O', 11, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #  TC lower critic grad
    worksheet.set_column('P:P', 11, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #  TC weighted avg qr
    worksheet.set_column('Q:Q', 12, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #   
    worksheet.set_column('R:R', 12, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #   
    worksheet.set_column('S:S', 12, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #   
    worksheet.set_column('T:T', 12, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #  
    worksheet.set_column('U:U', 12, format.format_float_with_2decimals      , {'level': 2, 'hidden': True}) #  
    worksheet.set_column('V:V', 12, None                                    , {'level': 2, 'hidden': True}) #
    worksheet.set_column('W:W', 15, None                                    , {'level': 2, 'hidden': True}) #
    worksheet.set_column('X:X', 15, None                                    , {'level': 2, 'hidden': True}) #
    worksheet.set_column('Y:Y', 15, None                                    , {'level': 2, 'hidden': True}) #
    
    last_column = 'Y'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))    
   
    #formula = "=INDEX('%s'!$I$2:$I$%s,MATCH(1,(E%s='%s'!$E$2:$E$%s) * (A%s='%s'!$A$2:$A$%s),0))" % (tab_rule_grade, str(nb_rows), row_num + 1, tab_rule_grade, str(nb_rows), row_num + 1, tab_rule_grade,  str(nb_rows))
    #print(formula)
    #worksheet.write_array_formula(1,9 - 1,nb_rows,9 - 1, formula)
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):

        worksheet.write_formula(row_num, 1 - 1, '=$B%d&$F%d' % (row_num + 1, row_num + 1))
        
        # simulation grade
        #=worksheet.write_formula(row_num, 8 - 1, "=VLOOKUP(E%d,'%s'!E:I,5,FALSE)" % (row_num + 1, tab_rule_grade))
        #formula = "=INDEX('%s'!$I$2:$I$%s,MATCH(1,(E%s='%s'!$E$2:$E$%s) * (A%s='%s'!$A$2:$A$%s),0))" % (tab_rule_grade, str(nb_rows), row_num + 1, tab_rule_grade, str(nb_rows), row_num + 1, tab_rule_grade,  str(nb_rows))
        formula = "=VLOOKUP(A%s,'%s'!A:J,10,FALSE)" %  (row_num + 1, tab_rule_grade)
        worksheet.write_formula(row_num, 9 - 1, formula)
        worksheet.write_formula(row_num, 10 - 1, '=$I%d*$G%d' % (row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 11 - 1, '=IF(L%d,IF(H%d,1,2),IF(AND(H%d,I%d<4),3,IF(I%d<4,4,"")))' % (row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        
        # Grade improvement opportunity
        #worksheet.write_formula(row_num, 12 - 1,  "=IF(AND(H%s,O%s<P%s),IF(I%s=O%s,TRUE,FALSE),IF(AND(O%s>0,O%s<P%s),FALSE,IF(S%s=_xlfn.MAXIFS(S:S,D:D,D%s),TRUE,FALSE)))" % (row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 12 - 1,  "=IF(Y%s>0,IF(AND(H%s,O%s<P%s),IF(I%s=O%s,TRUE,FALSE),IF(AND(O%s>0,O%s<P%s),FALSE,IF(S%s=_xlfn.MAXIFS(S:S,D:D,D%s),TRUE,FALSE))),FALSE)" % (row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        
        #worksheet.write_formula(row_num, 12 - 1, "=VLOOKUP(B%s,'%s'!A:H,5,FALSE)"% (row_num + 1, tab_tc_grade))
        #worksheet.write_formula(row_num, 12 - 1, "=INDIRECT(ADDRESS(MATCH(A%s&C%s,'%s'!A:A&'%s'!C:C,0),5,1,1,\"%s\"))" % (row_num + 1,row_num + 1, tab_tc_grade, tab_tc_grade, tab_tc_grade))
        #worksheet.write_formula(row_num, 12 - 1, "=INDIRECT(ADDRESS(2,1))")
        
        worksheet.write_formula(row_num, 14 - 1, "=SUMIF(C:C,C%s,G:G)"% (row_num + 1))
        worksheet.write_formula(row_num, 15 - 1, "=_xlfn.MINIFS(I:I,D:D,D%s,H:H,TRUE)"% (row_num + 1))
        worksheet.write_formula(row_num, 16 - 1, "=SUMIF(D:D,D%s,J:J)/SUMIF(D:D,D%s,G:G)" % (row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 17 - 1, "=IF(IF(I%s*1.1>4,I%s*1.05,I%s*1.1)>4,4,IF(I%s*1.1>4,I%s*1.05,I%s*1.1))"% (row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 18 - 1, "=Q%s*G%s" % (row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 19 - 1, "=R%s-J%s" % (row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 20 - 1, "=SUMIF(D:D,D%s,J:J)/SUMIF(D:D,D%s,G:G)"% (row_num + 1, row_num + 1))
        worksheet.write_formula(row_num, 21 - 1, "=_xlfn.MINIFS(Q:Q,D:D,D%s,H:H,TRUE)"% (row_num + 1))
        worksheet.write_formula(row_num, 22 - 1, "=COUNTIFS(D:D,D%s,L:L,TRUE)"% (row_num + 1))
        worksheet.write_formula(row_num, 23 - 1, "=SUMIFS(Y:Y,D:D,D%s,L:L,TRUE)" % (row_num + 1))
        worksheet.write_formula(row_num, 24 - 1, "=VLOOKUP(A%s,'%s'!A:O,13,FALSE)" % (row_num + 1, tab_rule_grade))
        worksheet.write_formula(row_num, 25 - 1, "=VLOOKUP(A%s,'%s'!A:Q,15,FALSE)" % (row_num + 1, tab_rule_grade))
        
    
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)

########################################################################

def format_table_mod_weight(workbook,worksheet,table,format):
    worksheet.set_tab_color(format.const_color_light_rose)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
    
    nb_rows = len(table.index.values)+1
    worksheet.set_column('A:A', 50, None) # Module name
    worksheet.set_column('B:B', 25, None) # Nb of artifacts 
   
    last_column = 'B'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))    
   

   
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)


########################################################################

def format_table_remediation_effort(workbook,worksheet,table,format):
    worksheet.set_tab_color(format.const_color_light_grey)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.set_zoom(85)
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    nb_rows = len(table.index.values)+1
    
    worksheet.set_column('A:A', 15, None) # QR Id
    worksheet.set_column('B:B', 100, None) # QR Name 
    worksheet.set_column('C:C', 50, None) # Remediation effort 
    last_column = 'C'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))    

    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)