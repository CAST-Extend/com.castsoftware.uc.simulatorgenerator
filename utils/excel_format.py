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
    # header color= green
    const_color_header_columns = '#D7E4BC'
    
    # Tab names
    const_TAB_README = 'README'
    const_TAB_BC_GRADES = 'BC Grades'
    const_TAB_TC_GRADES = 'TC Grades'
    const_TAB_RULES_GRADES = 'Rules Grades'
    const_TAB_VIOLATIONS = 'Violations'
    const_TAB_BC_CONTRIBUTIONS = 'BC contributions'    
    const_TAB_TC_CONTRIBUTIONS = 'TC contributions'
    const_TAB_REMEDIATION_EFFORT = 'Remediation effort'
    
    format_percentage = None
    format_int_thousands = None
    format_align_left = None
    
    format_green_percentage = None
    format_red_percentage = None
    format_grey_float_1decimal = None
    format_green_int = None
    format_red_int = None
    format_green_int = None

def generate_excelfile(logger, filepath, appName, snapshotversion, snapshotdate, loadviolations, listbusinesscriteria, dictechnicalcriteria, listbccontributions, listtccontributions, dictmetrics, dictapsummary, dicremediationabacus, listviolations, broundgrades, dictaptriggers):
    broundgrades = broundgrades
    format = ExcelFormat()
    pd.options.display.float_format = format.const_float_format.format
    
    logger.info("Loading data in Excel")
    
    #Readme Page content
    str_readme_content =  "Tab;Content;Comment\n"
    str_readme_content += "README;Read me;\n"
    str_readme_content += "BC Grades;Business Criteria current grade and simulation grade;Use this sheet to see the global impact on application grades and total estimated effort\n"
    str_readme_content += "TC Grades;Technical criteria current grade and simulation grade;\n"
    str_readme_content += "Rules Grades;Quality Rules, Distributions and Measures grades and simulation;Use this sheet to change the number of violations for action and see the impact on rules grades and estimated effort\n"
    if loadviolations:
        str_readme_content += "Violations;Violations list;Use this sheet to select your violations for action\n"
    str_readme_content += "BC contributions;Business Criteria contributors (Technical criteria);\n"
    str_readme_content += "TC Contributions;Technical Criteria contributors (Quality metrics);\n"
    str_readme_content += "Remediation effort;Quality rules unit remediation effort;Use to sheet to set or modify the unit remediation effort per quality rule\n"
    
    try: 
        df_readme = pd.read_csv(StringIO(StringUtils.remove_unicode_characters(str_readme_content)), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : readme',True)
    
    ###############################################################################
    # Data for the BC Grades Tab
    
    str_df_bc_grades = "Application Name;Business criterion;Metric Id;Grade;Simulation grade;Lowest critical grade;Weighted average of Technical criteria; Delta\n"
    for bc in listbusinesscriteria:
        if bc.applicationName == appName:
            str_df_bc_grades += appName + ";" + bc.name + ";" + bc.id + ";" + str(round_grades(broundgrades, bc.grade)) + ";;;;"
            str_df_bc_grades += '\n'
    
    
    emptyline = ";;;;;;;\n"
    # Summary
    str_df_bc_grades += emptyline+emptyline+emptyline
    str_df_bc_grades += ";Application name;" + appName + "\n"
    str_df_bc_grades += ";Version;" + snapshotversion + "\n"
    str_df_bc_grades += ";Date;" + snapshotdate + "\n"    
    str_df_bc_grades += ';Number of violations for action\n'
    str_df_bc_grades += ';Number of quality rules for action\n'
    str_df_bc_grades += ';Estimated effort (man.days)\n'

    if loadviolations:
        str_df_bc_grades += '\n'
        str_df_bc_grades += ';Number of action plans added\n'
        str_df_bc_grades += ';Number of action plans removed\n'
        #TODO: identify the action plan modified
        str_df_bc_grades += ';Number of action plans modified; <Not available>\n'
        str_df_bc_grades += ';JSON violations added\n'
        str_df_bc_grades += ';JSON violations removed\n'
        str_df_bc_grades += ';JSON violations modified\n'
    try: 
        str_df_bc_grades = StringUtils.remove_unicode_characters(str_df_bc_grades)
        df_bc_grades = pd.read_csv(StringIO(str_df_bc_grades), sep=";")
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_bc_grades %s ' % str_df_bc_grades,True)
    
    ###############################################################################
    # Data for the TC Grades Tab
    str_df_tc_grades = "Technical criterion name;Metric Id;Grade;Simulation grade;Lowest critical grade;Weighted average of quality rules;Delta grade (%)\n"
    for tc in dictechnicalcriteria:
        #print('tc grade 2=' + str(tc.grade) + str(type(tc.grade)))
        otc = dictechnicalcriteria[tc]
        str_line = otc.name + ';' + str(otc.id) + ';'+ str(round_grades(broundgrades,otc.grade)) + ';;;;' + '\n'
        str_df_tc_grades += str_line
    try: 
        str_df_tc_grades = StringUtils.remove_unicode_characters(str_df_tc_grades)
        df_tc_grades = pd.read_csv(StringIO(str_df_tc_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_tc_grades %s ' % str_df_tc_grades,True)

    ###############################################################################
    # Data for the Rules Grades Tab

    str_df_rules_grades = "Application Name;Snapshot Date;Snapshot version;Metric Name;Metric Id;Metric Type;Critical;Grade;Simulation grade;Grade Delta;Grade Delta (%);Nb of violations;Nb violations for action;Remaining violations;Unit effort (man.hours);Total effort (man.hours);Total effort (man.days);Total Checks;Compliance ratio;New compliance ratio;Thres.1;Thres.2;Thres.3;Thres.4;Educate (Mark for ...);Violations extracted\n"
    for qr in dictmetrics:
        oqr = dictmetrics[qr]
        str_line = ''
        str_line += appName
        str_line += ";" + str(snapshotdate) 
        str_line += ";" + str(snapshotversion) 
        str_line += ";" +  str(oqr.name)
        str_line += ";" + str(oqr.id) 
        #str_line += ";" 
        
        #TODO: formatting problem to fix ? 
        #str_line += ";" + str(oqr.type)
        str_line += ";" + str(oqr.type) 
        
        str_line += ";" + str(oqr.critical) 
        str_line += ";" + str(round_grades(broundgrades,oqr.grade)) 
        #simulation grade, grade delta%, grade delta%
        str_line += ';;;' 
        #failed checks
        str_line += ';'
        if oqr.failedchecks != None: str_line += str(oqr.failedchecks)
        #number of actions
        str_line += ';'
        if dictapsummary.get(oqr.id) != None and oqr.type == type_quality_rules:
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
        if oqr.totalchecks != None:
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
        if dictaptriggers.get(oqr.id) != None and oqr.type == type_quality_rules:
            if dictaptriggers.get(oqr.id):
                str_line += 'Action'
            else:
                str_line += 'Continuous improvement'
        #new compliance ratio
        str_line += ';'
        str_line += '\n'
        str_df_rules_grades += str_line
    #logger.debug(str_df_rules_grades)
    try: 
        str_df_rules_grades = StringUtils.remove_unicode_characters(str_df_rules_grades)
        df_rules_grades = pd.read_csv(StringIO(str_df_rules_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_rules_grades %s ' % str_df_rules_grades,True)

    ###############################################################################
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
            df_violations = pd.read_csv(StringIO(str_df_violations), sep=";",engine='python',quoting=csv.QUOTE_NONE) 
        except: 
            LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_violations %s ' % str_df_violations,True)
    
    ###############################################################################
    
    # Data for the BC Contributions Tab
    str_df_bc_contribution = 'Business criterion name;Business criterion Id;Technical criterion name;Technical criterion Id;Weight;Critical;Simulation grade;Weighted grade\n'
    for bcc in listbccontributions:
        hasresults = False
        for tc in dictechnicalcriteria:
            if tc == bcc.metricid: hasresults = True 
        # keep only the technical criteria that have results
        if not hasresults:
            continue        
        str_df_bc_contribution += bcc.parentmetricname + ';' + bcc.parentmetricid + ';' + bcc.metricname + ';' + bcc.metricid
        str_df_bc_contribution += ';' + str(bcc.weight) + ';' + str(bcc.critical) + ';;'
        str_df_bc_contribution += '\n'
    #logger.debug(str_df_bc_contribution)
    try: 
        str_df_bc_contribution = StringUtils.remove_unicode_characters(str_df_bc_contribution)
        df_bc_contribution = pd.read_csv(StringIO(str_df_bc_contribution), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_bc_contribution %s ' % str_df_bc_contribution,True)
        
    ###############################################################################
    # Data for the TC Contributions Tab
    str_df_tc_contribution = 'Technical criterion name;Technical criterion Id;Metric name;Metric Id;Weight;Critical;Grade simulation;Weighted grade\n'
    # for each contribution TC/QR 
    for tcc in listtccontributions:
        QRhasresults = False
        for met in dictmetrics:
            if str(met) == str(tcc.metricid): 
                QRhasresults = True
        # keep only the quality metrics that have metrics that have results 
        if QRhasresults:
            #print (tcc.metricid)
            str_df_tc_contribution += tcc.parentmetricname + ';' + tcc.parentmetricid + ';' + tcc.metricname + ';' + tcc.metricid
            str_df_tc_contribution += ';' + str(tcc.weight) + ';' + str(tcc.critical) + ';;'
            str_df_tc_contribution += '\n'
    try: 
        str_df_tc_contribution = StringUtils.remove_unicode_characters(str_df_tc_contribution)
        df_tc_contribution = pd.read_csv(StringIO(str_df_tc_contribution), sep=";",quoting=csv.QUOTE_NONE)
    #,engine='python'
    except: 
        LogUtils.logerror(logger, 'csv.Error: unexpected end of data : df_tc_contribution %s ' % str_df_tc_contribution, True)
    ###############################################################################
    # Data for the Remediation Tab
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
        
    ###############################################################################
    logger.info("Writing data in Excel")
    file = open(filepath, 'w')
    with pd.ExcelWriter(filepath,engine='xlsxwriter') as writer:
        df_readme.to_excel(writer, sheet_name=format.const_TAB_README, index=False)
        df_bc_grades.to_excel(writer, sheet_name=format.const_TAB_BC_GRADES, index=False)
        df_tc_grades.to_excel(writer, sheet_name=format.const_TAB_TC_GRADES, index=False)
        df_rules_grades.to_excel(writer, sheet_name=format.const_TAB_RULES_GRADES, index=False)
        if loadviolations:
            df_violations.to_excel(writer, sheet_name=format.const_TAB_VIOLATIONS, index=False)        
        df_bc_contribution.to_excel(writer, sheet_name=format.const_TAB_BC_CONTRIBUTIONS, index=False) 
        df_tc_contribution.to_excel(writer, sheet_name=format.const_TAB_TC_CONTRIBUTIONS, index=False)
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
        
        worksheet = writer.sheets[format.const_TAB_BC_GRADES]
        format_table_bc_grades(workbook,worksheet,df_bc_grades,format,loadviolations)   
    
        worksheet = writer.sheets[format.const_TAB_TC_GRADES]
        format_table_tc_grades(workbook,worksheet,df_tc_grades,format)  
    
        worksheet = writer.sheets[format.const_TAB_RULES_GRADES]
        format_table_rules_grades(workbook,worksheet,df_rules_grades,format,listmetricsinviolations)  
    
        if loadviolations:
            worksheet = writer.sheets[format.const_TAB_VIOLATIONS]
            format_table_violations(workbook,worksheet,df_violations,format)      
    
        worksheet = writer.sheets[format.const_TAB_BC_CONTRIBUTIONS]
        format_table_bc_contribution(workbook,worksheet,df_bc_contribution,format)     
        
        worksheet = writer.sheets[format.const_TAB_TC_CONTRIBUTIONS]
        format_table_tc_contribution(workbook,worksheet,df_tc_contribution,format)  

        worksheet = writer.sheets[format.const_TAB_REMEDIATION_EFFORT]        
        format_table_remediation_effort(workbook,worksheet,df_remediationeffort,format)  
        
        worksheet = writer.sheets[format.const_TAB_BC_GRADES]
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

def format_table_bc_grades(workbook,worksheet,table,format,loadviolations):
    worksheet.set_tab_color(format.const_color_light_blue)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
   
    # the last 6 lines don't have this formula    
    offset = 1
    if not loadviolations:
        nb_rows = len(table.index.values)+1 - 9
    else: 
        nb_rows = len(table.index.values)+1 - 15
    
    col_to_format = colnum_string(len(table.columns) + 1 + offset)    

    
    # 3 empty line + 3 lines for application name, snapshot version and date
    row_to_format_for_summary = nb_rows + 6

    start = "H2"
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    #print("range {}".format(range_to_format))
    
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '>','value': 0.000001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '<','value': -0.000001, 'format':   format.format_red_percentage})

    worksheet.set_column('A:A', 20, None) # Application column
    worksheet.set_column('B:B', 32, format.format_align_left) # BC name
    worksheet.set_column('C:C', 7.5, format.format_align_left) # Metric Id
    worksheet.set_column('D:D', 11, format.format_float_with_2decimals) # Grade 
    worksheet.set_column('E:E', 11, format.format_float_with_2decimals) # Simulated grade 
    # group and hide columns lowest critical grade and weighted average
    worksheet.set_column('F:F', 15, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) #  
    worksheet.set_column('G:G', 20, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) #  
    # group and hide columns lowest critical grade and weighted average
    #worksheet.set_column('F:F', None, None, {'level': 1, 'collapsed': True})
    #worksheet.set_column('G:G', None, None, {'level': 1, 'collapsed': True})    
    worksheet.set_column('H:H', 11, format.format_percentage) # delta %  
    last_column = 'H'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))      
    
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        # simulation grade
        worksheet.write_formula(row_num, 5-1, round_grades(broundgrades,'=IF(F%d=0,G%d,MIN(F%d,G%d))') % (row_num + 1, row_num + 1, row_num + 1, row_num + 1))
        # lowest critical
        worksheet.write_formula(row_num, 6-1,  round_grades(broundgrades,"=_xlfn.MINIFS('BC contributions'!G:G,'BC contributions'!B:B,'BC grades'!C%d,'BC contributions'!F:F,TRUE)") % (row_num + 1))
        # weighted average
        worksheet.write_formula(row_num, 7-1, round_grades(broundgrades,"=SUMIF('BC contributions'!B:B,C%d,'BC contributions'!H:H)/SUMIF('BC contributions'!B:B,C%d,'BC contributions'!E:E)") % (row_num + 1, row_num + 1))
        # Delta %
        worksheet.write_formula(row_num, 8-1, '=($E%d-$D%d)/$D%d' % (row_num + 1, row_num + 1, row_num + 1), format.format_percentage)
    #number of violations
    worksheet.write_formula(row_to_format_for_summary, 3-1, "=SUM('Rules Grades'!M:M)")
    #number of quality rules for action
    worksheet.write_formula(row_to_format_for_summary+1, 3-1, "=COUNTIF('Rules Grades'!M:M,\">0\")")
    #estimated effort m.d
    worksheet.write_formula(row_to_format_for_summary+2, 3-1, "=SUM('Rules Grades'!Q:Q)")
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
    id = ""
    hrefsplit = href.split('/')
    for elem in hrefsplit:
        # the last element is the id
        id = elem    
    return id  
    
########################################################################    
    
def format_table_tc_grades(workbook,worksheet,table,format):
    worksheet.set_tab_color(format.const_color_light_grey)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)

    offset = 1 
    nb_rows = len(table.index.values)+1
    col_to_format = colnum_string(len(table.columns) + 1 + offset)    

    #define the range to be formated in excel format
    start = "G2"
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '>','value': 0.000001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '<','value': -0.000001, 'format':   format.format_red_percentage})
            
    worksheet.set_column('A:A', 60, None) #  TC name
    worksheet.set_column('B:B', 8, format.format_align_left) # Id
    worksheet.set_column('C:C', 8, format.format_float_with_2decimals) # Grade
    worksheet.set_column('D:D', 10, format.format_float_with_2decimals) # Simulation grade
    # group and hide columns lowest critical grade and weighted average
    worksheet.set_column('E:E', 13, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) # 
    worksheet.set_column('F:F', 19, format.format_float_with_2decimals, {'level': 1, 'hidden': True}) # 
    worksheet.set_column('G:G', 12, format.format_percentage) # 
    last_column = 'G'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))     
 
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        #simulation grade
        worksheet.write_formula(row_num, 4-1, round_grades(broundgrades,"=IF(E%d=0,F%d,MIN(E%d,F%d))") % (row_num + 1, row_num + 1, row_num + 1, row_num + 1), format.format_float_with_2decimals)
        #lowest critical rule grade
        worksheet.write_formula(row_num, 5-1, round_grades(broundgrades,"=_xlfn.MINIFS('TC contributions'!G:G,'TC contributions'!B:B,'TC grades'!B%d,'TC contributions'!F:F,TRUE)") % (row_num + 1), format.format_float_with_2decimals)
        #weighted av
        worksheet.write_formula(row_num, 6-1, round_grades(broundgrades,"=SUMIF('TC contributions'!B:B,'TC grades'!B%d,'TC contributions'!H:H)/SUMIF('TC contributions'!B:B,'TC grades'!B%d,'TC contributions'!E:E)") % (row_num + 1, row_num + 1), format.format_float_with_2decimals)
        #delta %
        worksheet.write_formula(row_num, 7-1, "=($D%d-$C%d)/$C%d" % (row_num + 1, row_num + 1, row_num + 1), format.format_percentage)
    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format) 
 

 
########################################################################

def format_table_rules_grades(workbook,worksheet,table,format,listmetricsinviolations):
    worksheet.set_tab_color(format.const_color_light_blue)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1})
    worksheet.set_zoom(85)
    worksheet.freeze_panes(1, 0)  # Freeze the first row.    
    nb_rows = len(table.index.values)+1
    
    # conditional formating for the Grade delta column (red and green)
    col_to_format = 'K'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '>', 'value': 0.000001, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '<', 'value': -0.000001, 'format':   format.format_red_percentage})   

    # conditional formating for the number of violations for action
    col_to_format = 'M'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '>', 'value': 0.000001, 'format':   format.format_green_int})
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '<', 'value': -0.000001, 'format':   format.format_red_int})   

    # conditional formating for the unit effort column
    col_to_format = 'O'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})
    # conditional formating for the total effort column in hours
    col_to_format = 'P'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})
    # conditional formating for the total effort column in days
    col_to_format = 'Q'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '=', 'value': 0, 'format':   format.format_grey_float_1decimal})


    worksheet.set_column('A:A', 25, None) #  Application name
    worksheet.set_column('B:B', 12, format.format_align_left) # Application column 
    worksheet.set_column('C:C', 10, format.format_align_left) # Snapshot date
    worksheet.set_column('D:D', 60, None) # Metric name 
    worksheet.set_column('E:E', 8, None) # metric id
    worksheet.set_column('F:F', 18, None) #  
    worksheet.set_column('G:G', 6.5, None) #  
    worksheet.set_column('H:H', 9, format.format_float_with_2decimals) #   
    worksheet.set_column('I:I', 8, format.format_float_with_2decimals) #    
    worksheet.set_column('J:J', 8, format.format_float_with_2decimals)
    worksheet.set_column('K:K', 8, format.format_percentage) # % Grade delta
    
    worksheet.set_column('L:L', 10, None) # Nb violations
    worksheet.set_column('M:M', 11, None) # Nb violations for action
    worksheet.set_column('N:N', 10, None) # Reminaing vioaltions
    worksheet.set_column('O:O', 11, format.format_float_with_2decimals) # Unit effort
    worksheet.set_column('P:P', 11, format.format_float_with_2decimals) # Total effort mh
    worksheet.set_column('Q:Q', 11, format.format_float_with_2decimals) # Total effort md
    worksheet.set_column('R:R', 11, format.format_int_thousands) # total checks
    worksheet.set_column('S:S', 11, format.format_percentage) # % compliance ratio
    worksheet.set_column('T:T', 11, format.format_percentage) # % new compliance ratio
    worksheet.set_column('U:U', 6.5, None) # Thres 1   
    worksheet.set_column('V:V', 6.5, None) #
    worksheet.set_column('W:W', 6.5, None) #
    worksheet.set_column('X:X', 6.5, None) # Thres 4
    
    worksheet.set_column('Y:Y', 11, None) # Educate   
    worksheet.set_column('Z:Z', 11, None) # violations extracted ?
    last_column='Z'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))     

    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        metrictype = str(table.loc[row_num-1, 'Metric Type'])
        metricid = str(table.loc[row_num-1, 'Metric Id'])
        
        # formulas applicable only for quality-rules, not for quality-measures and quality-distributions 
        if metrictype == type_quality_rules:
            #simulation grade
            formula = round_grades(broundgrades,'=IF(T%d=0,H%d,IF(T%d<=U%d/100,1,IF(T%d<V%d/100,(T%d*100-U%d)/(V%d-U%d)+1,IF(T%d<W%d/100,(T%d*100-V%d)/(W%d-V%d)+2,IF(T%d<X%d/100,(T%d*100-W%d)/(X%d-W%d)+3,4)))))') % (row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1,row_num + 1)
            worksheet.write_formula(row_num, 9-1, formula)
            #grade delta
            worksheet.write_formula(row_num, 10-1, '=$I%d-$H%d' % (row_num + 1, row_num + 1))
            #grade delta %
            worksheet.write_formula(row_num, 11-1, '=$J%d/$H%d' % (row_num + 1, row_num + 1))
            
            #Nb violations for action, with a formula only if the violations are loaded
            # also only if we find the metric id in the list of violations
            if listmetricsinviolations != None and len(listmetricsinviolations) > 0 and metricid in listmetricsinviolations:
                formula = "=SUMIF(Violations!B:B,'Rules Grades'!E%d,Violations!N:N)"% (row_num + 1)
                #print(formula)
                worksheet.write_formula(row_num, 13-1, formula)
            #remaining violations
            worksheet.write_formula(row_num, 14-1, '=$L%d-$M%d' % (row_num + 1, row_num + 1))
            #unit effort
            worksheet.write_formula(row_num, 15-1, "=(VLOOKUP(E%d,'Remediation effort'!A:C,3,FALSE))/60" % (row_num + 1))        
            #total effort (mh)
            worksheet.write_formula(row_num, 16-1, "=O%d*M%d" % (row_num + 1, row_num + 1))
            #total effort (md)
            worksheet.write_formula(row_num, 17-1, "=P%d/8" % (row_num + 1))
            #new compliance ratio
            worksheet.write_formula(row_num, 20-1, '=($R%d-$N%d)/$R%d' % (row_num + 1, row_num + 1, row_num + 1))
            #Violations extracted ? Present in violations tab
            #TODO: fix Violations extracted formula
            #formula = '=IF(NOT(ISNA(VLOOKUP($E%d,Violations!B:B,1,FALSE))),TRUE,FALSE)' % (row_num + 1)
            formula = '=IF(NOT(ISNA(VLOOKUP($E%d,Violations!B:B,1,FALSE))),TRUE,FALSE)' % (row_num + 1)
            #print(formula)
            worksheet.write_formula(row_num, 26-1, formula)
            
        else:
            # simulation grade = grade
            worksheet.write_formula(row_num, 9-1, '=$H%d' % (row_num + 1))
            # grade delta
            worksheet.write_formula(row_num, 10-1, '=0')
            # grade delta % 
            worksheet.write_formula(row_num, 11-1, '=0')

        # Write the column headers with the defined format.
        for col_num, value in enumerate(table.columns.values):
            worksheet.write(0, col_num, value, header_format)

    # group and hide the context
    worksheet.set_column('A:A', None, None, {'level': 1, 'hidden': True})
    worksheet.set_column('B:B', None, None, {'level': 1, 'hidden': True})
    worksheet.set_column('C:C', None, None, {'level': 1, 'hidden': True})

    # group and hide the thresholds
    worksheet.set_column('V:V', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('W:W', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('X:X', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('Y:Y', None, None, {'level': 2, 'hidden': True})

########################################################################


def format_table_violations(workbook,worksheet,table,format):
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

def format_table_bc_contribution(workbook,worksheet,table, format):
    worksheet.set_tab_color(format.const_color_light_grey)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1})
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
    nb_rows = len(table.index.values)+1
    
    worksheet.set_column('A:A', 25, None) #  
    worksheet.set_column('B:B', 13, format.format_align_left) # Application column 
    worksheet.set_column('C:C', 60, format.format_align_left) # BC column 
    worksheet.set_column('D:D', 13, None) # Metric Id column 
    worksheet.set_column('E:E', 9, None) # HF column 
    worksheet.set_column('F:F', 9, None) # HF column 
    worksheet.set_column('G:G', 13, format.format_float_with_2decimals) # HF column 
    worksheet.set_column('H:H', 13, format.format_float_with_2decimals) # HF column  
    last_column = 'H'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))

    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        worksheet.write_formula(row_num, 7 - 1, "=VLOOKUP(D%d,'TC grades'!B:D,3,FALSE)" % (row_num + 1))
        worksheet.write_formula(row_num, 8 - 1, '=$G%d*$E%d' % (row_num + 1, row_num + 1))

    # Write the column headers with the defined format.
    for col_num, value in enumerate(table.columns.values):
        worksheet.write(0, col_num, value, header_format)

########################################################################

def format_table_tc_contribution(workbook,worksheet,table,format):
    worksheet.set_tab_color(format.const_color_light_grey)
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'middle','fg_color': format.const_color_header_columns,'border': 1}) 
    worksheet.freeze_panes(1, 0)  # Freeze the first row.
    worksheet.set_zoom(85)
    
    offset = 1 
    nb_rows = len(table.index.values)+1
    col_to_format = colnum_string(len(table.columns) + 1 + offset)    
    
    worksheet.set_column('A:A', 60, None) #  
    worksheet.set_column('B:B', 13, format.format_align_left)  
    worksheet.set_column('C:C', 70, format.format_align_left)  
    worksheet.set_column('D:D', 8, None) # 
    worksheet.set_column('E:E', 9, None) #  
    worksheet.set_column('F:F', 9, None) #  
    worksheet.set_column('G:G', 13, format.format_float_with_2decimals) #  grade simulation
    worksheet.set_column('H:H', 13, format.format_float_with_2decimals) #  weighted grade
    last_column = 'H'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))    
   
    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        worksheet.write_formula(row_num, 7 - 1, "=VLOOKUP(D%d,'Rules Grades'!E:I,5,FALSE)" % (row_num + 1))
        worksheet.write_formula(row_num, 8 - 1, '=$G%d*$E%d' % (row_num + 1, row_num + 1))

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