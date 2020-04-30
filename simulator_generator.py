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
from io import StringIO
import xlsxwriter
import xml.etree.ElementTree as ET
from utils.utils import RestUtils, AIPRestAPI, LogUtils, ObjectViolationMetric, RulePatternDetails, FileUtils, StringUtils, Metric, Contribution, Violation

'''
 Author : MMR & TGU
 March 2020
'''
########################################################################

# Total Quality Index,Security,Efficiency,Robustness,Transferability,Changeability,Coding Best Practices/Programming Practices,Documentation,Architectural Design
bcids = ["60017","60016","60016","60014","60013","60012","60011","66031","66032","66033"]

########################################################################

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
    requiredNamed.add_argument('-of', required=False, dest='outputfolder', help='output folder')    
    requiredNamed.add_argument('-effortcsvfilepath', required=False, dest='effortcsvfilepath', help='Inputs quality rules effort csv file path (default=CAST_QualityRulesEffort.csv)')    
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
    
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '>','value': 0, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type': 'cell','criteria': '<','value': 0, 'format':   format.format_red_percentage})

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
        worksheet.write_formula(row_num, 8-1, '=$E%d-$D%d' % (row_num + 1, row_num + 1), format.format_percentage)
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
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '>','value':    0, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':     'cell', 'criteria': '<','value':    0, 'format':   format.format_red_percentage})
            
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
        worksheet.write_formula(row_num, 7-1, "=$D%d-$C%d" % (row_num + 1, row_num + 1), format.format_percentage)
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
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '>', 'value': 0, 'format':   format.format_green_percentage})
    worksheet.conditional_format(range_to_format, {'type':'cell', 'criteria': '<', 'value': 0, 'format':   format.format_red_percentage})   

    # conditional formating for the number of violations for action
    col_to_format = 'M'    
    start = col_to_format + '2'
    #define the range to be formated in excel format
    range_to_format = "{}:{}{}".format(start,col_to_format,nb_rows)
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '>', 'value': 0, 'format':   format.format_green_int})
    worksheet.conditional_format(range_to_format, {'type': 'cell', 'criteria': '<', 'value': 0, 'format':   format.format_red_int})   

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
    worksheet.set_column('J:J', 8, format.format_float_with_2decimals) #    
    worksheet.set_column('K:K', 8, format.format_percentage) # % 
    worksheet.set_column('L:L', 10, None) #
    worksheet.set_column('M:M', 11, None) #
    
    worksheet.set_column('N:N', 10, None) #
    worksheet.set_column('O:O', 11, format.format_float_with_2decimals) #
    worksheet.set_column('P:P', 11, format.format_float_with_2decimals) #
    worksheet.set_column('Q:Q', 11, format.format_float_with_2decimals) #
    worksheet.set_column('R:R', 11, format.format_int_thousands) #
    worksheet.set_column('S:S', 11, format.format_percentage) # %
    worksheet.set_column('T:T', 11, format.format_percentage) # %
    worksheet.set_column('U:U', 6.5, None) # Thres 1   
    worksheet.set_column('V:V', 6.5, None) #
    worksheet.set_column('W:W', 6.5, None) #
    worksheet.set_column('X:X', 6.5, None) # Thres 4   
    worksheet.set_column('Y:Y', 11, None) # violations extracted ?
    last_column='Y'
    worksheet.autofilter('A1:' + last_column + str(nb_rows))     

    # Create a for loop to start writing the formulas to each row
    for row_num in range(1,nb_rows):
        metrictype = str(table.loc[row_num-1, 'Metric Type'])
        metricid = str(table.loc[row_num-1, 'Metric Id'])
        
        # formulas applicable only for quality-rules, not for quality-measures and quality-distributions 
        if metrictype == 'quality-rules':
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
            worksheet.write_formula(row_num, 25-1, formula)
            
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
    worksheet.set_column('U:U', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('V:V', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('W:W', None, None, {'level': 2, 'hidden': True})
    worksheet.set_column('X:X', None, None, {'level': 2, 'hidden': True})

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
    worksheet.set_column('K:K', 11, None) # Viol status
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
    
    worksheet.set_column('A:A', 45, None) #  
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

########################################################################

def generate_excelfile(logger, filepath, appName, snapshotversion, snapshotdate, loadviolations, listbusinesscriteria, listtechnicalcriteria, listbccontributions, listtccontributions, listmetrics, dictapsummary, dicremediationabacus, listviolations, broundgrades):
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
    for tc in listtechnicalcriteria:
        #print('tc grade 2=' + str(tc.grade) + str(type(tc.grade)))
        str_df_tc_grades += tc.name + ';' + str(tc.id) + ';'+ str(round_grades(broundgrades,tc.grade)) + ';;;;'  
        str_df_tc_grades  += '\n'
    try: 
        str_df_tc_grades = StringUtils.remove_unicode_characters(str_df_tc_grades)
        df_tc_grades = pd.read_csv(StringIO(str_df_tc_grades), sep=";",engine='python',quoting=csv.QUOTE_NONE)
    except: 
        LogUtils.logerror(logger,'csv.Error: unexpected end of data : df_tc_grades %s ' % str_df_tc_grades,True)

    ###############################################################################
    # Data for the Rules Grades Tab

    str_df_rules_grades = "Application Name;Snapshot Date;Snapshot version;Metric Name;Metric Id;Metric Type;Critical;Grade;Simulation grade;Grade Delta;Grade Delta (%);Nb of violations;Nb violations for action;Remaining violations;Unit effort (man.hours);Total effort (man.hours);Total effort (man.days);Total Checks;Compliance ratio;New compliance ratio;Thres.1;Thres.2;Thres.3;Thres.4;Violations extracted\n"
    for qr in listmetrics:
        str_df_rules_grades += appName
        str_df_rules_grades += ";" + str(snapshotdate) 
        str_df_rules_grades += ";" + str(snapshotversion) 
        str_df_rules_grades += ";" +  str(qr.name)
        str_df_rules_grades += ";" + str(qr.id) 
        str_df_rules_grades += ";" + str(qr.type) 
        str_df_rules_grades += ";" + str(qr.critical) 
        str_df_rules_grades += ";" + str(round_grades(broundgrades,qr.grade)) 
        #simulation grade, grade delta%, grade delta%
        str_df_rules_grades += ';;;' 
        #failed checks
        str_df_rules_grades += ';'
        if qr.failedchecks != None: str_df_rules_grades += str(qr.failedchecks)
        #number of actions
        str_df_rules_grades += ';'
        if dictapsummary.get(qr.id) != None and qr.type == 'quality-rules':
            str_df_rules_grades += str(dictapsummary.get(qr.id)) 
        #remaining violations
        str_df_rules_grades += ';'
        #unit effort mh, total effort mh, total effort md
        str_df_rules_grades += ';;;'
        #total checks 
        str_df_rules_grades += ';'
        if qr.totalchecks != None: 
            str_df_rules_grades += str(qr.totalchecks) 
        #compliance ratio
        str_df_rules_grades += ';'
        if qr.totalchecks != None:
            str_df_rules_grades += str(qr.ratio)
        #new compliance ratio
        str_df_rules_grades += ';'
        #4 thresholds new compliance ratio
        if qr.type == 'quality-rules':
            str_df_rules_grades += ';'+str(qr.threshold1)+';'+str(qr.threshold2)+';'+str(qr.threshold3)+';' + str(qr.threshold4)
        else:
            str_df_rules_grades += ';;;;;'
        str_df_rules_grades += '\n'
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
        for tc in listtechnicalcriteria:
            if tc.id == bcc.metricid: hasresults = True 
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
        for met in listmetrics:
            if str(met.id) == str(tcc.metricid): 
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
    for qr in listmetrics:
        # we are looking only at quality rules here, not distributions or measures 
        if qr.type == 'quality-rules':
            str_df_remediationeffort += str(qr.id) + ';' + str(qr.name) 
            # if the quality rule is not 
            if dicremediationabacus.get(qr.id) != None and dicremediationabacus.get(qr.id).get('uniteffortinhours'):
                #print (str(qr.id) + ' in the abacus')
                str_df_remediationeffort += ';' + str(dicremediationabacus.get(qr.id).get('uniteffortinhours'))
            else:
                #print (str(qr.id) + ' not in the abacus => N/A')
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

########################################################################

def get_excelfilepath(outputfolder, appName):
    fpath = ''
    if outputfolder != None:
        fpath = outputfolder + '/'
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
    loadonlyXmetrics = False    
    # round the grades or not
    broundgrades = False

    parser = init_parse_argument()
    args = parser.parse_args()
    restapiurl = args.restapiurl
    if restapiurl != None and restapiurl[-1:] == '/':
        # remove the trailing / 
        restapiurl = restapiurl[:-1] 
    edurl = restapiurl 
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
    effortcsvfilepath = "CAST_QualityRulesEffort.csv"
    if args.effortcsvfilepath != None:
        effortcsvfilepath = args.effortcsvfilepath 

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
            LogUtils.loginfo(logger,pluginfile,True)
            tree = ET.parse(pluginfile)
            root = tree.getroot()
            namespace = "{http://schemas.microsoft.com/packaging/2011/08/nuspec.xsd}"
            for versiontag in root.findall('{0}metadata/{0}version'.format(namespace)):
                script_version = versiontag.text
        except:
            None 
        
        protocol = 'Undefined'
        host = 'Undefined'
        warname = 'Undefined'
        
        # split the URL to extract the warname, host, protocol ... 
        rexURL = "(([hH][tT][tT][pP][sS]*)[:][/][/]([A-Za-z0-9_:\.-]+)([/]([A-Za-z0-9_\.-]+))*[/]*)"
        m0 = re.search(rexURL, restapiurl)
        if m0:
            protocol = m0.group(2)
            host = m0.group(3)
            warname = m0.group(5)
    
        # log params
        logger.info('********************************************')
        LogUtils.loginfo(logger,'log script_version='+script_version,True)
        logger.info('python version='+sys.version)
        logger.info('****************** params ******************')
        logger.info('restapiurl='+restapiurl)
        logger.info('edurl='+edurl)        
        logger.info('host='+host)
        logger.info('protocol='+protocol)
        logger.info('warname='+str(warname))
        logger.info('usr='+str(user))
        logger.info('pwd=*******')
        logger.info('apikey='+str(apikey))
        LogUtils.loginfo(logger,'log file='+log,True)
        logger.info('extensioninstallationfolder='+extensioninstallationfolder)
        logger.info('log level='+loglevel)
        logger.info('applicationfilter='+str(applicationfilter))
        logger.info('nbrows='+str(nbrows))
        logger.info('output folder='+str(outputfolder))
        logger.info('effortcsvfilepath='+str(effortcsvfilepath))
        logger.info('loadviolations='+str(loadviolations))
        logger.info('qridfilter='+str(qridfilter))
        logger.info('qrnamefilter='+str(qrnamefilter))
        logger.info('criticalrulesonlyfilter='+str(criticalrulesonlyfilter))
        logger.info('businesscriterionfilter='+str(businesscriterionfilter))
        logger.info('technofilter='+str(technofilter)) 
        logger.info('********************************************')
        
        LogUtils.loginfo(logger, 'Initialization', True) 
        rest_utils = RestUtils(logger, restapiurl, user, password, apikey)
        rest_utils.open_session()
        rest_service = AIPRestAPI(rest_utils) 
        
        # few checks on the server 
        json_server = rest_service.get_server()
        if json_server != None:
            logger.info('server status=' + json_server['status'])    
            servversion = json_server['version']
            logger.info('server version=' + servversion)
            #servversion2digits = servversion[-4:] 
            #if float(servversion2digits) <= 1.13 : 
            #    None
            logger.info('server memory (free)=' + str(json_server['memory']['freeMemory']))
            logger.info('********************************************')    
        
        # retrieve the domains & the applications in those domains 
        json_domains = rest_service.get_domains()
        if json_domains != None:
            bAEDdomainFound = False
            for item in json_domains:
                domain = item['href']
                if domain != 'AAD':
                    bAEDdomainFound = True
                    
            idomain = 0            
            for item in json_domains:
                idomain += 1
                domain = ''
                try:
                    domain = item['href']
                except KeyError:
                    pass
                
                LogUtils.loginfo(logger, "Domain " + domain + " | progress:" + str(idomain) + "/" + str(len(json_domains)), True)
 
                # only engineering domains, or AAD domain only in case there is no engineering domain, we prefer to have engineering domains containing of action plan summary
                if domain == 'AAD' and bAEDdomainFound:
                    logger.info("  Skipping domain " + domain + ", because we process in priority Engineering domains")
                    continue
                
                if domain != 'AAD' or not bAEDdomainFound:
                    json_apps = rest_service.get_applications(domain)
                    applicationid = -1
                    appHref = ''
                    appName = ''
                    for app in json_apps:
                        try:
                            appName = app['name']
                        except KeyError:
                            pass                        
                        try:
                            appHref = app['href']
                        except KeyError:
                            pass     
                        hrefsplit = appHref.split('/')
                        for elem in hrefsplit:
                            # the last element is the id
                            applicationid = elem
                            
                        #appName = 'eCommer.*'
                        if applicationfilter != None and not re.match(applicationfilter, appName):
                            logger.info('Skipping application : ' + appName)
                            continue                
                        elif applicationfilter == None or applicationfilter == '' or re.match(applicationfilter, appName):
                            LogUtils.loginfo(logger, "Processing application " + appName, True)
                            # testing if csv file can be written
                            fpath = get_excelfilepath(outputfolder, appName)
                            # if the output file is locked we move to next application
                            if checkoutputfilelocked(logger, fpath):
                                continue

                            listbusinesscriteria = []
                            dicremediationabacus = {}
                            # applications health factors for last snapshot
                            if (loaddata):
                                logger.info('Extracting the applications business criteria grades for last snapshot')
                                json_bc_grades = rest_service.get_businesscriteria_grades(domain)
                                if json_bc_grades != None:
                                    for res in json_bc_grades:
                                        for bc in res['applicationResults']:
                                            businesscriterion = Metric()
                                            businesscriterion.applicationName = res['application']['name'] 
                                            businesscriterion.name = bc['reference']['name']
                                            businesscriterion.id = bc['reference']['key']
                                            businesscriterion.grade = bc['result']['grade']
                                            #print('bc grade=' + str(businesscriterion.grade) + str(type(businesscriterion.grade)))
                                            if (businesscriterion.grade == None): 
                                                logger.warning("Business criterions has no grade, removing it from the list : " + businesscriterion.name)
                                            else:
                                                listbusinesscriteria.append(businesscriterion)
                                json_bc_grades = None

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
                            json_snapshots = rest_service.get_application_snapshots(domain, applicationid)
                            if json_snapshots != None:
                                for snap in json_snapshots:
                                    snapHref = ''
                                    snapshotid = -1
                                    try:
                                        snapHref = snap['href']
                                    except KeyError:
                                        pass                             
                                    hrefsplit = snapHref.split('/')
                                    for elem in hrefsplit:
                                        # the last element is the id
                                        snapshotid = elem
    
                                    snapshotversion = snap['annotation']['version']
                                    snapshotdate =  snap['annotation']['date']['isoDate']    
                                    logger.info("    Snapshot " + snapHref + '#' + snapshotid)
                                    ###################################################################
                                    listmetrics = []
                                    listtechnicalcriteria = []
                                    listbccontributions = []
                                    listtccontributions = []
                                    dictapsummary = {}
                                    
                                    try:
                                        tqiqm = {}
                                        if not loaddata:
                                            logger.info("NOT Extracting the snapshot quality model")                                           
                                        else:
                                            json_snapshot_quality_model = rest_service.get_snapshot_tqi_quality_model(domain, snapshotid)
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
                                                qrcrt = qmitem['critical']
                                                tqiqm[qrid] = {"critical":qrcrt,"tc":{},"maxWeight":maxWeight,"compoundedWeight":qrcompoundWeight,"compoundedWeightFormula":qrcompoundWeightFormula}
                                                #contains the technical criteria (might be several) for each rule, we keep the fist one
                                                for tccont in qmitem['compoundedWeightTerms']:
                                                    term = tccont['term'] 
                                                    #tqiqm[qrid] = {tccont['technicalCriterion']['key']: tccont['technicalCriterion']['name']} 
                                                    #TODO: add qrcompoundWeight and/or qrcompoundWeightFormula
                                                    tqiqm.get(qrid).get("tc").update({tccont['technicalCriterion']['key']: tccont['technicalCriterion']['name']})                                        
                                        
                                        
                                        json_apsummary = None
                                        if loaddata:
                                            logger.info("Extracting the action plan summary")
                                            json_apsummary = rest_service.get_actionplan_summary(domain, applicationid, snapshotid)
                                        if json_apsummary != None:
                                            for qrap in json_apsummary:
                                                qrhref = qrap['rulePattern']['href']
                                                qrid = ''
                                                hrefsplit = qrhref.split('/')
                                                for elem in hrefsplit:
                                                    # the last element is the id
                                                    qrid = elem
                                                addedissues = 0
                                                pendingissues = 0
                                                try:
                                                    addedissues  = qrap['addedIssues']
                                                except KeyError:
                                                    logger.warning('Error in extracting the addedIssues')             
                                                try:
                                                    pendingissues  = qrap['pendingIssues']
                                                except KeyError:
                                                    logger.warning('Error in extracting the pendingIssues')                                                    
                                                numberofactions = addedissues + pendingissues
                                                dictapsummary.update({qrid:numberofactions})
                                        json_apsummary = None 
                                    except:
                                        LogUtils.logwarning(logger, 'Not able to extract the action plan summary ***',True)
                                    
                                    json_qr_results = None
                                    if loaddata:
                                        json_qr_results = rest_service.get_qualitymetrics_results(domain, applicationid, False, nbrows)
                                    if json_qr_results != None:
                                        for res in json_qr_results:
                                            iCount = 0
                                            lastProgressReported = None
                                            for res2 in res['applicationResults']:
                                                iCount += 1
                                                metricssize = len(res['applicationResults'])
                                                imetricprogress = int(100 * (iCount / metricssize))
                                                if imetricprogress in (9,19,29,39,49,59,69,79,89,99) : 
                                                    if lastProgressReported == None or lastProgressReported != imetricprogress:
                                                        LogUtils.loginfo(logger,  ' ' + str(imetricprogress+1) + '% of the metrics processed',True)
                                                        lastProgressReported = imetricprogress
                                                # for testing purpose, limit to the X first to optimize the testing time
                                                if loadonlyXmetrics and iCount > 10:
                                                    break
                                                    
                                                metric = Metric()
                                                try:
                                                    metric.type = res2['type']
                                                except KeyError:
                                                    None
                                                try:
                                                    metric.grade = res2['result']['grade']
                                                except KeyError:
                                                    None
                                                try:
                                                    metric.id = res2['reference']['key']
                                                except KeyError:
                                                    None
                                                try:
                                                    metric.name = res2['reference']['name']
                                                except KeyError:
                                                    None                                                    
                                                try:
                                                    metric.critical = res2['reference']['critical']
                                                except KeyError:
                                                    None
                                                try:
                                                    metric.failedchecks = res2['result']['violationRatio']['failedChecks']
                                                except KeyError:
                                                    None                                                          
                                                try:
                                                    metric.successfulchecks = res2['result']['violationRatio']['successfulChecks']
                                                except KeyError:
                                                    None                                                             
                                                try:
                                                    metric.totalchecks = res2['result']['violationRatio']['totalChecks']
                                                except KeyError:
                                                    totalChecks = None                                                         
                                                try:
                                                    metric.ratio = res2['result']['violationRatio']['ratio']
                                                except KeyError:
                                                    None                                                            
                                                try:
                                                    metric.addedviolations = res2['result']['evolutionSummary']['addedViolations']                                              
                                                except KeyError:
                                                    None   
                                                try:
                                                    metric.removedviolations = res2['result']['evolutionSummary']['removedViolations']
                                                except KeyError:
                                                    None
                                                if metric.type in ("quality-measures","quality-distributions","quality-rules"):
                                                    if (metric.grade == None): 
                                                        logger.warning("Metric has no grade, removing it from the list : " + metric.name)
                                                    else:
                                                        listmetrics.append(metric)
                                                        if metric.type == "quality-rules":
                                                            json_thresholds = None
                                                            if loaddata:
                                                                #LogUtils.loginfo(logger,'Extracting the quality rules thresholds',True)
                                                                json_thresholds = rest_service.get_qualityrules_thresholds(domain, snapshotid, metric.id)   
                                                            if json_thresholds != None and json_thresholds['thresholds'] != None:
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
                                                        logger.warning("Technical criterion has no grade, removing it from the list : " + metric.name)
                                                    else:
                                                        listtechnicalcriteria.append(metric)
                                                #logger.debug(metric.id + ":" + str(metric.type) + ":" + str(metric.grade))
                                                metric = None
                                        logger.info('Extracting the technical criteria contributors')
                                        for tciterator in listtechnicalcriteria:
                                            json_metriccontributions = None
                                            if loaddata:
                                                json_metriccontributions = rest_service.get_metric_contributions(domain, tciterator.id, snapshotid)
                                            if json_metriccontributions != None:
                                                for contr in json_metriccontributions['gradeContributors']:
                                                    tccontribution = Contribution()
                                                    tccontribution.parentmetricname = json_metriccontributions['name']
                                                    tccontribution.parentmetricid = json_metriccontributions['key']
                                                    tccontribution.metricname = contr['name']
                                                    tccontribution.metricid = contr['key']
                                                    tccontribution.critical = contr['critical']
                                                    tccontribution.weight = contr['weight']
                                                    # add only the one that have results
                                                    listtccontributions.append(tccontribution)
                                            json_metriccontributions = None
                                        logger.info('Extracting the business criteria contributors')
                                        for bcid in bcids:
                                            json_metriccontributions = None
                                            if loaddata:                                            
                                                json_metriccontributions = rest_service.get_metric_contributions(domain, bcid, snapshotid)
                                            if json_metriccontributions != None:
                                                for contr in json_metriccontributions['gradeContributors']:
                                                    bccontribution = Contribution()
                                                    bccontribution.parentmetricname = json_metriccontributions['name']
                                                    bccontribution.parentmetricid = json_metriccontributions['key']
                                                    bccontribution.metricname = contr['name']
                                                    bccontribution.metricid = contr['key']
                                                    bccontribution.critical = contr['critical']
                                                    bccontribution.weight = contr['weight']
                                                    # we add only the technical criteria that have results in the contribution list
                                                    bfound = False
                                                    for tc in listtechnicalcriteria:
                                                        if tc.id == bccontribution.metricid:
                                                            bfound = True
                                                            break 
                                                    if bfound:
                                                        listbccontributions.append(bccontribution)
                                            json_metriccontributions = None
                                    
                                    json_violations = None
                                    listviolations = []
                                    ''' loaddata and ''' 
                                    if loadviolations: 
                                        LogUtils.loginfo(logger,'Extracting violations',True)
                                        LogUtils.loginfo(logger,'Loading violations & components data from the REST API',True)
                                        json_violations = rest_service.get_snapshot_violations(domain, applicationid, snapshotid, criticalrulesonlyfilter, None, businesscriterionfilter, technofilter, nbrows)                                            
                                    if json_violations != None:
                                        iCouterRestAPIViolations = 0
                                        lastProgressReported = None
                                        for violation in json_violations:
                                            objviol = Violation()
                                            iCouterRestAPIViolations += 1
                                            currentviolurl = ''
                                            violations_size = len(json_violations)
                                            imetricprogress = int(100 * (iCouterRestAPIViolations / violations_size))
                                            if iCouterRestAPIViolations==1 or iCouterRestAPIViolations==violations_size or iCouterRestAPIViolations%500 == 0:
                                                LogUtils.loginfo(logger,"processing violation " + str(iCouterRestAPIViolations) + "/" + str(violations_size)  + ' (' + str(imetricprogress) + '%)',True)
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
                                                LogUtils.logwarning(logger, 'Could not find the critical contribution for %s'% str(objviol.qrid), True)
                                                
                                            # filter on quality rule id or name, if the filter match
                                            if qridfilter != None and not re.match(qridfilter, str(qrid)):
                                                continue
                                            if qrnamefilter != None and not re.match(qrnamefilter, qrname):
                                                continue
                                            actionPlan = violation['remedialAction']
                                            try:               
                                                objviol.hasActionPlan = actionPlan != None
                                            except KeyError:
                                                logger.warning('Not able to extract the action plan')
                                            if objviol.hasActionPlan:
                                                try:               
                                                    objviol.actionplanstatus = actionPlan['status']
                                                    objviol.actionplantag = actionPlan['tag']
                                                    objviol.actionplancomment = actionPlan['comment']
                                                except KeyError:
                                                    logger.warning('Not able to extract the action plan details')
                                            try:                                    
                                                objviol.hasExclusionRequest = violation['exclusionRequest'] != None
                                            except KeyError:
                                                logger.warning('Not able to extract the exclusion request')
                                            # filter the violations already in the exclusion list 
                                            try:                                    
                                                objviol.violationstatus = violation['diagnosis']['status']
                                            except KeyError:
                                                logger.warning('Not able to extract the violation status')
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
                                                logger.warning('Not able to extract the componentShortName')                                     
                                            try:
                                                objviol.componentNameLocation = violation['component']['name']
                                            except KeyError:
                                                logger.warning('Not able to extract the componentNameLocation')
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
                                            currentviolfullurl = edurl + '/engineering/index.html#' + snapHref 
                                            currentviolfullurl += '/business/60017/qualityInvestigation/0/60017/' 
                                            currentviolfullurl += firsttechnicalcriterionid + '/' + objviol.qrid + '/' + objviol.componentid
                                            objviol.url = currentviolfullurl
                                    
                                            listviolations.append(objviol)
                                        
                                    # generated csv file if required                                    
                                    fpath = get_excelfilepath(outputfolder, appName)
                                    LogUtils.loginfo(logger,"Generating xlsx file " + fpath,True)
                                    generate_excelfile(logger, fpath, appName, snapshotversion, snapshotdate, loadviolations, listbusinesscriteria, listtechnicalcriteria, listbccontributions, listtccontributions, listmetrics, dictapsummary, dicremediationabacus, listviolations, broundgrades)
                                    
                                    json_qr_results = None
                                    # keep only last snapshot
                                    break
                                        
    except: # catch *all* exceptions
        tb = traceback.format_exc()
        #e = sys.exc_info()[0]
        LogUtils.logerror(logger, '  Error during the processing %s' % tb,True)

    LogUtils.loginfo(logger,'Done !',True)