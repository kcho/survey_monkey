# coding: utf-8
import pandas as pd
import os
import numpy as np
import openpyxl
import argparse
import textwrap
# Set pandas options
pd.set_option('max_rows',5000)
pd.set_option('max_columns',5000)

def surveyMonkeyToCCNC(args):

    template_df = load_template(args.surveyTemplate)
    real_df = excel_merge_rearrange(args.inputExcel)
    real_df['question name'] = template_df['question name']

    template_openpyxl = openpyxl.load_workbook(args.template, data_only=False)
    update_template_excel(template_openpyxl, real_df)
    template_openpyxl.save(args.output)

    print '=========='
    print 'Completed'
    print '=========='


def load_template(survey_monkey_template):
    '''
    Takes input of survey monkey template location.
    A row of question name should be in the excel file.
    '''
    # Load survey monkey template with question names
    template_df = pd.read_excel(survey_monkey_template).T.reset_index()
    template_df.columns = ['title','subtitle','question name','answers']
    template_df.ix[:8,'question name'] = 'personal_info'
    return template_df

def excel_merge_rearrange(excelList):
    '''
    Takes input of list of excel file locations.

    It extracts the last subject's data in the excel List,
    concatenating them to a pandas dataframe.
    Also, the 'title' column is decoded to unicode.
    '''

    all_sheets = pd.concat([pd.read_excel(x, index_col=range(9)) for x in excelList],
                           axis=1).reset_index()
    
    # Extracts the last subject before transposing
    real_df = all_sheets.ix[all_sheets.index[[0,-1]],:].T.reset_index()
    #real_df = all_sheets.T.reset_index()
    real_df.columns = ['title','subtitle','answers']

    real_df.title = real_df.title.str.replace('<[^>]+>','')
    real_df.title = real_df.title.map(unicode)
    return real_df


def update_template_excel(template_openpyxl, real_df):
    for num, question_name in enumerate(template_openpyxl.sheetnames[:-5]): # 뒤에서 5개를 제외한 템플레이트의 모든 sheet를 looping
        # 엑셀파일에 데이터가 바로 들어가는 것이 아니고,
        # 행 열, 4칸 3칸씩 띄우고 데이터가 들어감
        startRow = 5
        startCol = 4

        sheet = template_openpyxl.get_sheet_by_name(question_name) # sheet를 불러옴
        data = real_df.ix[real_df['question name'].str.strip()==question_name] # 서베이멍키에서 불러온 자료를 불러옴
        
        if question_name == u'기본정보':
            startRow -= 1
            startCol -= 1

            info_list = [u'응답자 ID', u'이름', 
                         u'사용자 정의 데이터', u'시작 날짜', 
                         u'IP 주소', u'컬렉터 ID', u'종료 날짜']

            for info, answer in [(x,str(real_df.ix[real_df.title==x,'answers'].get_values()[0])) for x in info_list]:
                if u'날짜' in info:
                    answer = answer[:10]
                try:
                    sheet.cell(row = startRow, column = startCol).value = answer
                except:
                    sheet.cell(row = startRow, column = startCol).value = ''
                startRow+=1


        elif question_name == '(H, I) SCL, EF': #이런 이름을 가진 sheet는
            for num,answer in enumerate(data.answers): #나머지를 프린트
                doubleCellWrite(sheet.cell, startRow, startCol, answer)
                if num==28:
                    startRow+=5 #세칸 띄우고
                else:
                    startRow+=1

        elif question_name == '(8) ELSQ': #이런 이름을 가진 sheet는
            # 3, 4, 3, 4가 반복되는 리스트를 만듬 (column이 3,4 순으로 들어감)
            columns = [startCol, startCol+1] * (len(data.answers)/2) 
            # make double numbered row list
            rows = sorted(range(startRow, len(data.answers)) * 2)

            for num, (row, column, answer) in enumerate(zip(rows, columns, data.answers)):
                sheet.cell(row=row, column=column).value = answer#데이터 입력

        else:
            for answer in data.answers:# 서베이멍키에서 불러온 각 문제의 자료들을 looping
                doubleCellWrite(sheet.cell, startRow, startCol, answer)
                startRow +=1

def doubleCellWrite(cell, row, column, answer):
    cell(row = row, column = column).value = answer
    cell(row = row, column = column+1).value = answer


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=textwrap.dedent('''\
            {codeName} : Saves the Survey monkey data to CCNC format
            ========================================
            '''.format(codeName=os.path.basename(__file__))))

    parser.add_argument(
        '-i', '--inputExcel',
        help='Survey Monkey exported data files',
        nargs='+',
        default=os.getcwd())

    parser.add_argument(
        '-g', '--group',
        help='Group fo the subject',
        )

    parser.add_argument(
        '-t', '--template',
        help='Excel template with all formulas',
        default = '/ccnc_bin/survery_monkey/template.xlsx')

    parser.add_argument(
        '-gt', '--surveyTemplate',
        help='Survey monkey template with question name column',
        default = '/ccnc_bin/survery_monkey/merged.xlsx')

    parser.add_argument(
        '-o', '--output',
        help='Excel ouput',
        default = 'out.xlsx')

    args = parser.parse_args()


    surveyMonkeyToCCNC(args)
