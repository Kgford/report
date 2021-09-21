from django.shortcuts import render
import xlwt
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from openpyxl import Workbook
from openpyxl import load_workbook
from django.http import HttpResponse
from django.contrib.auth.models import User
import os
from django import forms
from django.views import View
from report.overhead import TimeCode, Security, StringThings,Conversions
from report.reports import ExcelReports
from django.shortcuts import render
from django.http import HttpResponseRedirect
from django.http import JsonResponse
from django.core import serializers
from django.core.files import File
#from .forms import QAForm
from django.urls import reverse, reverse_lazy
from django.db.models import Q
from django.utils.timezone import make_aware
import re
import datetime
import io
from io import BytesIO
import shutil
import getpass
import subprocess
import sys
from base64 import *
from test_db.models import Specifications,Workstation,Workstation1,Testdata,Testdata3,Trace,Tracepoints,Tracepoints2,Effeciency




class ReportView(View):
    #~~~~~~~~~~~Load Item database from csv. must put this somewhere else later"
    contSuccess = 0
    
    template_name = "index.html"
    success_url = reverse_lazy('excel:reports')
    def get(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            status ='In Process'
            job_num=-1
            part_num=-1
            workstation=-1
            operator=-1
            start_date=-1
            search=-1
            end_date = -1
            
            job_list = []
            part_list = []
            workstation_list = []
            operator_list = []
           
            
            #  Equations to get today - days
            #~~~~~~~~~~~~~ Time ~~~~~~~~~~~~~~~~~
            days=30 # start_date is today - days 
            time_code = TimeCode(days)
            friday = time_code.friday()
            print('friday=',friday)
            today = datetime.datetime.today()
            today = make_aware(today)
            print('today =', today)
            #start_date  = time_code.today_minus() # end_date is today - days 
            #start_date = make_aware(start_date)
            #end_date = today
            print('start_date =',start_date)
            print('end_date =',end_date)
            year = time_code.this_year()
            month_num = time_code.this_month()
            month_string = time_code.month_string()
            day = time_code.this_day()
            hour = time_code.this_hour()
            minute = time_code.this_minute()
            sec = time_code.this_sec()
            print('Today=',day,'/',month_num,'/',year,'/ ',hour,':',minute,':',sec)
            print('Month=',month_string)
            #~~~~~~~~~~~~~ Time ~~~~~~~~~~~~~~~~~
            
            
            workstation_list = Workstation.objects.using('TEST').order_by('computername').values_list('computername', flat=True).distinct()
            operator_list = Effeciency.objects.using('TEST').order_by('operator').values_list('operator', flat=True).distinct()
            job_list = Testdata.objects.using('TEST').order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
            part_list = Testdata.objects.using('TEST').order_by('partnumber').values_list('partnumber', flat=True).distinct()
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e)     
        return render (self.request,"excel/index.html",{'job_num':job_num,'part_num':part_num,'workstation':workstation,'operator':operator,'start_date':start_date,'end_date':end_date,
                                                        'job_list':job_list,'part_list':part_list,'workstation_list':workstation_list,'operator_list':operator_list})    
    
    def post(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            status ='In Process'
            job_num=-1
            part_num=-1
            workstation=-1
            operator=-1
            start_date=-1
            search=-1
            end_date = -1
            
            job_list = []
            part_list = []
            workstation_list = []
            operator_list = []
            stat_list = []
           
            
            #  Equations to get today - days
            #~~~~~~~~~~~~~ Time ~~~~~~~~~~~~~~~~~
            days=30 # start_date is today - days 
            time_code = TimeCode(days)
            friday = time_code.friday()
            print('friday=',friday)
            today = datetime.datetime.today()
            today = make_aware(today)
            print('today =', today)
            #start_date  = time_code.today_minus() # end_date is today - days 
            #start_date = make_aware(start_date)
            #end_date = today
            print('start_date =',start_date)
            print('end_date =',end_date)
            year = time_code.this_year()
            month_num = time_code.this_month()
            month_string = time_code.month_string()
            day = time_code.this_day()
            hour = time_code.this_hour()
            minute = time_code.this_minute()
            sec = time_code.this_sec()
            print('Today=',day,'/',month_num,'/',year,'/ ',hour,':',minute,':',sec)
            print('Month=',month_string)
            #~~~~~~~~~~~~~ Time ~~~~~~~~~~~~~~~~~
            
            #~~~~~~~~~~Get Post Values~~~~~~~~~~~~~~~
            job_num = request.POST.get('_job', -1)
            print('job_num=',job_num)
            part_num = request.POST.get('_part', -1)
            workstation = request.POST.get('_workstation', -1)
            operator = request.POST.get('_operator', -1)
            start_date = request.POST.get('_start_date', -1)
            end_date = request.POST.get('_end_date', -1)
            report = request.POST.get('_report', -1)
            print('report=',report)
            analyze = request.POST.get('_analyze', -1)
            #~~~~~~~~~~Get Post Values~~~~~~~~~~~~~~~
            job_num = '38783-01'
            #https://openpyxl.readthedocs.io/en/stable/
            #https://www.softwaretestinghelp.com/python-openpyxl-tutorial/
            if job_num != -1 and report !=-1:
                reporting = ExcelReports(job_num,operator,workstation)
                spec_data = Specifications.objects.using('TEST').filter(jobnumber=job_num).first()
                if '90 degree coupler' in spec_data.spectype.lower():
                    reporting.coupler_90_deg()
            elif job_num != -1:
                job_list = Testdata.objects.using('TEST').filter(jobnumber=job_num).order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
                part_list = Testdata.objects.using('TEST').filter(jobnumber=job_num).order_by('partnumber').values_list('partnumber', flat=True).distinct()
                report_data = Testdata.objects.using('TEST').filter(jobnumber=job_num).all()
            else:
                job_list = Testdata.objects.using('TEST').order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
                part_list = Testdata.objects.using('TEST').order_by('partnumber').values_list('partnumber', flat=True).distinct()
            
            workstation_list = Workstation.objects.using('TEST').order_by('workstationname').values_list('workstationname', flat=True).distinct()
            operator_list = Workstation.objects.using('TEST').order_by('operator').values_list('operator', flat=True).distinct()
            
            
            
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e)     
        return render (self.request,"excel/index.html",{'job_num':job_num,'part_num':part_num,'workstation':workstation,'operator':operator,'start_date':start_date,'end_date':end_date,
                                                        'job_list':job_list,'part_list':part_list,'workstation_list':workstation_list,'operator_list':operator_list,'stat_list':stat_list}) 



def export_users_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="users.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Users Data') # this will make a sheet named Users Data

    # Sheet header, first row
    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    columns = ['Username', 'First Name', 'Last Name', 'Email Address', ]

    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style) # at 0 row 0 column 

    # Sheet body, remaining rows
    font_style = xlwt.XFStyle()

    rows = User.objects.all().values_list('username', 'first_name', 'last_name', 'email')
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    wb.save(response)

    return response
    
#This code will explain how to Style your Excel File. The bellow code will explain Wrap text in the cell, background color, border, and text color.    
def export_styling_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="users.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Styling Data') # this will make a sheet named Users Data - First Sheet
    styles = dict(
        bold = 'font: bold 1',
        italic = 'font: italic 1',
        # Wrap text in the cell
        wrap_bold = 'font: bold 1; align: wrap 1;',
        # White text on a blue background
        reversed = 'pattern: pattern solid, fore_color blue; font: color white;',
        # Light orange checkered background
        light_orange_bg = 'pattern: pattern fine_dots, fore_color white, back_color orange;',
        # Heavy borders
        bordered = 'border: top thick, right thick, bottom thick, left thick;',
        # 16 pt red text
        big_red = 'font: height 320, color red;',
    )

    for idx, k in enumerate(sorted(styles)):
        style = xlwt.easyxf(styles[k])
        ws.write(idx, 0, k)
        ws.write(idx, 1, styles[k], style)

    wb.save(response)

    return response

#The below code will explain how to write data in Exisiting excel file and the content inside it.
def export_write_xls(request):
    # EG: path = excel_app/sample.xls
    path = os.path.dirname(__file__)
    path = path + '/excel_templates/'
    print('path=',path)
    file = os.path.join(path, 'TestData.xlsx')
    print('file=',file)

    wb = load_workbook(file)
    print('wb=',wb)
    sheet = wb["Raw Data1"]
    print('sheet=',sheet)
    sheet['F2'] = '398789-02' 
    sheet['F3'] = 'IPP-89348' 
    sheet['F4'] = '90 degree coupler' 
    print("sheet['F2']=",sheet['F2'])
    print(wb.sheetnames)
    wb.save("C:/ATE Data/demo4.xlsx")
    wb = load_workbook("C:/ATE Data/demo4.xlsx")
    return response
    
    

def test_report(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="users.xls"'

    # EG: path = report/excel_templates/TestData.xls
    path = os.path.dirname(__file__)
    path = path + '/excel_templates/'
    print('path=',path)
    file = os.path.join(path, 'TestData.xls')

    rb = open_workbook(file, formatting_info=True)
    sh = rb.sheet_by_name('Data')

   

    row_num = 2 # index start from 0
    rows = User.objects.all().values_list('username', 'first_name', 'last_name', 'email')
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num])
    
    # wb.save(file) # will replace original file
    # wb.save(file + '.out' + os.path.splitext(file)[-1]) # will save file where the excel file is
    wb.save(response)
    return response
   