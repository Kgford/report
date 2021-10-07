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
from report.reports import ExcelReports,Statistics,Histogram_data,XY_Chart,X_Range
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
import pygal
from pygal.style import Style
from report.pygal_extended import LineHist,LineBar
from test_db.models import Specifications,Workstation,Workstation1,Testdata,Testdata3,Trace,Tracepoints,Tracepoints2,Effeciency

import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import norm
import statistics
# Create a class extending Graph


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
            spectype = -1
            spec1 = -1
            spec2 = -1
            spec3 = -1
            spec4 = -1
            spec5 = -1
            report_data = -1
            analyze =-1
            artwork = -1
            stat1_min = -1
            stat1_max = -1
            stat1_avg = -1
            stat1_std = -1
            stat2_min = -1
            stat2_max = -1
            stat2_avg = -1
            stat2_std = -1
            stat3_min = -1
            stat3_max = -1
            stat3_avg = -1
            stat3_std = -1
            stat4_min = -1
            stat4_max = -1
            stat4_avg = -1
            stat4_std = -1
            stat5_min = -1
            stat5_max = -1
            stat5_avg = -1
            stat5_std = -1
            il_histo_data = -1
            rl_histo_data = -1
            iso_histo_data = -1
            ab_histo_data = -1
            pb_histo_data = -1
            coup_histo_data = -1
            dir_histo_data = -1
            cb_histo_data = -1
            
            job_list = []
            part_list = []
            workstation_list = []
            operator_list = []
            spec_list = -1
            test1_list = []
            test2_list = []
            test3_list = []
            test4_list = []
            test5_list = []
            artwork_list = ['RawData 1',]
            passed1 = 0
            failed1 = 0
            failed_percent1 = 0
            passed2 = 0
            failed2 = 0
            failed_percent2 = 0
            passed3 = 0
            failed3 = 0
            failed_percent3 = 0
            passed4 = 0
            failed4 = 0
            failed_percent4 = 0
            passed5 = 0
            failed5 = 0
            failed_percent5 = 0
           
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
        return render (self.request,"excel/index.html",{'job_num':job_num,'part_num':part_num,'workstation':workstation,'operator':operator,'start_date':start_date,'end_date':end_date,'artwork_list':artwork_list,'artwork':artwork,
                                                        'job_list':job_list,'part_list':part_list,'workstation_list':workstation_list,'operator_list':operator_list,'spec1':spec1,'spec2':spec1,'spec3':spec3,'spectype':spectype,
                                                        'spec4':spec4,'spec5':spec5,'report_data':report_data,'test1_list':test1_list,'test2_list':test2_list,'test3_list':test3_list,'test4_list':test4_list,'test5_list':test5_list,
                                                        'stat1_min':stat1_min,'stat1_max':stat1_max,'stat1_avg':stat1_avg,'stat1_std':stat1_std,'stat2_min':stat2_min,'stat2_max':stat2_max,'stat2_avg':stat2_avg,'stat2_std':stat2_std,
                                                        'stat3_min':stat3_min,'stat3_max':stat3_max,'stat3_avg':stat3_avg,'stat3_std':stat3_std,'stat4_min':stat4_min,'stat4_max':stat4_max,'stat4_avg':stat4_avg,'stat4_std':stat4_std,
                                                        'stat5_min':stat3_min,'stat5_max':stat5_max,'stat5_avg':stat5_avg,'stat5_std':stat5_std,'analyze':analyze,'il_histo_data':il_histo_data,'rl_histo_data':rl_histo_data,
                                                        'iso_histo_data':iso_histo_data,'ab_histo_data':ab_histo_data,'pb_histo_data':pb_histo_data,'coup_histo_data':coup_histo_data,'iso_histo_data':iso_histo_data,'cb_histo_data':cb_histo_data,
                                                        'passed1':passed1,'failed1':failed1,'failed_percent1':failed_percent1,'passed2':passed2,'failed2':failed2,'failed_percent2':failed_percent2,'passed3':passed3,'failed3':failed3,'failed_percent3':failed_percent3,    
                                                        'passed4':passed4,'failed4':failed4,'failed_percent4':failed_percent4,'passed5':passed5,'failed5':failed5,'failed_percent5':failed_percent5})
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
            spectype = -1
            spec1 = -1
            spec2 = -1
            spec3 = -1
            spec4 = -1
            spec5 = -1
            artwork = -1
            report_data = -1
            stat1_min = -1
            stat1_max = -1
            stat1_avg = -1
            stat1_std = -1
            stat2_min = -1
            stat2_max = -1
            stat2_avg = -1
            stat2_std = -1
            stat3_min = -1
            stat3_max = -1
            stat3_avg = -1
            stat3_std = -1
            stat4_min = -1
            stat4_max = -1
            stat4_avg = -1
            stat4_std = -1
            stat5_min = -1
            stat5_max = -1
            stat5_avg = -1
            stat5_std = -1
            il_histo_data = -1
            rl_histo_data = -1
            iso_histo_data = -1
            ab_histo_data = -1
            pb_histo_data = -1
            coup_histo_data = -1
            dir_histo_data = -1
            cb_histo_data = -1
            il_histo_data2 = -1
            chart1 = -1
            chart2 = -1
            chart3 = -1
            chart4 = -1
            chart5 = -1
            
            job_list = []
            part_list = []
            workstation_list = []
            operator_list = []
            stat_list = []
            spec_list = -1
            artwork_list = ['RawData 1',]
            art_rev_list = []
            test1_list = []
            test2_list = []
            test3_list = []
            test4_list = []
            test4_list = []
            test5_list = []
            
            bad1_list = []
            bad2_list = []
            bad3_list = []
            bad4_list = []
            bad4_list = []
            bad5_list = []
           
            blank = 0
            total = 0
            bad_data1 = 0
            passed1 = 0
            failed1 = 0
            failed_percent1 = 0
            bad_data2 = 0
            passed2 = 0
            failed2 = 0
            failed_percent2 = 0
            bad_data3 = 0
            passed3 = 0
            failed3 = 0
            failed_percent3 = 0
            bad_data4 = 0
            passed4 = 0
            failed4 = 0
            failed_percent4 = 0
            bad_data5 = 0
            passed5 = 0
            failed5 = 0
            failed_percent5 = 0
                
            
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
            print('part_num=',part_num)
            workstation = request.POST.get('_workstation', -1)
            operator = request.POST.get('_operator', -1)
            start_date = request.POST.get('_start_date', -1)
            end_date = request.POST.get('_end_date', -1)
            report = request.POST.get('_report', -1)
            print('report123=',report)
            analyze = request.POST.get('_analyze', -1)
            trace = request.POST.get('_trace', -1)
            artwork = request.POST.get('_art', -1)
            
            if analyze != -1 :
                analyze = 1
            print('analyze=',analyze)
            spec1 = request.POST.get('_spec1', -1)
            spec2 = request.POST.get('_spec2', -1)
            spec3 = request.POST.get('_spec3', -1)
            spec4 = request.POST.get('_spec4', -1)
            spec5 = request.POST.get('_spec5', -1)
            print('spec1=',spec1,'spec2=',spec2,'spec3=',spec3,'spec4=',spec4,'spec5=',spec5)
            if spec1!=-1:
                spec1 = float(spec1)
                spec2 = float(spec2)
                spec3 = float(spec3)
                spec4 = float(spec4)
                spec5 = float(spec5)
            
            #~~~~~~~~~~Get Post Values~~~~~~~~~~~~~~~
            #https://openpyxl.readthedocs.io/en/stable/
            #https://www.softwaretestinghelp.com/python-openpyxl-tutorial/
            if job_num !=-1 and report !=-1:
                reporting = ExcelReports(job_num,operator,workstation)
                spec_data = Specifications.objects.using('TEST').filter(jobnumber=job_num).first()
                job_list = Testdata.objects.using('TEST').order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
                if '90 degree coupler' in spec_data.spectype.lower():
                    reporting.test_data()
            elif job_num !=-1:
                print('shit')
                job_list = Testdata.objects.using('TEST').order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
                part_list = Testdata.objects.using('TEST').filter(jobnumber=job_num).order_by('partnumber').values_list('partnumber', flat=True).distinct()
                artwork_list = Testdata.objects.using('TEST').filter(jobnumber=job_num).order_by('partnumber').values_list('artwork_rev', flat=True).distinct()
                #filter blanks
                temp_list = []
                for artwork_rev in artwork_list:
                    if not artwork_rev == '':
                        temp_list.append(artwork_rev)
                artwork_list = temp_list
                if artwork_list:
                    artwork = artwork_list[0]
                #print('artwork_list=',artwork_list)
                spec_data = Specifications.objects.using('TEST').filter(jobnumber=job_num).first()
                print('spec_data.vswr=',spec_data.vswr)
                conversions = Conversions(spec_data.vswr,'')
                spec_rl = round(conversions.vswr_to_rl(),3)
                print('spec_rl=',spec_rl)
                spectype = spec_data.spectype
                if spec1==-1:
                    if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                        spec1 = round(spec_data.insertionloss,3)
                        spec2 = spec_rl
                        spec3 = round(spec_data.isolation,3)
                        spec4 = round(spec_data.amplitudebalance,3)
                        spec5 = round(spec_data.phasebalance,3)
                    elif 'DIRECTIONAL COUPLER' in spectype: 
                        spec1 = round(spec_data.insertionloss,3)
                        spec2 = spec_rl
                        spec3 = round(spec_data.coupling,3)
                        spec4 = round(spec_data.directivity,3)
                        spec5 = round(spec_data.coupledflatness,3)
                    
                report_data = Testdata.objects.using('TEST').filter(jobnumber=job_num).all()
                part_num = report_data[0].partnumber
                workstation = report_data[0].workstation
                operator = report_data[0].operator
                print('part_num=',part_num)
                #print('report_data=',report_data)
                
                total=0
                temp_list = []
                for data in report_data:
                    if data.serialnumber[3] == " ":
                        temp_list.append(data)
                        art_rev_list.append(data.artwork_rev)
                        test1_list.append(data.insertionloss)
                        test2_list.append(data.returnloss)
                        if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                            test3_list.append(data.isolation)
                            test4_list.append(data.amplitudebalance)
                            test5_list.append(data.phasebalance)
                        else:
                            test3_list.append(data.coupling)
                            test4_list.append(data.directivity)
                            test5_list.append(data.coupledflatness)
                    else:
                        blank+=1
                    total+=1
                        
                if len(bad1_list)>2:
                    bad_data1 = round(statistics.mean(bad1_list),2)
                if len(bad2_list)>2:
                    bad_data2 = round(statistics.mean(bad2_list),2)
                if len(bad3_list)>2:
                    bad_data3 = round(statistics.mean(bad3_list),2)
                if len(bad4_list)>2:
                    bad_data4 = round(statistics.mean(bad4_list),2)
                if len(bad5_list)>2:
                    bad_data5 = round(statistics.mean(bad5_list),2)
                
                
                #print('art_rev_list=',art_rev_list)
                report_data = temp_list 
                
                if len(test1_list) > 1:# must have at least two tests
                    test_list = [test1_list,test2_list,test3_list,test4_list,test5_list]
                    spec_list = [spec1,spec2,spec3,spec4,spec5]
                    histo_data = Histogram_data(test_list,spec_list,'test1') 
                    il_histo_data = histo_data.Hist_data()
                    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~statistics~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    stat_data = Statistics(test1_list,test2_list,test3_list,test4_list,test5_list) 
                    stat_list = stat_data.get_stats()
                    #print('stat_list=',stat_list)
                    stat1_min = stat_list[0][0]
                    print('stat1_min=',stat1_min)
                    stat1_max = stat_list[0][1]
                    print('stat1_max=',stat1_max)
                    stat1_avg = stat_list[0][2]
                    print('stat1_avg=',stat1_avg)
                    stat1_std = stat_list[0][3]
                    print('stat1_std=',stat1_std)
                    stat2_min = stat_list[1][0]
                    stat2_max = stat_list[1][1]
                    stat2_avg = stat_list[1][2]
                    stat2_std = stat_list[1][3]
                    stat3_min = stat_list[2][0]
                    stat3_max = stat_list[2][1]
                    stat3_avg = stat_list[2][2]
                    stat3_std = stat_list[2][3]
                    stat4_min = stat_list[3][0]
                    stat4_max = stat_list[3][1]
                    stat4_avg = stat_list[3][2]
                    stat4_std = stat_list[3][3]
                    stat5_min = stat_list[4][0]
                    stat5_max = stat_list[4][1]
                    stat5_avg = stat_list[4][2]
                    stat5_std = stat_list[4][3]
                    for x in range(len(test1_list)):
                        if test1_list[x] <= spec1:
                            passed1 +=1
                        else:
                            failed1 +=1
                            
                        if test2_list[x] <= spec2:
                            passed2 +=1
                        else:
                            failed2 +=1
                        if test3_list[x] <= spec3:
                            passed3 +=1
                        else:
                            failed3 +=1
                        if test4_list[x] <= spec4:
                            passed4 +=1
                        else:
                            failed4 +=1
                        if test5_list[x] <= spec5:
                            passed5 +=1
                        else:
                            failed5 +=1
                            
                    if passed1==0:
                        failed_percent1 = '100%'
                    elif failed1==0:
                        failed_percent1 = '0%'
                    else:    
                        failed_percent1 = str(round((failed1/passed1)* 100,3)) + '%'
                    
                    if passed2==0:
                        failed_percent2 = '100%'
                    elif failed2==0:
                        failed_percent2 = '0%'
                    else:    
                        failed_percent2 = str(round((failed2/passed2)* 100,3)) + '%'
                    
                    if passed3==0:
                        failed_percent3 = '100%'
                    elif failed3==0:
                        failed_percent3 = '0%'
                    else:    
                        failed_percent3 = str(round((failed3/passed3)* 100,3)) + '%'
                    
                    if passed4==0:
                        failed_percent4 = '100%'
                    elif failed4==0:
                        failed_percent4 = '0%'
                    else:    
                        failed_percent4 = str(round((failed4/passed4)* 100,3)) + '%'
                    
                    if passed5==0:
                        failed_percent5 = '100%'
                    elif failed5==0:
                        failed_percent5 = '0%'
                    else:    
                        failed_percent5 = str(round((failed5/passed5)* 100,2)) + '%'
                    
                    mean = statistics.mean(test3_list)
                    sd = statistics.stdev(test3_list)
                    print('mean=',mean)
                    print('sd=',sd)   
                     
                                    
                    #hist.add("SD", [10, 20, 39, 40, 30, 20,10], plotas='line', secondary=True)
                    #data_max = max(il_histo_data)
                    #hist.add('', [ x[3] for x in il_histo_data_line], plotas='line', secondary=True)
                    #hist.add('IL', il_histo_data_line, plotas='bar')
                    #hist.x_labels = x_range_list
                    #hist.title = 'IL Histogram' 
                    
                    histo_data = Histogram_data(test_list,spec_list,'test1')
                    il_histo = histo_data.Hist_data()
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                    x_range_list = []
                    sd_x_range = []
                    x_range = X_Range(il_histo,stat1_min,stat1_max)                
                    x_range_list=x_range.list()
                    sd_x_range =x_range.list_expanded()
                    print('x_range_list=',x_range_list)
                    print('sd_x_range=',sd_x_range)
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                    
                    #~~~~~~~~~~make standard deviation line~~~~~~~~~~~~~
                    sd_list = []
                    for i in range(len(sd_x_range)):
                        item = spec1- abs(sd-sd_x_range[i])
                        if item <0:
                            item=0
                        
                        sd_list.append(item)
                    print('sd_list=',sd_list)
                    xy_data = []
                    x=0
                    for test in sd_list:
                        xy_data.append((sd_x_range[x],test))
                        x+=1
                    #~~~~~~~~~~make standard deviation line~~~~~~~~~~~~~
                    
                    
                    #print('rl_histo=',rl_histo)
                    custom_style = Style(colors=('#991593','#201599'),title_font_size=39, label_font_size=15)
                    hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                    hist.x_labels = x_range_list
                    hist.add('Insertion Loss', il_histo)
                    hist.title = 'IL Histogram' 
                    il_histo_data = hist.render_data_uri()
                    #print('il_histo_data=',il_histo_data)
                    
                    histo_data = Histogram_data(test_list,spec_list,'test1')
                    il_histo = histo_data.Hist_data()
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                    x_range_list = []
                    x_range = X_Range(il_histo,stat1_min,stat1_max)                
                    x_range_list=x_range.list()
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                                       
                     #print('chart_data=',chart_data)
                    custom_style = Style(colors=('#991593','#201599'),title_font_size=39, label_font_size=15)
                    xy_chart = pygal.XY(style=custom_style, background= 'transparent')
                    xy_chart.title = 'Insertion Loss'
                    xy_chart.add('IL Histogram', xy_data)
                    xy_chart.add('spec', [(spec1, 0), (spec1, 0.5)])
                    il_histo_data2 = xy_chart.render_data_uri()
                    
                    
                    charts = XY_Chart(test_list,spec_list,'test1') 
                    chart_data = charts.Chart_data()
                    #print('chart_data=',chart_data)
                    custom_style = Style(colors=('#991593','#201599'),title_font_size=39, label_font_size=15)
                    xy_chart = pygal.XY(style=custom_style, human_readable=True)
                    xy_chart.title = 'IL XY Plot'
                    xy_chart.add('Insertion Loss', chart_data)
                    xy_chart.add('spec', [(0, spec1), (len(chart_data), spec1)])
                    chart1 = xy_chart.render_data_uri()
                    
                    
                    histo_data = Histogram_data(test_list,spec_list,'test2')
                    rl_histo = histo_data.Hist_data()
                    
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                    x_range_list = []
                    x_range = X_Range(rl_histo,stat2_min,stat2_max)                
                    x_range_list=x_range.list()
                    #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                    #print('rl_histo=',rl_histo)
                    custom_style = Style(colors=('#47ff7b','#201599'),title_font_size=39, label_font_size=15)
                    hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                    hist.x_labels = x_range_list
                    hist.add('Return Loss', rl_histo)
                    hist.title = 'RL Histogram' 
                    rl_histo_data = hist.render_data_uri()
                    
                    charts = XY_Chart(test_list,spec_list,'test2') 
                    chart_data = charts.Chart_data()
                    #print('chart_data=',chart_data)
                    custom_style = Style(colors=('#47ff7b','#201599'),title_font_size=39, label_font_size=15)
                    xy_chart = pygal.XY(style=custom_style, human_readable=True)
                    xy_chart.title = 'RL XY Plot'
                    xy_chart.add('Return Loss', chart_data)
                    xy_chart.add('spec', [(0, spec2), (len(chart_data), spec2)])
                    chart2 = xy_chart.render_data_uri()
                    
                    #print('rl_histo_data=',rl_histo_data)
                    if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                        histo_data = Histogram_data(test_list,spec_list,'test3')
                        iso_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(iso_histo,stat3_min,stat3_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('iso_histo=',iso_histo)
                        custom_style = Style(colors=('#ffd138','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Isolation', iso_histo)
                        hist.title = 'ISO Histogram' 
                        iso_histo_data = hist.render_data_uri()
                        #print('iso_histo_data=',iso_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test3') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#ffd138','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style, human_readable=True)
                        xy_chart.title = 'ISO XY Plot'
                        xy_chart.add('Isolation', chart_data)
                        xy_chart.add('spec', [(0, 0-spec3), (len(chart_data), 0-spec3)])
                        chart3 = xy_chart.render_data_uri()
                    
                        histo_data = Histogram_data(test_list,spec_list,'test4')
                        ab_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(ab_histo,stat4_min,stat4_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('ab_histo=',ab_histo)
                        custom_style = Style(colors=('#130fff','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Amplitude Balance', ab_histo)
                        hist.title = 'AB Histogram' 
                        ab_histo_data = hist.render_data_uri()
                        #print('il_histo_data=',il_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test4') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#130fff','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style, human_readable=True)
                        xy_chart.title = 'AB XY Plot'
                        xy_chart.add('Amplitude_Balance', chart_data)
                        xy_chart.add('spec +', [(0, spec4), (len(chart_data), spec4)])
                        xy_chart.add('spec -', [(0, 0-spec4), (len(chart_data), 0-spec4)])
                        chart4 = xy_chart.render_data_uri()
                    
                        histo_data = Histogram_data(test_list,spec_list,'test5')
                        pb_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(pb_histo,stat5_min,stat5_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('pb_histo=',pb_histo)
                        custom_style = Style(colors=('#ff2617','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Phase Balance', pb_histo)
                        hist.title = 'PB Histogram' 
                        pb_histo_data = hist.render_data_uri()
                        #print('pb_histo_data=',pb_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test5') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#ff2617','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style, human_readable=True)
                        xy_chart.title = 'PB XY Plot'
                        xy_chart.add('Phase Balance', chart_data)
                        xy_chart.add('spec +', [(0, spec5), (len(chart_data), spec5)])
                        xy_chart.add('spec -', [(0, 0-spec5), (len(chart_data), 0-spec5)])
                        chart5 = xy_chart.render_data_uri()
                    
                    else:
                        histo_data = Histogram_data(test_list,spec_list,'test3')
                        coup_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(coup_histo,stat3_min,stat3_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('coup_histo=',coup_histo)
                        custom_style = Style(colors=('#ffd138','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Coupling', coup_histo)
                        hist.title = 'Coupling Histogram' 
                        il_histo_data = hist.render_data_uri()
                        #print('il_histo_data=',il_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test3') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#ffd138','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style, human_readable=True)
                        xy_chart.title = 'Coupling XY Plot'
                        xy_chart.add('Coupling', chart_data)
                        xy_chart.add('spec', [(0, spec3), (len(chart_data), spec3)])
                        chart3 = xy_chart.render_data_uri()
                    
                        histo_data = Histogram_data(test_list,spec_list,'test4')
                        dir_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(dir_histo,stat4_min,stat4_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('dir_histo=',dir_histo)
                        custom_style = Style(colors=('#130fff','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Directivity', dir_histo)
                        hist.title = 'DIR Histogram' 
                        dir_histo_data = hist.render_data_uri()
                        #print('dir_histo_data=',dir_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test4') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#130fff','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style, human_readable=True)
                        xy_chart.title = 'DIR XY Plot'
                        xy_chart.add('Directivity', chart_data)
                        xy_chart.add('spec +', [(0, spec4), (len(chart_data), spec4)])
                        xy_chart.add('spec -', [(0, 0-spec4), (len(chart_data), 0-spec4)])
                        chart4 = xy_chart.render_data_uri()
                    
                        histo_data = Histogram_data(test_list,spec_list,'test5')
                        cb_histo = histo_data.Hist_data()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        x_range_list = []
                        x_range = X_Range(cb_histo,stat5_min,stat5_max)                
                        x_range_list=x_range.list()
                        #~~~~~~~~~~make plot x-range~~~~~~~~~~~~~
                        #print('cb_histo=',cb_histo)
                        custom_style = Style(colors=('#ff2617','#201599'),title_font_size=39, label_font_size=15)
                        hist = pygal.Histogram(fill=True,style=custom_style, human_readable=True)
                        hist.x_labels = x_range_list
                        hist.add('Coupling Balance', cb_histo)
                        hist.title = 'CB Histogram' 
                        cb_histo_data = hist.render_data_uri()
                        #print('cb_histo_data=',cb_histo_data)
                        
                        charts = XY_Chart(test_list,spec_list,'test5') 
                        chart_data = charts.Chart_data()
                        #print('chart_data=',chart_data)
                        custom_style = Style(colors=('#ff2617','#201599'),title_font_size=39, label_font_size=15)
                        xy_chart = pygal.XY(style=custom_style)
                        xy_chart.title = 'BB XY Plot'
                        xy_chart.add('Coupling Balance', chart_data, human_readable=True)
                        xy_chart.add('spec +', [(0, spec5), (len(chart_data), spec5)])
                        xy_chart.add('spec -', [(0, 0-spec5), (len(chart_data), 0-spec5)])
                        chart5 = xy_chart.render_data_uri()
                                    
            else:
                job_list = Testdata.objects.using('TEST').order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
                part_list = Testdata.objects.using('TEST').order_by('partnumber').values_list('partnumber', flat=True).distinct()
            
            workstation_list = Workstation.objects.using('TEST').order_by('workstationname').values_list('workstationname', flat=True).distinct()
            operator_list = Workstation.objects.using('TEST').order_by('operator').values_list('operator', flat=True).distinct()
            
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e)     
        return render (self.request,"excel/index.html",{'job_num':job_num,'part_num':part_num,'workstation':workstation,'operator':operator,'start_date':start_date,'end_date':end_date,'artwork_list':artwork_list,'artwork':artwork,
                                                        'job_list':job_list,'part_list':part_list,'workstation_list':workstation_list,'operator_list':operator_list,'spec1':spec1,'spec2':spec1,'spec3':spec3,'spectype':spectype,
                                                        'spec4':spec4,'spec5':spec5,'report_data':report_data,'test1_list':test1_list,'test2_list':test2_list,'test3_list':test3_list,'test4_list':test4_list,'test5_list':test5_list,
                                                        'stat1_min':stat1_min,'stat1_max':stat1_max,'stat1_avg':stat1_avg,'stat1_std':stat1_std,'stat2_min':stat2_min,'stat2_max':stat2_max,'stat2_avg':stat2_avg,'stat2_std':stat2_std,
                                                        'stat3_min':stat3_min,'stat3_max':stat3_max,'stat3_avg':stat3_avg,'stat3_std':stat3_std,'stat4_min':stat4_min,'stat4_max':stat4_max,'stat4_avg':stat4_avg,'stat4_std':stat4_std,
                                                        'stat5_min':stat3_min,'stat5_max':stat5_max,'stat5_avg':stat5_avg,'stat5_std':stat5_std,'analyze':analyze,'il_histo_data':il_histo_data,'rl_histo_data':rl_histo_data,
                                                        'iso_histo_data':iso_histo_data,'ab_histo_data':ab_histo_data,'pb_histo_data':pb_histo_data,'coup_histo_data':coup_histo_data,'iso_histo_data':iso_histo_data,
                                                        'passed1':passed1,'failed1':failed1,'failed_percent1':failed_percent1,'passed2':passed2,'failed2':failed2,'failed_percent2':failed_percent2,'passed3':passed3,'failed3':failed3,   
                                                        'cb_histo_data':cb_histo_data,'failed_percent3':failed_percent3,'passed4':passed4,'failed4':failed4,'failed_percent4':failed_percent4,'passed5':passed5,'failed5':failed5,
                                                        'failed_percent5':failed_percent5,'bad_data1':bad_data1,'bad_data2':bad_data2,'bad_data3':bad_data3,'bad_data4':bad_data4,'bad_data5':bad_data5,'il_histo_data2':il_histo_data2,
                                                        'chart1':chart1,'chart2':chart2,'chart3':chart3,'chart4':chart4,'chart5':chart5,'blank':blank,'total':total,'art_rev_list':art_rev_list,'blank':blank})


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
   