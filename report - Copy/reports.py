from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill#Connect cell styles
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Fill#Connect styles for text
from openpyxl.styles import colors#Connect colors for text and cells

from django.http import HttpResponse
from report.overhead import TimeCode, Security, StringThings,Conversions
from test_db.models import Specifications,Workstation,Workstation1,Testdata,Testdata3,Trace,Tracepoints,Tracepoints2,Effeciency,ReportQueue
import os
import statistics 


class ExcelReports:
    def __init__ (self, job_num,operator,workstation):
        self.job_num = job_num
        self.operator = operator
        self.workstation = workstation
        #print('job_num=',self.job_num)
               
    def test_data(self):
        job_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
        part_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('partnumber').values_list('partnumber', flat=True).distinct()
        artwork_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('partnumber').values_list('artwork_rev', flat=True).distinct()
        report_data = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).all()
        print('running report')
        ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='running report')
        
        
        part_num = report_data[0].partnumber
        spec_data = Specifications.objects.using('TEST').filter(jobnumber=self.job_num).first()
        spectype = spec_data.spectype
        paths = ReportFiles(self.job_num,part_num,spectype)
        data_path = paths.data_path()
        template_path = paths.template()
        #print('template_path=',template_path)
        
        wb = load_workbook(template_path)
        #print('wb=',wb)
        
        print('artwork_list1',artwork_list)
        if not artwork_list:
            artwork_list = ['RawData 1',]
        
        print('artwork_list2',artwork_list)
        #filter blanks
        temp_list = []
        for artwork_rev in artwork_list:
            if not artwork_rev == '':
                temp_list.append(artwork_rev)
        artwork_list = temp_list
        print('artwork_list3',artwork_list)
        if not artwork_list:
            artwork_list = ['RawData 1',]
        print('artwork_list4',artwork_list)
        # datasheet can only handle 5 artworks----for now---
        #print('len(artwork_list)=',len(artwork_list))
        if len(artwork_list) >5:
            group = 5
        else:
            group = len(artwork_list)
        
        remove_extra = DeleteSheets(group,wb)
        remove_extra.remove()
        
        x=1
        z=1
        print('loading data')
        ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='loading data')
        for artwork_rev in artwork_list:
            if not artwork_rev:
                artwork_rev='RawData 1'
                report_data = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).all()
                data_count = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).count()
            elif 'RawData 1' in artwork_rev:
                report_data = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).all()
                data_count = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).count()
            else:
                report_data = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).filter(artwork_rev=artwork_rev).all()
                data_count = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).filter(artwork_rev=artwork_rev).count()
            
            conversions = Conversions(spec_data.vswr,'')
            spec_rl = round(conversions.vswr_to_rl(),2)
            
            #print('spec_rl=',spec_rl)
            #print('spec_data=',spec_data)
            #print('report_data=',report_data)
            if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                spec_list = [spec_data.insertionloss,spec_rl,spec_data.isolation,spec_data.amplitudebalance,spec_data.phasebalance,spec_data.ab_ex] 
            elif 'DIRECTIONAL COUPLER' in spectype: 
                spec_list = [spec_data.insertionloss,spec_rl,spec_data.coupling,spec_data.directivity,spec_data.coupledflatness]
                
                
                
            #print('spec_list=',spec_list)
            if report_data:
                part_num = report_data[0].partnumber
                print('part_num=',part_num)
                spectype = spec_data.spectype
                
                activesheet = "Raw Data" + str(z)
                sheet = wb[activesheet]
                print('sheet=',sheet)
                sheet['B4'] = self.operator 
                sheet['B5'] = self.workstation 
                sheet['H2'] = self.job_num
                sheet['H3'] = part_num 
                sheet['H4'] = spectype 
                sheet['H5'] = artwork_rev 
                
                #~~~~~~~~~~~~configure  the tests~~~~~~~~~~~~~
                if 'DIRECTIONAL COUPLER' in spectype:
                    sheet['F6'] = "Coupling"
                    sheet['H6'] = "Directivity"
                    sheet['J6'] = "Coupling Flatness"
                elif 'BALUN' in spectype:
                    sheet['F6'] = "No Test"
                    sheet['H6'] = "Amplitude Balance"
                    sheet['J6'] = "Phase Balance"
                else:
                    sheet['F6'] = "Isolation"
                    sheet['H6'] = "Amplitude Balance"
                    sheet['J6'] = "Phase Balance"
                #~~~~~~~~~~~~choose the tests~~~~~~~~~~~~~
                
                #~~~~~~~~~~~~~~~~~~~~~~~~~~format the sheet for data~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                #Mearge split cells for normal data
                if spec_data.ab_exp_tf:  # Dual Band AB only Don't mearge AB cels
                    for x in range(int(data_count) + 1):
                        sheet.merge_cells(start_row=x+7, start_column=2, end_row=x+7, end_column=3) #IL
                        sheet.merge_cells(start_row=x+7, start_column=4, end_row=x+7, end_column=5) #RL
                        sheet.merge_cells(start_row=x+7, start_column=6, end_row=x+7, end_column=7)  #ISO/Coup
                        sheet.merge_cells(start_row=x+7, start_column=10, end_row=x+7, end_column=11) #PB/COUP Flat
                else:
                    for x in range(int(data_count) + 1):
                        sheet.merge_cells(start_row=x+7, start_column=2, end_row=x+7, end_column=3) #IL
                        sheet.merge_cells(start_row=x+7, start_column=4, end_row=x+7, end_column=5) #RL
                        sheet.merge_cells(start_row=x+7, start_column=6, end_row=x+7, end_column=7)  #ISO/Coup
                        sheet.merge_cells(start_row=x+7, start_column=8, end_row=x+7, end_column=9) #AB/DIR
                        sheet.merge_cells(start_row=x+7, start_column=10, end_row=x+7, end_column=11) #PB/COUP Flat    
                #~~~~~~~~~~~~~~~~~~~~~~~~~~format the sheet for data~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~Load the specs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                if ('90 DEGREE COUPLER' in spectype or 'BALUN' in spectype) and spec_data.ab_exp_tf:  # Dual Band AB only
                    sheet['B7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['D7'] = str(spec_rl) + ' Max'
                    sheet['F7'] = str(spec_data.isolation) + ' Max'
                    sheet['H7'] = "+/- " + str(spec_data.amplitudebalance) + ' dB'
                    sheet['I7'] = "+/- " + str(spec_data.SpecAB_exp) + ' dB'
                    sheet['J7'] = "+/- " + str(spec_data.phasebalance) + ' deg'
                    sheet['N7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['O7'] = str(spec_rl) + ' Max'
                    sheet['P7'] = str(spec_data.isolation) + ' Max'
                    sheet['Q7'] = "+/- " + str(spec_data.amplitudebalance) + ' dB'
                    sheet['R7'] = "+/- " + str(spec_data.phasebalance) + ' deg'
                elif '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                    sheet['B7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['D7'] = str(spec_rl) + ' Max'
                    sheet['F7'] = str(spec_data.isolation) + ' Max'
                    sheet['H7'] = "+/- " + str(spec_data.amplitudebalance) + ' dB'
                    sheet['J7'] = "+/- " + str(spec_data.phasebalance) + ' deg'
                    sheet['N7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['O7'] = str(spec_rl) + ' Max'
                    sheet['P7'] = str(spec_data.isolation) + ' Max'
                    sheet['Q7'] = "+/- " + str(spec_data.amplitudebalance) + ' dB'
                    sheet['R7'] = "+/- " + str(spec_data.phasebalance) + ' deg'
                elif 'DIRECTIONAL COUPLER' in spectype:
                    sheet['B7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['D7'] = str(spec_rl) + ' Max'
                    sheet['F7'] = str(spec_data.coupling) + ' Max'
                    sheet['H7'] = "+/- " + str(spec_data.directivity) + ' dB'
                    sheet['J7'] = "+/- " + str(spec_data.coupledflatness) + ' deg'
                    sheet['N7'] = str(spec_data.insertionloss) + ' Max'
                    sheet['O7'] = str(spec_rl) + ' Max'
                    sheet['P7'] = str(spec_data.coupling) + ' Max'
                    sheet['Q7'] = "+/- " + str(spec_data.directivity) + ' dB'
                    sheet['R7'] = "+/- " + str(spec_data.coupledflatness) + ' deg'

                #Tabular data
                rownum = 8
                insertion_loss = []
                return_loss = []
                isolation = []
                coupling = []
                amplitude_balance = []
                phase_balance = []
                directivity = []
                coupledflatness = []
                
                stat_list = []
                il_pass = 0
                rl_pass = 0
                iso_pass = 0
                ab_pass = 0
                pb_pass = 0
                coup_pass = 0
                dir_pass = 0
                cf_pass = 0                
                il_fail = 0
                rl_fail = 0
                iso_fail = 0
                ab_fail = 0
                pb_fail = 0
                coup_fail = 0
                dir_fail = 0
                cf_fail = 0
                uut = 1
                sum_row = 5
                #print('report_data=',report_data)
                for data in report_data:
                    if data.serialnumber[3] == " ":
                        sheet.cell(row=rownum, column=1).value= 'UUT ' + str(uut)
                        #print('data.serialnumber=',data.serialnumber)
                        sheet.cell(row=rownum, column=2).value= round(data.insertionloss,2)
                        testdata1 = sheet.cell(row=rownum, column=2)#Created a variable that contains cell
                        insertion_loss.append(data.insertionloss)
                        if data.insertionloss <= spec_list[0]:
                            il_pass+=1
                        else:
                            il_fail+=1
                            testdata1.font = Font(color='FF3342', bold=True, italic=True) #W
                        
                        sheet.cell(row=rownum, column=4).value= round(data.returnloss,2)
                        testdata2 = sheet.cell(row=rownum, column=4)#Created a variable that contains cell
                        return_loss.append(data.returnloss)
                        if data.returnloss <= spec_list[1]:
                            rl_pass+=1
                        else:
                            rl_fail+=1
                            testdata2.font = Font(color='FF3342', bold=True, italic=True) #W
                        
                        if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                            sheet.cell(row=rownum, column=6).value= round(data.isolation,2)
                            testdata3 = sheet.cell(row=rownum, column=6)#Created a variable that contains cell
                            isolation.append(data.isolation)
                            if data.isolation <= spec_list[2]:
                                iso_pass+=1
                            else:
                                iso_fail+=1
                                testdata3.font = Font(color='FF3342', bold=True, italic=True) #W
                            ##~~~~~~~~~~~~~~~~~~~~~~~~AB Dual Band ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
                            if spec_data.ab_exp_tf:
                                sheet.cell(row=rownum, column=8).value= round(data.amplitudebalance1,2)
                                sheet.cell(row=rownum, column=9).value= round(data.amplitudebalance2,2)
                                testdata4a = sheet.cell(row=rownum, column=8)#Created a variable that contains cell
                                testdata4b = sheet.cell(row=rownum, column=8)#Created a variable that contains cell
                                amplitude_balance.append(data.amplitudebalance)
                                if data.amplitudebalance1 <= spec_list[3] and data.amplitudebalance2 <= spec_list[5]:
                                    ab_pass+=1
                                else:
                                    ab_fail+=1
                                    if data.amplitudebalance1 > spec_list[3]:
                                        testdata4a.font = Font(color='FF3342', bold=True, italic=True) #W
                                    if data.amplitudebalance2 > spec_list[5]:
                                        testdata4b.font = Font(color='FF3342', bold=True, italic=True) #W
                            ##~~~~~~~~~~~~~~~~~~~~~~~~AB Dual Band ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  
                            else:
                                sheet.cell(row=rownum, column=8).value= round(data.amplitudebalance,2)
                                testdata4 = sheet.cell(row=rownum, column=8)#Created a variable that contains cell
                                amplitude_balance.append(data.amplitudebalance)
                                if data.amplitudebalance <= spec_list[3]:
                                    ab_pass+=1
                                else:
                                    ab_fail+=1
                                    testdata4.font = Font(color='FF3342', bold=True, italic=True) #W

                            sheet.cell(row=rownum, column=10).value= round(data.phasebalance,2)
                            testdata5 = sheet.cell(row=rownum, column=10)#Created a variable that contains cell
                            phase_balance.append(data.phasebalance)
                            if data.phasebalance <= spec_list[4]:
                                pb_pass+=1
                            else:
                                pb_fail+=1
                        else:
                            sheet.cell(row=rownum, column=6).value= round(data.coupling,2)
                            testdata3 = sheet.cell(row=rownum, column=6)#Created a variable that contains cell
                            coupling.append(data.coupling)
                            if data.coupling <= spec_list[2]:
                                coup_pass+=1
                            else:
                                coup_fail+=1
                                testdata3 = sheet.cell(row=rownum, column=6)#Created a variable that contains cell
                            
                            sheet.cell(row=rownum, column=8).value= round(data.directivity,2)
                            testdata4 = sheet.cell(row=rownum, column=8)#Created a variable that contains cell
                            directivity.append(data.directivity)
                            if data.directivity <= spec_list[3]:
                                dir_pass+=1
                            else:
                                dir_fail+=1
                                testdata4.font = Font(color='FF3342', bold=True, italic=True) #W
                            
                            sheet.cell(row=rownum, column=10).value= round(data.coupledflatness,2)
                            testdata5 = sheet.cell(row=rownum, column=10)#Created a variable that contains cell
                            coupledflatness.append(data.coupledflatness)
                            if data.coupledflatness <= spec_list[4]:
                                cf_pass+=1
                            else:
                                cf_fail+=1
                                testdata5.font = Font(color='FF3342', bold=True, italic=True) #W
                        rownum+=1
                        uut+=1
                    
                #~~~~~~~~~~~~~~~~Statics and Summary ~~~~~~~~~~~~~~~~~~~~
                if len(insertion_loss) > 1:# must have at least two tests
                    list_names = ['Min','Max','Avg','Stdev']
                    #print('insertion_loss=',insertion_loss)
                    il_stdev = round(statistics.stdev(insertion_loss),2) #Standard deviation
                    il_var = round(statistics.variance(insertion_loss),2) #Variance
                    il_avg = round(statistics.mean(insertion_loss),2) #Mean Average
                    il_min = round(min(insertion_loss),2) #Min
                    il_max = round(max(insertion_loss),2) #Max
                    sheet['N8'] = il_avg
                    sheet['N9'] = il_min
                    sheet['N10'] = il_max
                    sheet['N11'] = il_stdev
                    sheet['N12'] = il_pass
                    sheet['N13'] = il_fail
                    sheet['N14'] = round(il_fail/rownum,2)
                    il_list = [il_min,il_max,il_avg,il_stdev]
                    #print('il_list=',il_list)

                    #print('return_loss=',return_loss)
                    rl_stdev = round(statistics.stdev(return_loss),2) #Standard deviation
                    if len(return_loss)>1:
                        rl_var = round(statistics.variance(return_loss),2) #Variance
                    else:
                        rl_var = 0 #Variance
                    rl_var = round(statistics.variance(return_loss),2) #Variance
                    rl_avg = round(statistics.mean(return_loss),2) #Mean Average
                    rl_min = round(min(return_loss),2) #Min
                    rl_max = round(max(return_loss),2) #Max
                    rl_list = [rl_min,rl_max,rl_avg,rl_stdev]
                    sheet['O8'] = rl_avg
                    sheet['O9'] = rl_min
                    sheet['O10'] = rl_max
                    sheet['O11'] = rl_stdev
                    sheet['O12'] = rl_pass
                    sheet['O13'] = rl_fail
                    sheet['O14'] = round(rl_fail/rownum,2)
                    #print('rl_list=',rl_list)

                    if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                        iso_stdev = round(statistics.stdev(isolation),2) #Standard deviation
                        if len(isolation)>1:
                            iso_var = round(statistics.variance(isolation),2) #Variance
                        else:
                            iso_var = 0 #Variance
                        iso_avg = round(statistics.mean(isolation),2) #Mean Average
                        iso_min = round(min(isolation),2) #Min
                        iso_max = round(max(isolation),2) #Max
                        iso_list = [iso_min,iso_max,iso_avg,iso_stdev]
                        sheet['P8'] = iso_avg
                        sheet['P9'] = iso_min
                        sheet['P10'] = iso_max
                        sheet['P11'] = iso_stdev
                        sheet['P12'] = iso_pass
                        sheet['P13'] = iso_fail
                        sheet['P14'] = round(iso_fail/rownum,2)
                        #print('iso_list=',iso_list)

                        ab_stdev = round(statistics.stdev(amplitude_balance),2) #Standard deviation
                        if len(amplitude_balance)>1:
                            ab_var = round(statistics.variance(amplitude_balance),2) #Variance
                        else:
                            ab_var = 0 #Variance
                        ab_avg = round(statistics.mean(amplitude_balance),2) #Mean Average
                        ab_min = round(min(amplitude_balance),2) #Min
                        ab_max = round(max(amplitude_balance),2) #Max
                        ab_list = [ab_min,ab_max,ab_avg,ab_stdev]
                        sheet['Q8'] = ab_avg
                        sheet['Q9'] = ab_min
                        sheet['Q10'] = ab_max
                        sheet['Q11'] = ab_stdev
                        sheet['Q12'] = ab_pass
                        sheet['Q13'] = ab_fail
                        sheet['Q14'] = round(ab_fail/rownum,2)
                        #print('ab_list=',ab_list)

                        pb_stdev = round(statistics.stdev(phase_balance),2) #Standard deviation
                        if len(phase_balance)>1:
                            pb_var = round(statistics.variance(phase_balance),2) #Variance
                        else:
                            pb_var = 0 #Variance
                        pb_avg = round(statistics.mean(phase_balance),2) #Mean Average
                        pb_min = round(min(phase_balance),2) #Min
                        pb_max = round(max(phase_balance),2) #Max
                        pb_list = [pb_min,pb_max,pb_avg,pb_stdev]
                        sheet['R8'] = pb_avg
                        sheet['R9'] = pb_min
                        sheet['R10'] = pb_max
                        sheet['R11'] = pb_stdev
                        sheet['R12'] = pb_pass
                        sheet['R13'] = pb_fail
                        sheet['R14'] = round(pb_fail/rownum,2)
                        #print('pb_list=',pb_list)
                        stat_list = [il_list,rl_list,iso_list,ab_list,pb_list]
                    else:
                        coup_stdev = round(statistics.stdev(coupling),2) #Standard deviation
                        if len(coupling)>1:
                            coup_var = round(statistics.variance(coupling),2) #Variance
                        else:
                            coup_var = 0 #Variance
                        coup_avg = round(statistics.mean(coupling),2) #Mean Average
                        coup_min = round(min(coupling),2) #Min
                        coup_max = round(max(coupling),2) #Max
                        coup_list = [coup_min,coup_max,coup_avg,coup_stdev]
                        sheet['P8'] = coup_avg
                        sheet['P9'] = coup_min
                        sheet['P10'] = coup_max
                        sheet['P11'] = coup_stdev
                        sheet['P12'] = coup_pass
                        sheet['P13'] = coup_fail
                        sheet['P14'] = round(coup_fail/rownum,2)
                        #print('iso_list=',iso_list)

                        dir_stdev = round(statistics.stdev(directivity),2) #Standard deviation
                        if len(directivity)>1:
                            dir_var = round(statistics.variance(directivity),2) #Variance
                        else:
                            dir_var = 0 #Variance
                        dir_avg = round(statistics.mean(directivity),2) #Mean Average
                        dir_min = round(min(directivity),2) #Min
                        dir_max = round(max(directivity),2) #Max
                        dir_list = [dir_min,dir_max,dir_avg,dir_stdev]
                        sheet['Q8'] = dir_avg
                        sheet['Q9'] = dir_min
                        sheet['Q10'] = dir_max
                        sheet['Q11'] = dir_stdev
                        sheet['Q12'] = dir_pass
                        sheet['Q13'] = dir_fail
                        sheet['Q14'] = round(dir_fail/rownum,2)
                        #print('ab_list=',ab_list)

                        cf_stdev = round(statistics.stdev(coupledflatness),2) #Standard deviation
                        if len(coupledflatness)>1:
                            cf_var = round(statistics.variance(coupledflatness),2) #Variance
                        else:
                            cf_var = 0 #Variance
                        cf_avg = round(statistics.mean(coupledflatness),2) #Mean Average
                        cf_min = round(min(coupledflatness),2) #Min
                        cf_max = round(max(coupledflatness),2) #Max
                        cf_list = [cf_min,cf_max,cf_avg,cf_stdev]
                        sheet['R8'] = cf_avg
                        sheet['R9'] = cf_min
                        sheet['R10'] = cf_max
                        sheet['R11'] = cf_stdev
                        sheet['R12'] = cf_pass
                        sheet['R13'] = cf_fail
                        sheet['R14'] = round(cf_fail/rownum,2)
                        #print('pb_list=',pb_list)
                        stat_list = [il_list,rl_list,coup_list,dir_list,cf_list]

                    
                    #print('stat_list=',stat_list)
                    sheet.title = artwork_rev
                    
                    #~~~~~~~~~~~~~~~~~~~~~~Summary sheet~~~~~~~~~~~~~~~~~~~~~~~~
                    sheet = wb["Summary"]
                    #print('sheet=',sheet)
                    
                    if '90 DEGREE COUPLER' in spectype or 'BALUN' in spectype:
                        #AVG
                        sheet['A' + str(sum_row)] = artwork_rev
                        sheet['B' + str(sum_row-1)] = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row-1)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row-1)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row-1)] = "'+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row-1)] = "'+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row)] = il_avg
                        sheet['C' + str(sum_row)] = rl_avg
                        sheet['D' + str(sum_row)] = iso_avg
                        sheet['E' + str(sum_row)] = ab_avg
                        sheet['F' + str(sum_row)] = pb_avg
                        sheet['G' + str(sum_row)] = rownum
                        
                     
                        #MIN
                        sheet['A' + str(sum_row + 14)] = artwork_rev
                        sheet['B' + str(sum_row + 13)] = spec_list[0]  = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row + 13)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row + 13)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row + 13)] = "+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row + 13)] = "+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row + 14)] = il_min
                        sheet['C' + str(sum_row + 14)] = rl_min
                        sheet['D' + str(sum_row + 14)] = iso_min
                        sheet['E' + str(sum_row + 14)] = ab_min
                        sheet['F' + str(sum_row + 14)] = pb_min
                        sheet['G' + str(sum_row + 14)] = rownum
                        #Max
                        sheet['A' + str(sum_row + 28)] = artwork_rev
                        sheet['B' + str(sum_row + 27)] = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row + 27)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row + 27)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row + 27)] = "+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row + 27)] = "+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row + 28)] = il_max
                        sheet['C' + str(sum_row + 28)] = rl_max
                        sheet['D' + str(sum_row + 28)] = iso_max
                        sheet['E' + str(sum_row + 28)] = ab_max
                        sheet['F' + str(sum_row + 28)] = pb_max
                        sheet['G' + str(sum_row + 28)] = rownum
                    else:
                                                #AVG
                        sheet['A' + str(sum_row)] = artwork_rev
                        sheet['B' + str(sum_row-1)] = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row-1)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row-1)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row-1)] = "'+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row-1)] = "'+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row)] = il_avg
                        sheet['C' + str(sum_row)] = rl_avg
                        sheet['D' + str(sum_row)] = coup_avg
                        sheet['E' + str(sum_row)] = dir_avg
                        sheet['F' + str(sum_row)] = cf_avg
                        sheet['G' + str(sum_row)] = rownum
                        
                     
                        #MIN
                        sheet['A' + str(sum_row + 14)] = artwork_rev
                        sheet['B' + str(sum_row + 13)] = spec_list[0]  = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row + 13)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row + 13)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row + 13)] = "+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row + 13)] = "+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row + 14)] = il_min
                        sheet['C' + str(sum_row + 14)] = rl_min
                        sheet['D' + str(sum_row + 14)] = coup_min
                        sheet['E' + str(sum_row + 14)] = dir_min
                        sheet['F' + str(sum_row + 14)] = cf_min
                        sheet['G' + str(sum_row + 14)] = rownum
                        #Max
                        sheet['A' + str(sum_row + 28)] = artwork_rev
                        sheet['B' + str(sum_row + 27)] = str(spec_list[0]) + ' Max'
                        sheet['C' + str(sum_row + 27)] = str(spec_list[1]) + ' Max'
                        sheet['D' + str(sum_row + 27)] = str(spec_list[2]) + ' Max'
                        sheet['E' + str(sum_row + 27)] = "+/- " + str(spec_list[3]) + ' dB'
                        sheet['F' + str(sum_row + 27)] = "+/- " + str(spec_list[4]) + ' deg'
                        sheet['B' + str(sum_row + 28)] = il_max
                        sheet['C' + str(sum_row + 28)] = rl_max
                        sheet['D' + str(sum_row + 28)] = coup_max
                        sheet['E' + str(sum_row + 28)] = dir_max
                        sheet['F' + str(sum_row + 28)] = cf_max
                        sheet['G' + str(sum_row + 28)] = rownum

                    #~~~~~~~~~~~~~~~~~~~~~~Summary sheet~~~~~~~~~~~~~~~~~~~~~~~~
                    #rename the sheet to the artworkrev
                    x+=1
                    sum_row+=1
            z+=1    
            #~~~~~~~~~~~~~~~~~~~~~~~~~charts~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='loading charts')
            trace_num = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title='Insertion Loss J3').count()
            loadcharts = Charts(len(artwork_list),self.job_num,part_num,spectype,self.operator,self.workstation,wb)
            loadcharts.load()
            print('Charts Loaded')
         #~~~~~~~~~~~~~~~~~~~~~~~~~Save~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        savenow = SaveReports(self.job_num,part_num,spectype,self.operator,self.workstation,wb)
        savenow.save()
        ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='report complete')
        print("Report for ",self.job_num, " is complete")


class ReportFiles:
    def __init__ (self, job_num,part_num,spec_type):
        self.job_num = job_num
        self.part_num = part_num
        self.spec_type = spec_type
    
    def data_path(self):
        top_folder = "\\\ippdc\\Test Automation\\Report Server Data\\"
        report_path = "90_Degree\\"
        if '90 DEGREE COUPLER SMD' in self.spec_type:
            report_path = '90_Degree_SMD\\'
        elif '90 DEGREE COUPLER' in self.spec_type:
            report_path = "90_Degree\\"
        elif 'BALUN' in self.spec_type:
            report_path = "Balun\\"
        elif 'DIRECTIONAL COUPLER SMD' in self.spec_type:
            report_path = "Directional_Couplers_SMD\\"
        elif 'DIRECTIONAL COUPLER' in self.spec_type:
            report_path = "Directional_Couplers\\"
        elif 'COMBINER/DIVIDER SMD' in self.spec_type:
            report_path = "Combiner-Divider_SMD\\"
        elif 'COMBINER/DIVIDER' in self.spec_type:
            report_path = "Combiner-Divider\\"
        
        #Create the path if it doesn't exist
        new_path = top_folder + report_path 
        #print('new_path=',new_path)
        if not os.path.exists(new_path):
            os.makedirs(new_path)
        return new_path
    
    def template(self):
        top_folder = "\\\ippdc\\Test Automation\\Excel_Templates\\"
        template = "90DEGREE_STANDARD.xlsx"
        if '90 DEGREE COUPLER SMD' in self.spec_type:
            template = '90DEGREE_STANDARD.xlsx'
        elif '90 DEGREE COUPLER' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
        elif 'BALUN' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
        elif 'DIRECTIONAL COUPLER SMD' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
        elif 'DIRECTIONAL COUPLER' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
        elif 'COMBINER/DIVIDER SMD' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
        elif 'COMBINER/DIVIDER' in self.spec_type:
            template = "90DEGREE_STANDARD.xlsx"
     
        new_path = top_folder + template 
        return new_path
        
class SaveReports:
    def __init__ (self, job_num,part_num,spec_type,operator,workstation,wb):
        self.job_num = job_num
        self.part_num = part_num
        self.spec_type = spec_type
        self.operator = operator
        self.workstation = workstation
        self.wb = wb
        
    def save(self):
        paths = ReportFiles(self.job_num,self.part_num,self.spec_type)
        data_path = paths.data_path()
        self.wb.save(data_path + "TestData " + self.job_num + ".xlsx")
        ReportQueue.objects.using('TEST').filter(reportstatus='in process').filter(jobnumber = self.job_num).filter(partnumber=self.part_num).filter(operator=self.operator).filter(workstation=self.workstation).update(reportstatus='complete')





class Charts:
    def __init__ (self, rev_num,job_num,part_num,spec_type,operator,workstation,wb):
        self.rev_num = rev_num
        self.job_num = job_num
        self.part_num = part_num
        self.spec_type = spec_type
        self.operator = operator
        self.workstation = workstation
        self.wb = wb
        print('loading charts')
        
    def load(self):
        if self.rev_num==1:
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            #print('charts=',charts)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
        elif self.rev_num==2:
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,2,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
        elif self.rev_num==3:
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,2,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,3,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
        elif self.rev_num==4:
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,2,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,3,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,4,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
        elif self.rev_num==5:
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            ccharts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,2,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,3,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,4,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,5,self.wb)
            charts.chart1()
            charts.chart2()
            charts.chart3()
            charts.chart4()




class LoadCharts:    
    def __init__ (self, job_num,part_num,spec_type,operator,workstation,chart_group,wb):
        self.job_num = job_num
        self.part_num = part_num
        self.spec_type = spec_type
        self.operator = operator
        self.workstation = workstation
        self.chart_group = chart_group
        self.wb = wb
        #print('loading charts')
    
    def chart1(self): 
        print('self.spec_type=',self.spec_type)
        if '90 DEGREE COUPLER' in self.spec_type or 'BALUN' in self.spec_type:
            title1='Insertion Loss J3'
            title2='Insertion Loss J4'
        else:
            title='Insertion Loss'
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            if '90 DEGREE COUPLER' in self.spec_type or 'BALUN' in self.spec_type:
                trace_id1 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title1).filter(serialnumber=serialnumber).values_list('id').first()
                trace_id2 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title2).filter(serialnumber=serialnumber).values_list('id').first()
                #~~~~~~~~~~~~~~~~~~~~~~Insertion Loss J3~~~~~~~~~~~~~~~~~~~~~~~~
                #print('trace_id=',trace_id1)
                trace_points = []
                if trace_id1:
                    if trace_id1[0] > 171666:
                        trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id1[0]).all()
                    else:
                        trace_points = Tracepoints.objects.using('TEST').filter(traceid=trace_id1[0]).all()
                    #print('trace_points=',trace_points)
                    rownum=56
                    for point in trace_points:
                        sheet.cell(row=rownum, column=1).value= round(point.xdata,0)
                        sheet.cell(row=rownum, column=2).value= round(point.ydata,0)
                        #print('rownum=',rownum,' point.xdata=',point.xdata)
                        rownum+=1
                 #~~~~~~~~~~~~~~~~~~~~~~Insertion Loss J4~~~~~~~~~~~~~~~~~~~~~~~~
                trace_points = []
                if trace_id2:
                    if trace_id2[0] > 171666:
                        trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id2[0]).all()
                    else:
                        trace_points = Tracepoints.objects.using('TEST').objects.using('TEST').filter(traceid=trace_id2[0]).all()
                    rownum=56
                    for point in trace_points:
                        sheet.cell(row=rownum, column=3).value= round(point.ydata,0)
                        rownum+=1 
            else:
                #~~~~~~~~~~~~~~~~~~~~~~Insertion Loss ~~~~~~~~~~~~~~~~~~~~~~~~
                trace_id = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title).filter(serialnumber=serialnumber).values_list('id').first()
                #print('trace_id=',trace_id[0])
                trace_points = []
                if trace_id:
                    if trace_id[0] > 171666:
                        trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id[0]).all()
                    else:
                        trace_points = Tracepoints.objects.using('TEST').filter(traceid=trace_id[0]).all()
                    #print('trace_points=',trace_points)
                    rownum=56
                    for point in trace_points:
                        sheet.cell(row=rownum, column=1).value= round(point.xdata,0)
                        sheet.cell(row=rownum, column=2).value= round(point.ydata,0)
                        #print('rownum=',rownum,' point.xdata=',point.xdata)
                        rownum+=1
            
       
    def chart2(self):
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            #~~~~~~~~~~~~~~~~~~~~~~Return Loss~~~~~~~~~~~~~~~~~~~~~~~~
            trace_id = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title='Return Loss').filter(serialnumber=serialnumber).values_list('id').first()       
            trace_points = []
            if trace_id:
                if trace_id[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').objects.using('TEST').filter(traceid=trace_id[0]).all()
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=4).value= round(point.xdata,0)
                    sheet.cell(row=rownum, column=5).value= round(point.ydata,0)
                    rownum+=1
        
          
    def chart3(self): 
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            #~~~~~~~~~~~~~~~~~~~~~~isolation~~~~~~~~~~~~~~~~~~~~~~~~
            trace_id = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title='Return Loss').filter(serialnumber=serialnumber).values_list('id').first()    
            trace_points = []
            if trace_id:
                if trace_id[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').objects.using('TEST').filter(traceid=trace_id[0]).all()
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=6).value= round(point.xdata,0)
                    sheet.cell(row=rownum, column=7).value= round(point.ydata,0)
                    rownum+=1
            
    
    def chart4(self):
        if '90 DEGREE COUPLER' in self.spec_type or 'BALUN' in self.spec_type:
            title1='Phase Balance J3'
            title2='Phase Balance J4'
        else:
            title1='Coupled Flatness J3'
            title2='Coupled Flatness J4'
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            
            trace_id1 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title1).filter(serialnumber=serialnumber).values_list('id').first()
            trace_id2 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title2).filter(serialnumber=serialnumber).values_list('id').first()
        
            
            #~~~~~~~~~~~~~~~~~~~~~~Phase Balance J3~~~~~~~~~~~~~~~~~~~~~~~~
            trace_points = []
            if trace_id1:
                if trace_id1[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id1[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').objects.using('TEST').filter(traceid=trace_id1[0]).all()
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=8).value= round(point.xdata,0)
                    sheet.cell(row=rownum, column=9).value= round(point.ydata,0)
                    rownum+=1
            #~~~~~~~~~~~~~~~~~~~~~~Phase Balance J4~~~~~~~~~~~~~~~~~~~~~~~~
            trace_points = []
            if trace_id2:
                if trace_id2[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id2[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').objects.using('TEST').filter(traceid=trace_id2[0]).all()
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=10).value= round(point.ydata,0)
                    rownum+=1    
              
class DeleteSheets:    
    def __init__ (self,chart_group,wb):
        self.chart_group = chart_group
        self.wb = wb 
        #print('IN delete self.chart_group=',self.chart_group)
            
    def remove(self):
        if self.chart_group==1:    
            # Clean up the template
            #print(self.wb.sheetnames)
            sheetDelete = self.wb["Raw Data2"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data3"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data4"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data5"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 6"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 7"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 8"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 9"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 10"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 11"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 12"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 13"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 14"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 15"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 16"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 17"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 18"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 19"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 20"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 21"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 22"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 23"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 24"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 25"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
        elif self.chart_group==2:
            sheetDelete = self.wb["Raw Data3"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data4"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data5"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 11"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 12"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 13"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 14"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 15"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 16"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 17"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 18"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 19"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 20"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 21"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 22"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 23"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 24"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 25"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
        elif self.chart_group==4:
            sheetDelete = self.wb["Raw Data4"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["Raw Data5"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 11"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 12"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 13"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 14"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
            sheetDelete = self.wb["UUT 15"]
            self.wb.remove(sheetDelete)  #Sheet will be deleted
        #~~~~~~~~~~~~~~~~~~~~~~Clean up the Template~~~~~~~~~~~~~~~~~~~~~~~~

    
                
        
class get_serial_num:
    def __init__ (self,chart_group,idx):
        self.chart_group = chart_group
        self.idx = idx
        #print('self.idx=',self.idx)       
            
    def uut(self):
        serialnumber = "UUT 1"
        if self.chart_group==1 and self.idx ==0:
            serialnumber = "UUT 1"
        elif self.chart_group==1 and self.idx ==1:
            serialnumber = "UUT 2"
        elif self.chart_group==1 and self.idx ==2:
            serialnumber = "UUT 3"
        elif self.chart_group==1 and self.idx ==3:
            serialnumber = "UUT 4"
        elif self.chart_group==1 and self.idx ==4:
            serialnumber = "UUT 5"
        elif self.chart_group==2 and self.idx ==0:
            serialnumber = "UUT 6"
        elif self.chart_group==2 and self.idx ==1:
            serialnumber = "UUT 7"
        elif self.chart_group==2 and self.idx ==2:
            serialnumber = "UUT 8" 
        elif self.chart_group==2 and self.idx ==3:
            serialnumber = "UUT 9" 
        elif self.chart_group==2 and self.idx ==4:
            serialnumber = "UUT 10" 
        elif self.chart_group==3 and self.idx ==0:
            serialnumber = "UUT 11"
        elif self.chart_group==3 and self.idx ==1:
            serialnumber = "UUT 12"
        elif self.chart_group==3 and self.idx ==2:
            serialnumber = "UUT 13"
        elif self.chart_group==3 and self.idx ==3:
            serialnumber = "UUT 14"
        elif self.chart_group==3 and self.idx ==4:
            serialnumber = "UUT 15"
        elif self.chart_group==4 and self.idx ==0:
            serialnumber = "UUT 16"
        elif self.chart_group==4 and self.idx ==1:
            serialnumber = "UUT 17"
        elif self.chart_group==4 and self.idx ==2:
            serialnumber = "UUT 18"
        elif self.chart_group==4 and self.idx ==3:
            serialnumber = "UUT 19"
        elif self.chart_group==4 and self.idx ==4:
            serialnumber = "UUT 20"    
        elif self.chart_group==5 and self.idx ==0:
            serialnumber = "UUT 21"
        elif self.chart_group==5 and self.idx ==1:
            serialnumber = "UUT 22"
        elif self.chart_group==5 and self.idx ==2:
            serialnumber = "UUT 23"
        elif self.chart_group==5 and self.idx ==3:
            serialnumber = "UUT 24"
        elif self.chart_group==5 and self.idx ==4:
            serialnumber = "UUT 25"    
                
        return serialnumber
  
class Statistics:  
    def __init__(self,test1,test2,test3,test4,test5):
        self.test1 = test1
        self.test2 = test2
        self.test3 = test3
        self.test4 = test4
        self.test5 = test5
       
        
    
    def get_stats(self):
     #~~~~~~~~~~~~~~~~Statics and Summary ~~~~~~~~~~~~~~~~~~~~
        if len(self.test1) > 1:# must have at least two tests
            #print('insertion_loss=',self.test1)
            s1_stdev = round(statistics.stdev(self.test1),2) #Standard deviation
            s1_var = round(statistics.variance(self.test1),2) #Variance
            s1_avg = round(statistics.mean(self.test1),2) #Mean Average
            s1_min = round(min(self.test1),2) #Min
            s1_max = round(max(self.test1),2) #Max
            s1_list = [s1_min,s1_max,s1_avg,s1_stdev]
            #print('il_list=',il_list)

            #print('return_loss=',self.test2)
            s2_stdev = round(statistics.stdev(self.test2),2) #Standard deviation
            s2_var = round(statistics.variance(self.test2),2) #Variance
            s2_avg = round(statistics.mean(self.test2),2) #Mean Average
            s2_min = round(min(self.test2),2) #Min
            s2_max = round(max(self.test2),2) #Max
            s2_list = [s2_min,s2_max,s2_avg,s2_stdev]
            

            s3_stdev = round(statistics.stdev(self.test3),2) #Standard deviation
            s3_var = round(statistics.variance(self.test3),2) #Variance
            s3_avg = round(statistics.mean(self.test3),2) #Mean Average
            s3_min = round(min(self.test3),2) #Min
            s3_max = round(max(self.test3),2) #Max
            s3_list = [s3_min,s3_max,s3_avg,s3_stdev]
           
            s4_stdev = round(statistics.stdev(self.test4),2) #Standard deviation
            s4_var = round(statistics.variance(self.test4),2) #Variance
            s4_avg = round(statistics.mean(self.test4),2) #Mean Average
            s4_min = round(min(self.test4),2) #Min
            s4_max = round(max(self.test4),2) #Max
            s4_list = [s4_min,s4_max,s4_avg,s4_stdev]
           
            s5_stdev = round(statistics.stdev(self.test5),2) #Standard deviation
            s5_var = round(statistics.variance(self.test5),2) #Variance
            s5_avg = round(statistics.mean(self.test5),2) #Mean Average
            s5_min = round(min(self.test5),2) #Min
            s5_max = round(max(self.test5),2) #Max
            s5_list = [s5_min,s5_max,s5_avg,s5_stdev]
            
            stat_list = [s1_list,s2_list,s3_list,s4_list,s5_list]
           
            return stat_list
           
class XY_Chart:  
    def __init__(self,tests,specs,test):
        self.test1 = tests[0]
        self.test2 = tests[1]
        self.test3 = tests[2]
        self.test4 = tests[3]
        self.test5 = tests[4]
        self.spec1 = specs[0]
        self.spec2 = specs[1]
        self.spec3 = 0-specs[2]
        self.spec4 = specs[3]
        self.spec5 = specs[4]
        self.test = test
        
    def Chart_data(self):
        chart = []
        if self.test == 'test1':
            testing = self.test1
            spec = self.spec1
        if self.test == 'test2':
            testing = self.test2
            spec = self.spec2
        if self.test == 'test3':
            testing = self.test3
            spec = self.spec3    
        if self.test == 'test4':
            testing = self.test4
            spec = self.spec4
        if self.test == 'test5':
            testing = self.test5
            spec = self.spec5
        
        x=0
        for test in testing:
            chart.append((x,test))
            x+=1
        
        return chart

class XY_Hist:  
    def __init__(self,sd_list,sd_x_range):
        self.sd_list = sd_list
        self.sd_x_range = sd_x_range
            
    def data(self):
        xy = []
        x=0
        for test in self.sd_x_range:
            xy.append((self.sd_x_range[x],test))
            x+=1
        
        return xy        
            

class SDEV_Dist:
    def __init__(self,spec,data,sdev,Min,Max,mean):
        #print('##########################################spec=',spec)
        self.Min = Min 
        self.Max = Max
        self.data = data
        self.sdev = sdev
        self.mean = mean
        span = self.Max - self.Min
        step = int(len(data))
        self.stepsize = (span/step)
        self.height = max(data)
        
        
    def matlab(self): # pur matlab
        from math import exp, pow
        temp_list = []
        try:
            import numpy as np
            variance = pow(self.sdev, 2)
            x = np.linspace(-3 * self.sdev +  self.mean, 3 * self.sdev + self.mean , 100)
            x = np.arange(0.4,0.8,self.stepsize)
            temp_list = np.exp(-np.square(x-self.mean)/2*variance)/(np.sqrt(2*np.pi*variance))
        except ZeroDivisionError as e:
           print('matlab list error=',e)
        return temp_list
    
    def linspace(self): # includes linspace
        import numpy as np
        temp_list = np.linspace(-3 * self.sdev +  self.mean, 3 * self.sdev + self.mean , 100)
        return temp_list
        
    
    def gauss(self): # includes height
        chart = []
        try:      
            #gaussian distribution
            import numpy as np
            variance = np.square(self.sdev)
            x = np.arange(self.Min,self.Max,self.stepsize)
            f = (-np.square(x-self.mean)/2*variance)/(np.sqrt(2*np.pi*variance))
            y=0
            for sd in f:
                chart.append((x[y],sd))
                y+=1
        except ZeroDivisionError as e:
           print('gaussian dist error=',e)
        return chart
    
    def gauss_min_max(self): # includes height
        chart = []
        answer = [0,0]      
        try:
            #gaussian distribution
            import numpy as np
            variance = np.square(self.sdev)
            x = np.arange(self.Min,self.Max,self.stepsize)
            f = (-np.square(x-self.mean)/2*variance)/(np.sqrt(2*np.pi*variance))
            y=0
            for sd in f:
                chart.append(sd)
                
            Min = min(chart)
            Max = max(chart)
            answer = [Min,Max]
        except ZeroDivisionError as e:
           print('gaussian dist error=',e)
        return answer   
     


class X_Range:
    def __init__(self,data,spec,Min,Max,Mean):
        self.data = data
        if spec < Min:
            self.Min = spec
        else:
            self.Min = Min 
            
        if spec > Max:
            self.Max = spec
        else:
            self.Max = Max
            
        span = self.Max - self.Min
        step = int(len(data))
        self.stepsize = (span/step)
        #print('self.stepsize=',self.stepsize)
        self.Mean = Mean
        #print('self.Min=',self.Min)
        #print('self.Max=',self.Max)
        #print('self.Mean=',self.Mean)
        self.step = int(step) + 2

    def list(self):
        #print('stepsize =',self.stepsize)
        import numpy as np
        lower_list = np.linspace(self.Min, self.Mean, self.step)
        upper_list = np.linspace(self.Mean+self.stepsize,self.Max, self.step)
        #print('lower_lis=',lower_list)
        #print('upper_list=',upper_list)
        x_range_list = np.concatenate((lower_list,upper_list),axis=None)
        temp_list = []
        for temp in x_range_list:
            temp_list.append(round(temp,3))
        x_range_list = temp_list    
        return x_range_list


        


class Histogram_data:  
    def __init__(self,tests,specs,test):
        self.test1 = tests[0]
        self.test2 = tests[1]
        self.test3 = tests[2]
        self.test4 = tests[3]
        self.test5 = tests[4]
        self.spec1 = specs[0]
        self.spec2 = specs[1]
        self.spec3 = specs[2]
        self.spec4 = specs[3]
        self.spec5 = specs[4]
        self.test = test
        
    def Hist_data(self):
        bin1 = []
        bin2 = []
        bin3 = []
        bin4 = []
        bin5 = []
        bin6 = []
        bin7 = []
        bin8 = []
        bin9 = []
        bin10 = []
        bin11 = []
        bin12 = []
        bin13 = []
        bin14 = []
        bin15 = []
        bin16 = []
        bin17 = []
        bin18 = []
        bin19 = []
        bin20 = []
        bin21 = []
        bin22 = []
        bin23 = []
        bin24 = []
        bin25 = []
        bin26 = []
        bin27 = []
        bin28 = []
        bin29 = []
        bin30 = []
        bin31 = []
        bin32 = []
        bin33 = []
        bin34 = []
        bin35 = []
        bin36 = []
        bin37 = []
        bin38 = []
        bin39 = []
        bin40 = []
        bin41 = []
        bin42 = []
        bin43 = []
        bin44 = []
        bin45 = []
        bin46 = []
        bin47 = []
        bin48 = []
        bin49 = []
        bin50 = []
        bin51 = []
        bin52 = []
        bin53 = []
        bin54 = []
        bin55 = []
        bin56 = []
        bin57 = []
        bin58 = []
        bin59 = []
        bin60 = []
        bin61 = []
        bin62 = []
        bin63 = []
        bin64 = []
        bin65 = []
        bin66 = []
        bin67 = []
        bin68 = []
        bin69 = []
        bin70 = []
        bin71 = []
        bin72 = []
        bin73 = []
        bin74 = []
        bin75 = []
        bin76 = []
        bin77 = []
        bin78 = []
        bin79 = []
        bin80 = []
        bin81 = []
        bin82 = []
        bin83 = []
        bin84 = []
        bin85 = []
        bin86 = []
        bin87 = []
        bin88 = []
        bin89 = []
       
        
        if self.test == 'test1':
            testing = self.test1
            spec = self.spec1
        if self.test == 'test2':
            testing = self.test2
            spec = self.spec2
        if self.test == 'test3':
            testing = self.test3
            spec = self.spec3    
        if self.test == 'test4':
            testing = self.test4
            spec = self.spec4
        if self.test == 'test5':
            testing = self.test5
            spec = self.spec5
            
        for tst in testing:
            if abs(tst) < (spec*0.01) - spec:
                bin1.append(abs(tst))
            elif abs(tst) < (spec*0.02) - spec:
                bin2.append(tst)
            elif abs(tst) < (spec*0.03) - spec:
                bin3.append(tst)
            elif abs(tst) < (spec*0.04) - spec:
                bin4.append(tst)
            elif abs(tst) < (spec*0.05) - spec:
                bin5.append(tst)
            elif abs(tst) < (spec*0.06) - spec:
                bin6.append(tst)
            elif abs(tst) < (spec*0.07) - spec:
                bin7.append(tst)
            elif abs(tst) < (spec*0.08) - spec:
                bin8.append(tst)
            elif abs(tst) < (spec*0.09) - spec:
                bin9.append(tst)
            elif abs(tst) < (spec*0.10) - spec:
                bin10.append(tst)
            elif abs(tst) < (spec*0.11) - spec:
                bin11.append(tst)
            elif abs(tst) < (spec*0.12) - spec:
                bin12.append(tst)
            elif abs(tst) < (spec*0.13) - spec:
                bin13.append(tst)
            elif abs(tst) < (spec*0.14) - spec:
                bin14.append(tst)
            elif abs(tst) < (spec*0.15) - spec:
                bin15.append(tst)
            elif abs(tst) < (spec*0.16) - spec:
                bin15.append(tst)
            elif abs(tst) < (spec*0.17) - spec:
                bin16.append(tst)
            elif abs(tst) < (spec*0.18) - spec:
                bin17.append(tst)
            elif abs(tst) < (spec*0.19) - spec:
                bin18.append(tst)
            elif abs(tst) < (spec*0.2) - spec:
                bin19.append(tst)
            elif abs(tst) < (spec*0.3) - spec:
                bin20.append(tst)
            elif abs(tst) < (spec*0.4) - spec:
                bin21.append(tst)
            elif abs(tst) < (spec*0.5) - spec:
                bin22.append(tst)
            elif abs(tst) < (spec*0.6) - spec:
                bin23.append(tst)
            elif abs(tst) < (spec*0.7) - spec:
                bin24.append(tst)
            elif abs(tst) < (spec*0.8) - spec:
                bin25.append(tst)
            elif abs(tst) < (spec*0.9) - spec:
                bin26.append(tst)
            elif abs(tst) < (spec*1) - spec:
                bin27.append(tst)
            elif abs(tst) < (spec*1.1) - spec:
                bin28.append(tst)
            elif abs(tst) < (spec*1.2) - spec:
                bin29.append(tst)
            elif abs(tst) < (spec*1.3) - spec:
                bin30.append(tst)
            elif abs(tst) < (spec*1.4) - spec:
                bin31.append(tst)
            elif abs(tst) < (spec*1.5) - spec:
                bin32.append(tst)
            elif abs(tst) < (spec*1.6) - spec:
                bin33.append(tst)
            elif abs(tst) < (spec*1.7) - spec:
                bin34.append(tst)
            elif abs(tst) < (spec*1.8) - spec:
                bin35.append(tst)
            elif abs(tst) < (spec*1.9) - spec:
                bin36.append(tst)
            elif abs(tst) < (spec*2) - spec:
                bin37.append(tst)
            elif abs(tst) < (spec*2.1) - spec:
                bin38.append(tst)
            elif abs(tst) < (spec*2.2) - spec:
                bin39.append(tst)
            elif abs(tst) < (spec*2.3) - spec:
                bin40.append(tst)
            elif abs(tst) < (spec*2.4) - spec:
                bin41.append(tst)
            elif abs(tst) < (spec*2.5) - spec:
                bin42.append(tst)
            elif abs(tst) < (spec*2.6) - spec:
                bin43.append(tst)
            elif abs(tst) < (spec*2.7) - spec:
                bin44.append(tst)
            elif abs(tst) < (spec*2.8) - spec:
                bin45.append(tst)
            elif abs(tst) < (spec*2.9) - spec:
                bin46.append(tst)
            elif abs(tst) < (spec*3) - spec:
                bin47.append(tst)
            elif abs(tst) > (spec*3) + spec:
                bin48.append(tst)
            elif abs(tst) > (spec*2.9) + spec:
                bin49.append(tst)
            elif abs(tst) > (spec*2.8) + spec:
                bin50.append(tst)
            elif abs(tst) > (spec*2.7) + spec:
                bin51.append(tst)
            elif abs(tst) > (spec*2.6) + spec:
                bin52.append(tst)
            elif abs(tst) > (spec*2.5) + spec:
                bin53.append(tst)
            elif abs(tst) > (spec*2.4) + spec:
                bin54.append(tst)
            elif abs(tst) > (spec*2.3) + spec:
                bin55.append(tst)
            elif abs(tst) > (spec*2.2) + spec:
                bin56.append(tst)
            elif abs(tst) > (spec*2.1) + spec:
                bin57.append(tst)
            elif abs(tst) > (spec*2) + spec:
                bin58.append(tst)
            elif abs(tst) > (spec*1.9) + spec:
                bin59.append(tst)
            elif abs(tst) > (spec*1.8) + spec:
                bin60.append(tst)
            elif abs(tst) > (spec*1.7) + spec:
                bin61.append(tst)
            elif abs(tst) > (spec*1.6) + spec:
                bin62.append(tst)
            elif abs(tst) > (spec*1.5) + spec:
                bin63.append(tst)
            elif abs(tst) > (spec*2.4) + spec:
                bin64.append(tst)
            elif abs(tst) > (spec*1.3) + spec:
                bin65.append(tst)
            elif abs(tst) > (spec*1.2) + spec:
                bin66.append(tst)
            elif abs(tst) > (spec*1.1) + spec:
                bin67.append(tst)
            elif abs(tst) > (spec*1) + spec:
                bin68.append(tst)
            elif abs(tst) > (spec*0.9) + spec:
                bin69.append(tst)
            elif abs(tst) > (spec*0.8) + spec:
                bin70.append(tst)
            elif abs(tst) > (spec*0.7) + spec:
                bin71.append(tst)
            elif abs(tst) > (spec*0.6) + spec:
                bin72.append(tst)
            elif abs(tst) > (spec*0.5) + spec:
                bin73.append(tst)
            elif abs(tst) > (spec*0.4) + spec:
                bin74.append(tst)
            elif abs(tst) > (spec*0.3) + spec:
                bin75.append(tst)
            elif abs(tst) > (spec*0.2) + spec:
                bin76.append(tst)
            elif abs(tst) > (spec*0.19) + spec:
                bin77.append(tst)
            elif abs(tst) > (spec*0.18) + spec:
                bin78.append(tst)
            elif abs(tst) > (spec*0.17) + spec:
                bin79.append(tst)
            elif abs(tst) > (spec*0.16) + spec:
                bin80.append(tst)
            elif abs(tst) > (spec*0.15) + spec:
                bin81.append(tst)
            elif abs(tst) > (spec*0.14) + spec:
                bin82.append(tst)
            elif abs(tst) > (spec*0.13) + spec:
                bin83.append(tst)
            elif abs(tst) > (spec*0.12) + spec:
                bin84.append(tst)
            elif abs(tst) > (spec*0.11) + spec:
                bin85.append(tst)
            elif abs(tst) > (spec*0.1) + spec:
                bin86.append(tst)
            elif abs(tst) > (spec*0.05) + spec:
                bin87.append(tst)
            elif abs(tst) > (spec*0.02) + spec:
                bin88.append(tst)
            elif abs(tst) > (spec*0.01) + spec:
                bin89.append(tst)   
        
        
        hist = []
        if len(bin1)>0:
            hist.append((len(bin1),min(bin1),max(bin1)))
        if len(bin2)>0:
            hist.append((len(bin2),min(bin2),max(bin2)))
        if len(bin3)>0:
            hist.append((len(bin3),min(bin3),max(bin3)))
        if len(bin4)>0:
            hist.append((len(bin4),min(bin4),max(bin4)))
        if len(bin5)>0:
            hist.append((len(bin5),min(bin5),max(bin5)))
        if len(bin6)>0:
            hist.append((len(bin6),min(bin6),max(bin6)))
        if len(bin7)>0:
            hist.append((len(bin7),min(bin7),max(bin7)))
        if len(bin8)>0:
            hist.append((len(bin8),min(bin8),max(bin8)))
        if len(bin9)>0:
            hist.append((len(bin9),min(bin9),max(bin9)))
        if len(bin10)>0:
            hist.append((len(bin10),min(bin10),max(bin10)))
        if len(bin11)>0:
            hist.append((len(bin11),min(bin11),max(bin11)))
        if len(bin12)>0:
            hist.append((len(bin12),min(bin12),max(bin12)))
        if len(bin13)>0:
            hist.append((len(bin13),min(bin13),max(bin13)))
        if len(bin14)>0:
            hist.append((len(bin14),min(bin14),max(bin14)))
        if len(bin15)>0:
            hist.append((len(bin15),min(bin15),max(bin15)))
        if len(bin16)>0:
            hist.append((len(bin16),min(bin16),max(bin16)))
        if len(bin17)>0:
            hist.append((len(bin17),min(bin17),max(bin17)))
        if len(bin18)>0:
            hist.append((len(bin18),min(bin18),max(bin18)))
        if len(bin19)>0:
            hist.append((len(bin19),min(bin19),max(bin19)))
        if len(bin20)>0:
            hist.append((len(bin20),min(bin20),max(bin20)))
        if len(bin21)>0:
            hist.append((len(bin21),min(bin21),max(bin21)))
        if len(bin22)>0:
            hist.append((len(bin22),min(bin22),max(bin22)))
        if len(bin23)>0:
            hist.append((len(bin23),min(bin23),max(bin23)))
        if len(bin24)>0:
            hist.append((len(bin24),min(bin24),max(bin24)))
        if len(bin25)>0:
            hist.append((len(bin25),min(bin25),max(bin25)))
        if len(bin26)>0:
            hist.append((len(bin26),min(bin26),max(bin26)))
        if len(bin27)>0:
            hist.append((len(bin27),min(bin27),max(bin27)))
        if len(bin28)>0:
            hist.append((len(bin28),min(bin28),max(bin28)))
        if len(bin29)>0:
            hist.append((len(bin29),min(bin29),max(bin29)))
        if len(bin30)>0:
            hist.append((len(bin30),min(bin30),max(bin30)))
        if len(bin31)>0:
            hist.append((len(bin31),min(bin31),max(bin31)))
        if len(bin32)>0:
            hist.append((len(bin32),min(bin32),max(bin32)))
        if len(bin33)>0:
            hist.append((len(bin33),min(bin33),max(bin33)))
        if len(bin34)>0:
            hist.append((len(bin34),min(bin34),max(bin34)))
        if len(bin35)>0:
            hist.append((len(bin35),min(bin35),max(bin35)))
        if len(bin36)>0:
            hist.append((len(bin36),min(bin36),max(bin36)))
        if len(bin37)>0:    
            hist.append((len(bin37),min(bin37),max(bin37)))
        if len(bin38)>0:
            hist.append((len(bin38),min(bin38),max(bin38)))
        if len(bin39)>0:
            hist.append((len(bin39),min(bin39),max(bin39)))
        if len(bin40)>0:
            hist.append((len(bin40),min(bin40),max(bin40)))
        if len(bin41)>0:
            hist.append((len(bin41),min(bin41),max(bin41)))
        if len(bin42)>0:
            hist.append((len(bin42),min(bin42),max(bin42)))
        if len(bin43)>0:
            hist.append((len(bin43),min(bin43),max(bin43)))
        if len(bin44)>0:
            hist.append((len(bin44),min(bin44),max(bin44)))
        if len(bin45)>0:
            hist.append((len(bin45),min(bin45),max(bin45)))
        if len(bin46)>0:
            hist.append((len(bin46),min(bin46),max(bin46)))
        if len(bin47)>0:    
            hist.append((len(bin47),min(bin47),max(bin47)))
        if len(bin48)>0:
            hist.append((len(bin48),min(bin48),max(bin48)))
        if len(bin49)>0:
            hist.append((len(bin49),min(bin49),max(bin49)))
        if len(bin50)>0:
            hist.append((len(bin50),min(bin50),max(bin50)))
        if len(bin51)>0:
            hist.append((len(bin51),min(bin51),max(bin51)))
        if len(bin42)>0:
            hist.append((len(bin42),min(bin42),max(bin42)))
        if len(bin43)>0:
            hist.append((len(bin43),min(bin43),max(bin43)))
        if len(bin44)>0:
            hist.append((len(bin44),min(bin44),max(bin44)))
        if len(bin45)>0:
            hist.append((len(bin45),min(bin45),max(bin45)))
        if len(bin46)>0:
            hist.append((len(bin46),min(bin46),max(bin46)))
        if len(bin47)>0:    
            hist.append((len(bin47),min(bin47),max(bin47)))
        if len(bin48)>0:
            hist.append((len(bin48),min(bin48),max(bin48)))
        if len(bin49)>0:
            hist.append((len(bin49),min(bin49),max(bin49)))
        if len(bin50)>0:
            hist.append((len(bin50),min(bin50),max(bin50)))
        if len(bin51)>0:
            hist.append((len(bin51),min(bin51),max(bin51)))
        if len(bin42)>0:
            hist.append((len(bin42),min(bin42),max(bin42)))
        if len(bin43)>0:
            hist.append((len(bin43),min(bin43),max(bin43)))
        if len(bin44)>0:
            hist.append((len(bin44),min(bin44),max(bin44)))
        if len(bin45)>0:
            hist.append((len(bin45),min(bin45),max(bin45)))
        if len(bin46)>0:
            hist.append((len(bin46),min(bin46),max(bin46)))
        if len(bin47)>0:    
            hist.append((len(bin47),min(bin47),max(bin47)))
        if len(bin48)>0:
            hist.append((len(bin48),min(bin48),max(bin48)))
        if len(bin49)>0:
            hist.append((len(bin49),min(bin49),max(bin49)))
        if len(bin50)>0:
            hist.append((len(bin50),min(bin50),max(bin50)))
        if len(bin51)>0:
            hist.append((len(bin51),min(bin51),max(bin51)))
        if len(bin52)>0:
            hist.append((len(bin52),min(bin52),max(bin52)))
        if len(bin53)>0:
            hist.append((len(bin53),min(bin53),max(bin53)))
        if len(bin54)>0:
            hist.append((len(bin54),min(bin54),max(bin54)))
        if len(bin55)>0:
            hist.append((len(bin55),min(bin55),max(bin55)))
        if len(bin56)>0:
            hist.append((len(bin56),min(bin56),max(bin56)))
        if len(bin57)>0:    
            hist.append((len(bin57),min(bin57),max(bin57)))
        if len(bin58)>0:
            hist.append((len(bin58),min(bin58),max(bin58)))
        if len(bin59)>0:
            hist.append((len(bin59),min(bin59),max(bin59)))
        if len(bin60)>0:
            hist.append((len(bin60),min(bin60),max(bin60)))
        if len(bin61)>0:
            hist.append((len(bin61),min(bin61),max(bin61)))
        if len(bin62)>0:
            hist.append((len(bin42),min(bin42),max(bin62)))
        if len(bin63)>0:
            hist.append((len(bin63),min(bin63),max(bin63)))
        if len(bin64)>0:
            hist.append((len(bin64),min(bin64),max(bin64)))
        if len(bin65)>0:
            hist.append((len(bin65),min(bin65),max(bin65)))
        if len(bin66)>0:
            hist.append((len(bin66),min(bin66),max(bin66)))
        if len(bin67)>0:    
            hist.append((len(bin67),min(bin67),max(bin67)))
        if len(bin68)>0:
            hist.append((len(bin68),min(bin68),max(bin68)))
        if len(bin69)>0:
            hist.append((len(bin69),min(bin69),max(bin69)))
        if len(bin70)>0:
            hist.append((len(bin70),min(bin70),max(bin70)))
        if len(bin57)>0:
            hist.append((len(bin71),min(bin71),max(bin71)))
        if len(bin72)>0:
            hist.append((len(bin72),min(bin72),max(bin72)))
        if len(bin73)>0:
            hist.append((len(bin73),min(bin73),max(bin73)))
        if len(bin74)>0:
            hist.append((len(bin74),min(bin74),max(bin74)))
        if len(bin75)>0:
            hist.append((len(bin75),min(bin75),max(bin75)))
        if len(bin76)>0:
            hist.append((len(bin76),min(bin76),max(bin76)))
        if len(bin77)>0:    
            hist.append((len(bin77),min(bin77),max(bin77)))
        if len(bin78)>0:
            hist.append((len(bin78),min(bin78),max(bin78)))
        if len(bin79)>0:
            hist.append((len(bin79),min(bin79),max(bin79)))
        if len(bin80)>0:
            hist.append((len(bin80),min(bin80),max(bin80)))
        if len(bin81)>0:
            hist.append((len(bin81),min(bin81),max(bin81)))
        if len(bin82)>0:
            hist.append((len(bin82),min(bin82),max(bin82)))
        if len(bin83)>0:
            hist.append((len(bin83),min(bin83),max(bin83)))
        if len(bin84)>0:
            hist.append((len(bin84),min(bin84),max(bin84)))
        if len(bin85)>0:
            hist.append((len(bin85),min(bin85),max(bin85)))
        if len(bin86)>0:
            hist.append((len(bin86),min(bin86),max(bin86)))
        if len(bin87)>0:    
            hist.append((len(bin87),min(bin87),max(bin87)))
        if len(bin88)>0:
            hist.append((len(bin88),min(bin88),max(bin88)))
        if len(bin49)>0:
            hist.append((len(bin89),min(bin89),max(bin89)))
       
        return hist
        
    
    
