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
        print('job_num=',self.job_num)
               
    def test_data(self):
        job_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('jobnumber').values_list('jobnumber', flat=True).distinct()
        part_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('partnumber').values_list('partnumber', flat=True).distinct()
        artwork_list = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).order_by('partnumber').values_list('artwork_rev', flat=True).distinct()
        report_data = Testdata.objects.using('TEST').filter(jobnumber=self.job_num).all()
        #ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='running report')
        
        
        part_num = report_data[0].partnumber
        spec_data = Specifications.objects.using('TEST').filter(jobnumber=self.job_num).first()
        spectype = spec_data.spectype
        paths = ReportFiles(self.job_num,part_num,spectype)
        data_path = paths.data_path()
        template_path = paths.template()
        #print('template_path=',template_path)
        
        wb = load_workbook(template_path)
        #print('wb=',wb)
        
        if not artwork_list:
            artwork_list = ['RawData 1',]
        
        #filter blanks
        temp_list = []
        for artwork_rev in artwork_list:
            if not artwork_rev == '':
                temp_list.append(artwork_rev)
        artwork_list = temp_list
        # datasheet can only handle 5 artworks
        #print('len(artwork_list)=',len(artwork_list))
        if len(artwork_list) >5:
            group = 5
        else:
            group = len(artwork_list)
        
        remove_extra = DeleteSheets(group,wb)
        remove_extra.remove()
        
        x=1
        print('artwork_list=',artwork_list)
        for artwork_rev in artwork_list:
            if 'RawData 1' in artwork_rev:
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
                spec_list = [spec_data.insertionloss,spec_data_rl,spec_data.coupling,spec_data.directivity,spec_data.coupledflatness]
                
                
                
            #print('spec_list=',spec_list)
            if report_data:
                part_num = report_data[0].partnumber
                print('part_num=',part_num)
                spectype = spec_data.spectype
                
                activesheet = "Raw Data" + str(x)
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
                il_fail = 0
                rl_fail = 0
                iso_fail = 0
                ab_fail = 0
                pb_fail = 0
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
                                iso_pass+=1
                            else:
                                iso_fail+=1
                                testdata3 = sheet.cell(row=rownum, column=6)#Created a variable that contains cell
                            
                            sheet.cell(row=rownum, column=8).value= round(data.directivity,2)
                            testdata4 = sheet.cell(row=rownum, column=8)#Created a variable that contains cell
                            directivity.append(data.directivity)
                            if data.directivity <= spec_list[3]:
                                ab_pass+=1
                            else:
                                ab_fail+=1
                                testdata4.font = Font(color='FF3342', bold=True, italic=True) #W
                            
                            sheet.cell(row=rownum, column=10).value= round(data.coupledflatness,2)
                            testdata5 = sheet.cell(row=rownum, column=10)#Created a variable that contains cell
                            coupledflatness.append(data.coupledflatness)
                            if data.coupledflatness <= spec_list[4]:
                                pb_pass+=1
                            else:
                                pb_fail+=1
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

                    iso_stdev = round(statistics.stdev(isolation),2) #Standard deviation
                    iso_var = round(statistics.variance(isolation),2) #Variance
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
                    ab_var = round(statistics.variance(amplitude_balance),2) #Variance
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
                    pb_var = round(statistics.variance(phase_balance),2) #Variance
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
                    #print('stat_list=',stat_list)
                    sheet.title = artwork_rev
                    
                    #~~~~~~~~~~~~~~~~~~~~~~Summary sheet~~~~~~~~~~~~~~~~~~~~~~~~
                    sheet = wb["Summary"]
                    #print('sheet=',sheet)
                    
                    
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
                    sheet['E' + str(sum_row + 28)]= ab_max
                    sheet['F' + str(sum_row + 28)] = pb_max
                    sheet['G' + str(sum_row + 28)] = rownum
                    #~~~~~~~~~~~~~~~~~~~~~~Summary sheet~~~~~~~~~~~~~~~~~~~~~~~~
                    #rename the sheet to the artworkrev
                    x+=1
                    sum_row+=1
                
            #~~~~~~~~~~~~~~~~~~~~~~~~~charts~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            trace_num = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title='Insertion Loss J3').count()
            loadcharts = Charts(len(artwork_list),self.job_num,part_num,spectype,self.operator,self.workstation,wb)
            loadcharts.load()
             #~~~~~~~~~~~~~~~~~~~~~~~~~Save~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            savenow = SaveReports(self.job_num,part_num,spectype,self.operator,self.workstation,wb)
            savenow.save()
            #ReportQueue.objects.using('TEST').filter(jobnumber=self.job_num).filter(workstation=self.workstation).update(reportstatus='report complete')


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
        print('we are here now1')
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
        print('loading charts1',self.rev_num)
        
    def load(self):
        if self.rev_num==1:
            print('loading charts12')
            charts = LoadCharts(self.job_num,self.part_num,self.spec_type,self.operator,self.workstation,1,self.wb)
            print('charts=',charts)
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
            print('in chart1')
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
                print('trace_id=',trace_id1)
                if trace_id1[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id1[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').filter(traceid=trace_id1[0]).all()
                print('trace_points=',trace_points)
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=1).value= round(point.xdata,0)
                    sheet.cell(row=rownum, column=2).value= round(point.ydata,0)
                    print('rownum=',rownum,' point.xdata=',point.xdata)
                    rownum+=1
                 #~~~~~~~~~~~~~~~~~~~~~~Insertion Loss J4~~~~~~~~~~~~~~~~~~~~~~~~
                
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
                print('trace_id=',trace_id[0])
                if trace_id[0] > 171666:
                    trace_points = Tracepoints2.objects.using('TEST').filter(traceid=trace_id[0]).all()
                else:
                    trace_points = Tracepoints.objects.using('TEST').filter(traceid=trace_id[0]).all()
                print('trace_points=',trace_points)
                rownum=56
                for point in trace_points:
                    sheet.cell(row=rownum, column=1).value= round(point.xdata,0)
                    sheet.cell(row=rownum, column=2).value= round(point.ydata,0)
                    print('rownum=',rownum,' point.xdata=',point.xdata)
                    rownum+=1
            
       
    def chart2(self):
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            #~~~~~~~~~~~~~~~~~~~~~~Return Loss~~~~~~~~~~~~~~~~~~~~~~~~
            trace_id = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title='Return Loss').filter(serialnumber=serialnumber).values_list('id').first()       
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
            title='Insertion Loss'
        for idx in range(5): 
            getser = get_serial_num(self.chart_group,idx)
            serialnumber = getser.uut()
            sheet = self.wb[serialnumber]
            
            trace_id1 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title1).filter(serialnumber=serialnumber).values_list('id').first()
            trace_id2 = Trace.objects.using('TEST').filter(jobnumber=self.job_num).filter(title=title2).filter(serialnumber=serialnumber).values_list('id').first()
        
            
            #~~~~~~~~~~~~~~~~~~~~~~Phase Balance J3~~~~~~~~~~~~~~~~~~~~~~~~
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
        print('self.idx=',self.idx)       
            
    def uut(self):
        serialnumber = "UUT 1"
        if self.chart_group==1 and self.idx ==0:
            serialnumber = "UUT 1"
        elif self.chart_group==1 and self.idx ==1:
            serialnumber = "UUT 2"
        elif self.chart_group==1 and self.idx ==3:
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

  