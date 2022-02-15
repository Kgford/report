from django.apps import AppConfig
from threading import Thread
import time


class TestThread(Thread):

    def run(self):
        a = 1
        while a == 1:#Checking for reports in queue in endless loop
            from test_db.models import ReportQueue,Specifications
            from report.reports import ExcelReports
            
            queue = ReportQueue.objects.using('TEST').filter(reportstatus='report queue').values_list('jobnumber','operator','workstation').all()
            #print('checking Report queue')
            for jobnumber,operator,workstation in queue:
                print('jobnumber=',jobnumber)
                reporting = ExcelReports(jobnumber,operator,workstation)
                reporting.test_data()
            
            time.sleep(10)

class ExcelConfig(AppConfig):
    name = 'excel'
   
    def ready(self):
        TestThread().start()       
