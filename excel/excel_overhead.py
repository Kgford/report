import math
import os, requests
from django.http import request
import sys
from datetime import date, datetime, timedelta
from django.utils import timezone
from report import settings
#from users.models import UserProfileInfo
from django.contrib.auth.models import User
from django.db.models import Q
#from atspublic.models import Visitor
from django.shortcuts import render
mport xlwt
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from django.http import HttpResponse
from django.contrib.auth.models import User
import os


#https://data-flair.training/blogs/django-send-email/
class Email:
    def __init__ (self, recepient_list,subject,message):
        self.subject = subject
        self.message = message
        self.recepient = recepient_list
        print('recepient=',self.recepient)
        if not isinstance(self.recepient, list):
            self.recepient = [self.recepient]
            print('recepient=',self.recepient)
    
    def send_email(self):
        print('EMAIL_HOST_USER=',settings.EMAIL_HOST_USER)
        res = send_mail(self.subject, self.message, settings.EMAIL_HOST_USER, self.recepient, fail_silently = False)
        print('response=',res)
        return res

