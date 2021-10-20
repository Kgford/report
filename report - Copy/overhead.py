import math
import os, requests
from django.http import request
#from client.models import Clients
#from contractors.models import Contractors
#from googlevoice import Voice
#from googlevoice.util import input
from six.moves import input
import sys
from datetime import date, datetime, timedelta
from django.utils import timezone
from report import settings
from django.core.mail import send_mail
from users.models import UserProfileInfo
from django.contrib.auth.models import User
from django.db.models import Q
#from atspublic.models import Visitor
from django.shortcuts import render


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




#https://sphinxdoc.github.io/pygooglevoice/examples.html#send-sms-messages
class Comunication:
    def __init__ (self, number,message):
        num_list = []
        if isinstance(number, list):
            print('list number=',number)
            for num in number:
                if '+1' in num:
                    num_list.append(num)
                    print('this phone number in this list is already properly formated')
                else:
                    print('number=',num)
                    numeric_filter = filter(str.isdigit, str(num))
                    number = "".join(numeric_filter)
                    print('numeric_string=',number)
                    print('numeric_string=',number[0])
                    if not number[0] == '1':
                        print('adding 1')
                        num_list.append('+1' + number)
                    else:
                        print('not adding 1')
                        num_list.append('+' + number)
        else:
            if '+1' in number:
                num_list.append(number)
                print('phone number is already properly formated')
            else:
                print('number=',number)
                numeric_filter = filter(str.isdigit, str(number))
                number = "".join(numeric_filter)
                print('numeric_string=',number)
                print('numeric_string=',number[0])
                if not number[0] == '1':
                    print('adding 1')
                    num_list.append('+1' + number)
                else:
                    print('not adding 1')
                    num_list.append('+' + number)
                
            
        self.message = message
        self.number = num_list
        print('message=',self.message)
        print('num_list in com =',self.number)
                

    def send_sms(self):
        till_username = settings.TILL_USERNAME
        till_api_key = settings.TILL_API_KEY
                        
        requests.post(
            "https://platform.tillmobile.com/api/send?username=%s&api_key=%s" % (
                till_username,
                till_api_key
            ), 
            json={
                "phone": self.number,
                "text" : self.message,
                "tag": "New User Alert"
            }
        )
    def send_sms_question(self):
        till_username = settings.TILL_USERNAME
        till_api_key = settings.TILL_API_KEY

        requests.post(
            "https://platform.tillmobile.com/api/send?username=%s&api_key=%s" % (
                till_username,
                till_api_key
            ), 
            json={
                "phone": [self.number],
                "introduction": "Hello from Till.",
                "questions" : [{
                    "text": "Do you have a few moments to answer a question or two?",
                    "tag": "have_a_few_moments",
                    "responses": ["Yes", "No"],
                    "conclude_on": "No",
                    "webhook": "https://yourapp.com/have_a_few_moments_results/"
                },
                {
                    "text": "What is your favorite color?",
                    "tag": "favorite_color",
                    "responses": ["Red", "Green", "Yellow"],
                    "webhook": "https://yourapp.com/favorite_color_results/"
                },
                {
                    "text": "Who is you favorite Star Wars character?",
                    "tag": "favorite_star_wars_character",
                    "webhook": "https://yourapp.com/favorite_star_wars_character_results/"
                }],
                "conclusion": "Thank you for your time"
            }
        )
        
    
    
    def call(self):
        user = 'atetestalerts@gmail.com'
        password = 'Gadget2021'
        print('self.number=',self.number)
        voice = Voice()
        print('voice=',voice)
        print(voice.login(user, password))
        print('in communication')
        outgoingNumber = input('Number to call: ')
        forwardingNumber = input('Number to call from [optional]: ') or None

        voice.call(outgoingNumber, forwardingNumber)

        if input('Calling now... cancel?[y/N] ').lower() == 'y':
            voice.cancel(outgoingNumber, forwardingNumber)
    
    def voice_mails(self):
        user = 'atetestalerts@gmail.com'
        password = 'Gadget2021'
        print('self.number=',self.number)
        voice = Voice()
        print('voice=',voice)
        print(voice.login(user, password))
        print('in communication')
        for message in voice.voicemail().messages:
            util.print_(message)
        
    def delete_messages(self):
        user = 'atetestalerts@gmail.com'
        password = 'Gadget2021'
        print('self.number=',self.number)
        voice = Voice()
        print('voice=',voice)
        print(voice.login(user, password))
        print('in communication')
        for message in voice.sms().messages:
            if message.isRead:
                message.delete()
        
 


class StringThings:
    def __init__ (self, this_string):
        self.this_string = this_string
        
    def number_count(self):
        numCount=0
        for item in self.this_string:
            if item.isdigit():
                numCount +=1
        return numCount
 

#https://www.programiz.com/python-programming/modules/math
class TimeCode:
    def __init__ (self,days):
        timestamp = datetime.today() 
        self.today = timezone.now()
        self.year = timestamp.year
        self.month = timestamp.month
        self.day = timestamp.day
        self.hour = timestamp.hour
        self.minute = timestamp.minute
        self.sec = timestamp.second
        self.days = days
        self.monthstr = {1: "Jan", 2: "Feb",  3: "March", 4: 'April', 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sept", 10: "Oct", 11: "Nov", 12: "Dec"}
        self.monthnum = {"Jan":1,"Feb":2,"March":3,'April':4,"May":5,"Jun":6,"Jul":7, "Aug":8,"Sept":9,"Oct":10,"Nov":11,"Dec":11}
        self.month_list = ["Jan","Feb","March",'April',"May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"]
        
    
    def this_today(self):
        return timezone.now()
        
    def this_year(self):
        return self.year
        
    def this_month(self):
        return self.month
        
    def this_day(self):
        return self.day
    
    def this_hour(self):
        return self.hour   
    
    def this_minute(self):
        return self.minute
        
    def this_sec(self):
        return self.sec
     
    def leap_year(self):
        if (self.year % 4) == 0:
           if (self.year % 100) == 0:
               if (self.year % 400) == 0:
                   return True
               else:
                   return False
           else:
               return True
        else:
           return False
    
    def month_length(self):
        leap = self.leap_year()
        if self.month==2 and leap:
            return 29
        elif self.month==2 and not leap:
            return 28
        elif self.month==4 or self.month ==6 or self.month==9 or self.month==11:
            return 30
        else:
            return 31
            
    def last_month_length(self):
        leap = self.leap_year()
        month=self.month-1
        if month==2 and leap:
            return 29
        elif month==2 and not leap:
            return 28
        elif month==4 or month ==6 or month==9 or month==11:
            return 30
        else:
            return 31
            
    def month_string(self):        
        return self.monthstr[self.month]
       
    def month_number(self):        
        return self.monthnum[self.month]
    
    def today_plus(self): 
        #Equations to get today + x days
        leap_year = self.leap_year()
        print(self.days)
        print('look here',self.year, self.month, self.day+1)
        if self.month == 2 and leap_year and self.day == 29-self.days:
            return datetime(self.year, self.month+1, 1, 19 )
        elif self.month == 2 and not leap_self.year and self.day ==28-self.days:
            return datetime(self.year, self.month+1, 1, 19 )
        elif (self.month == 4 or self.month == 6 or self.month == 9 or self.month == 11) and self.day ==31-self.days:
            return datetime(self.year, self.month+1, 1, 19 )
        elif self.month == 12 and self.day ==32-self.days:
            return datetime(self.year + 1, self.month+1, 1, 19 )
        elif self.day ==31-self.days:
            return datetime(self.year, self.month+1, 1, 19 ) 
        elif self.day ==31:
            return datetime(self.year, self.month, self.day, 19 ) 
        else:
            return datetime(self.year, self.month, self.day+1, 19 )  
    
    def today_minus(self):
        leap_year = self.leap_year()
        last_month_days = self.last_month_length()
        newday=last_month_days+self.day-self.days
        print('last_month_days',last_month_days)
        print('self.day=',self.day)
        print('self.days=',self.days)
        if self.month == 2 and leap_year and 28-self.days >=1:
            return datetime(self.year, self.month-1, newday, 19 )
        elif self.month == 2 and not leap_year and self.day==2-self.days:
            return datetime(self.year, self.month-1, newday, 19 )
        elif (self.month == 4 or self.month == 6 or self.month==9 or self.month == 11) and self.day ==31-self.days:
            print('look here',self.year, self.month-1, newday)
            return datetime(self.year, self.month-1, newday, 19 )
        elif self.month == 12 and self.day ==32-self.days:
            return datetime(self.year - 1, self.month-1, newday, 19 )
        elif self.day <= self.days:
            return datetime(self.year, self.month-1, newday, 19 ) 
        else:
            return datetime(self.year, self.month, self.day-self.days, 19 ) 
    
    def sunday(self):
        date_str = str(self.year) + '-' + str(self.month) + '-' + str(self.day)
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date_obj - timedelta(days=date_obj.weekday())  # Monday
        end_of_week = start_of_week + timedelta(days=4)  # Sunday
        print(start_of_week)
        print(end_of_week)
        return end_of_week
    
    def friday(self):
        date_str = str(self.year) + '-' + str(self.month) + '-' + str(self.day)
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date_obj - timedelta(days=date_obj.weekday())  # Monday
        end_of_week = start_of_week + timedelta(days=5)  # Friday
        print(start_of_week)
        print(end_of_week)
        return end_of_week
        
    def monday(self):
        date_str = str(self.year) + '-' + str(self.month) + '-' + str(self.day)
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        start_of_week = date_obj - timedelta(days=date_obj.weekday())  # Monday
        end_of_week = start_of_week + timedelta(days=6)  # Sunday
        print(start_of_week)
        print(end_of_week)
        return start_of_week
    
    def week(self):
        ww=datetime.date(self.today).strftime("%V")
        return ww
        
 
class GPSCode:
    def __init__ (lat1,lon1,lat2,lon2,unit):
        self.lat1 = lat1
        self.lon1 = lon1
        self.lat2 = lat2
        self.lon2 = lon2
        self.unit = unit
              
    def distance(self):
        try:
            if self.lat1==self.lat2 and self.lon1==lself.on2:
                return 0
            else:
                theta = self.lon1 - self.lon2
                
                dist = math.sin(self.deg2rad(self.lat1)) * math.sin(self.deg2rad(self.lat2)) + math.cos(self.deg2rad(self.lat1)) * math.cos(self.deg2rad(self.lat2)) * math.cos(self.deg2rad(theta))
                dist = math.acos(dist)
                dist = self.rad2deg(dist)
                dist = dist * 60 * 1.1515
                if unit == "K":
                    dist = dist * 1.609344
                elif unit == "N":
                    dist = dist * 0.8684
                return dist
        except IOError as e:
            print('error = ',e) 
            return 0            
            
   
    def deg2rad (self,deg):
        ans = deg * math.pi / 180.0
        return ans
        
    def rad2deg (self,rad):
        ans = rad / math.pi * 180.0
        return ans
        
    def get_num(self,x):
        return float(''.join(ele for ele in x if ele.isdigit() or ele == '.'))
        
        

class Security:
    def __init__ (self, request, page):
        self.page = page
        self.request = request
        print('In security')
        print('self.page=',self.page)
        print('self.request=',self.request)
        
    def visitor_monitor(self):
        timestamp = date.today()
        visitor =  self.get_visitor()
        print('visitor-',visitor)
        client_id=self.get_client_id()
        user_agent=self.get_user_agent()
        session_id = self.get_session_id()
        visitor_ip = self.get_visitor_ip()
        phone_list = self.get_security_phone_list()
        email_list = self.get_security_email_list()
        cookie = self.get_cookie()
        email = self.get_email()
        print('In visitor_monitor')
        reason = -1
        error_message =-1
        #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Check Database~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        print('visitor+',visitor)
        if Visitor.objects.filter(Q(client_id=client_id) | Q(visitor=visitor) | Q(user_agent=user_agent) | Q(visitor_ip=visitor_ip)).exists():
            isthere = Visitor.objects.filter(Q(client_id=client_id) | Q(visitor=visitor) | Q(user_agent=user_agent) | Q(visitor_ip=visitor_ip))
            print('isthere=',isthere)
            if isthere:
                print('isthere=',isthere)
                visitor = isthere[0].visitor
                email = isthere[0].email
                reason = isthere[0].blocked_reason
                isblocked = isthere[0].blocked
                print('blocked=',isblocked)
                print('visitor=',visitor)
                if isblocked:
                    error_message = isthere[0].blocked_reason
                    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~Send Message to staff ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    subject = 'Blocked user visiting ATS ' + self.page + ' page'
                    email_body = visitor + ' is attempting re-entry onto this page after being blocked\n\nBlocked Reason: ' + reason + '\n\nvisitor_ip: ' + visitor_ip + '\n\nClient_id: ' + client_id + '\n\nCookie: ' + cookie + '\n\nUser Agent: ' + user_agent
                    email=Email(email_list,subject, email_body)
                    print('email=',email)
                    email.send_email()
                    mes= 'Blocked user visiting  ' + self.page + ' page' + ' Check your email' 
                    com=Comunication(phone_list,mes)
                    print('com=',com)
                    com.send_sms()
                    blocked = True
        else:
            #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Save New visitor info to Database~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Visitor.objects.create(visitor=visitor,email=email,session_id=session_id,client_id=client_id,
                                user_agent=user_agent,visitor_ip=visitor_ip,created_on=timestamp,last_entry=timestamp)
        return error_message
    

    def get_cookie(self):
        cookie = str(self.request.headers.get('Cookie'))
        print('cookie=',cookie)
        return cookie
        
    def get_visitor(self):
        visitor = str(self.request.user)
        client_id=self.get_client_id()
        if Visitor.objects.filter(client_id=client_id).exists():
            isthere = Visitor.objects.filter(client_id=client_id)
            visitor = isthere[0].visitor
        
        print('visitor=',visitor)
        return visitor
        
    def get_email(self):
        if str(self.request.user) != 'AnonymousUser':
            email = str(self.request.user.email)
        else:
            email = 'N/A'
        print('email=',email)
        return email
        
    def get_client_id(self):
        cookie_array =[]
        inner_array =[]
        email='N/A'
        cookie = str(self.request.headers.get('Cookie'))
        cookie_array= cookie.split( ';',-1)
        print('Cookie=',cookie)
        inner_array=cookie_array[0].split( '=',-1) 
        print('inner_array=',inner_array)
        print('len(inner_array=',len(inner_array))
        if len(inner_array) >=2:
            client_id=inner_array[1]
        else:
            client_id='Unknown'
        print('client_id=',client_id)
        return client_id    
    
    def get_session_id(self):
        cookie_array =[]
        inner_array =[]
        email='N/A'
        cookie = str(self.request.headers.get('Cookie'))
        cookie_array= cookie.split( ';',-1)
        visitor = str(self.request.user)
        session_id ='N/A'
        if visitor!='AnonymousUser':
            inner_array=cookie_array[2].split( '=',-1) 
            if len(inner_array) >=2:
                session_id=inner_array[1]
            else:
                session_id='Unknown'
            print('session_id=',session_id)
 
        return session_id    
    
    def get_user_agent(self):
        user_agent = str(self.request.headers.get('User-Agent'))
        return user_agent    
        
   
    def get_contactus_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) |  Q(alerts_security=True) | Q(alerts_sales=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_contactus_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_help_desk=True) | Q(alerts_security=True) | Q(alerts_sales=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        email_list=[]
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_sales_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_sales=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_sales_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_help_desk=True)| Q(alerts_sales=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_marketing_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_marketing_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
        
    def get_security_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) |  Q(alerts_security=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_security_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) |  Q(alerts_security=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_newuser_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_sales=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_newuser_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True) | Q(alerts_sales=True) | Q(alerts_marketing=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_monitor_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_monitor_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_developer=True)).all()
        print('profiles=',profiles)
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_security_phone_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_security=True) | Q(alerts_developer=True)).all()
        print('profiles=',profiles)
        phone_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                phone_list.append(staff.phone)
        print('phone_list=',phone_list)
        return phone_list
    
    def get_security_email_list(self):
        #~~~~~~~~~~~~~~~~~~~~Get Web_monitor email/phone list/security ~~~~~~~~~~~~~~~~~~~~~~~~
        profiles = UserProfileInfo.objects.filter(Q(alerts_web_monitor=True) | Q(alerts_security=True) | Q(alerts_developer=True)).all()
        print('profiles=',profiles)
        email_list=[]
        print('profiles[0]=',profiles[0].address)
        for staff in profiles:
            if staff.alerts_web_monitor:
                email_list.append(staff.email)
        print('email_list=',email_list)
        return email_list
    
    def get_visitor_ip(self):
        PRIVATE_IPS_PREFIX = ('10.', '172.', '192.', )
        remote_address = self.request.META.get('HTTP_X_FORWARDED_FOR') or self.request.META.get('REMOTE_ADDR')
        print("remote_address=",remote_address)
        ip = remote_address
        x_forwarded_for = self.request.META.get('HTTP_X_FORWARDED_FOR')
        print("x_forwarded_for=",x_forwarded_for)
        if x_forwarded_for:
            proxies = x_forwarded_for.split(',')
            while (len(proxies) > 0 and proxies[0].startswith(PRIVATE_IPS_PREFIX)):
                proxies.pop(0)
                if len(proxies) > 0:
                    ip = proxies[0]
                    print("IP Address",ip)
        visitor_ip = ip
        print('visitor_ip=',visitor_ip)
        return visitor_ip
 

class Conversions:
    def __init__ (self, data,unit):
        self.data = data
        self.unit = unit
        #print('*******look here *******self.data=',self.data)
       
    def vswr_to_rl(self):
        try:
            rl = 20 * math.log10((self.data-1)/(self.data +1))
            return rl
        except ValueError as e:
            print('error = ',e) 
            return 0     





class Style:
    BOLD = '\x1b[1m'
    DIM = '\x1b[2m'
    NORMAL = '\x1b[22m'
    ITALIC = '\x1b[2m'
    UNDERLINE = '\x1b[4m'
    DBL_UNDERLINE = '\x1b[21m'
    NO_UNDERLINE = '\x1b[24m'
    OVERLINE = '\x1b[53m'
    NOT_OVERLINE = '\x1b[55m'
    SLOW_BLINK = '\x1b[5m'
    FAST_BLINK = '\x1b[6m'
    NO_BLINK = '\x1b[25m'
    REVERSE= '\x1b[7m'
    NO_REVERSE = '\x1b[27m'
    STRIKE = '\x1b[9m'
    NO_STRIKE = '\x1b[29m'
    FONT1 = '\x1b[10m'
    FONT2 = '\x1b[11m'
    FONT3 = '\x1b[12m'
    FONT4 = '\x1b[13m'
    FONT5 = '\x1b[14m'
    FONT6 = '\x1b[15m'
    FONT7 = '\x1b[16m'
    FONT8 = '\x1b[17m'
    FONT9 = '\x1b[18m'
    FONT10 = '\x1b[19m'
    ITALIC_UNDERLINE = '\ x1b[2;4m'
    END = '\x1b[0m'
    RED = '\x1b[31m'
    GREEN = '\x1b[32m'
    BLUE = '\x1b[34m'
    YELLOW = '\x1b[33m'
    MAGENTA = '\x1b[35m'
    CYAN = '\x1b[36m'
    BLACK = '\x1b[30m'
    WHITE = '\x1b[37m'
    RED_BG = '\x1b[41m'
    GREEN_BG = '\x1b[42m'
    BLUE_BG = '\x1b[44m'
    YELLOW_BG = '\x1b[43m'
    MAGENTA_BG = '\x1b[45m'
    CYAN_BG = '\x1b[46m'
    BLACK_BG = '\x1b[40m'
    WHITE = '\x1b[37m'
    INVERSE = '\x1b[37;40m'
    INVERSE_BOLD= '\x1b[37;40m'
    BOLD_RED = '\x1b[1;31m'
    BOLD_GREEN = '\x1b[1;32m'
    HILITE = '\x1b[43m'
    BOLD_HILITE = '\x1b[1;43m'
    
        