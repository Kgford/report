from django import forms
from django.shortcuts import render
from django.http import HttpResponseRedirect
from django.http import JsonResponse
from django.core import serializers
from django.core.files import File

from django.urls import reverse, reverse_lazy
from E2.models import Order_Detail,Order_Header,PartNumberWarehouse,PackingListDetail, PackingListHeader,Action
from .models import Trouble_Ticket
from django.views import View
import time
from report.overhead import TimeCode, Security, StringThings, Email
from django.contrib.auth.decorators import login_required
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
from report import settings
import ast
import os

class OpenTicketView(View):
    template_name = "open_ticket.html"
    success_url = reverse_lazy('trouble:open_tkt')
    def get(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            item = -1
            open_tickets = []
            item_id = request.GET.get('item_id', -1)
            print('item_id =',item_id )
            if item_id !=-1:
                item = Trouble_Ticket.objects.filter(id=item_id).first()
                print('item=',item)
            
            
            open_tickets = Trouble_Ticket.objects.filter(trouble_status='Open').all()
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            print('open_tickets=',open_tickets)
        except IOError as e:
            open_tickets = Trouble_Ticket.objects.filter(trouble_status='Open').all()
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"trouble/open_ticket.html",{'item':item,'open_tickets':open_tickets,'closed_tickets':closed_tickets})

    def post(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            item = -1
            open_tickets = []
            item_id = request.POST.get('_item_id', -1)
            trouble_type = request.POST.get('_trouble_type', -1)
            application_type = request.POST.get('_application_type', -1)
            trouble_status = request.POST.get('_trouble_status', -1)
            print('trouble_status=',trouble_status)
            update_by = request.POST.get('_item_update_by', -1)
            print('update_bywww=',update_by)
            open_time = request.POST.get('_open_date', -1)
            location = request.POST.get('_location', -1)
            job_number = request.POST.get('_item_job', '')
            if job_number==-1:
                job_number='Not Entered'
            part_number = request.POST.get('_item_part', '')
            if part_number==-1:
                part_number='Not Entered'
            packing_slip = request.POST.get('_item_pack', '')
            if packing_slip==-1:
                packing_slip='Not Entered'
            description = request.POST.get('_item_desc', -1)
            if description==-1:
                description='Not Entered'
            reported_by = request.POST.get('_item_created_by', -1)
            print('reported_by=',reported_by)
            
            save = request.POST.get('_save', -1)
            close_time = request.POST.get('_close_date', -1)
            print('close_time=',close_time)
            solution = request.POST.get('_item_solution', -1)
            update = request.POST.get('_update', -1)
            delete = request.POST.get('_delete', -1)
            print('trouble_status=',trouble_status)
            print('application_type=',application_type)
            print('trouble_type=',trouble_type)
            if save != -1:
                Trouble_Ticket.objects.create(job_number=job_number, part_number=part_number, application_type=application_type, packing_slip=packing_slip, trouble_status=trouble_status,
                                              trouble_type=trouble_type, open_time=open_time, description=description, reported_by=reported_by, location=location)
                
                #~~~~~~~~~~~~~~~~~~~~~~~~~~~~Send Message ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                subject = '***ATE Trouble Ticket Alert*** ' + str(location)
                email_body = 'ATE Trouble Ticket Alert!\n\nTrouble Type: ' + str(trouble_type) + '\n\nLocation: ' + str(location) + '\n\nMessage: ' + str(description) + '\n\nReported by ' + str(reported_by)
                print(email_body)
                email_list = ['automatedtestsolutions@gmail.com','mford@innovativepp.com','apapocchia@innovativepp.com','jhoang@innovativepp.com','tjdowling@innovativepp.com']
                email=Email(email_list,subject, email_body)
                print('email=',email)
                email.send_email() 
                
            
            if update != -1:
                if 'midnight' in close_time:
                    Trouble_Ticket.objects.filter(id=item_id).update(job_number=job_number,part_number=part_number,application_type=application_type,packing_slip=packing_slip,trouble_status=trouble_status,
                                                  trouble_type=trouble_type,description=description,location=location, solution=solution,update_by=update_by)
                else:
                    Trouble_Ticket.objects.filter(id=item_id).update(job_number=job_number,part_number=part_number,application_type=application_type,packing_slip=packing_slip,trouble_status=trouble_status,
                                                  trouble_type=trouble_type,description=description,location=location, solution=solution, close_time=close_time, update_by=update_by)
                    
                    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~Send Message ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    subject = '***ATE Trouble Ticket# ' + str(item_id) + ' has been closed'
                    email_body = 'ATE Trouble Ticket Alert!\n\nTrouble Type: ' + str(trouble_type) + '\n\nLocation: ' + str(location) + '\n\nMessage: ' + str(description) + '\n\nReported by ' + str(update_by) + ' has been closed at: '  + str(close_time) + '\n\nSolution ' + str(solution)
                    print(email_body)
                    email_list = ['automatedtestsolutions@gmail.com','mford@innovativepp.com','apapocchia@innovativepp.com','jhoang@innovativepp.com','tjdowling@innovativepp.com']
                    email=Email(email_list,subject, email_body)
                    print('email=',email)
                    email.send_email()  
                    
                    
            if delete != -1:
                Trouble_Ticket.objects.filter(id=item_id).delete
                        
            open_tickets = Trouble_Ticket.objects.filter(trouble_status='Open').all()
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            
        except IOError as e:
            open_tickets = Trouble_Ticket.objects.filter(trouble_status='Open').all()
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"trouble/open_ticket.html",{'item':item,'open_tickets':open_tickets,'closed_tickets':closed_tickets})


class ClosedTicketView(View):
    template_name = "closed_ticket.html"
    success_url = reverse_lazy('trouble:close_tkt')
    def get(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            item=1
            closed_tickets = []
            item_id = request.GET.get('item_id', -1)
            print('item_id =',item_id )
            if item_id !=-1:
                item = Trouble_Ticket.objects.filter(id=item_id).first()
                print('item=',item)
            
            
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            
            
            
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"trouble/closed_ticket.html",{'item':item,'closed_tickets':closed_tickets})
        
    def post(self, request, *args, **kwargs):
        operator = self.request.user
        form = 0
        try:
            item = -1
            closed_tickets = []
            item_id = request.POST.get('_item_id', -1)
            trouble_type = request.POST.get('_trouble_type', -1)
            application_type = request.POST.get('_application_type', -1)
            trouble_status = request.POST.get('_trouble_status', -1)
            print('trouble_status=',trouble_status)
            update_by = request.POST.get('_item_updated_by', -1)
            print('update_by=',update_by)
            print('update_by=',update_by)
            open_time = request.POST.get('_open_date', -1)
            location = request.POST.get('_location', -1)
            job_number = request.POST.get('_item_job', '')
            if job_number==-1:
                job_number='Not Entered'
            part_number = request.POST.get('_item_part', '')
            if part_number==-1:
                part_number='Not Entered'
            packing_slip = request.POST.get('_item_pack', '')
            if packing_slip==-1:
                packing_slip='Not Entered'
            description = request.POST.get('_item_desc', -1)
            if description==-1:
                description='Not Entered'
            reported_by = request.POST.get('_item_created_by', -1)
            print('reported_by=',reported_by)
                        
            #print('reported_by=',reported_by)
            save = request.POST.get('_save', -1)
            close_time = request.POST.get('_close_date', -1)
            #print('close_time=',close_time)
            solution = request.POST.get('_item_solution', -1)
            update = request.POST.get('_update', -1)
            delete = request.POST.get('_delete', -1)
            #print('trouble_status=',trouble_status)
            #print('application_type=',application_type)
            #print('trouble_type=',trouble_type)
            if save != -1:
                Trouble_Ticket.objects.create(job_number=job_number, part_number=part_number, application_type=application_type, packing_slip=packing_slip, trouble_status=trouble_status,
                                              trouble_type=trouble_type, open_time=open_time, description=description, reported_by=reported_by, location=location)
            if update != -1:
                if 'midnight' in close_time: # simple update. not closing.
                    Trouble_Ticket.objects.filter(id=item_id).update(job_number=job_number,part_number=part_number,application_type=application_type,packing_slip=packing_slip,trouble_status=trouble_status,
                                                  trouble_type=trouble_type,description=description,location=location, solution=solution,update_by=update_by)
                                                  
                else:
                    Trouble_Ticket.objects.filter(id=item_id).update(job_number=job_number,part_number=part_number,application_type=application_type,packing_slip=packing_slip,trouble_status=trouble_status,
                                                  trouble_type=trouble_type,description=description,location=location, solution=solution, close_time=close_time, update_by=update_by)
                                                  
                              
            
            
            if delete != -1:
                Trouble_Ticket.objects.filter(id=item_id).delete
                        
            closed_tickets = Trouble_Ticket.objects.filter(trouble_status='Closed').all()
            
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"trouble/closed_ticket.html",{'item':item,'closed_tickets':closed_tickets})    

# Create your views here.
