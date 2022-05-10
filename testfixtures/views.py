from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib.auth.models import User
import os
from django import forms
from django.views import View
from django.urls import reverse, reverse_lazy
from report.overhead import TimeCode, Security, StringThings,Conversions
from test_db.models import TestFixtures,Testdata

class FixtureView(View):
    template_name = "index.html"
    success_url = reverse_lazy('testfixtures:index')
    def get(self, request, *args, **kwargs):
        try:
            active_fixture_id = request.GET.get('active_fixture_id', -1)
            print('active_fixture_id=',active_fixture_id)
            active_part=-1
            fixture=-1
            plunger=-1
            revision=-1
            fix_num=-1
            print('active_fixture_id=',active_fixture_id)
            if active_fixture_id!=-1:
                part=TestFixtures.objects.using('TEST').filter(pk=active_fixture_id).last()
                if part:
                    active_part=part.partnumber
                    print('active_part=',active_part)
                    fixture=part.fixturenumber
                    plunger=part.plunger
                    revision=part.revision
                    fix_num=part.fixnum
            
            
            test_fixtures = TestFixtures.objects.using('TEST').all()
            print('test_fixtures=',test_fixtures)
            part_list = Testdata.objects.using('TEST').order_by('partnumber').values_list('partnumber', flat=True).distinct()
            print('test_fixtures=',test_fixtures)
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"testfixtures/index.html",{'test_fixtures':test_fixtures,'active_fixture_id':active_fixture_id,'part_list':part_list,'active_part':active_part,
                                                                'fixture':fixture,'plunger':plunger,'revision':revision,'fix_num':fix_num})
    def post(self, request, *args, **kwargs):
        try:
            active_fixture_id = request.POST.get('_active_fixture_id', -1)
            print('active_fixture_id=',active_fixture_id)
            active_part = request.POST.get('_part', -1)
            print('active_part=',active_part)
            fixture = request.POST.get('_fixture', -1)
            plunger = request.POST.get('_plunger', -1)
            if plunger ==-1:
                plunger='N/A'
            revision = request.POST.get('_revision', -1)
            if revision==-1:
                revision=-"N/A"
            fix_num = request.POST.get('_fix_num', -1)
            save = request.POST.get('_save', -1)
            update = request.POST.get('_update', -1)
            delete = request.POST.get('_delete', -1)
            if save!=-1:
                TestFixtures.objects.using('TEST').create(partnumber=active_part,fixturenumber=fixture,plunger=plunger,revision=revision,fixnum=fix_num)
                return redirect('testfixtures:index')
            if update!=-1:
                print('in update. active_fixture_id=',active_fixture_id)
                TestFixtures.objects.using('TEST').filter(pk=active_fixture_id).update(partnumber=active_part,fixturenumber=fixture,plunger=plunger,revision=revision,fixnum=fix_num)
                return redirect('testfixtures:index')
            if delete!=-1:
                TestFixtures.objects.using('TEST').filter(pk=active_fixture_id).delete()
                return redirect('testfixtures:index')
                
                
            test_fixtures = TestFixtures.objects.using('TEST').all()
            part_list = Testdata.objects.using('TEST').order_by('partnumber').values_list('partnumber', flat=True).distinct()
            print('test_fixtures=',test_fixtures)
        except IOError as e:
            print ("Lists load Failure ", e)
            print('error = ',e) 
        return render (self.request,"testfixtures/index.html",{'test_fixtures':test_fixtures,'active_fixture_id':active_fixture_id,'part_list':part_list,'active_part':active_part,
                                                                'fixture':fixture,'plunger':plunger,'revision':revision,'fix_num':fix_num})
