from django.contrib.auth.models import User, Group
from rest_framework import serializers
from test_db.models import ReportQueue
from .models import SPCData
        
class ReportQueueSerializer(serializers.HyperlinkedModelSerializer):
    class Meta:
        model = ReportQueue
        fields = ['reportname', 'reporttype', 'reportstatus','jobnumber','workstation','partnumber','operator','activedate']
        
    
class SPCQueueSerializer(serializers.HyperlinkedModelSerializer):
    part_number = serializers.CharField(max_length=20)
    job_number = serializers.CharField(max_length=20)
    closed_date = serializers.DateField()
    artwork = serializers.CharField(default='N/A')
    panel = serializers.CharField(default='N/A')
    quadrant = serializers.CharField(default='N/A')
    quantity = serializers.IntegerField()
    failed = serializers.IntegerField()
    failed_percent = serializers.FloatField()
    pcb_damage = serializers.IntegerField()
    misc_mechanical = serializers.IntegerField()
    IL_failure = serializers.IntegerField()
    AB_failure = serializers.IntegerField()
    VSWR_failure = serializers.IntegerField()
    ISO_failure = serializers.IntegerField()

    class Meta:
        model = SPCData
        fields = ['part_number', 'job_number', 'closed_date','artwork','panel','quadrant','quantity','failed','failed_percent','pcb_damage','misc_mechanical','IL_failure','AB_failure','VSWR_failure','ISO_failure']