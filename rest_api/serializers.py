from django.contrib.auth.models import User, Group
from rest_framework import serializers
from test_db.models import ReportQueue
        
class ReportQueueSerializer(serializers.HyperlinkedModelSerializer):
    class Meta:
        model = ReportQueue
        fields = ['reportname', 'reporttype', 'reportstatus','jobnumber','workstation','partnumber','operator','activedate']