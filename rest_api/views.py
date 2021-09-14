from test_db.models import ReportQueue
from rest_framework import viewsets
from rest_framework import permissions
from rest_framework.views import APIView
from rest_framework import status
from rest_framework.response import Response
from rest_api.serializers import ReportQueueSerializer





#https://www.django-rest-framework.org/tutorial/quickstart/
#https://forums.asp.net/t/2100314.aspx?Calling+a+WEB+API+using+VB+Net
#https://www.django-rest-framework.org/api-guide/requests/#authentication

   
class ReportQueueViewSet(viewsets.ModelViewSet):
    """
    API endpoint that allows groups to be viewed or edited.
    """
    queryset = ReportQueue.objects.all()
    serializer_class = ReportQueueSerializer
    

    