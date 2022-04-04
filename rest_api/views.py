from test_db.models import ReportQueue
from rest_framework import viewsets
from rest_framework import permissions
from rest_framework.views import APIView
from rest_framework import status
from rest_framework.response import Response
from rest_api.serializers import ReportQueueSerializer,SPCQueueSerializer
from rest_framework.decorators import action
from test_db.models import ReportQueue
from .models import SPCData



#https://www.django-rest-framework.org/tutorial/quickstart/
#https://forums.asp.net/t/2100314.aspx?Calling+a+WEB+API+using+VB+Net
#https://www.django-rest-framework.org/api-guide/requests/#authentication


class ReportQueueViews(viewsets.ModelViewSet):
    
    serializer_class = ReportQueueSerializer
    queryset = ReportQueue.objects.using('TEST').filter(reportstatus='in process').all()
    def create(self, request, *args, **kwargs):
        bill_data = request.data
        print(bill_data)
        return bill_data
    
    
class ExcelReportStartView(APIView):
    def post(self, request):
        serializer_class = ReportQueueSerializer
        if serializer_class.is_valid():
            #serializer.save()
            return Response({"status": "success", "data": serializer.data}, status=status.HTTP_200_OK)
        else:
            return Response({"status": "error", "data": serializer.errors}, status=status.HTTP_400_BAD_REQUEST)
    
    
class SPCQueueViews(viewsets.ModelViewSet):
    
    serializer_class = SPCQueueSerializer
    queryset = SPCData.objects.all()
    def create(self, request, *args, **kwargs):
        bill_data = request.data
        print(bill_data)
        return bill_data


class SPCDataView(APIView):
    def post(self, request):
        serializer = CartItemSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response({"status": "success", "data": serializer.data}, status=status.HTTP_200_OK)
        else:
            return Response({"status": "error", "data": serializer.errors}, status=status.HTTP_400_BAD_REQUEST)