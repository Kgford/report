from django.urls import include, path

from rest_api import views
from .views import ReportQueueViews,SPCQueueViews



# Wire up our API using automatic URL routing.
# Additionally, we include login URLs for the browsable API.
urlpatterns = [
    
    #path('report_queue/', include('rest_framework.urls', namespace='rest_framework'))
    path('report_queue/', ReportQueueViews.as_view()),
    path('spc_queue/', SPCQueueViews.as_view()),
    path('api/', , ReportQueueView.as_view()),
    path('spc/', , SPCDataView.as_view()),
    
 ]