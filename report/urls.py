from django.contrib import admin
from django.urls import path
from django.conf.urls import url,include
from django.conf.urls.static import static
from django.conf import settings
from rest_framework import routers
from rest_api.views import ReportQueueViews
from rest_api.views import ExcelReportStartView,SPCQueueViews



router = routers.DefaultRouter()
router.register(r'report_queue', ReportQueueViews)
router.register(r'spc_queue', SPCQueueViews)
#router.register(r'excel_report_start', ExcelReportStartView, basename='excel_report_start')
admin.site.site_header = "Innovative Power Products Database"
admin.site.site_title = "Innovative Power Products Database Portal"
admin.site.index_title = "Welcome to Innovative Power Products SPC Portal"


urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include(router.urls)),
    path('api-auth/', include('rest_framework.urls', namespace='rest_framework')),
    path('staff/', include("users.urls")),
    path('test/', include("excel.urls")),
    path('testfixtures/', include("testfixtures.urls")),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
