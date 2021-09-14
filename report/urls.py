from django.contrib import admin
from django.urls import path
from django.conf.urls import url,include
from django.conf.urls.static import static
from django.conf import settings
from rest_framework import routers
from users import views
from rest_api import views

router = routers.DefaultRouter()
router.register(r'report queue', views.ReportQueueViewSet)

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include(router.urls)),
    path('api-auth/', include('rest_framework.urls', namespace='rest_framework')),
    path('staff/', include("users.urls")),
    path('test/', include("excel.urls")),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
