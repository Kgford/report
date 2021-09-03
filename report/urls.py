from django.contrib import admin
from django.urls import path
from django.conf.urls import url,include
from django.conf.urls.static import static
from django.conf import settings
from excel import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('excel/', include("excel.urls")),
    #path('', views.ExcelPageView.as_view(), name='home'), 
   
]
