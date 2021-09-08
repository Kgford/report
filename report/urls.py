from django.contrib import admin
from django.urls import path
from django.conf.urls import url,include
from django.conf.urls.static import static
from django.conf import settings
#from excel import views
from users import views

urlpatterns = [
    path('admin/', admin.site.urls),
    url(r'^$',views.index,name='index'),
    url(r'^special/',views.special,name='special'),
    url(r'^users/',include('users.urls')),
    url(r'^logout/$', views.user_logout, name='logout'),
    url(r'^user_login/$', views.user_login, name='login'), 
    path('test/', include("excel.urls")),
    
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
