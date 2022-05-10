from django.urls import path
from django.conf.urls import url,include
from django.conf.urls.static import static
from django.conf import settings
from django.contrib import admin
from django.contrib.auth.decorators import login_required, permission_required
from testfixtures.views import (
    FixtureView,
)


app_name = "testfixtures"

urlpatterns =[
  path('', FixtureView.as_view(template_name="index.html"), name='index'),
    ]+ static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

