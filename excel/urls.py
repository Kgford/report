from django.urls import pathfrom django.conf.urls import urlfrom . import views#from django.contrib.auth.decorators import login_required, permission_required'''from qa.views import (    QAView,    QAClosedView,    SearchView,    CertReportView,    CertUpdateView,    LabelReportView,    LabelUpdateView,)'''urlpatterns = [    path('export/excel', views.export_users_xls, name='export_excel'),    path('export/excel-styling', views.export_styling_xls, name='export_styling_excel'),    path('export/export-write-xls', views.export_write_xls, name='export_write_xls'),]