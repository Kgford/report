from django.urls import path
from django.conf.urls import url
from . import views
from django.contrib.auth.decorators import login_required, permission_required
from trouble.views import (
    OpenTicketView,
    ClosedTicketView,
)


app_name = "trouble"

urlpatterns =[
    path('trouble_ticket_open/', OpenTicketView.as_view(template_name="open_ticket.html"), name='open_tkt'),
    path('trouble_ticket_closed/', ClosedTicketView.as_view(template_name="closed_ticket.html"), name='close_tkt'),
    ]