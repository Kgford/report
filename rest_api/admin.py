from django.contrib import admin
from .models import SPCData,SPCDataAPIKey
from rest_framework_api_key.admin import APIKeyModelAdmin

# Register your models here.
admin.site.register(SPCData)

@admin.register(SPCDataAPIKey)
class SPCDataAPIKeyModelAdmin(APIKeyModelAdmin):
    pass
