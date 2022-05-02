from django.db import models
from rest_framework_api_key.models import AbstractAPIKey

class SPCData(models.Model):
    part_number = models.CharField("part_number",max_length=20,null=True,unique=False,default='N/A')
    job_number = models.CharField("job_number",max_length=20,null=True,unique=False,default='N/A')
    gl_account = models.CharField("gl_account",max_length=20,null=True,unique=False,default='N/A')
    closed_date = models.DateField("closed_date",null=True, blank=True)
    artwork = models.CharField("artwork",max_length=20,null=True,unique=False,default='N/A')
    panel = models.CharField("panel",max_length=20,null=True,unique=False,default='N/A')
    sector = models.CharField("sector",max_length=20,null=True,unique=False,default='N/A')
    lot = models.CharField("lot",max_length=20,null=True,unique=False,default='N/A')
    quantity = models.IntegerField("quantity",null=True)
    failed = models.IntegerField("failed",null=True)
    failed_percent = models.FloatField("failed_percent",null=True)
    pcb_damage = models.IntegerField("pcb_damage",null=False,default=0)
    misc_mechanical = models.IntegerField("misc_mechanical",null=False,default=0)
    IL_failure = models.IntegerField("IL_failure",null=False,default=0)
    AB_failure = models.IntegerField("AB_failure",null=False,default=0)
    PB_failure = models.IntegerField("PB_failure",null=False,default=0)
    VSWR_failure = models.IntegerField("vswr_failure",null=False,default=0)
    ISO_failure = models.IntegerField("ISO_failure",null=False,default=0)
    
class SPCDataAPIKey(AbstractAPIKey):
    spcdata = models.ForeignKey(SPCData,on_delete=models.CASCADE, related_name="api_keys", )
 
    
