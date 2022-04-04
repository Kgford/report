from django.db import models

class SPCData(models.Model):
    part_number = models.CharField("part_number",max_length=20,null=True,unique=False,default='N/A')
    job_number = models.CharField("part_number",max_length=20,null=True,unique=False,default='N/A')
    gl_account = models.CharField("gl_account",max_length=20,null=True,unique=False,default='N/A')
    closed_date = models.DateField("closed_date",null=True, blank=True)
    artwork = models.CharField("artwork",max_length=20,null=True,unique=False,default='N/A')
    panel = models.CharField("panel",max_length=20,null=True,unique=False,default='N/A')
    quadrant = models.CharField("quadrant",max_length=20,null=True,unique=False,default='N/A')
    quantity = models.IntegerField("quantity",null=True)
    failed = models.IntegerField("failed",null=True)
    failed_percent = models.FloatField("failed_percent",null=True)
    pcb_damage = models.IntegerField("pcb_damage",null=True)
    misc_mechanical = models.IntegerField("misc_mechanical",null=True)
    IL_failure = models.IntegerField("IL_failure",null=True)
    AB_failure = models.IntegerField("AB_failure",null=True)
    VSWR_failure = models.IntegerField("vswr_failure",null=True)
    ISO_failure = models.IntegerField("ISO_failure",null=True)
    