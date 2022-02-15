from django.db import models
from django.utils import timezone

class Trouble_Ticket(models.Model):
    TYPE_CHOICES = (
        ('Server Error 500', 'Server Error 500'),
        ('Server Not Running', 'Server Not Running'),
        ('PDF not Moving to Network', 'PDF not Moving to Network'),
        ('Item not Closing', 'Item not Closing'),
        ('Date Error', 'Date Error'),
        ('Date Code Error', 'Date Code Error'),
        ('C of C Scaling', 'C of C Scaling'),
        ('Label Scaling', 'Label Scaling'),
        ('Label Maker Software', 'Label Maker Software'),
        ('E2 Failure', 'E2 Failure'),
        ('Database Error', 'Database Error'),
        ('Printer Error', 'Printer Error'),
        ('Printer Low', 'Printer Low'),
        ('Network Error', 'Network Error'),
        ('Windows Error', 'Windows Error'),
        ('Misc', 'Misc')
    )
    APP_CHOICES = (
        ('ATE Application', 'ATE Application'),
        ('Windows Application', 'Windows Application'),
        ('Web Application', 'Web Application'),
        ('Database Application', 'Database Application'),
        ('E2 Application', 'E2 Application'),
        ('QA Process Priority', 'QA Process Priority'),
        ('Stockroom Process Priority', 'Stockroom Process Priority'),
        ('Windows Label Maker', 'Windows Label Maker'),
        ('Robot Test Fixture', 'Robot Test Fixture'),
        ('ATE Test Fixture', 'ATE Test Fixture'),
        ('Manfacturing Test Fixture', 'Manfacturing Test Fixture'),
        ('Inventory Application', 'Inventory Application')
    )
    LOCATION_CHOICES = (
        ('Stock Room', 'Stock Room'),
        ('Inspection', 'Inspection'),
        ('Shipping', 'Shipping'),
        ('Sales', 'Sales'),
        ('Test', 'Test'),
        ('Manfacturing', 'Manfacturing'),
        ('Machine Shop', 'Machine Shop'),
    )
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    application_type = models.CharField(db_column='application_type', choices = APP_CHOICES, max_length=100)  # Field name made lowercase.
    trouble_status = models.CharField(db_column='trouble_status', max_length=20)  # Field name made lowercase.
    trouble_type = models.CharField(db_column='trouble_type', choices = TYPE_CHOICES, max_length=100)  # Field name made lowercase.
    job_number = models.CharField("job_number",max_length=50,null=True,unique=False,default='N/A')  
    part_number = models.CharField("part_number",max_length=50,null=True,unique=False,default='N/A')
    packing_slip = models.CharField("packing_slip",max_length=50,null=True,unique=False,default='N/A')
    open_time = models.DateTimeField(db_column='open_time', blank=True, null=True)  # Field name made lowercase.
    description = models.CharField("description",max_length=500,null=True,unique=False,default='N/A') 
    solution = models.CharField("solution",max_length=500,null=True,unique=False,default='N/A')  
    reported_by = models.CharField("reported_by",max_length=50,null=True,unique=False,default='N/A')    
    location = models.CharField("location",max_length=100,null=True,unique=False, choices = LOCATION_CHOICES, default='N/A')
    close_time = models.DateTimeField(db_column='close_time', blank=True, null=True)  # Field name made lowercase.
    update_by = models.CharField("update_by",max_length=50,null=False,unique=False,default='N/A')  
    timestamp = models.DateTimeField(default=timezone.now)
    def __str__(self):
        return "%s %s %s" % (self.trouble_type, self.job_number, self.part_number)
    
    class Meta:
        managed = True
        db_table = 'trouble_ticket' 

# Create your models here.
