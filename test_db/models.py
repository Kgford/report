# This is an auto-generated Django model module.
# You'll have to do the following manually to clean this up:
#   * Rearrange models' order
#   * Make sure each model has one field with primary_key=True
#   * Make sure each ForeignKey and OneToOneField has `on_delete` set to the desired behavior
#   * Remove `managed = False` lines if you wish to allow Django to create, modify, and delete the table
# Feel free to rename the models, but don't rename db_table values or field names.
from django.db import models
from django.utils import timezone


class ReportQueue(models.Model):    
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    reportname = models.CharField(db_column='ReportName', max_length=50, blank=True, null=True)  # Field name made lowercase.
    reporttype = models.CharField(db_column='ReportType', max_length=50, blank=True, null=True)  # Field name made lowercase.
    reportstatus = models.CharField(db_column='ReportStatus', max_length=50, blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=50, blank=True, null=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=50, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=50, blank=True, null=True)  # Field name made lowercase.
    operator = models.CharField(db_column='Operator', max_length=50, blank=True, null=True)  # Field name made lowercase.
    activedate = models.DateTimeField(db_column='ActiveDate', blank=True, null=True)  # Field name made lowercase.
    percentcomplete = models.IntegerField(db_column='percentcomplete')  # Field name made lowercase. 
    value = models.IntegerField(db_column='value')  # Field name made lowercase. 
    maxvalue = models.IntegerField(db_column='maxvalue')  # Field name made lowercase. 
    
    class Meta:
        managed = False
        db_table = 'ReportQueue'


class Coeffcal(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    stringname = models.CharField(db_column='StringName', max_length=50, blank=True, null=True)  # Field name made lowercase.
    string = models.CharField(db_column='String', max_length=255, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'CoeffCal'


class Devices(models.Model):
    devname = models.CharField(db_column='DevName', max_length=100, blank=True, null=True)  # Field name made lowercase.
    manufacturer = models.CharField(db_column='Manufacturer', max_length=100, blank=True, null=True)  # Field name made lowercase.
    devtype = models.CharField(db_column='DevType', max_length=100, blank=True, null=True)  # Field name made lowercase.
    defaultaddr = models.FloatField(db_column='DefaultAddr', blank=True, null=True)  # Field name made lowercase.
    timeout_ms = models.FloatField(db_column='Timeout_ms', blank=True, null=True)  # Field name made lowercase.
    termin = models.CharField(db_column='Termin', max_length=100, blank=True, null=True)  # Field name made lowercase.
    termout = models.CharField(db_column='Termout', max_length=100, blank=True, null=True)  # Field name made lowercase.
    eoi = models.BooleanField(db_column='EOI', blank=True, null=True)  # Field name made lowercase.
    idnresponse = models.CharField(db_column='IDNResponse', max_length=100, blank=True, null=True)  # Field name made lowercase.
    notes = models.CharField(db_column='Notes', max_length=100, blank=True, null=True)  # Field name made lowercase.
    manual = models.BooleanField(db_column='Manual', blank=True, null=True)  # Field name made lowercase.
    safecommand = models.CharField(db_column='SafeCommand', max_length=100, blank=True, null=True)  # Field name made lowercase.
    saferesponse = models.CharField(db_column='SafeResponse', max_length=100, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'DEVICES'


class Effeciency(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=255, blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=255, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=255, blank=True, null=True)  # Field name made lowercase.
    operator = models.CharField(db_column='Operator', max_length=255, blank=True, null=True)  # Field name made lowercase.
    activedate = models.CharField(db_column='ActiveDate', max_length=255, blank=True, null=True)  # Field name made lowercase.
    totaluuts = models.BigIntegerField(db_column='TotalUUTs', blank=True, null=True)  # Field name made lowercase.
    completeuuts = models.BigIntegerField(db_column='CompleteUUTs', blank=True, null=True)  # Field name made lowercase.
    effeciencystatus = models.CharField(db_column='EffeciencyStatus', max_length=255, blank=True, null=True)  # Field name made lowercase.
    runstatus = models.CharField(db_column='RunStatus', max_length=255, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'Effeciency'


class Graphdb3Db(models.Model):
    statename = models.CharField(db_column='StateName', max_length=255, blank=True, null=True)  # Field name made lowercase.
    pass_field = models.IntegerField(db_column='PASS', blank=True, null=True)  # Field name made lowercase. Field renamed because it was a Python reserved word.
    fail = models.IntegerField(db_column='Fail', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'GRAPHDB_3dB'


class GraphdbDir(models.Model):
    statename = models.CharField(max_length=255, blank=True, null=True)
    pass_field = models.IntegerField(db_column='PASS', blank=True, null=True)  # Field name made lowercase. Field renamed because it was a Python reserved word.
    fail = models.IntegerField(db_column='Fail', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'GRAPHDB_Dir'


class Portconfig(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j1j1 = models.CharField(db_column='J1J1', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j1j2 = models.CharField(db_column='J1J2', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j1j3 = models.CharField(db_column='J1J3', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j1j4 = models.CharField(db_column='J1J4', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j1j5 = models.CharField(db_column='J1J5', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j2j1 = models.CharField(db_column='J2J1', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j2j2 = models.CharField(db_column='J2J2', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j2j3 = models.CharField(db_column='J2J3', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j2j4 = models.CharField(db_column='J2J4', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j3j1 = models.CharField(db_column='J3J1', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j3j2 = models.CharField(db_column='J3J2', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j3j3 = models.CharField(db_column='J3J3', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j3j4 = models.CharField(db_column='J3J4', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j4j1 = models.CharField(db_column='J4J1', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j4j2 = models.CharField(db_column='J4J2', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j4j3 = models.CharField(db_column='J4J3', max_length=100, blank=True, null=True)  # Field name made lowercase.
    j4j4 = models.CharField(db_column='J4J4', max_length=100, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'PortConfig'


class Specifications(models.Model):
    specid = models.AutoField(db_column='SpecID', primary_key=True)  # Field name made lowercase.
    spectype = models.CharField(db_column='SpecType', max_length=100, blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    title = models.CharField(db_column='Title', max_length=100, blank=True, null=True)  # Field name made lowercase.
    quantity = models.IntegerField(db_column='Quantity', blank=True, null=True)  # Field name made lowercase.
    startfreqmhz = models.FloatField(db_column='StartFreqMHz', blank=True, null=True)  # Field name made lowercase.
    stopfreqmhz = models.FloatField(db_column='StopFreqMHz', blank=True, null=True)  # Field name made lowercase.
    cutofffreqmhz = models.FloatField(db_column='CutOffFreqMHz', blank=True, null=True)  # Field name made lowercase.
    outputportnumber = models.FloatField(db_column='OutputPortNumber', blank=True, null=True)  # Field name made lowercase.
    vswr = models.FloatField(db_column='VSWR', blank=True, null=True)  # Field name made lowercase.
    insertionloss = models.FloatField(db_column='InsertionLoss', blank=True, null=True)  # Field name made lowercase.
    isolation = models.FloatField(db_column='Isolation', blank=True, null=True)  # Field name made lowercase.
    isolation2 = models.FloatField(db_column='Isolation2', blank=True, null=True)  # Field name made lowercase.
    amplitudebalance = models.FloatField(db_column='AmplitudeBalance', blank=True, null=True)  # Field name made lowercase.
    coupling = models.FloatField(db_column='Coupling', blank=True, null=True)  # Field name made lowercase.
    coupplusminus = models.FloatField(db_column='COUPPlusMinus', blank=True, null=True)  # Field name made lowercase.
    directivity = models.FloatField(db_column='Directivity', blank=True, null=True)  # Field name made lowercase.
    phasebalance = models.FloatField(db_column='PhaseBalance', blank=True, null=True)  # Field name made lowercase.
    coupledflatness = models.FloatField(db_column='CoupledFlatness', blank=True, null=True)  # Field name made lowercase.
    power = models.FloatField(db_column='Power', blank=True, null=True)  # Field name made lowercase.
    temperature = models.FloatField(db_column='Temperature', blank=True, null=True)  # Field name made lowercase.
    offset1 = models.FloatField(db_column='Offset1', blank=True, null=True)  # Field name made lowercase.
    offset2 = models.FloatField(db_column='Offset2', blank=True, null=True)  # Field name made lowercase.
    offset3 = models.FloatField(db_column='Offset3', blank=True, null=True)  # Field name made lowercase.
    offset4 = models.FloatField(db_column='Offset4', blank=True, null=True)  # Field name made lowercase.
    offset5 = models.FloatField(db_column='Offset5', blank=True, null=True)  # Field name made lowercase.
    test1 = models.IntegerField(db_column='Test1', blank=True, null=True)  # Field name made lowercase.
    test2 = models.IntegerField(db_column='Test2', blank=True, null=True)  # Field name made lowercase.
    test3 = models.IntegerField(db_column='Test3', blank=True, null=True)  # Field name made lowercase.
    test4 = models.IntegerField(db_column='Test4', blank=True, null=True)  # Field name made lowercase.
    test5 = models.IntegerField(db_column='Test5', blank=True, null=True)  # Field name made lowercase.
    pph = models.FloatField(db_column='PPH')  # Field name made lowercase.
    po = models.CharField(db_column='PO', max_length=1000, blank=True, null=True)  # Field name made lowercase.
    datecode = models.CharField(db_column='DateCode', max_length=100, blank=True, null=True)  # Field name made lowercase.
    bypass = models.IntegerField(db_column='Bypass')  # Field name made lowercase.
    password = models.CharField(db_column='Password', max_length=150, blank=True, null=True)  # Field name made lowercase.
    globalfail = models.FloatField(db_column='GlobalFail', blank=True, null=True)  # Field name made lowercase.
    testfail = models.FloatField(db_column='TestFail', blank=True, null=True)  # Field name made lowercase.
    retestfail = models.FloatField(db_column='RetestFail', blank=True, null=True)  # Field name made lowercase.
    failpercent = models.FloatField(db_column='FailPercent', blank=True, null=True)  # Field name made lowercase.
    ab_exp_tf = models.BooleanField(db_column='AB_exp_tf', blank=True, null=True)  # Field name made lowercase.
    ab_ex = models.FloatField(db_column='AB_ex', blank=True, null=True)  # Field name made lowercase.
    ab_start1 = models.FloatField(db_column='AB_start1', blank=True, null=True)  # Field name made lowercase.
    ab_start2 = models.FloatField(db_column='AB_start2', blank=True, null=True)  # Field name made lowercase.
    ab_stop1 = models.FloatField(db_column='AB_stop1', blank=True, null=True)  # Field name made lowercase.
    ab_stop2 = models.FloatField(db_column='AB_stop2', blank=True, null=True)  # Field name made lowercase.
    ab_tf = models.IntegerField(db_column='AB_tf', blank=True, null=True)  # Field name made lowercase.
    pb_exp_tf = models.BooleanField(db_column='PB_exp_tf', blank=True, null=True)  # Field name made lowercase.
    pb_ex = models.FloatField(db_column='PB_ex', blank=True, null=True)  # Field name made lowercase.
    pb_start1 = models.FloatField(db_column='PB_start1', blank=True, null=True)  # Field name made lowercase.
    pb_start2 = models.FloatField(db_column='PB_start2', blank=True, null=True)  # Field name made lowercase.
    pb_stop1 = models.FloatField(db_column='PB_stop1', blank=True, null=True)  # Field name made lowercase.
    pb_stop2 = models.FloatField(db_column='PB_stop2', blank=True, null=True)  # Field name made lowercase.
    pb_tf = models.IntegerField(db_column='PB_tf', blank=True, null=True)  # Field name made lowercase.
    iso_exp_tf = models.BooleanField(db_column='ISO_exp_tf', blank=True, null=True)  # Field name made lowercase.
    iso_ex = models.FloatField(db_column='ISO_ex', blank=True, null=True)  # Field name made lowercase.
    iso_start1 = models.FloatField(db_column='ISO_start1', blank=True, null=True)  # Field name made lowercase.
    iso_start2 = models.FloatField(db_column='ISO_start2', blank=True, null=True)  # Field name made lowercase.
    iso_stop1 = models.FloatField(db_column='ISO_stop1', blank=True, null=True)  # Field name made lowercase.
    iso_stop2 = models.FloatField(db_column='ISO_stop2', blank=True, null=True)  # Field name made lowercase.
    iso_tf = models.IntegerField(db_column='ISO_tf', blank=True, null=True)  # Field name made lowercase.
    coup_exp_tf = models.BooleanField(db_column='COUP_exp_tf', blank=True, null=True)  # Field name made lowercase.
    coup_ex = models.FloatField(db_column='COUP_ex', blank=True, null=True)  # Field name made lowercase.
    coup_start1 = models.FloatField(db_column='COUP_start1', blank=True, null=True)  # Field name made lowercase.
    coup_start2 = models.FloatField(db_column='COUP_start2', blank=True, null=True)  # Field name made lowercase.
    coup_stop1 = models.FloatField(db_column='COUP_stop1', blank=True, null=True)  # Field name made lowercase.
    coup_stop2 = models.FloatField(db_column='COUP_stop2', blank=True, null=True)  # Field name made lowercase.
    coup_tf = models.IntegerField(db_column='COUP_tf', blank=True, null=True)  # Field name made lowercase.
    dir_exp_tf = models.BooleanField(db_column='DIR_exp_tf', blank=True, null=True)  # Field name made lowercase.
    dir_ex = models.FloatField(db_column='DIR_ex', blank=True, null=True)  # Field name made lowercase.
    dir_start1 = models.FloatField(db_column='DIR_start1', blank=True, null=True)  # Field name made lowercase.
    dir_start2 = models.FloatField(db_column='DIR_start2', blank=True, null=True)  # Field name made lowercase.
    dir_stop1 = models.FloatField(db_column='DIR_stop1', blank=True, null=True)  # Field name made lowercase.
    dir_stop2 = models.FloatField(db_column='DIR_stop2', blank=True, null=True)  # Field name made lowercase.
    dir_tf = models.IntegerField(db_column='DIR_tf', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'Specifications'


class Testdata(models.Model):
    testid = models.AutoField(db_column='TestID', primary_key=True)  # Field name made lowercase.
    specid = models.IntegerField(db_column='SpecID', blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    serialnumber = models.CharField(db_column='SerialNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=100, blank=True, null=True)  # Field name made lowercase.
    insertionloss = models.FloatField(db_column='InsertionLoss', blank=True, null=True)  # Field name made lowercase.
    returnloss = models.FloatField(db_column='ReturnLoss', blank=True, null=True)  # Field name made lowercase.
    coupling = models.FloatField(db_column='Coupling', blank=True, null=True)  # Field name made lowercase.
    isolation = models.FloatField(db_column='Isolation', blank=True, null=True)  # Field name made lowercase.
    directivity = models.FloatField(db_column='Directivity', blank=True, null=True)  # Field name made lowercase.
    amplitudebalance = models.FloatField(db_column='AmplitudeBalance', blank=True, null=True)  # Field name made lowercase.
    coupledflatness = models.FloatField(db_column='CoupledFlatness', blank=True, null=True)  # Field name made lowercase.
    phasebalance = models.FloatField(db_column='PhaseBalance', blank=True, null=True)  # Field name made lowercase.
    failurelog = models.CharField(db_column='FailureLog', max_length=255, blank=True, null=True)  # Field name made lowercase.
    artwork_rev = models.CharField(db_column='artwork_rev',max_length=50, blank=True, null=True)
    amplitudebalance1 = models.FloatField(db_column='AmplitudeBalance1', blank=True, null=True)  # Field name made lowercase.
    amplitudebalance2 = models.FloatField(db_column='AmplitudeBalance2', blank=True, null=True)  # Field name made lowercase.
    isolation2 = models.FloatField(db_column='Isolation2', blank=True, null=True)  # Field name made lowercase.
    operator = models.CharField(db_column='Operator', max_length=100, blank=True, null=True)  # Field name made lowercase.
    returnloss2 = models.FloatField(db_column='ReturnLoss2', blank=True, null=True)  # Field name made lowercase.
    coupling2 = models.FloatField(db_column='Coupling2', blank=True, null=True)  # Field name made lowercase.
    directivity2 = models.FloatField(db_column='Directivity2', blank=True, null=True)  # Field name made lowercase.
    phasebalance2 = models.FloatField(db_column='PhaseBalance2', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TestData'


class Testdata3(models.Model):
    testid = models.AutoField(db_column='TestID', primary_key=True)  # Field name made lowercase.
    specid = models.IntegerField(db_column='SpecID', blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    partnumber = models.CharField(db_column='PartNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    serialnumber = models.CharField(db_column='SerialNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=100, blank=True, null=True)  # Field name made lowercase.
    insertionloss = models.FloatField(db_column='InsertionLoss', blank=True, null=True)  # Field name made lowercase.
    returnloss = models.FloatField(db_column='ReturnLoss', blank=True, null=True)  # Field name made lowercase.
    coupling = models.FloatField(db_column='Coupling', blank=True, null=True)  # Field name made lowercase.
    isolation = models.FloatField(db_column='Isolation', blank=True, null=True)  # Field name made lowercase.
    directivity = models.FloatField(db_column='Directivity', blank=True, null=True)  # Field name made lowercase.
    amplitudebalance = models.FloatField(db_column='AmplitudeBalance', blank=True, null=True)  # Field name made lowercase.
    coupledflatness = models.FloatField(db_column='CoupledFlatness', blank=True, null=True)  # Field name made lowercase.
    phasebalance = models.FloatField(db_column='PhaseBalance', blank=True, null=True)  # Field name made lowercase.
    failurelog = models.FloatField(db_column='FailureLog', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TestData3'


class Trace(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    testid = models.IntegerField(db_column='TestID', blank=True, null=True)  # Field name made lowercase.
    specid = models.IntegerField(db_column='SpecID', blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    title = models.CharField(db_column='Title', max_length=100, blank=True, null=True)  # Field name made lowercase.
    serialnumber = models.CharField(db_column='SerialNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=100, blank=True, null=True)  # Field name made lowercase.
    points = models.IntegerField(db_column='Points', blank=True, null=True)  # Field name made lowercase.
    activedate = models.DateTimeField(db_column='ActiveDate', blank=True, null=True)  # Field name made lowercase.
    rfpower = models.FloatField(db_column='RFPower', blank=True, null=True)  # Field name made lowercase.
    temperature = models.FloatField(db_column='Temperature', blank=True, null=True)  # Field name made lowercase.
    calibrationdate = models.DateTimeField(db_column='CalibrationDate', blank=True, null=True)  # Field name made lowercase.
    instrumentcaldue = models.DateTimeField(db_column='InstrumentCalDue', blank=True, null=True)  # Field name made lowercase.
    progtitle = models.CharField(db_column='ProgTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    progversion = models.CharField(db_column='ProgVersion', max_length=100, blank=True, null=True)  # Field name made lowercase.
    xtitle = models.CharField(db_column='XTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    ytitle = models.CharField(db_column='YTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    notes = models.CharField(db_column='Notes', max_length=255, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'Trace'


class Traceimage(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    testid = models.IntegerField(db_column='TestID', blank=True, null=True)  # Field name made lowercase.
    specid = models.IntegerField(db_column='SpecID', blank=True, null=True)  # Field name made lowercase.
    jobnumber = models.CharField(db_column='JobNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    title = models.CharField(db_column='Title', max_length=100, blank=True, null=True)  # Field name made lowercase.
    serialnumber = models.CharField(db_column='SerialNumber', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstation = models.CharField(db_column='WorkStation', max_length=100, blank=True, null=True)  # Field name made lowercase.
    points = models.IntegerField(db_column='Points', blank=True, null=True)  # Field name made lowercase.
    activedate = models.DateTimeField(db_column='ActiveDate', blank=True, null=True)  # Field name made lowercase.
    rfpower = models.FloatField(db_column='RFPower', blank=True, null=True)  # Field name made lowercase.
    temperature = models.FloatField(db_column='Temperature', blank=True, null=True)  # Field name made lowercase.
    calibrationdate = models.DateTimeField(db_column='CalibrationDate', blank=True, null=True)  # Field name made lowercase.
    instrumentcaldue = models.DateTimeField(db_column='InstrumentCalDue', blank=True, null=True)  # Field name made lowercase.
    progtitle = models.CharField(db_column='ProgTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    progversion = models.CharField(db_column='ProgVersion', max_length=100, blank=True, null=True)  # Field name made lowercase.
    xtitle = models.CharField(db_column='XTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    ytitle = models.CharField(db_column='YTitle', max_length=100, blank=True, null=True)  # Field name made lowercase.
    notes = models.CharField(db_column='Notes', max_length=255, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TraceImage'


class Traceimagepoints(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    traceid = models.IntegerField(db_column='TraceID', blank=True, null=True)  # Field name made lowercase.
    idx = models.IntegerField(db_column='Idx', blank=True, null=True)  # Field name made lowercase.
    xdata = models.FloatField(db_column='Xdata', blank=True, null=True)  # Field name made lowercase.
    ydata = models.FloatField(db_column='Ydata', blank=True, null=True)  # Field name made lowercase.
    coeffcal = models.FloatField(db_column='CoeffCal', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TraceImagePoints'


class Tracepoints(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    traceid = models.IntegerField(db_column='TraceID', blank=True, null=True)  # Field name made lowercase.
    idx = models.IntegerField(db_column='Idx', blank=True, null=True)  # Field name made lowercase.
    xdata = models.FloatField(db_column='Xdata', blank=True, null=True)  # Field name made lowercase.
    ydata = models.FloatField(db_column='Ydata', blank=True, null=True)  # Field name made lowercase.
    coeffcal = models.FloatField(db_column='CoeffCal', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TracePoints'


class Tracepoints2(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    traceid = models.IntegerField(db_column='TraceID', blank=True, null=True)  # Field name made lowercase.
    idx = models.IntegerField(db_column='Idx', blank=True, null=True)  # Field name made lowercase.
    xdata = models.FloatField(db_column='Xdata', blank=True, null=True)  # Field name made lowercase.
    ydata = models.FloatField(db_column='Ydata', blank=True, null=True)  # Field name made lowercase.
    coeffcal = models.FloatField(db_column='CoeffCal', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TracePoints2'


class Tracestr(models.Model):
    id  = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    traceid = models.IntegerField(db_column='TraceID')  # Field name made lowercase.
    xdata = models.TextField(db_column='XData', blank=True, null=True)  # Field name made lowercase.
    ydata = models.TextField(db_column='YData', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'TraceStr'


class Workstation(models.Model):
    id = models.AutoField(db_column='ID', primary_key=True)  # Field name made lowercase.
    computername = models.CharField(db_column='ComputerName', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstationname = models.CharField(db_column='WorkstationName', max_length=100, blank=True, null=True)  # Field name made lowercase.
    vnatype = models.CharField(db_column='VNAType', max_length=100, blank=True, null=True)  # Field name made lowercase.
    operator = models.CharField(db_column='Operator', max_length=50, blank=True, null=True)  # Field name made lowercase.
    vnafreq = models.FloatField(db_column='VNAFreq', blank=True, null=True)  # Field name made lowercase.
    password = models.CharField(db_column='Password', max_length=150, blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'WorkStation'


class Workstation1(models.Model):
    id = models.IntegerField(db_column='ID', primary_key=True)  # Field name made lowercase.
    computername = models.CharField(db_column='ComputerName', max_length=100, blank=True, null=True)  # Field name made lowercase.
    workstationname = models.CharField(db_column='WorkstationName', max_length=100, blank=True, null=True)  # Field name made lowercase.
    vnatype = models.CharField(db_column='VNAType', max_length=100, blank=True, null=True)  # Field name made lowercase.
    operator = models.CharField(db_column='Operator', max_length=50, blank=True, null=True)  # Field name made lowercase.
    vnafreq = models.FloatField(db_column='VNAFreq', blank=True, null=True)  # Field name made lowercase.

    class Meta:
        managed = False
        db_table = 'WorkStation1'


class Sysdiagrams(models.Model):
    name = models.CharField(max_length=128)
    principal_id = models.IntegerField()
    diagram_id = models.AutoField(primary_key=True)
    version = models.IntegerField(blank=True, null=True)
    definition = models.BinaryField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'sysdiagrams'
        unique_together = (('principal_id', 'name'),)



# Create your models here.
class Test_Events(models.Model):
    job_number = models.CharField("job_number",max_length=50,null=True,unique=False,default='N/A')  
    part_number = models.CharField("part_number",max_length=50,null=True,unique=False,default='N/A')
    order_list_header_id = models.IntegerField(db_column='order_list_header_id')  # Field name made lowercase. 
    order_list_detail_id = models.IntegerField(db_column='order_list_detail_id ')  # Field name made lowercase.
    event_type = models.CharField(db_column='event_type', max_length=20)  # Field name made lowercase.
    event_time = models.DateTimeField(db_column='event_time', blank=True, null=True)  # Field name made lowercase.
    description = models.CharField("description",max_length=50,null=True,unique=False,default='N/A')  
    priority = models.CharField(db_column='Priority', max_length=50, blank=True, null=True)  # Field name made lowercase.
    severity = models.CharField(db_column='Severity', max_length=50, blank=True, null=True)  # Field name made lowercase.
    impact = models.CharField(db_column='Impact', max_length=50, blank=True, null=True)  # Field name made lowercase.
    cost = models.DecimalField(db_column='cost', max_digits=18, decimal_places=8, blank=True, null=True)  # Field name made lowercase.
    update_by = models.CharField("update_by",max_length=50,null=False,unique=False,default='N/A')  
    timestamp = models.DateTimeField(default=timezone.now)
    def __str__(self):
        return "%s %s" % (self.event_type, self.event_time)
    
    class Meta:
        managed = True
        db_table = 'test_events' 