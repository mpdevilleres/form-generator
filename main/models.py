from decimal import Decimal
from django.db import models

# Create your models here.

class PurchaseOrder(models.Model):

    po_num = models.CharField(blank=True, max_length=255)
    po_date = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    po_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    contractor = models.CharField(blank=True, max_length=255)
    budget = models.CharField(blank=True, max_length=255)
    capex_commitment_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    capex_expenditure_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    opex_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    revenue_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    task_num = models.CharField(blank=True, max_length=255)
    renew_status = models.CharField(blank=True, max_length=255)
    renew_po_no = models.CharField(blank=True, max_length=255)
    po_status = models.CharField(blank=True, max_length=255)
    po_remarks = models.TextField(blank=True)
    po_type = models.CharField(blank=True, max_length=255)

    def __str__(self):
        return self.po_num


class PurchaseOrderLine(models.Model):

    po_num = models.ForeignKey(PurchaseOrder, null=True, on_delete=models.CASCADE)

    line_num = models.CharField(blank=True, max_length=255)
    line_duration = models.IntegerField(default=0)
    line_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    line_revise_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    line_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    line_status = models.CharField(blank=True, max_length=255)
    line_actuals = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    capex_percent = models.IntegerField(default=0)
    opex_percent = models.IntegerField(default=0)
    revenue_percent = models.IntegerField(default=0)

    @property
    def rate(self):
        return self.line_revise_rate if self.line_revise_rate != Decimal('0.00') else self.line_rate

    def __str__(self):
        return self.line_num


class PurchaseOrderLineDetail(models.Model):

    po_line = models.ForeignKey(PurchaseOrderLine, null=True, on_delete=models.CASCADE)
    po_os_ref = models.CharField(blank=True, max_length=255)
    po_position = models.CharField(blank=True, max_length=255)
    po_level = models.CharField(blank=True, max_length=255)
    po_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    po_revise_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    division = models.CharField(blank=True, max_length=255)
    section = models.CharField(blank=True, max_length=255)
    sub_section = models.CharField(blank=True, max_length=255)
    director_name = models.CharField(blank=True, max_length=255)
    frozen_status = models.CharField(blank=True, max_length=255)
    approval_ref_num = models.CharField(blank=True, max_length=255)
    approval_reason = models.TextField(blank=True)
    kpi_2016 = models.CharField(blank=True, max_length=255)
    capex_percent = models.IntegerField(default=0)
    opex_percent = models.IntegerField(default=0)
    revenue_percent = models.IntegerField(default=0)
    rate_diff_percent = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))

    @property
    def rate(self):
        return self.po_revise_rate if self.po_revise_rate != Decimal('0.00') else self.po_rate

    def __str__(self):
        return self.po_os_ref


class Resource(models.Model):

    po_line_detail = models.ForeignKey(PurchaseOrderLineDetail, null=True, on_delete=models.CASCADE)
    po = models.ForeignKey(PurchaseOrder, null=True, on_delete=models.CASCADE)

    res_type = models.CharField(blank=True, max_length=255)
    res_type_class = models.CharField(blank=True, max_length=255)
    agency_ref_num = models.CharField(blank=True, max_length=255)
    po_os_ref = models.CharField(blank=True, max_length=255)
    res_emp_num = models.CharField(blank=True, max_length=255)
    res_full_name = models.CharField(blank=True, max_length=255)
    date_of_join = models.DateField(blank=True, null=True)
    po_position = models.CharField(blank=True, max_length=255)
    po_level = models.CharField(blank=True, max_length=255)
    division = models.CharField(blank=True, max_length=255)
    section = models.CharField(blank=True, max_length=255)
    manager = models.CharField(blank=True, max_length=255)
    rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    capex_percent = models.IntegerField(default=0)
    capex_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    opex_percent = models.IntegerField(default=0)
    opex_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    revenue_percent = models.IntegerField(default=0)
    revenue_rate = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    remarks = models.TextField(blank=True)
    po_num = models.CharField(blank=True, max_length=255)
    po_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    capex_commitment_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    opex_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    revenue_value = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    contractor = models.CharField(blank=True, max_length=255)

class UnitPrice(models.Model):

    contractor = models.CharField(blank=True, max_length=255)
    po_position = models.CharField(blank=True, max_length=255)
    po_level = models.CharField(blank=True, max_length=255)
    amount = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    percent = models.IntegerField(default=0)


class Invoice(models.Model):

    resource = models.ForeignKey(Resource, null=True, on_delete=models.CASCADE)
    po_line_detail = models.ForeignKey(PurchaseOrderLineDetail, null=True, on_delete=models.CASCADE)

    invoice_date = models.DateField(blank=True, null=True)
    invoice_hour = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    invoice_claim = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    invoice_cert_amount = models.DecimalField(max_digits=20, decimal_places=2, default=Decimal('0.00'))
    remarks = models.TextField(blank=True)