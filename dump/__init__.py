# -*- coding: utf-8 -*-

import datetime as dt
import os, csv

# models
from decimal import Decimal, InvalidOperation

import subprocess

from main.models import PurchaseOrder, PurchaseOrderLine, PurchaseOrderLineDetail, Resource, UnitPrice, Invoice

dumpdir = os.path.dirname(os.path.realpath(__file__))


def clean_csv(csvpath):
    csvbak = '{}.bak'.format(csvpath)
    os.rename(csvpath, csvbak)
    subprocess.call("""cat {} | sed -e '/[^"]$/N' -e 's/\n//g' >> {}""".format(csvbak, csvpath))
    os.remove(csvbak)


def id_reset_to(model=None, pk=2001):
    table_name = model._meta.db_table

    from django.db import connection
    cursor = connection.cursor()
    cursor.execute("ALTER SEQUENCE {0}_id_seq RESTART WITH {1};".format(table_name, pk))


def to_bool(data):
    if isinstance(data, str):
        data = data.lower()
        if data == "0" or data == "false":
            return False
        elif data == "1" or data == "true":
            return True

    return NotImplemented


def to_date_format(string):
    # remove time element
    string = string.split(' ')[0]
    try:
        return dt.datetime.strptime(string, '%d/%m/%Y')
    except ValueError:
        return None


def to_dec(data):
    try:
        return Decimal(data)
    except InvalidOperation:
        return Decimal('0')


def dump_purchase_order(file='TechPO.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            po = PurchaseOrder.objects.filter(id=row["POID"])

            if len(po) != 0:
                print('Purchase Order ID# %s Exist' % row["POID"])
                continue
            else:
                po = PurchaseOrder()

                po.po_num = row["PONum"]

                # Need Date Fixing
                # po.po_date = to_date_format(row["PODate"])
                po.pk = row["POID"]
                po.po_value = to_dec(row["POValue"])
                po.contractor = row["Contractor"]
                po.budget = row["Budget"]
                po.capex_commitment_value = to_dec(row["CPXComValue"])
                po.capex_expenditure_value = to_dec(row["CPXExpValue"])
                po.opex_value = to_dec(row["OPXValue"])
                po.revenue_value = to_dec(row["REVValue"])
                po.task_num = row["TaskNum"]
                po.renew_status = row["RenewStatus"]
                po.renew_po_no = row["RenewPONum"]
                po.po_status = row["POStatus"]
                po.po_remarks = row["PORemarks"]
                po.po_type = row["POType"]
                po.save()


def dump_purchase_order_line(file='TechPOLine.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            po_line = PurchaseOrderLine.objects.filter(id=row["POLineID"])
            if len(po_line) != 0:
                print('Purchase Order Line ID# %s Exist' % row["POID"])
                continue
            else:

                po = PurchaseOrder.objects.filter(id=row["POID"])
                if len(po) == 0:
                    print("Purchase Order ID# %s Does't Exist" % row["POID"])
                    continue
                else:
                    po_line = PurchaseOrderLine()

                    po_line.po_num = po.first()

                    po_line.pk = row["POLineID"]
                    po_line.line_num = row["POLineNum"]
                    po_line.line_duration = row["POLineDuration"]
                    po_line.line_value = to_dec(row["POLineValue"])
                    po_line.line_revise_rate = to_dec(row["POLineRValue"])
                    po_line.line_rate = to_dec(row["POLineRate"])
                    po_line.line_status = row["POLineStatus"]
                    po_line.line_actuals = row["POLineActuals"]
                    po_line.capex_percent = row["CPXPercent"]
                    po_line.opex_percent = row["OPXPercent"]
                    po_line.revenue_percent = row["REVPercent"]
                    po_line.save()


def dump_purchase_order_line_details(file='TechPOLineDetail.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            po_line_detail = PurchaseOrderLineDetail.objects.filter(id=row["PODetID"])
            if len(po_line_detail) != 0:
                print('Purchase Order Line Detail ID# %s Exist' % row["PODetID"])
                continue
            else:

                po_line = PurchaseOrderLine.objects.filter(id=row["POLineID"])
                if len(po_line) == 0:
                    print("Purchase Order Line ID# %s Does't Exist" % row["POLineID"])
                    continue
                else:
                    po_line_detail = PurchaseOrderLineDetail()

                    po_line_detail.po_line = po_line.first()

                    po_line_detail.pk = row["PODetID"]
                    po_line_detail.po_os_ref = row["POOSRef"]
                    po_line_detail.po_position = row["POPosition"]
                    po_line_detail.po_level = row["POLevel"]
                    po_line_detail.po_rate = to_dec(row["PORate"])
                    po_line_detail.po_revise_rate = to_dec(row["PORRate"])
                    po_line_detail.division = row["Division"]
                    po_line_detail.section = row["Section"]
                    po_line_detail.sub_section = row["SubSection"]
                    po_line_detail.director_name = row["DirName"]
                    po_line_detail.frozen_status = row["FrozenStatus"]
                    po_line_detail.approval_ref_num = row["ApprovalRefNum"]
                    po_line_detail.approval_reason = row["ApprovalReason"]
                    po_line_detail.kpi_2016 = row["2016KPI"]
                    po_line_detail.capex_percent = row["CPX%Age"]
                    po_line_detail.opex_percent = row["OPX%Age"]
                    po_line_detail.revenue_percent = row["REV%Age"]
                    po_line_detail.rate_diff_percent = to_dec(row['PercentDiffRate'])
                    po_line_detail.save()


def dump_resources(file='RPT01_ Monthly Accruals - Mobillized.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath, encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            resource = Resource.objects.filter(id=row["ResID"]).first()
            if resource:
                print('Resource ID# %s Exist' % row["ResID"])
                continue
            else:
                resource = Resource()

#                resource.po_id = row["ResID"]
                resource.po_line_detail_id = row["ResID"]

                resource.pk = row["ResID"]
                resource.res_type = row['ResType']
                resource.res_type_class = row['ResTypeClass']
                resource.agency_ref_num = row['AgencyRefNum']
                resource.po_os_ref = row['POOSRef']
                resource.res_emp_num = row['ResEmpNum']
                resource.res_full_name = row['ResFullName']
                resource.date_of_join = to_date_format(row['DoJ'])
                resource.po_position = row['POPosition']
                resource.po_level = row['POLevel']
                resource.division = row['Division']
                resource.section = row['Section']
                resource.manager = row['Manager']
                resource.rate = row['Rate']
                resource.capex_percent = row['CPXPercent']
                resource.capex_rate = row['CAPEXRate']
                resource.opex_percent = row['OPXPercent']
                resource.opex_rate = row['OPEXRate']
                resource.revenue_percent = row['REVPercent']
                resource.revenue_rate = row['REVENUERate']
                resource.remarks = row['Remarks']
                resource.po_num = row['PONum']
                resource.po_value = row['POValue']
                resource.capex_commitment_value = row['CPXComValue']
                resource.opex_value = row['OPXValue']
                resource.revenue_value = row['REVValue']
                resource.contractor = row['Contractor']
                resource.save()


def dump_unit_price(file='UnitPrice.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            unit_price = UnitPrice.objects.filter(id=row["id"]).first()
            if unit_price:
                print('Resource ID# %s Exist' % row["id"])
                continue
            else:
                unit_price = UnitPrice()

                unit_price.pk = row["id"]
                unit_price.po_level = row['po_level']
                unit_price.po_position = row['po_position']
                unit_price.contractor = row['contractor']
                unit_price.amount = to_dec(row['amount'])
                unit_price.percent = row['percent']
                unit_price.save()


def dump_invoice(file='TechInvoiceMgmt.csv'):
    csvpath = os.path.join(dumpdir, file)
    #    clean_csv(csvpath)
    with open(csvpath) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            invoice = Invoice.objects.filter(id=row["InvID"]).first()
            if invoice:
                print('Invoice ID# %s Exist' % row["InvID"])
                continue
            else:
                invoice = Invoice()

                invoice.pk = row['InvID']
                invoice.resource_id = row['ResID']
                invoice.po_line_detail_id = row['PODetID']
                invoice.invoice_date = to_date_format(row['InvMonth'])
                invoice.invoice_hour = row['InvMonthHrs']
                invoice.invoice_claim = row['InvClaimHrs']
                invoice.invoice_cert_amount = Decimal(row['InvCertAmt'])
                invoice.remarks = row['Remarks']
                invoice.save()


def start():
    dump_purchase_order()
    dump_purchase_order_line()
    dump_purchase_order_line_details()
    dump_resources()
    dump_unit_price()
    dump_invoice()

if __name__ == '__main__':
    from ma import *
    PurchaseOrderLineDetail.objects.delete()
    dump_purchase_order_line_details()