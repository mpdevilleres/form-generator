# -*- coding: utf-8 -*-
# Generated by Django 1.10 on 2016-08-16 09:28
from __future__ import unicode_literals

from decimal import Decimal
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('main', '0006_invoice'),
    ]

    operations = [
        migrations.AddField(
            model_name='purchaseorderlinedetail',
            name='rate_diff_percent',
            field=models.DecimalField(decimal_places=2, default=Decimal('0.00'), max_digits=20),
        ),
    ]
