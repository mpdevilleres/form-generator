# -*- coding: utf-8 -*-
# Generated by Django 1.10 on 2016-08-11 08:30
from __future__ import unicode_literals

from decimal import Decimal
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('main', '0004_resource_po_os_ref'),
    ]

    operations = [
        migrations.CreateModel(
            name='UnitPrice',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('contractor', models.CharField(blank=True, max_length=255)),
                ('po_position', models.CharField(blank=True, max_length=255)),
                ('po_level', models.CharField(blank=True, max_length=255)),
                ('amount', models.DecimalField(decimal_places=2, default=Decimal('0.00'), max_digits=20)),
                ('percent', models.IntegerField(default=0)),
            ],
        ),
    ]
