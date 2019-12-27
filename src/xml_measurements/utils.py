import os

import xlrd
from django.conf import settings
from django.http import HttpResponse, Http404
from django.shortcuts import get_object_or_404
from xml_measurements.models import Configuration
from xml_measurements.xcel import generate_new_xml


def prepare_rules(rules):
    for rule in rules:
        rule.names = [[el.strip() for el in line.split(',')] for line in rule.names.splitlines()]


def download(slug, conf_pk, sheet_num):
    file_path = os.path.join(settings.MEDIA_ROOT, slug)
    if os.path.exists(file_path):
        conf = get_object_or_404(Configuration, pk=conf_pk)

        rules = conf.rule_set.all()
        prepare_rules(rules)
        out_path = f'{file_path}_out.xlsx'
        generate_new_xml(file_path, conf.rules_direction, rules, out_path, sheet_num - 1)

        with open(out_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(out_path)
            return response
    raise Http404
