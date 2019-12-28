import os
import re

from django.conf import settings
from django.core.exceptions import ValidationError
from django.forms import modelform_factory, inlineformset_factory
from django import forms
from xml_measurements.models import Rule, Configuration
from djangoformsetjs.utils import formset_media_js
from xml_measurements.utils import prepare_rule
from xml_measurements.xcel import generate_new_xml, FakeRule
from openpyxl.utils import FORMULAE


class RuleForm(forms.ModelForm):
    class Meta:
        model = Rule
        fields = 'names', 'rule'

    class Media:
        js = formset_media_js

    def clean(self):
        names_set = prepare_rule(self.cleaned_data.get('names', ''))
        for i, names in enumerate(names_set):
            for name in names:
                for j in range(len(names_set)):
                    if j != i and name in names_set[j]:
                        raise ValidationError("You are using the same name for 2 cells in the same rule")

        rule = self.cleaned_data.get('rule', '')
        match = re.match(r'=(.*)\(.+\)', rule)
        err = ValidationError("")
        if match:
            f_name = match.group(1)
            if f_name.upper() not in FORMULAE and f_name:
                raise err
        else:
            raise err

        # TODO: this validation is bullshit, should we remove it ?
        in_file = settings.TMP_FILE
        out = f'{settings.TMP_FILE[:-5]}_out'
        rules = [FakeRule(names=names_set, rule=rule, pk=1)]
        try:
            generate_new_xml(in_file, "COL", rules, out, 0)
        except Exception as ex:
            raise ValidationError("The rule is not xcel valid or is not supported")
        finally:
            os.remove(out)


RuleFormset = inlineformset_factory(Configuration, Rule, min_num=1, extra=0, can_delete=True, form=RuleForm)
ConfigurationForm = modelform_factory(Configuration, fields=('name', 'rules_direction'))


class UploadXlsxFileForm(forms.Form):
    sheet_num = forms.IntegerField()
    file = forms.FileField()
    configuration = forms.ModelChoiceField(queryset=Configuration.objects.all())
