from django.core.exceptions import ValidationError
from django.forms import modelform_factory, inlineformset_factory
from django import forms
from xml_measurements.models import Rule, Configuration
from djangoformsetjs.utils import formset_media_js
from xml_measurements.utils import prepare_rule


class RuleForm(forms.ModelForm):
    class Meta:
        model = Rule
        fields = 'names', 'rule'

    class Media:
        js = formset_media_js

    def clean(self):
        names_set = prepare_rule(self.cleaned_data['names'])
        for i, names in enumerate(names_set):
            for name in names:
                for j in range(len(names_set)):
                    if j != i and name in names_set[j]:
                        raise ValidationError("You are using the same name for 2 cells in the same rule")


RuleFormset = inlineformset_factory(Configuration, Rule, min_num=1, extra=0, can_delete=True, form=RuleForm)
ConfigurationForm = modelform_factory(Configuration, fields=('name', 'rules_direction'))


class UploadXlsxFileForm(forms.Form):
    sheet_num = forms.IntegerField()
    file = forms.FileField()
    configuration = forms.ModelChoiceField(queryset=Configuration.objects.all())
