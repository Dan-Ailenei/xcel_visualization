from django.forms import modelform_factory, formset_factory, inlineformset_factory
from django import forms
from xml_measurements.models import Rule, Configuration
from djangoformsetjs.utils import formset_media_js


class RuleForm(forms.ModelForm):
    class Meta:
        model = Rule
        fields = 'names', 'rule'

    class Media:
        js = formset_media_js


RuleFormset = inlineformset_factory(Configuration, Rule, min_num=1, extra=0, can_delete=True, form=RuleForm)
ConfigurationForm = modelform_factory(Configuration, fields=('name', 'rules_direction'))


class UploadXlsxFileForm(forms.Form):
    sheet_num = forms.IntegerField()
    file = forms.FileField()
    configuration = forms.ModelChoiceField(queryset=Configuration.objects.all())
