from django import forms

class UploadFileForm(forms.Form):
    email_id = forms.EmailField(widget=forms.TextInput())
    file = forms.FileField()
    fallout_report=forms.BooleanField(required=False)
