# blog/forms.py
from django import forms
from .models import PolicyIssue  # Replace with your actual model name


class PolicyIssueForm(forms.ModelForm):
    class Meta:
        model = PolicyIssue  # Replace with your actual model name
        fields = '__all__'  # Include all fields or specify the ones you want to include
