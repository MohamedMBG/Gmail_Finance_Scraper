from __future__ import annotations

from datetime import date

from django import forms

CURRENCY_CHOICES = [
    ("MAD", "MAD"),
    ("USD", "USD"),
    ("EUR", "EUR"),
    ("GBP", "GBP"),
]

CONTACT_CATEGORY_CHOICES = [
    ("cleaning", "Cleaning"),
    ("repair", "Repair"),
    ("promotion", "Promotion"),
    ("maintenance", "Maintenance"),
    ("other", "Other"),
]


class ExpenseForm(forms.Form):
    date = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={"type": "date", "class": "form-control"}),
        help_text="Leave empty to use today's date.",
    )
    category = forms.CharField(
        max_length=100,
        widget=forms.TextInput(attrs={"class": "form-control"}),
    )
    description = forms.CharField(
        max_length=255,
        required=False,
        widget=forms.Textarea(attrs={"rows": 2, "class": "form-control"}),
    )
    amount = forms.DecimalField(
        max_digits=12,
        decimal_places=2,
        min_value=0,
        widget=forms.NumberInput(attrs={"class": "form-control", "step": "0.01"}),
    )
    currency = forms.ChoiceField(
        choices=CURRENCY_CHOICES,
        initial="MAD",
        widget=forms.Select(attrs={"class": "form-select"}),
    )


class ContactForm(forms.Form):
    name = forms.CharField(
        max_length=120,
        widget=forms.TextInput(attrs={"class": "form-control"}),
    )
    company = forms.CharField(
        max_length=120,
        required=False,
        widget=forms.TextInput(attrs={"class": "form-control"}),
        help_text="Optional company or organization name.",
    )
    category = forms.ChoiceField(
        choices=CONTACT_CATEGORY_CHOICES,
        initial="cleaning",
        widget=forms.Select(attrs={"class": "form-select"}),
    )
    phone = forms.CharField(
        max_length=50,
        widget=forms.TextInput(attrs={"class": "form-control"}),
    )
    email = forms.EmailField(
        required=False,
        widget=forms.EmailInput(attrs={"class": "form-control"}),
    )
    notes = forms.CharField(
        required=False,
        widget=forms.Textarea(attrs={"rows": 2, "class": "form-control"}),
    )

    def clean(self):
        cleaned = super().clean()
        phone = cleaned.get("phone")
        email = cleaned.get("email")
        if not phone and not email:
            raise forms.ValidationError("Provide at least a phone number or an email address.")
        return cleaned


def default_expense_initial() -> dict[str, date | str]:
    return {"date": date.today()}
