from pathlib import Path
import json
from datetime import date
from decimal import Decimal

import pandas as pd

from django.http import FileResponse, HttpResponse
from django.shortcuts import redirect, render
from django.contrib.auth import logout

from .forms import ContactForm, ExpenseForm, default_expense_initial
from .data_utils import (
    Contact,
    Expense,
    calculate_expense_totals,
    load_contacts,
    load_expenses,
    save_contacts,
    save_expenses,
    today_iso,
)
from .gmail_amounts_to_excel import (
    run_scraper,
    load_existing_excel,
    write_financial_summary,
)

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR.parent / "email_amounts.xlsx"
TOKENS_DIR = Path(__file__).resolve().parent / "tokens"


def classify_service(subject: str) -> str:
    """Return a coarse service category based on the email subject."""
    if not subject:
        return "Other"
    s = subject.lower()
    if any(k in s for k in ["invoice", "facture"]):
        return "Invoice"
    if any(k in s for k in ["quote", "devis", "quotation"]):
        return "Quote"
    if any(k in s for k in ["payment", "paiement", "paid"]):
        return "Payment"
    return "Other"


def extract_project(subject: str) -> str:
    """Heuristically extract a project name from the email subject."""
    if not subject:
        return "Unknown"
    for sep in [" - ", " – ", ":"]:
        if sep in subject:
            part = subject.split(sep, 1)[1].strip()
            return part or "Unknown"
    return subject.strip()


def is_connected() -> bool:
    """Return True if an OAuth token file exists."""
    return any(TOKENS_DIR.glob("token-*.json"))


def refresh_financial_summary() -> None:
    """Regenerate the financial summary sheet when expenses change."""
    if EXCEL_PATH.exists():
        df = load_existing_excel(str(EXCEL_PATH))
        write_financial_summary(EXCEL_PATH, df)


def revenue_totals() -> dict[str, Decimal]:
    """Return a mapping of currency to revenue total from the Excel file."""
    if not EXCEL_PATH.exists():
        return {}
    df = load_existing_excel(str(EXCEL_PATH))
    if df.empty or "amount_currency" not in df.columns:
        return {}
    totals = (
        df.groupby("amount_currency")["amount_value"].sum().to_dict()
    )
    return {str(currency): Decimal(str(total)) for currency, total in totals.items()}


def home(request):
    """Display results and allow the user to run the scraper."""
    connected = is_connected()
    today = date.today()
    current_period = today.strftime("%Y-%m")
    if request.session.get("period") != current_period:
        request.session["period"] = current_period
        request.session["total_amount"] = 0.0
    total_amount = request.session.get("total_amount", 0.0)
    context = {"connected": connected, "period": current_period}
    if request.method == "POST":
        if not connected:
            context["error"] = "Please log in first."
        else:
            try:
                df = run_scraper()
                if not df.empty:
                    df = df.copy()
                    if "sender_name" not in df.columns:
                        df["sender_name"] = ""
                    df["sender_name"] = df["sender_name"].fillna("").astype(str).str.strip()
                    if "client_name" not in df.columns:
                        df["client_name"] = ""
                    df["client_name"] = df["client_name"].fillna("").astype(str).str.strip()
                    if "sender_email" not in df.columns:
                        df["sender_email"] = ""
                    df["sender_email"] = df["sender_email"].fillna("").astype(str).str.strip()

                    missing_name = df["sender_name"] == ""
                    df.loc[missing_name, "sender_name"] = df.loc[missing_name, "sender_email"]
                    df["sender_name"] = df["sender_name"].replace("", "Unknown")

                    df["client_display"] = df["client_name"].replace("", pd.NA).fillna("Unknown")

                    display_df = df.copy()
                    name_idx = (
                        display_df.columns.get_loc("client_name")
                        if "client_name" in display_df.columns
                        else 0
                    )
                    display_df.insert(name_idx, "Client", display_df["client_display"])
                    display_df.drop(
                        columns=[
                            col
                            for col in [
                                "sender_name",
                                "sender_email",
                                "client_name",
                                "client_display",
                            ]
                            if col in display_df.columns
                        ],
                        inplace=True,
                    )
                    context["table_html"] = display_df.to_html(
                        classes="table table-striped table-hover table-sm w-100", index=False
                    )
                    totals = df.groupby("amount_currency")["amount_value"].sum().reset_index()
                    context["totals_html"] = totals.to_html(
                        classes="table table-striped table-hover table-sm w-100", index=False
                    )

                    df["service"] = df["subject"].apply(classify_service)
                    df["project"] = df["subject"].apply(extract_project)

                    clients = (
                        df.groupby("client_display")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    clients_display = clients.rename(columns={"client_display": "Client"})
                    context["clients_html"] = clients_display.to_html(
                        classes="table table-striped table-hover table-sm w-100", index=False
                    )
                    context["clients_chart"] = json.dumps(
                        {
                            "labels": clients_display["Client"].tolist(),
                            "values": clients_display["amount_value"].tolist(),
                        }
                    )

                    projects = (
                        df.groupby("project")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["projects_html"] = projects.to_html(
                        classes="table table-striped table-hover table-sm w-100", index=False
                    )

                    services = (
                        df.groupby("service")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["services_html"] = services.to_html(
                        classes="table table-striped table-hover table-sm w-100", index=False
                    )
                    context["services_chart"] = json.dumps(
                        {
                            "labels": services["service"].tolist(),
                            "values": services["amount_value"].tolist(),
                        }
                    )
                    # Each run should reflect only the current month's total, not a
                    # cumulative sum across runs.  Replace the session value with the
                    # freshly computed total rather than adding to it.
                    total_amount = df["amount_value"].sum()
                    request.session["total_amount"] = total_amount
            except Exception as exc:
                context["error"] = str(exc)
    context["total_amount"] = total_amount
    badge_thresholds = [
        ("Bronze", 5000, "fa-solid fa-medal", "#cd7f32"),
        ("Silver", 10000, "fa-solid fa-award", "#c0c0c0"),
        ("Gold", 20000, "fa-solid fa-trophy", "#ffd700"),
        ("Platinum", 30000, "fa-solid fa-crown", "#e5e4e2"),
        ("Diamond", 50000, "fa-solid fa-gem", "#b9f2ff"),
    ]
    badges = []
    current_badge = None
    next_badge = None
    next_target = None
    for name, threshold, icon, color in badge_thresholds:
        unlocked = total_amount >= threshold
        badges.append({"name": name, "unlocked": unlocked, "icon": icon, "color": color})
        if unlocked:
            current_badge = name
        elif not next_badge:
            next_badge = name
            next_target = threshold
    context["badge"] = current_badge
    context["badges"] = badges
    context["next_target"] = next_target
    context["next_badge"] = next_badge
    return render(request, "scraper/home.html", context)


def expenses(request):
    expenses_list = load_expenses()
    if request.method == "POST":
        action = request.POST.get("action")
        if action == "delete":
            try:
                index = int(request.POST.get("index", "-1"))
            except ValueError:
                index = -1
            if 0 <= index < len(expenses_list):
                expenses_list.pop(index)
                save_expenses(expenses_list)
                refresh_financial_summary()
            return redirect("expenses")

        form = ExpenseForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            expense = Expense(
                date=(data["date"].isoformat() if data["date"] else today_iso()),
                category=data["category"],
                description=data.get("description", ""),
                amount=data["amount"],
                currency=data["currency"],
            )
            expenses_list.append(expense)
            save_expenses(expenses_list)
            refresh_financial_summary()
            return redirect("expenses")
        expense_form = form
    else:
        expense_form = ExpenseForm(initial=default_expense_initial())

    expense_totals = calculate_expense_totals(expenses_list)
    revenue = revenue_totals()
    currencies = sorted(set(expense_totals) | set(revenue))
    net_profit = {
        currency: revenue.get(currency, Decimal("0")) - expense_totals.get(currency, Decimal("0"))
        for currency in currencies
    }

    context = {
        "form": expense_form,
        "expenses": list(enumerate(expenses_list)),
        "expense_totals": sorted(expense_totals.items()),
        "revenue_totals": sorted(revenue.items()),
        "net_profit": sorted(net_profit.items()),
        "has_excel": EXCEL_PATH.exists(),
    }
    return render(request, "scraper/expenses.html", context)


def contacts(request):
    contact_list = load_contacts()
    if request.method == "POST":
        action = request.POST.get("action")
        if action == "delete":
            try:
                index = int(request.POST.get("index", "-1"))
            except ValueError:
                index = -1
            if 0 <= index < len(contact_list):
                contact_list.pop(index)
                save_contacts(contact_list)
            return redirect("contacts")

        form = ContactForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            contact = Contact(
                name=data["name"],
                company=data.get("company", ""),
                category=data["category"],
                phone=data["phone"],
                email=data.get("email", ""),
                notes=data.get("notes", ""),
            )
            contact_list.append(contact)
            save_contacts(contact_list)
            return redirect("contacts")
        contact_form = form
    else:
        contact_form = ContactForm()

    context = {
        "form": contact_form,
        "contacts": list(enumerate(contact_list)),
    }
    return render(request, "scraper/contacts.html", context)


def download_excel(request):
    if EXCEL_PATH.exists():
        return FileResponse(open(EXCEL_PATH, "rb"), as_attachment=True, filename="email_amounts.xlsx")
    return HttpResponse("File not found", status=404)


def login_view(request):
    """Trigger OAuth login and redirect to home."""
    from .gmail_amounts_to_excel import load_creds_for_account

    load_creds_for_account(None)
    return redirect("home")


def logout_view(request):
    """Log the user out, remove tokens, and redirect to the home page."""
    logout(request)
    for token_file in TOKENS_DIR.glob("token-*.json"):
        try:
            token_file.unlink()
        except OSError:
            pass
    return redirect("home")
