from pathlib import Path
import json

from django.http import FileResponse, HttpResponse
from django.shortcuts import redirect, render
from django.contrib.auth import logout

from .gmail_amounts_to_excel import run_scraper

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


def extract_tool(sender_email: str) -> str:
    """Return a simplified tool name from the sender's email domain."""
    if not sender_email:
        return "Unknown"
    return sender_email.split("@")[-1].lower()

def is_connected() -> bool:
    """Return True if an OAuth token file exists."""
    return any(TOKENS_DIR.glob("token-*.json"))


def home(request):
    """Display results and allow the user to run the scraper."""
    connected = is_connected()
    context = {"connected": connected}
    if request.method == "POST":
        if not connected:
            context["error"] = "Please log in first."
        else:
            try:
                df = run_scraper()
                if not df.empty:
                    context["table_html"] = df.to_html(
                        classes="table table-striped table-hover", index=False
                    )
                    totals = df.groupby("amount_currency")["amount_value"].sum().reset_index()
                    context["totals_html"] = totals.to_html(
                        classes="table table-striped table-hover", index=False
                    )

                    df["service"] = df["subject"].apply(classify_service)
                    df["project"] = df["subject"].apply(extract_project)

                    clients = (
                        df.groupby("sender_name")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["clients_html"] = clients.to_html(
                        classes="table table-striped table-hover", index=False
                    )
                    context["clients_chart"] = json.dumps(
                        {
                            "labels": clients["sender_name"].tolist(),
                            "values": clients["amount_value"].tolist(),
                        }
                    )

                    projects = (
                        df.groupby("project")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["projects_html"] = projects.to_html(
                        classes="table table-striped table-hover", index=False
                    )

                    services = (
                        df.groupby("service")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["services_html"] = services.to_html(
                        classes="table table-striped table-hover", index=False
                    )
                    context["services_chart"] = json.dumps(
                        {
                            "labels": services["service"].tolist(),
                            "values": services["amount_value"].tolist(),
                        }
                    )

                    df["tool"] = df["sender_email"].apply(extract_tool)
                    tools = (
                        df.groupby("tool")["amount_value"].sum()
                        .reset_index()
                        .sort_values("amount_value", ascending=False)
                    )
                    context["tools_html"] = tools.to_html(
                        classes="table table-striped table-hover", index=False
                    )
                    if not tools.empty:
                        context["top_tool"] = tools.iloc[0].to_dict()
            except Exception as exc:
                context["error"] = str(exc)
    return render(request, "scraper/home.html", context)


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
