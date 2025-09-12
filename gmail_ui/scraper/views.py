from pathlib import Path

from django.http import FileResponse, HttpResponse
from django.shortcuts import redirect, render
from django.contrib.auth import logout

from .gmail_amounts_to_excel import run_scraper

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR.parent / "email_amounts.xlsx"
TOKENS_DIR = Path(__file__).resolve().parent / "tokens"


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
                    context["table_html"] = df.to_html(classes="table table-striped", index=False)
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
