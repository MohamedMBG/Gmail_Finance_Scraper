from pathlib import Path

from django.http import FileResponse, HttpResponse
from django.shortcuts import redirect, render
from django.contrib.auth import logout

from .gmail_amounts_to_excel import run_scraper

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR.parent / "email_amounts.xlsx"


def home(request):
    """Display results and allow the user to run the scraper."""
    context = {}
    if request.method == "POST":
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


def logout_view(request):
    """Log the user out and redirect to the home page."""
    logout(request)
    return redirect("home")
