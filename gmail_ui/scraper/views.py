import subprocess
import sys
from pathlib import Path

import pandas as pd
from django.http import FileResponse, HttpResponse
from django.shortcuts import redirect, render

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR.parent / "email_amounts.xlsx"
SCRIPT_PATH = BASE_DIR.parent / "gmail_amounts_to_excel.py"


def home(request):
    """Display results and allow the user to run the scraper."""
    context = {}
    if request.method == "POST":
        # Run the scraping script and surface any errors to the user
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH)],
            capture_output=True,
            text=True,
            cwd=BASE_DIR.parent,
        )
        if result.returncode == 0:
            return redirect("home")
        context["error"] = result.stderr or "Scraper failed to run"

    if EXCEL_PATH.exists():
        df = pd.read_excel(EXCEL_PATH)
        context["table_html"] = df.to_html(classes="table table-striped", index=False)
    return render(request, "scraper/home.html", context)


def download_excel(request):
    if EXCEL_PATH.exists():
        return FileResponse(open(EXCEL_PATH, "rb"), as_attachment=True, filename="email_amounts.xlsx")
    return HttpResponse("File not found", status=404)
