import typer
from .scanner import run_scraper

app = typer.Typer(help="Gmail finance scraper")

@app.command()
def scan(
    days: int = typer.Option(30, help="Search in the last N days. 0 = all mail."),
    exclude_label: list[str] = typer.Option(None, help="Gmail labels to exclude"),
    min: float = typer.Option(0.0, help="Minimum amount to include"),
    out: str = typer.Option("excel", help="Output format (excel)"),
):
    """Scan the inbox and extract monetary amounts."""
    run_scraper(days=days, exclude_labels=exclude_label, min_amount=min)
    if out.lower() == "excel":
        typer.echo("Saved results to email_amounts.xlsx")

if __name__ == "__main__":
    app()
