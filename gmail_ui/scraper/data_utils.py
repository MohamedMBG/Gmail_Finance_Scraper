from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Iterable, List

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

EXPENSES_FILE = DATA_DIR / "expenses.json"
CONTACTS_FILE = DATA_DIR / "contacts.json"


@dataclass
class Expense:
    date: str
    category: str
    description: str
    amount: Decimal
    currency: str


@dataclass
class Contact:
    name: str
    company: str
    category: str
    phone: str
    email: str
    notes: str


def _read_json(path: Path, default):
    try:
        if not path.exists():
            return default
        with path.open("r", encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return default


def _write_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2, ensure_ascii=False)


def load_expenses() -> List[Expense]:
    raw = _read_json(EXPENSES_FILE, [])
    expenses: List[Expense] = []
    for item in raw:
        try:
            amount = Decimal(str(item.get("amount", "0")))
        except Exception:
            amount = Decimal("0")
        date_str = item.get("date") or ""
        expenses.append(
            Expense(
                date=date_str,
                category=str(item.get("category", "")),
                description=str(item.get("description", "")),
                amount=amount,
                currency=str(item.get("currency", "MAD")) or "MAD",
            )
        )
    expenses.sort(key=lambda e: e.date or "", reverse=True)
    return expenses


def save_expenses(expenses: Iterable[Expense]) -> None:
    payload = [
        {
            "date": expense.date,
            "category": expense.category,
            "description": expense.description,
            "amount": str(expense.amount),
            "currency": expense.currency,
        }
        for expense in sorted(expenses, key=lambda e: e.date or "", reverse=True)
    ]
    _write_json(EXPENSES_FILE, payload)


def calculate_expense_totals(expenses: Iterable[Expense]) -> dict[str, Decimal]:
    totals: dict[str, Decimal] = {}
    for expense in expenses:
        currency = expense.currency or "MAD"
        totals[currency] = totals.get(currency, Decimal("0")) + expense.amount
    return totals


def load_contacts() -> List[Contact]:
    raw = _read_json(CONTACTS_FILE, [])
    contacts: List[Contact] = []
    for item in raw:
        contacts.append(
            Contact(
                name=str(item.get("name", "")),
                company=str(item.get("company", "")),
                category=str(item.get("category", "other")) or "other",
                phone=str(item.get("phone", "")),
                email=str(item.get("email", "")),
                notes=str(item.get("notes", "")),
            )
        )
    contacts.sort(key=lambda c: (c.category.lower(), c.name.lower()))
    return contacts


def save_contacts(contacts: Iterable[Contact]) -> None:
    payload = [
        {
            "name": contact.name,
            "company": contact.company,
            "category": contact.category,
            "phone": contact.phone,
            "email": contact.email,
            "notes": contact.notes,
        }
        for contact in sorted(contacts, key=lambda c: (c.category.lower(), c.name.lower()))
    ]
    _write_json(CONTACTS_FILE, payload)


def today_iso() -> str:
    return date.today().isoformat()
