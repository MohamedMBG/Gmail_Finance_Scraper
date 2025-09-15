from django.test import RequestFactory, TestCase
from unittest.mock import patch
import pandas as pd

from .views import home
from .gmail_amounts_to_excel import extract_amounts


class HomeViewTests(TestCase):
    def setUp(self):
        self.factory = RequestFactory()

    @patch('scraper.views.is_connected', return_value=True)
    @patch('scraper.views.run_scraper')
    def test_month_total_does_not_accumulate(self, mock_run_scraper, mock_is_connected):
        df = pd.DataFrame({
            'amount_value': [10, 20],
            'amount_currency': ['USD', 'USD'],
            'subject': ['Invoice - ProjectX', 'Payment - ProjectY'],
            'sender_name': ['Client1', 'Client2'],
        })
        mock_run_scraper.return_value = df

        request1 = self.factory.post('/')
        request1.session = {}
        home(request1)
        self.assertEqual(request1.session['total_amount'], 30)

        request2 = self.factory.post('/')
        request2.session = request1.session
        home(request2)
        self.assertEqual(request2.session['total_amount'], 30)


class AmountParsingTests(TestCase):
    def test_space_and_nbsp_equivalence(self):
        texts = [
            'Total: 8000 USD',
            'Total: 8 000 USD',
            'Total: 8\u00a0000 USD',
        ]
        for t in texts:
            amts = extract_amounts(t)
            self.assertEqual(len(amts), 1)
            self.assertEqual(amts[0]['value'], 8000.0)
            self.assertEqual(amts[0]['currency'], 'USD')
