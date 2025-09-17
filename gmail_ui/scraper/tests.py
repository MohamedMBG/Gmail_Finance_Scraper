import json

from django.http import HttpResponse
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

    @patch('scraper.views.render')
    @patch('scraper.views.run_scraper')
    @patch('scraper.views.is_connected', return_value=True)
    def test_client_name_preferred_over_email(self, mock_is_connected, mock_run_scraper, mock_render):
        df = pd.DataFrame({
            'amount_value': [100, 200],
            'amount_currency': ['USD', 'USD'],
            'subject': ['Invoice A', 'Invoice B'],
            'sender_name': ['Client One', ''],
            'sender_email': ['client1@example.com', 'client2@example.com'],
        })
        mock_run_scraper.return_value = df

        captured = {}

        def render_side_effect(request, template_name, context):
            captured['context'] = context
            return HttpResponse('ok')

        mock_render.side_effect = render_side_effect

        request = self.factory.post('/')
        request.session = {}
        home(request)

        context = captured['context']
        table_html = context['table_html']
        self.assertIn('Client One', table_html)
        self.assertNotIn('client1@example.com', table_html)

        clients_chart = json.loads(context['clients_chart'])
        self.assertIn('Client One', clients_chart['labels'])
        self.assertIn('client2@example.com', clients_chart['labels'])
        self.assertNotIn('client1@example.com', clients_chart['labels'])


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
