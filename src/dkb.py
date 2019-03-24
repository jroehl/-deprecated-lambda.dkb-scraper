#!/usr/bin/env python
# encoding: utf-8

import csv
import re
from datetime import date
import requests
from lxml.html import fromstring
from fints.client import FinTS3PinTanClient

from config import get_config
from src.utils import normalize_currency


class DKBSession(object):
    """
    DKB Session
    Usage
    -----
    >>> dkbs = DKBSession()
    >>> dkbs.login('user', 'pass')
    >>> res = dkbs.get_transactions()
    >>> dkbs.logout()
    """

    def __init__(self, username, password, verbose=True):

        # Initialize HTTP session
        self.s = requests.Session()
        self.s.headers = {
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'
        }

        self.verbose = verbose
        self.__username = username
        self.__password = password
        self.__dkb_cfg = get_config('dkb')

    def login(self):
        """
        Login to DKB Online Banking
        """

        if self.verbose:
            print('Login to DKB Online Banking')

        # Get DKB Banking login page
        r = self.s.get(self.__dkb_cfg['base_url'])
        login_page = fromstring(r.text)
        banking_form = next(
            form for form in login_page.forms if form.action == '/banking')

        banking_form.fields['j_username'] = self.__username
        banking_form.fields['j_password'] = self.__password

        # Post login
        r = self.s.post(
            self.__dkb_cfg['base_url'],
            data=dict(banking_form.fields),
            params={'$javascript': 'disabled'}
        )

        # Parse returned page
        page = fromstring(r.text)
        if len(page.xpath('//*[text()="Finanzstatus"]')) == 0:
            raise RuntimeError('Login to DKB Online Banking failed.\n')
        elif self.verbose:
            print('Logged in to DKB Online Banking\n')

        self.__logout_url = page.xpath('//*/a[@id="logout"]/@href')[0]
        return True

    def logout(self):
        """
        Logout from DKB Online Banking
        """

        if self.verbose:
            print('Log out from DKB Online Banking')

        r = self.s.get(self.__dkb_cfg['base_url'] + self.__logout_url)
        ret = r.status_code == 200
        self.s.close()

        if self.verbose:
            print('Logged out from DKB Online Banking')

        return ret

    def query(self, start_date, end_date=date.today()):
        print('Querying transactions and balances between "{}" and "{}"'.format(
            start_date.date(), end_date.date()
        ))

        client = FinTS3PinTanClient(
            self.__dkb_cfg['blz'],  # Your bank's BLZ
            self.__username,  # Your login name
            self.__password,  # Your banking PIN
            self.__dkb_cfg['fints_url']
        )

        depot_accounts = {}
        credit_accounts = {}
        sepa_accounts = {}
        accounts = {}
        with client:
            info = client.get_information()
            for account in info['accounts']:
                del account['supported_operations']
                del account['bank_identifier']
                account_number = account['account_number']
                if account['type'] == 30:
                    depot_accounts[account_number] = {
                        **account,
                        'account_type': 'DEPOT',
                    }
                elif account['type'] == 50:
                    credit_accounts[account_number] = {
                        **account,
                        'account_type': 'CREDIT',
                    }
                else:
                    sepa_accounts[account_number] = {
                        **account,
                        'account_type': 'SEPA',
                    }

        client.deconstruct()

        for i, account in enumerate(depot_accounts):
            qp = {
                'slPortfolio': i,
                '$event': 'search',
                '$javascript': 'disabled'
            }
            # Parse transactions page
            res = self.__parse_csv(
                data=depot_accounts[account],
                params=qp,
            )

            accounts[account] = {
                **depot_accounts[account],
                **res,
            }

        r = self.s.get(
            get_config('SEPA.url', self.__dkb_cfg),
            params={'$event': 'init'}
        )
        init_page = fromstring(r.text)

        date_format = get_config('formats.date', self.__dkb_cfg)

        from_date = start_date.strftime(date_format)
        to_date = end_date.strftime(date_format)

        for account in credit_accounts:
            retained = account[:4] + (
                (len(account) - 8) * '*'
            ) + account[len(account) - 4:]
            found = init_page.xpath(
                '//*/option[contains(text(), "{}")]/@value'.format(retained)
            )
            if len(found) == 0:
                continue
            qp = {
                'slAllAccounts': found[0],
                'slTransactionStatus': 0,
                'slSearchPeriod': 4,
                'filterType': 'DATE_RANGE',
                'postingDate': from_date,
                'toPostingDate': to_date,
                '$event': 'search',
                '$javascript': 'disabled'
            }

            # Parse transactions page
            res = self.__parse_csv(
                data=credit_accounts[account],
                params=qp,
            )

            accounts[account] = {
                **credit_accounts[account],
                **res,
            }

        for account in sepa_accounts:
            found = init_page.xpath(
                '//*/option[contains(translate(text(), " ", ""), "{}")]/@tid'.format(account))
            if len(found) == 0:
                continue
            qp = {
                'slAllAccounts': found[0],
                'slTransactionStatus': 0,
                'slSearchPeriod': 1,
                'searchPeriodRadio': 1,
                'transactionDate': from_date,
                'toTransactionDate': to_date,
                '$event': 'search',
                '$javascript': 'disabled'
            }

            # Parse transactions page
            accounts[account] = self.__parse_csv(
                data=sepa_accounts[account],
                params=qp,
            )

        return {
            'info': {
                'start_date': from_date, 'end_date': to_date,
                'request_date': date.today().strftime(date_format)
            },
            'accounts': accounts
        }

    def __sanitize_transactions(self, currency_indices, transactions):
        for transaction in transactions:
            for i in currency_indices:
                transaction[i] = normalize_currency(transaction[i])

        return transactions

    def __parse_csv(self, data, params):
        account_type = data['account_type']
        cfg = self.__dkb_cfg[account_type]

        endpoint = cfg['url']
        total_key = cfg['keys']['total']
        fieldnames = cfg['fieldnames']

        url_string = endpoint + '?'
        ps = []
        for key in params:
            ps.append('{}={}'.format(key, params[key]))
        url_string += '&'.join(ps)
        # print(url_string)

        self.s.get(endpoint, params=params)
        params['$event'] = 'csvExport'
        download = self.s.get(endpoint, params=params)

        cr = csv.reader(
            download.text.splitlines(),
            delimiter=';'
        )

        total = ''
        transaction_begin = -1
        csv_rows = list(cr)
        for i, row in enumerate(csv_rows):
            if len(row) > 0:
                if total_key in row[0]:
                    total = row[1]
                elif row[0] == fieldnames[0]:
                    for x, cell in enumerate(row):
                        if cell != fieldnames[x]:
                            raise Exception(
                                "Header row fields differ from fieldnames array in config"
                            )
                    transaction_begin = i + 1
                    break

        indices = {}
        for i, name in enumerate(fieldnames):
            for key, value in cfg['keys'].items():
                if name in value and name != '':
                    indices[key] = [
                        i] if key not in indices else indices[key] + [i]

        transactions = self.__sanitize_transactions(
            indices['currency'],
            csv_rows[transaction_begin:]
        )
        return {
            **data,
            'total': normalize_currency(total),
            'url_string': url_string,
            'indices': indices,
            'total_key': total_key,
            'fieldnames': fieldnames,
            'transactions': transactions
        }
