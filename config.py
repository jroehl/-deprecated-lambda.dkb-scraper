from os import environ
import locale
import jmespath
from dotenv import load_dotenv

load_dotenv()
env = environ.get('STAGE', None)


def __merge(source, destination):
    for key, value in source.items():
        if isinstance(value, dict):
            # get node or create one
            node = destination.setdefault(key, {})
            __merge(value, node)
        else:
            destination[key] = value

    return destination


def get_config(path=None, data=None):
    cfg = config['default']
    if env in config:
        cfg = __merge(config[env], cfg)
    if path and data:
        return jmespath.search(path, data)
    if path:
        return jmespath.search(path, cfg)
    return cfg


__base_url = 'https://www.dkb.de/banking'
config = {
    'default': {
        'env': env,
        'needed_env_vars': [
            'DKB_USER',
            'DKB_PASSWORD',
            'GOOGLE_SHEET_WRITER',
            'CREDS_CLIENT_EMAIL',
            'CREDS_PRIVATE_KEY',
            'CREDS_CLIENT_ID',
            'CREDS_PRIVATE_KEY_ID',
            'CREDS_TOKEN_URI'
        ],
        'dkb': {
            'currency': environ.get('DKB_CURRENCY', '€'),
            'creds': {
                'username': environ.get('DKB_USER', None),
                'password': environ.get('DKB_PASSWORD', None),
            },
            'formats': {
                'date': '%d.%m.%Y',
                'datetime': '%d.%m.%Y %H:%M:%S',
            },
            'blz': '12030000',
            'fints_url': 'https://banking-dkb.s-fints-pt-dkb.de/fints30',
            'base_url': __base_url,
            'CREDIT': {
                'merge_values': True,
                'keys': {
                    'total': 'Saldo:',
                    'date': 'Wertstellung',
                    'currency': 'Betrag (EUR)'
                },
                'url': '{}/finanzstatus/kreditkartenumsaetze'.format(__base_url),
                'fieldnames': [
                    'Umsatz abgerechnet und nicht im Saldo enthalten',
                    'Wertstellung',
                    'Belegdatum',
                    'Beschreibung',
                    'Betrag (EUR)',
                    'Ursprünglicher Betrag',
                    ''
                ]
            },
            'SEPA': {
                'merge_values': True,
                'keys': {
                    'total': 'Kontostand',
                    'date': 'Buchungstag',
                    'currency': 'Betrag (EUR)'
                },
                'url': '{}/finanzstatus/kontoumsaetze'.format(__base_url),
                'fieldnames': [
                    'Buchungstag',
                    'Wertstellung',
                    'Buchungstext',
                    'Auftraggeber / Begünstigter',
                    'Verwendungszweck',
                    'Kontonummer',
                    'BLZ',
                    'Betrag (EUR)',
                    'Gläubiger-ID',
                    'Mandatsreferenz',
                    'Kundenreferenz',
                    ''
                ]
            },
            'DEPOT': {
                'merge_values': False,
                'keys': {
                    'total': 'Depotgesamtwert',
                    'currency': [
                        'Kurs',
                        'Gewinn / Verlust',
                        'Einstandswert',
                        'Dev. Kurs',
                        'Kurswert in Euro'
                    ]
                },
                'url': '{}/depotstatus'.format(__base_url),
                'fieldnames': [
                    'Bestand',
                    '',
                    'ISIN / WKN',
                    'Bezeichnung',
                    'Kurs',
                    'Gewinn / Verlust',
                    '',
                    'Einstandswert',
                    '',
                    'Dev. Kurs',
                    'Kurswert in Euro',
                    'Verfügbarkeit',
                    ''
                ]
            }
        },
        'gsheet': {
            'creds': {
                'private_key': environ.get('CREDS_PRIVATE_KEY', None),
                'client_id': environ.get('CREDS_CLIENT_ID', None),
                'client_email': environ.get('CREDS_CLIENT_EMAIL', None),
                'private_key_id': environ.get('CREDS_PRIVATE_KEY_ID', None),
                'token_uri': environ.get('CREDS_TOKEN_URI', None),
            },
            'locale':  locale.setlocale(locale.LC_ALL, ''),
            'sheet_name': environ.get('GOOGLE_SHEET_NAME', 'dkb-finance-dashboard'),
            'generated_values_ws_name': environ.get('GOOGLE_SHEET_GENVALUES_WS', 'GENERATED VALUES'),
            'sheet_writer': environ.get('GOOGLE_SHEET_WRITER', None),
            'formats': {
                'currency': '[<0][Red]-#,##0.00;[>0][Green]#,##0.00;[Blue]#,##0.00;'
            },
            'scope': [
                'https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive'
            ],
            'whitelisted': [
                'account_number',
                'customer_id',
                'currency',
                'owner_name',
                'product_name',
                'account_type',
                'total'
            ]
        }
    },
    'dev': {
    }
}
