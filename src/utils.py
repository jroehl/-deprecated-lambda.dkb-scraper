from os import environ
from datetime import datetime, timedelta
import re
import locale

from config import get_config


def init():
    """
    load env variables and check mandatory vars
    """

    for var in get_config('needed_env_vars'):
        env_var = environ.get(var, None)
        if not env_var:
            raise ValueError(
                '"{}" environment variable must be set'.format(var)
            )


def parse_range(time_span, end_date, date_format):
    try:
        return end_date - timedelta(days=int(time_span))
    except:
        return datetime.strptime(time_span, date_format)


def normalize_currency(amount):
    res = re.search(r'(.*?)(?:[\.\,]{0,1})(\d+)\s*[a-zA-Z]*$', amount)
    nr = ''
    suffix = '00'
    if not res:
        nr = amount
    elif not res.group(1):
        nr = res.group(2)
    else:
        nr = res.group(1)
        suffix = res.group(2)
    return locale.currency(float('{}.{}'.format(
        nr.replace(',', '').replace('.', ''),
        suffix
    )), grouping=True, symbol=False)


def format_pattern(pattern, suffix):
    parts = pattern.split(';')
    formatted = ''
    for part in parts:
        if part != '':
            formatted += '{}"{}";'.format(part, suffix)
    return formatted


def get_format_request(sheet_id, max_row, max_col, repeat_cells=[], borders=[]):
    requests = [
        # resize everything
        {
            'autoResizeDimensions': {
                'dimensions': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': 0,
                    'endIndex': max_row
                }
            }
        },
        {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 0,
                    'endRowIndex': max_row,
                    'startColumnIndex': 0,
                    'endColumnIndex': max_col
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'RIGHT',
                    }
                },
                'fields': 'userEnteredFormat(horizontalAlignment)'
            }
        },
    ]

    for border in borders:
        requests.append({
            'updateBorders': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': border['start_row'],
                    'endRowIndex': border['end_row'],
                    'startColumnIndex': border['start_col'],
                    'endColumnIndex': border['end_col']
                },
                **border['borders']
            }
        })

    for cells in repeat_cells:
        requests.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': cells['start_row'],
                    'endRowIndex': cells['end_row'],
                    'startColumnIndex': cells['start_col'],
                    'endColumnIndex': cells['end_col']
                },
                'cell': {
                    **cells['formats']
                },
                'fields': cells['fields']
            }
        })

    return {
        'includeSpreadsheetInResponse': False,
        'requests': requests
    }
