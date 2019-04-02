from datetime import datetime
import gspread
from gspread.utils import rowcol_to_a1
from gspread.urls import SPREADSHEETS_API_V4_BASE_URL
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client.crypt import Signer

from config import get_config
from src.utils import format_pattern, get_format_request

DRIVE_V3_URL = 'https://www.googleapis.com/drive/v3/files'
ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def parse_time(row, indices):
    row = row.split(';')
    date_idx = indices['date'][0]
    try:
        return datetime.strptime(row[date_idx], '%x')
    except ValueError as e:
        print(e)
        print(row, date_idx)
        return datetime.fromtimestamp(0)

class GSheet(object):
    """
    Google Sheet
    Usage
    -----
    >>> gsheet = GSheet()
    >>> gsheet.add_data(data)
    """

    def __init__(self, verbose=True):
        self.verbose = verbose
        cfg = get_config()
        self.__sheet_cfg = cfg['gsheet']
        self.__dkb_cfg = cfg['dkb']
        self.__creds = self.__authenticate()
        self.__gc = gspread.authorize(self.__creds)

        self.currency_pattern = format_pattern(
            get_config('formats.currency', self.__sheet_cfg),
            ' ' + self.__dkb_cfg['currency']
        )

        sheet_name = self.__sheet_cfg['sheet_name']
        try:
            # self.__delete_spreadsheet(sheet_name)
            self.__sh = self.__gc.open(sheet_name)
        except:

            _id = self.__create_spreadsheet(sheet_name)
            self.__sh = self.__gc.open_by_key(_id)
            print('Created spreadsheet "{}" please check your mail to gain access'.format(
                sheet_name)
            )

    def __authenticate(self):
        if self.verbose:
            print('Authenticating for Google Sheets')

        creds = self.__sheet_cfg['creds']
        signer = Signer.from_string(creds['private_key'])
        scope = self.__sheet_cfg['scope']

        return ServiceAccountCredentials(
            client_id=creds['client_id'],
            service_account_email=creds['client_email'],
            private_key_id=creds['private_key_id'],
            token_uri=creds['token_uri'],
            scopes=scope,
            signer=signer
        )

    def __delete_spreadsheet(self, title):
        res = self.__gc.list_spreadsheet_files()
        files = filter(
            lambda x: x['name'] == title or x['name'] == 'dkb-finances',
            res
        )

        for res in files:
            url = '{0}/{1}'.format(
                DRIVE_V3_URL, res['id']
            )
            res = self.__gc.request('delete', url)

    def __create_spreadsheet(self, title):
        res = self.__gc.request('post', SPREADSHEETS_API_V4_BASE_URL, json={
            'properties': {
                'locale':  self.__sheet_cfg['locale'].split('.')[0],
                'title': title,
            },
            'sheets': [
                {
                    'properties': {
                        'title': 'DASHBOARD',
                        'sheetType': 'GRID',
                        'gridProperties': {
                            'rowCount': 100,
                            'columnCount': 20,
                        },
                        'tabColor': {
                            'red': 0,
                            'green': 1,
                            'blue': 0,
                        }
                    },
                },
                {
                    'properties': {
                        'title': self.__sheet_cfg['generated_values_ws_name'],
                        'sheetType': 'GRID',
                        'gridProperties': {
                            'rowCount': 20,
                            'columnCount': 20,
                            'hideGridlines': True
                        },
                        'tabColor': {
                            'red': 0,
                            'green': 0,
                            'blue': 1,
                        }
                    },
                }
            ],
        })
        json = res.json()

        spreadsheet_id = json['spreadsheetId']
        url = '{0}/{1}/permissions'.format(
            'https://www.googleapis.com/drive/v3/files', spreadsheet_id
        )
        self.__gc.request('post', url, json={
            'type': 'user',
            'role': 'writer',
            'emailAddress': self.__sheet_cfg['sheet_writer'],
        })
        return json['spreadsheetId']

    def update_dashboard(self, data):
        title = self.__sheet_cfg['generated_values_ws_name']
        if self.verbose:
            print('Updating dashboard "{}"'.format(title))

        ws = self.__sh.worksheet(title=title)

        whitelisted = self.__sheet_cfg['whitelisted']

        account_values = []
        row = -1
        accounts = data['accounts']
        header = list(filter(
            lambda x: x in whitelisted,
            list(accounts[list(accounts.keys())[0]].keys())
        ))

        now = datetime.now().strftime('%x %X')
        sheet_header = [
            title,
            '(changes will be overwritten)'
        ]
        sheet_header = sheet_header + \
            ([''] * (len(header) - len(sheet_header)))
        sheet_header[len(sheet_header) - 1] = now
        sheet_header[len(sheet_header) - 2] = 'updated_at'

        rows = [
            sheet_header,
            [],
            header,
        ]

        start_row_format = None
        max_row = None
        start_col_format = None
        end_col_format = None

        start_idx = len(rows)
        query_row = []
        query_header = []
        for i, account in enumerate(accounts):
            row = i + (start_idx)
            rows.append([])
            col = 0
            account_values = accounts[account]
            account_cfg = self.__dkb_cfg[account_values['account_type']]
            if account_cfg['display_sums']:
                indices = account_values['indices']
                query = '''=QUERY(
                        {{ARRAYFORMULA(IF(LEN('{sheet}'!{date_col}2:{date_col}){arg_separator} EOMONTH('{sheet}'!{date_col}2:{date_col};0){arg_separator} "")){array_separator}'{sheet}'!{amount_col}2:{amount_col}}}{arg_separator}
                        "{formula}"
                    )'''.format(**{
                    'sheet': account_values['title'],
                    'date_col': ALPHABET[indices['date'][0]],
                    'amount_col': ALPHABET[indices['currency'][0]],
                    'formula': 'select Col1, sum(Col2) group by Col1 order by Col1 desc label Col1\'MONTHS\', sum(Col2)\'SUM\' format Col1\'MMMM YYYY\'',
                    'array_separator': ' \\ ' if account_values['has_decimal_comma'] else ', ',
                    'arg_separator': ';' if account_values['has_decimal_comma'] else ','
                })
                query_header = query_header + [account, '', '']
                query_row = query_row + [query, '', '']

            for key in header:
                if key in account_values:
                    rows[row].append(str(account_values[key]))
                if key == 'total':
                    if not start_row_format:
                        start_row_format = row
                    max_row = row
                    start_col_format = col
                    end_col_format = col + 1
                col += 1

        max_col = max(len(header), len(sheet_header))

        footer = [''] * max_col
        sum_formula = '=SUM({}:{})'.format(
            rowcol_to_a1(start_row_format + 1, end_col_format),
            rowcol_to_a1(len(rows), end_col_format)
        )
        footer[len(footer) - 1] = sum_formula
        footer[len(footer) - 2] = 'SUM'
        rows.append(footer)

        rows.append([])
        rows.append(query_header)
        rows.append(query_row)

        max_row = len(rows)
        block_range = 'A1:{}'.format(
            rowcol_to_a1(max_row, max(max_col, len(query_header)))
        )

        cells = ws.range(block_range)

        for cell in cells:
            x = cell.row - 1
            y = cell.col - 1
            try:
                value = rows[x][y]
                cell.value = value
            except:
                pass

        ws.clear()

        ws.update_cells(
            cells,
            value_input_option='USER_ENTERED'
        )

        query_end_col = len(query_header)

        repeat_cells = [
            {
                'start_row': start_row_format,
                'end_row': max_row - 2,
                'start_col': start_col_format,
                'end_col': max_col,
                'fields': 'userEnteredFormat.numberFormat',
                'formats': {
                    'userEnteredFormat': {
                        'numberFormat': {
                            'type': 'CURRENCY',
                            'pattern': self.currency_pattern,
                        }
                    }
                }
            },
            {
                'start_row': 0,
                'end_row': 1,
                'start_col': 0,
                'end_col': max_col,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.8,
                            'green': 0.8,
                            'blue': 0.8,
                        }
                    }
                }
            },
            {
                'start_row': 2,
                'end_row': 3,
                'start_col': 0,
                'end_col': max_col,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.8,
                            'green': 0.8,
                            'blue': 0.8,
                        }
                    }
                }
            },
            {
                'start_row': max_row - 2,
                'end_row': max_row - 1,
                'start_col': 0,
                'end_col': query_end_col,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.9,
                            'green': 0.9,
                            'blue': 0.9,
                        }
                    }
                }
            },
            {
                'start_row': max_row - 1,
                'end_row': max_row,
                'start_col': 0,
                'end_col': query_end_col,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.9,
                            'green': 0.9,
                            'blue': 0.9,
                        }
                    }
                }
            },
            # SUM ROW
            {
                'start_row': max_row - 4,
                'end_row': max_row - 3,
                'start_col': end_col_format - 2,
                'end_col': max_col,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.95,
                            'green': 0.95,
                            'blue': 0.95,
                        }
                    }
                }
            }
        ]

        for i, row in enumerate(query_row):
            if i % 3 == 1:
                repeat_cells.append(
                    {
                        'start_row': max_row,
                        'end_row': 999999,
                        'start_col': i,
                        'end_col': i + 1,
                        'fields': 'userEnteredFormat.numberFormat',
                        'formats': {
                            'userEnteredFormat': {
                                'numberFormat': {
                                    'type': 'CURRENCY',
                                    'pattern': self.currency_pattern,
                                }
                            }
                        }
                    }
                )

        req = get_format_request(
            ws.id,
            max_row=max_row,
            max_col=max_col,
            repeat_cells=repeat_cells,
            borders=[
                {
                    'start_row': 2,
                    'end_row': 3,
                    'start_col': 0,
                    'end_col': max_col,
                    'borders': {
                        'bottom': {
                            'style': 'DOUBLE',
                            'width': 1,
                            'color': {
                                'red': 0,
                                'green': 0,
                                'blue': 0,
                            }
                        }
                    }
                },
                {
                    'start_row': max_row - 4,
                    'end_row': max_row - 3,
                    'start_col': max_col - 2,
                    'end_col': max_col,
                    'borders': {
                        'top': {
                            'style': 'SOLID',
                            'width': 1,
                            'color': {
                                'red': 0,
                                'green': 0,
                                'blue': 0,
                            }
                        }
                    }
                },
                {
                    'start_row': max_row - 1,
                    'end_row': max_row,
                    'start_col': 0,
                    'end_col': query_end_col,
                    'borders': {
                        'bottom': {
                            'style': 'SOLID',
                            'width': 1,
                            'color': {
                                'red': 0,
                                'green': 0,
                                'blue': 0,
                            }
                        }
                    }
                },
            ]
        )

        self.__sh.batch_update(req)

    def add_data(self, data):
        if self.verbose:
            print('Updating worksheet data')

        for account in data['accounts']:
            account_values = data['accounts'][account]
            if 'transactions' in account_values:
                self.__add(account_values)

    def __add(self, data):
        ws = None
        title = data['title']
        if self.verbose:
            print('Adding data to {}'.format(title))

        try:
            ws = self.__sh.worksheet(title=title)
        except:
            ws = self.__sh.add_worksheet(title=title, rows='100', cols='20')

        header = data['fieldnames']
        max_col = len(header)

        suffix = ' ' + self.__dkb_cfg['currency']

        new_rows = list(
            map(lambda x: ';'.join(x), data['transactions'])
        )
        account_cfg = self.__dkb_cfg[data['account_type']]

        indices = data['indices']
        if account_cfg['merge_values']:
            existing_rows = list(
                map(
                    lambda x: ';'.join(
                        map(lambda x: x.replace(suffix, ''), x)
                    ) + ';',
                    ws.get_all_values()[1:]
                )
            )
            seen = set()
            unique_rows = [
                x for x in (new_rows + existing_rows)
                if not (x in seen or seen.add(x))
            ]
            unique_rows = sorted(
                unique_rows,
                key=lambda x: parse_time(x, indices),
                reverse=True
            )
        else:
            unique_rows = new_rows

        unique_rows = [';'.join(header)] + unique_rows

        max_row = len(unique_rows)
        block_range = 'A1:{}'.format(
            rowcol_to_a1(max_row, max_col)
        )

        user_entered = []
        raw = []
        for cell in ws.range(block_range):
            x = cell.row - 1
            y = cell.col - 1
            try:
                value = unique_rows[x].split(';')[y]
                cell.value = value
            except:
                cell.value = ''

            if y in indices['currency'] or ('date' in indices and y in indices['date']):
                user_entered.append(cell)
            else:
                raw.append(cell)

        ws.clear()
        ws.update_cells(
            user_entered,
            value_input_option='USER_ENTERED'
        )

        ws.update_cells(
            raw,
            value_input_option='RAW'
        )

        repeat_cells = [
            {
                'start_row': 0,
                'end_row': 1,
                'start_col': 0,
                'end_col': max_col - 1,
                'fields': 'userEnteredFormat.backgroundColor',
                'formats': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.9,
                            'green': 0.9,
                            'blue': 0.9,
                        }
                    }
                }
            }
        ]

        if 'date' in indices:
            for i in indices['date']:
                repeat_cells.append({
                    'start_row': 1,
                    'end_row': max_row,
                    'start_col': i,
                    'end_col': i + 1,
                    'fields': 'userEnteredFormat.numberFormat',
                    'formats': {
                        'userEnteredFormat': {
                            'numberFormat': {
                                    'type': 'DATE',
                                    }
                        }
                    }
                })

        for i in indices['currency']:
            repeat_cells.append({
                'start_row': 1,
                'end_row': max_row,
                'start_col': i,
                'end_col': i + 1,
                'fields': 'userEnteredFormat.numberFormat',
                'formats': {
                    'userEnteredFormat': {
                        'numberFormat': {
                                'type': 'CURRENCY',
                                'pattern': self.currency_pattern,
                                }
                    }
                }
            })

        req = get_format_request(
            ws.id,
            max_row=max_row,
            max_col=max_col,
            repeat_cells=repeat_cells,
            borders=[
                {
                    'start_row': 0,
                    'end_row': 1,
                    'start_col': 0,
                    'end_col': max_col - 1,
                    'borders': {
                        'bottom': {
                            'style': 'DOUBLE',
                            'width': 1,
                            'color': {
                                'red': 0,
                                'green': 0,
                                'blue': 0,
                            }
                        }
                    }
                },
            ]
        )

        self.__sh.batch_update(req)
