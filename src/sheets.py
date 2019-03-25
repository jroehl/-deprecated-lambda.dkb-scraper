from datetime import datetime
import gspread
from gspread.utils import rowcol_to_a1
from gspread.urls import SPREADSHEETS_API_V4_BASE_URL
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client.crypt import Signer

from config import get_config
from src.utils import format_pattern, get_format_request

DRIVE_V3_URL = 'https://www.googleapis.com/drive/v3/files'


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

        max_len = len(header) if len(
            header) > len(sheet_header) else len(sheet_header)

        rows = [
            sheet_header,
            [''] * max_len,
            header,
        ]

        start_row_format = None
        end_row_format = None
        start_col_format = None
        end_col_format = None

        start_idx = len(rows)
        for i, account in enumerate(accounts):
            row = i + (start_idx)
            rows.append([])
            col = 0
            account_values = accounts[account]
            for key in header:
                if key in account_values:
                    rows[row].append(str(account_values[key]))
                if key == 'total':
                    if not start_row_format:
                        start_row_format = row
                    end_row_format = row
                    start_col_format = col
                    end_col_format = col + 1
                col += 1

        footer = [''] * max_len
        sum_formula = '=SUM({}:{})'.format(
            rowcol_to_a1(start_row_format + 1, end_col_format),
            rowcol_to_a1(len(rows), end_col_format)
        )
        footer[len(footer) - 1] = sum_formula
        footer[len(footer) - 2] = 'SUM'
        rows.append(footer)

        end_row_format = len(rows)
        block_range = 'A1:{}'.format(
            rowcol_to_a1(end_row_format, max_len)
        )

        cells = ws.range(block_range)

        for cell in cells:
            x = cell.row - 1
            y = cell.col - 1
            try:
                value = rows[x][y]
                cell.value = value
            except:
                cell.value = 'row {} col {}'.format(cell.row, cell.col)

        ws.clear()

        ws.update_cells(
            cells,
            value_input_option='USER_ENTERED'
        )

        req = get_format_request(
            ws.id,
            max_row=end_row_format,
            max_col=max_len,
            repeat_cells=[
                {
                    'start_row': start_row_format,
                    'end_row': end_row_format,
                    'start_col': start_col_format,
                    'end_col': max_len,
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
                    'end_col': max_len,
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
                    'end_col': max_len,
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
                    'start_row': end_row_format - 1,
                    'end_row': end_row_format,
                    'start_col': end_col_format - 2,
                    'end_col': max_len,
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
            ],
            borders=[
                {
                    'start_row': 2,
                    'end_row': 3,
                    'start_col': 0,
                    'end_col': max_len,
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
                    'start_row': end_row_format - 1,
                    'end_row': end_row_format,
                    'start_col': max_len - 2,
                    'end_col': max_len,
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
                }
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
        title = '{} / {}'.format(data['product_name'], data['account_number'])
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
                key=lambda x: datetime.strptime(
                    x.split(';')[indices['date'][0]], '%x'
                ),
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
