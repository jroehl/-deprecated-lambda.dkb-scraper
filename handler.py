import json
import traceback
from datetime import date, timedelta, datetime
from dotenv import load_dotenv

from config import get_config
from src.dkb import DKBSession
from src.sheets import GSheet
from src.utils import init, parse_range

load_dotenv()
init()


def scrape(event, context):
    try:
        time_span_string = ''
        end_date_string = ''
        if 'pathParameters' in event:
            time_span_string = event['pathParameters']['time_span']
            if 'end_date' in event['pathParameters']:
                end_date_string = event['pathParameters']['end_date']
        elif 'time_span' in event:
            time_span_string = event['time_span']
            if 'end_date' in event:
                end_date_string = event['end_date']
        else:
            raise Exception("start_date has to be specified")

        dkb_cfg = get_config('dkb')
        session = DKBSession(
            username=dkb_cfg['creds']['username'],
            password=dkb_cfg['creds']['password'],
            verbose=True
        )

        date_format = dkb_cfg['formats']['date']
        end_date = datetime.now()
        if end_date_string:
            end_date = datetime.strptime(end_date_string, date_format)

        start_date = parse_range(time_span_string, end_date, date_format)
        if start_date < (datetime.today() - timedelta(days=((3*365) + 1))):
            raise ValueError('start_date can only be 3 years in the past')

        if end_date < start_date:
            raise ValueError('start_date must be after end_date')

        session.login()
        res = session.query(start_date, end_date)
        session.logout()

        gsheet = GSheet()
        gsheet.update_dashboard(res)
        gsheet.add_data(res)

        for account in res['accounts']:
            account_values = res['accounts'][account]
            if 'transactions' in account_values:
                del account_values['transactions']

        response = {
            "statusCode": 200,
            "message": "query successful",
            "body": json.dumps(res)
        }
        return response
    except Exception as e:
        traceback.print_exc()
        response = {
            "statusCode": 400,
            "err": str(e),
            "body": json.dumps({'err': str(e), 'info': traceback.format_exc().split('\n')})
        }
        return response


if __name__ == "__main__":
    scrape({'time_span': '1'}, '')
