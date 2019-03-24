# lambda dkb-scraper <PoC>

Scrape dkb banking website and store transactions in google sheet.

## Set up

Create `.env` file in project root or set environment variables manually

```bash
export GOOGLE_SHEET_NAME="dkb-finances"
export GOOGLE_SHEET_WRITER="mail@johannroehl.de"

# use bank account login credentials
export DKB_USER="user"
export DKB_PASSWORD="password"
export DKB_CURRENCY="€"

# create service account to programmatically use google sheets
export CREDS_CLIENT_EMAIL=""
export CREDS_PRIVATE_KEY=""
export CREDS_CLIENT_ID=""
export CREDS_PRIVATE_KEY_ID=""
export CREDS_TOKEN_URI=""
```
