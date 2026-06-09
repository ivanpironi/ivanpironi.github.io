#!/usr/bin/env python3
"""
pipeline-sheets.py — Claude Code read/write tool for the EY Pipeline Google Sheet.

Usage:
  python3 data/pipeline-sheets.py read <Tab> <Range>
  python3 data/pipeline-sheets.py setRow <Tab> <match_json> <data_json>
  python3 data/pipeline-sheets.py appendRow <Tab> <data_json>
  python3 data/pipeline-sheets.py deleteRow <Tab> <match_json>

Examples:
  python3 data/pipeline-sheets.py read Engagements A1:G5
  python3 data/pipeline-sheets.py setRow Engagements '{"account_key":"qiddiya","type":"A","saleMonth":"3"}' '{"prob":"0.55"}'
  python3 data/pipeline-sheets.py appendRow Engagements '{"account_key":"expo","type":"CP","saleMonth":"5","duration":"4","grossK":"80","prob":"0.8","isY3Origination":"FALSE"}'
  python3 data/pipeline-sheets.py deleteRow Engagements '{"account_key":"expo","type":"CP","saleMonth":"5"}'
"""

import json, sys, requests
from google.oauth2.service_account import Credentials
import google.auth.transport.requests
import warnings
warnings.filterwarnings('ignore')  # suppress SSL warnings

SHEET_ID   = '1rdQItcwoLnKBMzcpEpdWKk6eevwI8bQAqmICl_syoaw'
CREDS_FILE = '/home/user/ivanpironi.github.io/data/.pipeline-credentials.json'
SCOPES     = ['https://www.googleapis.com/auth/spreadsheets']
BASE_URL   = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'

def auth():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    session = requests.Session()
    session.verify = False
    creds.refresh(google.auth.transport.requests.Request(session=session))
    session.headers.update({
        'Authorization': f'Bearer {creds.token}',
        'Content-Type': 'application/json'
    })
    return session

def get_sheet(session, tab):
    r = session.get(f'{BASE_URL}/values/{tab}')
    r.raise_for_status()
    vals = r.json().get('values', [])
    if not vals:
        return [], []
    headers = vals[0]
    rows = vals[1:]
    return headers, rows

def cmd_read(session, tab, rng):
    r = session.get(f'{BASE_URL}/values/{tab}!{rng}')
    r.raise_for_status()
    print(json.dumps(r.json().get('values', []), indent=2))

def cmd_set_row(session, tab, match, data):
    headers, rows = get_sheet(session, tab)
    for i, row in enumerate(rows):
        row_dict = dict(zip(headers, row))
        if all(str(row_dict.get(k,'')) == str(v) for k,v in match.items()):
            # found — apply updates cell by cell
            for key, val in data.items():
                if key in headers:
                    col = headers.index(key) + 1
                    row_num = i + 2  # 1-indexed + header row
                    rng = f'{tab}!{col_letter(col)}{row_num}'
                    r = session.put(f'{BASE_URL}/values/{rng}?valueInputOption=RAW',
                                    json={'values': [[val]], 'majorDimension': 'ROWS'})
                    r.raise_for_status()
            print(f'✓ setRow OK — updated row {i+2} in {tab}')
            return
    print(f'✗ No matching row found in {tab}', file=sys.stderr)
    sys.exit(1)

def cmd_append_row(session, tab, data):
    headers, _ = get_sheet(session, tab)
    new_row = [str(data.get(h, '')) for h in headers]
    r = session.post(f'{BASE_URL}/values/{tab}!A1:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS',
                     json={'values': [new_row], 'majorDimension': 'ROWS'})
    r.raise_for_status()
    print(f'✓ appendRow OK — added to {tab}')

def cmd_delete_row(session, tab, match):
    headers, rows = get_sheet(session, tab)
    for i, row in enumerate(rows):
        row_dict = dict(zip(headers, row))
        if all(str(row_dict.get(k,'')) == str(v) for k,v in match.items()):
            row_num = i + 2  # 1-indexed + header row
            # Get sheet ID for batchUpdate
            meta = session.get(f'{BASE_URL}?fields=sheets.properties').json()
            sheet_id = next(s['properties']['sheetId']
                           for s in meta['sheets']
                           if s['properties']['title'] == tab)
            body = {'requests': [{'deleteDimension': {
                'range': {'sheetId': sheet_id, 'dimension': 'ROWS',
                          'startIndex': row_num-1, 'endIndex': row_num}
            }}]}
            r = session.post(f'{BASE_URL}:batchUpdate', json=body)
            r.raise_for_status()
            print(f'✓ deleteRow OK — removed row {row_num} from {tab}')
            return
    print(f'✗ No matching row found in {tab}', file=sys.stderr)
    sys.exit(1)

def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    cmd  = sys.argv[1]
    tab  = sys.argv[2]
    sess = auth()

    if cmd == 'read':
        rng = sys.argv[3] if len(sys.argv) > 3 else 'A1:Z200'
        cmd_read(sess, tab, rng)
    elif cmd == 'setRow':
        cmd_set_row(sess, tab, json.loads(sys.argv[3]), json.loads(sys.argv[4]))
    elif cmd == 'appendRow':
        cmd_append_row(sess, tab, json.loads(sys.argv[3]))
    elif cmd == 'deleteRow':
        cmd_delete_row(sess, tab, json.loads(sys.argv[3]))
    else:
        print(f'Unknown command: {cmd}', file=sys.stderr)
        sys.exit(1)
