import json
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
from http.server import BaseHTTPRequestHandler

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning)

# Configuration
ONEDRIVE_URL = "https://1drv.ms/x/c/625f71beaac1d94c/IQBWvpCT0VcyTY78VpTy0OSaAXOgo-VKMSu4hFV6_26vveo?e=lSFwiX&download=1"
MAKER_SHEETS = ["GREYSAGE", "ARVIND", "MIDSEN", "HASAN", "RAMA", "HAKIM", "RAMU", "ANIL", "SINU"]

# Global cache (Persists between warm invocations on Vercel)
_cache = {'data': None, 'timestamp': None}
CACHE_DURATION = 300 

def load_excel():
    global _cache
    now = time.time()
    if _cache['data'] is not None and (_cache['timestamp'] and now - _cache['timestamp'] < CACHE_DURATION):
        return _cache['data']
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    response = requests.get(ONEDRIVE_URL, timeout=20, headers=headers)
    response.raise_for_status()
    
    excel_file = pd.ExcelFile(BytesIO(response.content), engine='openpyxl')
    _cache['data'] = excel_file
    _cache['timestamp'] = now
    return excel_file

def find_header_row(df):
    try:
        for idx in range(min(20, len(df))):
            if any(str(val).upper().strip() == 'CLIENT' for val in df.iloc[idx]):
                return idx
        return 0
    except: return 0

def process_sheet_fast(args):
    sheet_name, excel_file = args
    try:
        if sheet_name not in excel_file.sheet_names: return None
        raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20)
        header_row = find_header_row(raw)
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row, na_values=[''], keep_default_na=False)
        
        if df.empty: return None
        df.columns = df.columns.str.strip().str.upper()
        if 'CLIENT' not in df.columns: return None
        
        cols_needed = [c for c in ['CLIENT', 'WASHING', 'PCS', 'WASH ED'] if c in df.columns]
        df = df[cols_needed].copy()
        df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce').fillna(0).astype(int)
        
        for col in ['CLIENT', 'WASHING', 'WASH ED']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        df = df[(df['CLIENT'] != '') & (df['CLIENT'] != 'nan')].copy()
        if df.empty: return None

        washing_empty = df['WASHING'].isin(['', 'nan'])
        wash_ed_empty = df['WASH ED'].isin(['', 'nan'])
        df['MAKING'] = df['PCS'].where(washing_empty, 0).astype(int)
        df['IN_WASHING'] = df['PCS'].where(~washing_empty, 0).astype(int)
        df['OUT_WASHING'] = df['PCS'].where(~washing_empty & ~wash_ed_empty, 0).astype(int)
        return df
    except: return None

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        start_time = time.time()
        try:
            excel_file = load_excel()
            all_rows = []
            
            with ThreadPoolExecutor(max_workers=len(MAKER_SHEETS)) as executor:
                futures = [executor.submit(process_sheet_fast, (s, excel_file)) for s in MAKER_SHEETS]
                for future in as_completed(futures):
                    res = future.result()
                    if res is not None: all_rows.append(res)

            if not all_rows:
                self._send_json({"rows": [], "total_pcs": 0}, 200)
                return

            master = pd.concat(all_rows, ignore_index=True)
            client_summary = master.groupby('CLIENT', as_index=False).agg({'MAKING': 'sum', 'IN_WASHING': 'sum', 'OUT_WASHING': 'sum'}).sort_values('MAKING', ascending=False)
            
            washer_mask = ~master['WASHING'].isin(['', 'nan'])
            washer_data = master[washer_mask]
            if not washer_data.empty:
                washer_summary = washer_data.groupby('WASHING', as_index=False).agg({'IN_WASHING': 'sum', 'OUT_WASHING': 'sum'}).rename(columns={'WASHING': 'WASHER'})
                washer_summary['PENDING'] = (washer_summary['IN_WASHING'] - washer_summary['OUT_WASHING']).astype(int)
                washer_summary = washer_summary.sort_values('PENDING', ascending=False)
            else:
                washer_summary = pd.DataFrame(columns=['WASHER', 'IN_WASHING', 'OUT_WASHING', 'PENDING'])

            essential_cols = ['CLIENT', 'PCS', 'WASHING', 'WASH ED']
            master_filtered = master[essential_cols]

            response_data = {
                "columns": essential_cols,
                "rows": master_filtered.values.tolist(),
                # "rows": master.to_dict(orient='records'),
                "client_summary": client_summary.to_dict(orient='records'),
                "washer_summary": washer_summary.to_dict(orient='records'),
                "total_pcs": int(master['PCS'].sum()),
                "timestamp": datetime.now().isoformat(),
                "processing_time": round(time.time() - start_time, 2)
            }
            self._send_json(response_data, 200)
            
            # self._send_json({"message": "Test Functionality."}, 200)

        except Exception as e:
            self._send_json({"error": str(e)}, 500)

    def _send_json(self, data, status):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())
