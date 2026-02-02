import os
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
ONEDRIVE_URL = os.getenv("ONEDRIVE_FILE_URL")
MAKER_SHEETS = ["GREYSAGE", "ARVIND", "MIDSEN", "HASAN", "RAMA", "HAKIM", "RAMU", "ANIL", "SINU"]

# Global cache (Persists between warm invocations on Vercel)
_cache = {'data': None, 'timestamp': None}
CACHE_DURATION = 300 

def load_excel():
    """Load Excel file with caching"""
    global _cache
    
    now = time.time()
    if _cache['data'] is not None and (_cache['timestamp'] and now - _cache['timestamp'] < CACHE_DURATION):
        return _cache['data']
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(ONEDRIVE_URL, timeout=20, headers=headers)
        response.raise_for_status()
        
        excel_file = pd.ExcelFile(BytesIO(response.content), engine='openpyxl')
        _cache['data'] = excel_file
        _cache['timestamp'] = now
        return excel_file
    except Exception as e:
        raise Exception(f"Failed to load Excel: {str(e)}")


def find_header_row_fast(df):
    """Find header row - ultra fast"""
    # Convert first column to string array for vectorized search
    first_col = df.iloc[:, 0].astype(str).str.upper().str.strip()
    matches = first_col == 'CLIENT'
    if matches.any():
        return matches.idxmax()
    return 0


def process_sheet_optimized(args):
    """Process a single sheet - maximum optimization"""
    sheet_name, excel_file = args
    
    try:
        if sheet_name not in excel_file.sheet_names:
            return None
        
        # Read minimal data to find header
        raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20, dtype=str)
        header_row = find_header_row_fast(raw)
        
        # Read all columns first (can't filter before normalization)
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=header_row
        )
        
        if df.empty:
            return None
        
        # Normalize column names once
        df.columns = df.columns.str.strip().str.upper()
        
        # Now filter to only required columns
        required_cols = ['CLIENT', 'WASHING', 'PCS', 'WASH ED']
        available_cols = [col for col in required_cols if col in df.columns]
        
        if 'CLIENT' not in available_cols:
            return None
        
        df = df[available_cols]
        
        # Handle PCS column efficiently
        if 'PCS' in df.columns:
            df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce').fillna(0).astype('int32')
        else:
            df['PCS'] = 0
        
        # Fill missing columns
        for col in ['WASHING', 'WASH ED']:
            if col not in df.columns:
                df[col] = ''
        
        # Clean string columns in one pass - use fillna to handle NaN
        df['CLIENT'] = df['CLIENT'].fillna('').astype(str).str.strip()
        df['WASHING'] = df['WASHING'].fillna('').astype(str).str.strip()
        df['WASH ED'] = df['WASH ED'].fillna('').astype(str).str.strip()
        
        # Filter valid rows
        df = df[df['CLIENT'].str.len() > 0]
        
        if df.empty:
            return None
        
        # Vectorized status calculation - much faster
        washing_empty = (df['WASHING'] == '') | (df['WASHING'] == 'nan')
        wash_ed_empty = (df['WASH ED'] == '') | (df['WASH ED'] == 'nan')
        
        df['MAKING'] = df['PCS'].where(washing_empty, 0).astype('int32')
        df['IN_WASHING'] = df['PCS'].where(~washing_empty, 0).astype('int32')
        df['OUT_WASHING'] = df['PCS'].where(~washing_empty & ~wash_ed_empty, 0).astype('int32')
        
        return df[['CLIENT', 'WASHING', 'PCS', 'MAKING', 'IN_WASHING', 'OUT_WASHING']]
    
    except Exception as e:
        print(f"Error in sheet {sheet_name}: {str(e)}")
        return None


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        start_time = time.time()
        try:
            # Load Excel once
            excel_file = load_excel()
            load_time = time.time() - start_time
            
            # Process all sheets in parallel with optimal worker count
            all_rows = []
            
            process_start = time.time()
            with ThreadPoolExecutor(max_workers=len(MAKER_SHEETS)) as executor:
                # Submit all tasks at once
                futures = {
                    executor.submit(process_sheet_optimized, (sheet, excel_file)): sheet
                    for sheet in MAKER_SHEETS
                }
                
                # Collect results as they complete
                for future in as_completed(futures):
                    result = future.result()
                    if result is not None and not result.empty:
                        all_rows.append(result)
            
            process_time = time.time() - process_start
            
            # Handle empty case
            if not all_rows:
                response_data = {
                    "rows": [],
                    "client_summary": [],
                    "washer_summary": [],
                    "total_pcs": 0,
                    "total_making": 0,
                    "total_in_washing": 0,
                    "total_out_washing": 0,
                    "timestamp": datetime.now().isoformat(),
                    "cached": (_cache['timestamp'] is not None),
                    "processing_time": round(time.time() - start_time, 2),
                    "load_time": round(load_time, 2),
                    "process_time": round(process_time, 2)
                }
                self._send_json(response_data, 200)
                return
            
            # Combine results efficiently
            concat_start = time.time()
            master = pd.concat(all_rows, ignore_index=True, copy=False)
            concat_time = time.time() - concat_start
            
            # Client Summary - optimized aggregation
            summary_start = time.time()
            client_summary = (
                master.groupby('CLIENT', as_index=False, sort=False)
                .agg({
                    'MAKING': 'sum',
                    'IN_WASHING': 'sum',
                    'OUT_WASHING': 'sum'
                })
                .sort_values('MAKING', ascending=False)
            )
            
            # Washer Summary - optimized
            washer_data = master[master['WASHING'].str.len() > 0]
            
            if not washer_data.empty:
                washer_summary = (
                    washer_data.groupby('WASHING', as_index=False, sort=False)
                    .agg({
                        'IN_WASHING': 'sum',
                        'OUT_WASHING': 'sum'
                    })
                    .rename(columns={'WASHING': 'WASHER'})
                )
                washer_summary['PENDING'] = (
                    washer_summary['IN_WASHING'] - washer_summary['OUT_WASHING']
                ).astype('int32')
                washer_summary = washer_summary.sort_values('PENDING', ascending=False)
            else:
                washer_summary = pd.DataFrame(columns=['WASHER', 'IN_WASHING', 'OUT_WASHING', 'PENDING'])
            
            summary_time = time.time() - summary_start
            
            # Calculate totals - single pass with optimized dtype
            total_pcs = int(master['PCS'].sum())
            total_making = int(master['MAKING'].sum())
            total_in_washing = int(master['IN_WASHING'].sum())
            total_out_washing = int(master['OUT_WASHING'].sum())
            
            # Convert to JSON efficiently
            json_start = time.time()
            response_data = {
                "rows": master.to_dict(orient='records'),
                "client_summary": client_summary.to_dict(orient='records'),
                "washer_summary": washer_summary.to_dict(orient='records'),
                "total_pcs": total_pcs,
                "total_making": total_making,
                "total_in_washing": total_in_washing,
                "total_out_washing": total_out_washing,
                "timestamp": datetime.now().isoformat(),
                "cached": (_cache['timestamp'] is not None and time.time() - _cache['timestamp'] < CACHE_DURATION),
                "processing_time": round(time.time() - start_time, 2),
                "timing": {
                    "load": round(load_time, 2),
                    "process_sheets": round(process_time, 2),
                    "concat": round(concat_time, 2),
                    "summaries": round(summary_time, 2),
                    "json": round(time.time() - json_start, 2)
                }
            }
            json_time = time.time() - json_start
            response_data["timing"]["json"] = round(json_time, 2)
            
            self._send_json(response_data, 200)

        except Exception as e:
            self._send_json({"error": str(e)}, 500)

    def _send_json(self, data, status):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())