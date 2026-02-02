"""
Vercel Serverless Function: /api/process.py
Optimized for speed with parallel processing and efficient Excel reading
"""

import json
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

# Suppress openpyxl warnings
warnings.filterwarnings('ignore', category=UserWarning)

# Configuration
ONEDRIVE_URL = "https://1drv.ms/x/c/625f71beaac1d94c/IQBWvpCT0VcyTY78VpTy0OSaAXOgo-VKMSu4hFV6_26vveo?e=lSFwiX&download=1"

MAKER_SHEETS = [
    "GREYSAGE", "ARVIND", "MIDSEN", "HASAN",
    "RAMA", "HAKIM", "RAMU", "ANIL", "SINU"
]

# Global cache
_cache = {'data': None, 'timestamp': None}
CACHE_DURATION = 300  # 5 minutes


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


def find_header_row(df):
    """Find header row efficiently"""
    try:
        # Check first 20 rows
        for idx in range(min(20, len(df))):
            if any(str(val).upper().strip() == 'CLIENT' for val in df.iloc[idx]):
                return idx
        return 0
    except:
        return 0


def process_sheet_fast(args):
    """Process a single sheet - optimized for speed"""
    sheet_name, excel_file = args
    
    try:
        # Skip if sheet doesn't exist
        if sheet_name not in excel_file.sheet_names:
            return None
        
        # Find header row
        raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20)
        header_row = find_header_row(raw)
        
        # Read sheet with proper header
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=header_row,
            na_values=[''],  # Treat empty strings as NaN
            keep_default_na=False
        )
        
        if df.empty:
            return None
        
        # Normalize columns - fast operation
        df.columns = df.columns.str.strip().str.upper()
        
        # Check if CLIENT column exists
        if 'CLIENT' not in df.columns:
            return None
        
        # Select only needed columns - avoid copying unnecessary data
        cols_needed = []
        for col in ['CLIENT', 'WASHING', 'PCS', 'WASH ED']:
            if col in df.columns:
                cols_needed.append(col)
        
        if not cols_needed or 'CLIENT' not in cols_needed:
            return None
        
        df = df[cols_needed]
        
        # Fast type conversion
        df['PCS'] = pd.to_numeric(df['PCS'], errors='coerce').fillna(0).astype(int)
        
        # Convert to string once - avoid multiple operations
        for col in ['CLIENT', 'WASHING', 'WASH ED']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        # Filter invalid rows in one operation
        df = df[(df['CLIENT'] != '') & (df['CLIENT'] != 'nan')].copy()
        
        if df.empty:
            return None
        
        # Add status columns - vectorized operations
        washing_empty = df['WASHING'].isin(['', 'nan'])
        wash_ed_empty = df['WASH ED'].isin(['', 'nan'])
        
        df['MAKING'] = df['PCS'].where(washing_empty, 0).astype(int)
        df['IN_WASHING'] = df['PCS'].where(~washing_empty, 0).astype(int)
        df['OUT_WASHING'] = df['PCS'].where(~washing_empty & ~wash_ed_empty, 0).astype(int)
        
        return df
    
    except Exception as e:
        print(f"Error in sheet {sheet_name}: {str(e)}")
        return None


def app(request):
    """Main Vercel handler"""
    try:
        start_time = time.time()
        
        # Load Excel once
        excel_file = load_excel()
        
        # Process all sheets in parallel
        all_rows = []
        
        with ThreadPoolExecutor(max_workers=4) as executor:
            # Submit all tasks
            futures = [
                executor.submit(process_sheet_fast, (sheet, excel_file))
                for sheet in MAKER_SHEETS
            ]
            
            # Collect results as they complete
            for future in as_completed(futures):
                result = future.result()
                if result is not None:
                    all_rows.append(result)
        
        # Combine results
        if not all_rows:
            return {
                "statusCode": 200,
                "body": json.dumps({
                    "rows": [],
                    "client_summary": [],
                    "washer_summary": [],
                    "total_pcs": 0,
                    "total_making": 0,
                    "total_in_washing": 0,
                    "total_out_washing": 0,
                    "timestamp": datetime.now().isoformat(),
                    "cached": False,
                    "processing_time": time.time() - start_time
                })
            }
        
        master = pd.concat(all_rows, ignore_index=True)
        
        # Client Summary - vectorized
        client_summary = (
            master.groupby('CLIENT', as_index=False)
            .agg({'MAKING': 'sum', 'IN_WASHING': 'sum', 'OUT_WASHING': 'sum'})
            .sort_values('MAKING', ascending=False)
        )
        
        # Washer Summary - vectorized
        washer_mask = ~master['WASHING'].isin(['', 'nan'])
        washer_data = master[washer_mask]
        
        if not washer_data.empty:
            washer_summary = (
                washer_data.groupby('WASHING', as_index=False)
                .agg({'IN_WASHING': 'sum', 'OUT_WASHING': 'sum'})
                .rename(columns={'WASHING': 'WASHER'})
            )
            washer_summary['PENDING'] = (
                washer_summary['IN_WASHING'] - washer_summary['OUT_WASHING']
            ).astype(int)
            washer_summary = washer_summary.sort_values('PENDING', ascending=False)
        else:
            washer_summary = pd.DataFrame(columns=['WASHER', 'IN_WASHING', 'OUT_WASHING', 'PENDING'])
        
        # Calculate totals - single pass
        total_pcs = int(master['PCS'].sum())
        total_making = int(master['MAKING'].sum())
        total_in_washing = int(master['IN_WASHING'].sum())
        total_out_washing = int(master['OUT_WASHING'].sum())
        
        processing_time = time.time() - start_time
        
        response_data = {
            "rows": master.to_dict(orient='records'),
            "client_summary": client_summary.to_dict(orient='records'),
            "washer_summary": washer_summary.to_dict(orient='records'),
            "total_pcs": total_pcs,
            "total_making": total_making,
            "total_in_washing": total_in_washing,
            "total_out_washing": total_out_washing,
            "timestamp": datetime.now().isoformat(),
            "cached": False,
            "processing_time": round(processing_time, 2)
        }
        
        return {
            "statusCode": 200,
            "headers": {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*",
                "Cache-Control": "public, max-age=300"
            },
            "body": json.dumps(response_data)
        }
    
    except Exception as e:
        return {
            "statusCode": 500,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({
                "error": str(e),
                "rows": [],
                "client_summary": [],
                "washer_summary": []
            })
        }
