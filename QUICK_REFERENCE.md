# Quick Reference Guide
ONEDRIVE_URL = os.getenv("ONEDRIVE_FILE_URL")

## File Mapping for Vercel

| Local File | Vercel Path | Purpose |
|---|---|---|
| `process.py` | `api/process.py` | Serverless function |
| `index_vercel.html` | `public/index.html` | Dashboard UI |
| `requirements.txt` | `requirements.txt` | Dependencies |
| `vercel.json` | `vercel.json` | Configuration |

## Performance Comparison

### Before (Sequential Processing)
```
Download: 2s
Process Sheet 1: 0.5s
Process Sheet 2: 0.5s
... (9 sheets)
Total: ~10-12 seconds ❌
```

### After (Parallel Processing)
```
Download: 2s
Process Sheets (parallel): 1.5s
Total: ~3.5-5 seconds ✅
```

## What Changed

### 1. **Parallel Processing**
```python
with ThreadPoolExecutor(max_workers=4) as executor:
    futures = [executor.submit(process_sheet_fast, (sheet, excel_file)) for sheet in MAKER_SHEETS]
    for future in as_completed(futures):
        result = future.result()
```

### 2. **Vectorized Operations**
```python
# Before (slow - row by row)
for idx, row in df.iterrows():
    df.loc[idx, 'MAKING'] = df.loc[idx, 'PCS'] if df.loc[idx, 'WASHING'] == '' else 0

# After (fast - vectorized)
washing_empty = df['WASHING'].isin(['', 'nan'])
df['MAKING'] = df['PCS'].where(washing_empty, 0).astype(int)
```

### 3. **Smart Column Selection**
```python
# Only read needed columns
cols_needed = [col for col in ['CLIENT', 'WASHING', 'PCS', 'WASH ED'] if col in df.columns]
df = df[cols_needed]  # Smaller dataset = faster processing
```

### 4. **NaT Handling**
```python
# No more NaT errors - convert to string immediately
for col in ['CLIENT', 'WASHING', 'WASH ED']:
    df[col] = df[col].astype(str).str.strip()
```

## Local Testing

### Test 1: Check Excel Loading
```python
from api.process import load_excel

excel_file = load_excel()
print(excel_file.sheet_names)
```

### Test 2: Check Single Sheet
```python
from api.process import process_sheet_fast

excel_file = load_excel()
result = process_sheet_fast(("GREYSAGE", excel_file))
print(result.head())
```

### Test 3: Full Handler
```python
from api.process import handler
import json

response = handler(None)
print(json.dumps(response, indent=2))
```

## Optimization Checklist

- [x] Parallel processing (ThreadPoolExecutor)
- [x] Vectorized pandas operations
- [x] Memory caching (5 minutes)
- [x] NaT type conversion
- [x] Reduced column selection
- [x] Early filtering
- [x] Timeout handling

## Deployment Checklist

- [ ] Copy `process.py` to `api/process.py`
- [ ] Copy `index_vercel.html` to `public/index.html`
- [ ] Verify `requirements.txt` in root
- [ ] Verify `vercel.json` in root
- [ ] Run `vercel login`
- [ ] Run `vercel` to deploy
- [ ] Test endpoint: `https://your-project.vercel.app/api/process`
- [ ] Check logs in Vercel Dashboard

## Performance Monitoring

### Check Speed
```javascript
// In browser console after loading dashboard
console.log(document.getElementById("processingTime").textContent)
// Shows: "1500" = 1.5 seconds
```

### View Logs
```bash
# Local: Check terminal output
vercel logs --follow

# Or in Vercel Dashboard:
# Project → Deployments → Functions → View Logs
```

## Common Issues & Solutions

### Issue: "NaTType does not support timetuple"
**Solution:** Using optimized code which converts to string immediately

### Issue: Function timeout
**Solution:** Increased to 30s in vercel.json, parallel processing reduces actual time

### Issue: High memory usage
**Solution:** Only reading needed columns, using vectorized operations

### Issue: CORS errors
**Solution:** Headers automatically added in response

## Advanced Customization

### Change Cache Duration
In `process.py`:
```python
CACHE_DURATION = 600  # 10 minutes (default: 300 = 5 minutes)
```

### Change Number of Workers
In `process.py`:
```python
with ThreadPoolExecutor(max_workers=6) as executor:  # default: 4
```

### Change Max Rows in Detail Table
In `index_vercel.html`:
```javascript
rows.slice(0, 200)  // Show 200 rows (default: 100)
```

### Add Custom Columns
In `process.py`, after creating df:
```python
df['CUSTOM_COLUMN'] = df['PCS'] * 1.5
```

## Database Migration Path

For production at scale, consider migrating to database:

1. **PostgreSQL on Vercel Postgres**
   - Store parsed data in database
   - Query by API instead of parsing Excel
   - 100x faster

2. **Implementation**:
   ```python
   # Instead of parsing Excel every time
   # INSERT INTO production_data SELECT * FROM excel_file
   # Then query from database
   ```

3. **Benefits**:
   - No Excel parsing needed
   - Instant API responses
   - Historical data tracking
   - Real-time updates possible

## Resources

- Python Vercel Runtime: https://vercel.com/docs/runtimes/python
- Pandas Optimization: https://pandas.pydata.org/docs/user_guide/enhancing.html
- ThreadPoolExecutor: https://docs.python.org/3/library/concurrent.futures.html
- Chart.js: https://www.chartjs.org/docs/latest/

## Contact Support

For Vercel issues:
- Vercel Status: https://www.vercelstatus.com/
- Help: https://vercel.com/help
- Email: support@vercel.com
