import yfinance as yf
import pandas as pd
import datetime

# --- 1. CONFIGURATION ---
tickers = [
    'ADANIENT.NS', 'ADANIPORTS.NS', 'APOLLOHOSP.NS', 'ASIANPAINT.NS', 'AXISBANK.NS',
    'BAJAJ-AUTO.NS', 'BAJFINANCE.NS', 'BAJAJFINSV.NS', 'BEL.NS', 'BPCL.NS',
    'BHARTIARTL.NS', 'BRITANNIA.NS', 'CIPLA.NS', 'COALINDIA.NS', 'DIVISLAB.NS',
    'DRREDDY.NS', 'EICHERMOT.NS', 'GRASIM.NS', 'HCLTECH.NS', 'HDFCBANK.NS',
    'HDFCLIFE.NS', 'HEROMOTOCO.NS', 'HINDALCO.NS', 'HINDUNILVR.NS', 'ICICIBANK.NS',
    'ITC.NS', 'INDUSINDBK.NS', 'INFY.NS', 'JSWSTEEL.NS', 'KOTAKBANK.NS',
    'LTIM.NS', 'LT.NS', 'M&M.NS', 'MARUTI.NS', 'NESTLEIND.NS',
    'NTPC.NS', 'ONGC.NS', 'POWERGRID.NS', 'RELIANCE.NS', 'SBILIFE.NS',
    'SBIN.NS', 'SHRIRAMFIN.NS', 'SUNPHARMA.NS', 'TATASTEEL.NS', 'TCS.NS',
    'TATACONSUM.NS', 'TECHM.NS', 'TITAN.NS', 'TRENT.NS',
    'ULTRACEMCO.NS', 'WIPRO.NS', 'PIDILITIND.NS', 'VEDL.NS'
]
benchmark_ticker = "^NSEI"

print(f"--- 🚀 STARTING PRO-GRADE SCAN ({len(tickers)} Stocks) ---")
print("Fetching live data and calculating RSI (Manual Algorithm)...")

analysis_report = []
start_date = datetime.date.today() - datetime.timedelta(days=365)
end_date = datetime.date.today()

# --- HELPER: DATA CLEANER ---
def get_clean_series(data):
    if data is None or data.empty: return None
    if isinstance(data, pd.Series): return data
    if isinstance(data, pd.DataFrame):
        for col in ['Close', 'Adj Close']:
            if col in data.columns: return data[col].squeeze()
        return data.iloc[:, 0].squeeze()
    return None

# --- HELPER: MANUAL RSI CALCULATOR (No External Libs) ---
def calculate_rsi(series, period=14):
    delta = series.diff()
    gain = (delta.where(delta > 0, 0))
    loss = (-delta.where(delta < 0, 0))
    
    # Use Exponential Weighted Moving Average (Wilder's Smoothing)
    avg_gain = gain.ewm(com=period-1, min_periods=period).mean()
    avg_loss = loss.ewm(com=period-1, min_periods=period).mean()
    
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return rsi

# --- 1. GET BENCHMARK ---
try:
    bench_raw = yf.download(benchmark_ticker, start=start_date, end=end_date, progress=False)
    bench_series = get_clean_series(bench_raw)
    bench_ret = bench_series.pct_change().dropna()
except:
    print("❌ Critical Error: Internet or API down.")
    exit()

# --- 2. THE ANALYST LOOP ---
for ticker in tickers:
    try:
        # Get Data
        stock_raw = yf.download(ticker, start=start_date, end=end_date, progress=False)
        stock_series = get_clean_series(stock_raw)
        
        if stock_series is None or len(stock_series) < 50: continue

        # --- A. ADVANCED CALCULATIONS ---
        # 1. RSI (Calculated Manually now)
        rsi_series = calculate_rsi(stock_series, period=14)
        if rsi_series.dropna().empty: continue
        current_rsi = rsi_series.iloc[-1]
        
        # 2. Risk Metrics
        stock_ret = stock_series.pct_change().dropna()
        common_dates = stock_ret.index.intersection(bench_ret.index)
        aligned_stock = stock_ret.loc[common_dates]
        aligned_bench = bench_ret.loc[common_dates]
        
        if aligned_stock.empty: continue
        
        beta = aligned_stock.cov(aligned_bench) / aligned_bench.var()
        
        # 3. Signals
        curr_price = float(stock_series.iloc[-1])
        sma_50 = float(stock_series.rolling(window=50).mean().iloc[-1])
        
        trend = "BULLISH" if curr_price > sma_50 else "BEARISH"
        
        # RSI Interpretation
        rsi_signal = "NEUTRAL"
        if current_rsi > 70: rsi_signal = "OVERBOUGHT (Risk)"
        if current_rsi < 30: rsi_signal = "OVERSOLD (Value)"

        analysis_report.append({
            'Ticker': ticker.replace('.NS', ''),
            'Price': curr_price,
            'Trend': trend,
            'RSI': round(current_rsi, 2),
            'RSI_Status': rsi_signal,
            'Beta': round(beta, 2)
        })
        print(f"✅ {ticker.replace('.NS', '')}: {trend} | RSI: {round(current_rsi, 0)}")

    except Exception as e:
        # print(e) # Uncomment to debug specific errors
        pass

# --- 3. EXCEL AUTOMATION (THE "NO HUMAN WASTE" PART) ---
if analysis_report:
    df = pd.DataFrame(analysis_report)
    
    output_file = 'Pro_Dashboard.xlsx'
    
    # Using xlsxwriter for formatting
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Market_Scan', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Market_Scan']
    
    # --- AUTOMATIC FORMATTING ---
    # Formats
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # Apply Header Format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    
    # Conditional Formatting: TREND (Green for Bullish, Red for Bearish)
    # Applied to Column C (Trend)
    worksheet.conditional_format('C2:C100', {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'BULLISH',
                                            'format': green_fmt})
    worksheet.conditional_format('C2:C100', {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'BEARISH',
                                            'format': red_fmt})

    # Conditional Formatting: RSI Status (Red for Overbought, Green for Oversold)
    # Applied to Column E (RSI_Status)
    worksheet.conditional_format('E2:E100', {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'OVERBOUGHT',
                                            'format': red_fmt})
    worksheet.conditional_format('E2:E100', {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'OVERSOLD',
                                            'format': green_fmt})

    # Adjust Column Widths automatically
    worksheet.set_column('A:A', 15) # Ticker
    worksheet.set_column('B:B', 10) # Price
    worksheet.set_column('C:C', 15) # Trend
    worksheet.set_column('D:D', 8)  # RSI
    worksheet.set_column('E:E', 20) # RSI Status
    worksheet.set_column('F:F', 8)  # Beta
    
    writer.close()
    print(f"\n📊 SUCCESS! '{output_file}' created with auto-formatting.")
    
    # Auto-open the Excel file
    try:
        import os
        os.startfile(output_file)
    except:
        pass

else:
    print("No data collected.")