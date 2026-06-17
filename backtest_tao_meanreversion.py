"""
TAO 5-Minute Mean-Reversion Backtest
Strategy: TAO/USDT mean-reversion on 5-min chart with BTC 1-hour permission gate.
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# DATA — synthetic OHLCV (realistic GBM + regimes)
# ─────────────────────────────────────────────

def _generate_ohlcv(
    n_bars: int,
    start: str,
    freq: str,
    s0: float,
    mu_annual: float,
    sigma_daily: float,
    vol_bars_per_day: float,
    seed: int = 0,
) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    idx = pd.date_range(start, periods=n_bars, freq=freq, tz="UTC")
    dt = 1.0 / vol_bars_per_day
    # Regime: occasionally switch to bear/bull via drift shifts
    regime_len = int(vol_bars_per_day * 7)  # ~1-week regimes
    regimes = rng.choice([-0.8, -0.2, 0.0, 0.3, 0.8], size=n_bars // regime_len + 1)
    drift_arr = np.repeat(regimes, regime_len)[:n_bars] * 0.05 / vol_bars_per_day
    mu_dt = mu_annual / 365 * dt
    sigma_dt = sigma_daily * np.sqrt(dt)
    log_ret = mu_dt + drift_arr + sigma_dt * rng.standard_normal(n_bars)
    price = s0 * np.exp(np.cumsum(log_ret))
    # Build OHLCV from close
    noise = sigma_dt * 0.5
    high = price * np.exp(np.abs(rng.normal(0, noise, n_bars)))
    low  = price * np.exp(-np.abs(rng.normal(0, noise, n_bars)))
    open_ = np.roll(price, 1)
    open_[0] = s0
    volume = np.abs(rng.normal(1_000, 300, n_bars)) * price / 100
    df = pd.DataFrame({"open": open_, "high": high, "low": low, "close": price, "volume": volume}, index=idx)
    df["high"] = df[["open", "high", "close"]].max(axis=1)
    df["low"]  = df[["open", "low",  "close"]].min(axis=1)
    return df

def generate_data():
    """Generate ~18 months of 5-min TAO and 1-hour BTC data."""
    n_5m  = 18 * 30 * 24 * 12   # 18 months of 5-min bars
    n_1h  = 18 * 30 * 24         # 18 months of 1-hour bars
    start = "2024-01-01"
    tao = _generate_ohlcv(n_5m, start, "5min",  s0=100.0,  mu_annual=0.30, sigma_daily=0.06, vol_bars_per_day=12*24, seed=42)
    btc = _generate_ohlcv(n_1h, start, "1h",    s0=42000.0, mu_annual=0.15, sigma_daily=0.025, vol_bars_per_day=24,    seed=7)
    return tao, btc

# ─────────────────────────────────────────────
# INDICATORS
# ─────────────────────────────────────────────

def sma(series: pd.Series, n: int) -> pd.Series:
    return series.rolling(n).mean()

def ema(series: pd.Series, n: int) -> pd.Series:
    return series.ewm(span=n, adjust=False).mean()

def rsi(series: pd.Series, n: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0).rolling(n).mean()
    loss = (-delta.clip(upper=0)).rolling(n).mean()
    rs = gain / loss.replace(0, np.nan)
    return 100 - 100 / (1 + rs)

def stoch_rsi(series: pd.Series, rsi_len: int = 14, stoch_len: int = 14,
              smooth_k: int = 3, smooth_d: int = 3):
    r = rsi(series, rsi_len)
    lo = r.rolling(stoch_len).min()
    hi = r.rolling(stoch_len).max()
    k = 100 * (r - lo) / (hi - lo).replace(0, np.nan)
    k = k.rolling(smooth_k).mean()
    d = k.rolling(smooth_d).mean()
    return k, d  # k=blue RSI line, d=orange signal line

def bollinger(series: pd.Series, n: int = 20, std: float = 2.0):
    mid = series.rolling(n).mean()
    dev = series.rolling(n).std()
    return mid + std * dev, mid, mid - std * dev  # upper, mid, lower

def vwap_daily(df: pd.DataFrame) -> pd.Series:
    tp = (df["high"] + df["low"] + df["close"]) / 3
    tpv = tp * df["volume"]
    date_key = df.index.date
    date_s = pd.Series(date_key, index=df.index)
    cum_tpv = tpv.groupby(date_s).cumsum()
    cum_vol = df["volume"].groupby(date_s).cumsum()
    return cum_tpv / cum_vol

def macd(series: pd.Series, fast: int = 12, slow: int = 26, sig: int = 9):
    m = ema(series, fast) - ema(series, slow)
    s = ema(m, sig)
    return m - s  # histogram

def adx(df: pd.DataFrame, n: int = 14) -> pd.Series:
    h, l, c = df["high"], df["low"], df["close"]
    tr = pd.concat([h - l, (h - c.shift()).abs(), (l - c.shift()).abs()], axis=1).max(axis=1)
    atr = tr.rolling(n).mean()
    dmp = h.diff().clip(lower=0)
    dmm = (-l.diff()).clip(lower=0)
    dmp = dmp.where(dmp > dmm, 0)
    dmm = dmm.where(dmm > dmp.shift(), 0)  # slight approximation
    di_p = 100 * dmp.rolling(n).mean() / atr
    di_m = 100 * dmm.rolling(n).mean() / atr
    dx = 100 * (di_p - di_m).abs() / (di_p + di_m).replace(0, np.nan)
    return dx.rolling(n).mean()

def slope(series: pd.Series, n: int = 5) -> pd.Series:
    """Normalized slope over n bars (% change per bar)."""
    return series.pct_change(n) / n * 100

# ─────────────────────────────────────────────
# BUILD FEATURE FRAMES
# ─────────────────────────────────────────────

def build_btc_1h(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["ma20"] = sma(d["close"], 20)       # 20-hour MA
    d["ma50"] = sma(d["close"], 50)       # 50-hour MA
    d["ma20_slope"] = slope(d["ma20"], 3)
    d["adx"] = adx(d, 14)
    d["adx_slope"] = d["adx"].diff(3)
    d["vwap"] = vwap_daily(d)
    k, dd = stoch_rsi(d["close"])
    d["srsi_k"] = k
    d["srsi_d"] = dd
    d["srsi_slope"] = d["srsi_k"].diff(3)
    return d

def build_tao_5m(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    # MAs in minutes → periods on 5-min chart
    d["ma20m"] = sma(d["close"], 4)       # 20-min MA  (4 × 5min)
    d["ma50m"] = sma(d["close"], 10)      # 50-min MA (10 × 5min)
    d["ma200m"] = sma(d["close"], 40)     # 200-min MA (40 × 5min)
    d["ma20m_slope"] = slope(d["ma20m"], 3)
    d["ma50m_slope"] = slope(d["ma50m"], 3)
    d["death_cross_dist"] = d["ma50m"] - d["ma200m"]  # negative = death cross
    d["death_cross_slope"] = d["death_cross_dist"].diff(6)  # narrowing = approaching
    d["vwap"] = vwap_daily(d)
    k, dd = stoch_rsi(d["close"])
    d["srsi_k"] = k          # blue line
    d["srsi_d"] = dd         # orange signal line
    bb_u, bb_m, bb_l = bollinger(d["close"], 20, 2)
    d["bb_upper"] = bb_u
    d["bb_mid"] = bb_m
    d["bb_lower"] = bb_l
    d["macd_hist"] = macd(d["close"])
    d["macd_hist_slope"] = d["macd_hist"].diff(3)
    d["body"] = (d["close"] - d["open"]).abs()
    d["range"] = d["high"] - d["low"]
    d["is_red"] = d["close"] < d["open"]
    d["strong_red"] = d["is_red"] & (d["body"] > 0.5 * d["range"])
    d["consec_red"] = d["strong_red"].rolling(3).sum()
    return d

# ─────────────────────────────────────────────
# PERMISSION GATES
# ─────────────────────────────────────────────

def btc_permission(btc_row) -> bool:
    """Return True if BTC 1-hour chart allows trading."""
    try:
        # 1. 20-hr MA flat / gently down / ranging up (slope > -0.03% per bar)
        if btc_row["ma20_slope"] < -0.03:
            return False
        # 2. 20-hr MA not crossing below or beneath 50-hr MA
        if btc_row["ma20"] < btc_row["ma50"]:
            return False
        # 3. Stoch RSI not zeroed (< 2) or steeply dropping from OB (> 80 & slope < -5)
        if btc_row["srsi_k"] < 2:
            return False
        if btc_row["srsi_k"] > 80 and btc_row["srsi_slope"] < -5:
            return False
        # 4. ADX not rising above 20 while price below 20-hr MA
        if btc_row["close"] < btc_row["ma20"] and btc_row["adx"] > 20 and btc_row["adx_slope"] > 0:
            return False
        # 5. BTC price must be BELOW session VWAP
        if btc_row["close"] >= btc_row["vwap"]:
            return False
    except Exception:
        return False
    return True

def tao_standdown(t, tao_3) -> bool:
    """Return True if we must stand down (DO NOT trade)."""
    try:
        # 1. 50-min MA heading toward Death Cross with 200-min MA (and close)
        if t["death_cross_dist"] < 0.5 and t["death_cross_slope"] < 0:
            return True
        # 2. 20-min MA steeply sloping down (slope < -0.05%)
        if t["ma20m_slope"] < -0.05:
            return True
        # 3. Candles below ALL 3 MAs for 3+ candles
        all_below_3 = all(
            r["close"] < r["ma20m"] and r["close"] < r["ma50m"] and r["close"] < r["ma200m"]
            for _, r in tao_3.iterrows()
        )
        if all_below_3:
            return True
        # 4. 3+ strong consecutive red candles
        if t["consec_red"] >= 3:
            return True
        # 5. MACD histogram intensifying red (negative & more negative)
        if t["macd_hist"] < 0 and t["macd_hist_slope"] < 0:
            return True
    except Exception:
        return False
    return False

# ─────────────────────────────────────────────
# BACKTEST
# ─────────────────────────────────────────────

POSITION_SIZE_USD = 50_000
MAKER_FEE = 0.001   # 0.1% round-trip (Binance taker)

def run_backtest(tao_5m: pd.DataFrame, btc_1h: pd.DataFrame) -> pd.DataFrame:
    # Align BTC 1h to TAO 5m: forward-fill hourly signal onto 5-min index
    btc_reindexed = btc_1h.reindex(tao_5m.index, method="ffill")

    trades = []
    in_trade = False
    entry_price = 0.0
    entry_time = None
    entry_qty = 0.0
    was_oversold = False   # track that stoch RSI dipped below 20 recently
    recent_os_window = 12  # candles (~1 hour) to remember oversold

    srsi_k = tao_5m["srsi_k"]
    srsi_d = tao_5m["srsi_d"]

    for i in range(50, len(tao_5m)):  # warm-up period
        t = tao_5m.iloc[i]
        b = btc_reindexed.iloc[i]
        tao_3 = tao_5m.iloc[max(0, i - 3):i]

        # ── EXIT LOGIC ──────────────────────────────────────
        if in_trade:
            k_now = srsi_k.iloc[i]
            d_now = srsi_d.iloc[i]
            k_prev = srsi_k.iloc[i - 1]
            d_prev = srsi_d.iloc[i - 1]

            # Exit when signal (orange/d) crosses above RSI (blue/k) and either overbought or k > 25
            signal_cross_above = (d_prev < k_prev) and (d_now >= k_now)  # d crossed above k
            if signal_cross_above and (k_now > 80 or k_now > 25):
                pnl_pct = (t["close"] - entry_price) / entry_price
                pnl_usd = POSITION_SIZE_USD * pnl_pct - POSITION_SIZE_USD * MAKER_FEE
                trades.append({
                    "entry_time": entry_time,
                    "exit_time": t.name if hasattr(t, "name") else tao_5m.index[i],
                    "entry_price": entry_price,
                    "exit_price": t["close"],
                    "pnl_pct": pnl_pct * 100,
                    "pnl_usd": pnl_usd,
                    "bars_held": i - entry_bar,
                })
                in_trade = False
                was_oversold = False
            continue

        # ── ENTRY LOGIC ──────────────────────────────────────
        k = srsi_k.iloc[i]
        d = srsi_d.iloc[i]
        k_prev = srsi_k.iloc[i - 1]
        d_prev = srsi_d.iloc[i - 1]

        if pd.isna(k) or pd.isna(d):
            continue

        # Track oversold condition (within recent window)
        if k < 20:
            was_oversold = True
        # Reset oversold memory if stoch RSI has gone well above 50
        if k > 50:
            was_oversold = False

        # Skip if stoch RSI is zeroed (< 2)
        if k < 2:
            continue

        # Entry trigger: was oversold, now signal (d=orange) is BELOW k (blue) [bullish cross occurred]
        # Bullish cross: d crossed below k → d < k AND previous bar d >= k
        bullish_cross = (d_prev >= k_prev) and (d < k)
        if not (was_oversold and bullish_cross):
            continue

        # BTC permission gate
        if not btc_permission(b):
            continue

        # TAO stand-down check
        if tao_standdown(t, tao_3):
            continue

        # All conditions met → ENTER
        in_trade = True
        entry_price = t["close"]
        entry_time = tao_5m.index[i]
        entry_bar = i
        entry_qty = POSITION_SIZE_USD / entry_price
        was_oversold = False

    return pd.DataFrame(trades)

# ─────────────────────────────────────────────
# REPORTING
# ─────────────────────────────────────────────

def report(trades: pd.DataFrame, tao_5m: pd.DataFrame):
    if trades.empty:
        print("No trades generated.")
        return

    n = len(trades)
    winners = trades[trades["pnl_usd"] > 0]
    losers  = trades[trades["pnl_usd"] <= 0]
    win_rate = len(winners) / n * 100
    total_pnl = trades["pnl_usd"].sum()
    avg_win = winners["pnl_usd"].mean() if len(winners) else 0
    avg_loss = losers["pnl_usd"].mean() if len(losers) else 0
    profit_factor = (winners["pnl_usd"].sum() / -losers["pnl_usd"].sum()
                     if len(losers) and losers["pnl_usd"].sum() != 0 else float("inf"))

    # Equity curve
    equity = POSITION_SIZE_USD + trades["pnl_usd"].cumsum()
    peak = equity.cummax()
    drawdown = (equity - peak) / peak * 100
    max_dd = drawdown.min()

    # Annualised return (rough: total bars / bars_per_year)
    total_bars = len(tao_5m)
    bars_per_year = 365 * 24 * 12  # 5-min bars
    years = total_bars / bars_per_year
    ann_return = (total_pnl / POSITION_SIZE_USD) / years * 100 if years > 0 else 0

    print("=" * 52)
    print("  TAO 5-MIN MEAN-REVERSION — BACKTEST RESULTS")
    print("=" * 52)
    print(f"  Period        : {tao_5m.index[0].date()} → {tao_5m.index[-1].date()}")
    print(f"  Total trades  : {n}")
    print(f"  Win rate      : {win_rate:.1f}%")
    print(f"  Profit factor : {profit_factor:.2f}")
    print(f"  Avg win       : ${avg_win:,.0f}")
    print(f"  Avg loss      : ${avg_loss:,.0f}")
    print(f"  Total P&L     : ${total_pnl:,.0f}")
    print(f"  Ann. return   : {ann_return:.1f}%  (on ${POSITION_SIZE_USD:,} base)")
    print(f"  Max drawdown  : {max_dd:.1f}%")
    print(f"  Avg hold(bars): {trades['bars_held'].mean():.1f}  ({trades['bars_held'].mean()*5:.0f} min)")
    print("=" * 52)

    print("\nTop 5 wins:")
    print(winners.nlargest(5, "pnl_usd")[["entry_time","exit_time","entry_price","exit_price","pnl_usd","bars_held"]].to_string(index=False))
    print("\nTop 5 losses:")
    print(losers.nsmallest(5, "pnl_usd")[["entry_time","exit_time","entry_price","exit_price","pnl_usd","bars_held"]].to_string(index=False))

    trades.to_csv("/home/user/demo/backtest_trades.csv", index=False)
    print("\nAll trades saved → backtest_trades.csv")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print("Generating synthetic TAO/USDT 5-min + BTC/USDT 1-hour data...")
    tao_raw, btc_raw = generate_data()
    print(f"  TAO: {len(tao_raw)} 5-min bars  ({tao_raw.index[0].date()} → {tao_raw.index[-1].date()})")
    print(f"  BTC: {len(btc_raw)} 1-hour bars")

    print("Calculating indicators...")
    tao = build_tao_5m(tao_raw)
    btc = build_btc_1h(btc_raw)

    print("Running backtest...")
    trades = run_backtest(tao, btc)

    report(trades, tao)
