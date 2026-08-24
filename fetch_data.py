import datetime
import json
import os
import sys
import pandas as pd
import requests
import numpy as np

# 定义需要从 FRED 数据库下载的数据系列 ID 映射表 (转为直接从纽约联储官网抓取 ACM)
SERIES_IDS = {
    # 名义美债常数到期收益率 (Nominal Yields)
    "dgs_3m": "DGS3MO",        # 3个月期名义美债收益率
    "dgs_6m": "DGS6MO",        # 6个月期名义美债收益率
    "dgs_1y": "DGS1",         # 1年期名义美债收益率
    "dgs_2y": "DGS2",         # 2年期名义美债收益率
    "dgs_3y": "DGS3",         # 3年期名义美债收益率
    "dgs_5y": "DGS5",         # 5年期名义美债收益率
    "dgs_7y": "DGS7",         # 7年期名义美债收益率
    "dgs_10y": "DGS10",       # 10年期名义美债收益率
    "dgs_20y": "DGS20",       # 20年期名义美债收益率
    "dgs_30y": "DGS30",       # 30年期名义美债收益率
    
    # 实际美债收益率 (TIPS Real Yields)
    "real_5y": "DFII5",        # 5年期通胀保值国债(TIPS)收益率
    "real_10y": "DFII10",      # 10年期通胀保值国债(TIPS)收益率
    
    # 盈亏平衡通胀率 (BEI - Breakeven Inflation Rate)
    "bei_5y": "T5YIE",         # 5年期盈亏平衡通胀率
    "bei_10y": "T10YIE",       # 10年期盈亏平衡通胀率
    
    # 美联储联邦基金目标利率区间 (用于计算当前基准利率中枢)
    "fed_target_upper": "DFEDTARU", # 目标利率区间上限
    "fed_target_lower": "DFEDTARL"  # 目标利率区间下限
}

# 设置数据文件的保存目录为当前脚本所在的绝对路径 (自适应本地运行和 GitHub Actions 运行环境)
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
# 拼接最终生成的结构化 JS 数据文件路径
DATA_JS_PATH = os.path.join(OUTPUT_DIR, "data.js")

def get_fred_csv(series_id):
    """从 FRED 官方 API 下载指定数据系列的 CSV 文件并转换为 DataFrame"""
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
    print(f"  正在下载 FRED 数据系列: {series_id}...")
    try:
        response = requests.get(url, timeout=15)
        if response.status_code == 200:
            lines = response.content.decode('utf-8').split('\n')
            data = []
            for line in lines[1:]:
                parts = line.strip().split(',')
                if len(parts) == 2:
                    date_str, val_str = parts
                    if val_str != '.':
                        try:
                            data.append({'Date': date_str, series_id: float(val_str)})
                        except ValueError:
                            pass
            return pd.DataFrame(data)
        else:
            print(f"  [警告] 下载失败 {series_id}，HTTP 状态码: {response.status_code}")
            return pd.DataFrame(columns=['Date', series_id])
    except Exception as e:
        print(f"  [警告] 下载 {series_id} 时出现异常: {e}")
        return pd.DataFrame(columns=['Date', series_id])

def fetch_acm_term_premium_ny_fed():
    """从纽约联储官网直接下载并解析 ACM 期限溢价模型日频 Excel 数据 (解决 FRED 没有 ACM 模型的问题)"""
    url = "https://www.newyorkfed.org/medialibrary/media/research/data_indicators/ACMTermPremium.xls"
    local_path = os.path.join(OUTPUT_DIR, "ACMTermPremium.xls")
    
    print("正在从纽约联储官网获取 ACM 期限溢价日频数据 (10MB)...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=35)
        if response.status_code == 200 and len(response.content) > 100000:
            with open(local_path, "wb") as f:
                f.write(response.content)
            print("  下载成功，已写入本地缓存。")
        else:
            print(f"  [警告] 纽约联储下载返回状态码: {response.status_code}。尝试读取本地缓存...")
    except Exception as e:
        print(f"  [警告] 联网下载 ACM 失败: {e}。将尝试读取本地已有缓存文件...")
        
    if not os.path.exists(local_path):
        print("  [错误] 纽约联储 ACM 期限溢价本地文件不存在且下载失败！")
        return pd.DataFrame(columns=['Date', 'acm_nom_10y', 'acm_tp_10y', 'acm_exp_10y'])
        
    try:
        print("  正在解析 Excel 'ACM Daily' 工作表...")
        df = pd.read_excel(local_path, sheet_name='ACM Daily', engine='xlrd')
        # 解析日期 DATE (14-Jun-1961 格式) 转换为 YYYY-MM-DD
        df['Date'] = pd.to_datetime(df['DATE'], format='%d-%b-%Y').dt.strftime('%Y-%m-%d')
        # 提取 ACMY10 (10Y零息拟合名义利率), ACMTP10 (10Y期限溢价), ACMRNY10 (10Y期望无风险短期利率)
        df_clean = df[['Date', 'ACMY10', 'ACMTP10', 'ACMRNY10']].copy()
        df_clean = df_clean.rename(columns={
            'ACMY10': 'acm_nom_10y',
            'ACMTP10': 'acm_tp_10y',
            'ACMRNY10': 'acm_exp_10y'
        })
        print(f"  成功提取了 {len(df_clean)} 条日频 ACM 模型数据。")
        return df_clean
    except Exception as e:
        print(f"  [错误] 解析 ACM Excel 失败: {e}")
        return pd.DataFrame(columns=['Date', 'acm_nom_10y', 'acm_tp_10y', 'acm_exp_10y'])

def fetch_gold_price_history():
    """从雅虎财经接口直接获取现货黄金价格（由GC=F黄金期货模拟）的历史收盘价序列，绕开FRED的版权下载限制"""
    print("  正在从雅虎财经获取黄金历史价格 (GC=F)...")
    try:
        import yfinance as yf
        gold = yf.Ticker("GC=F")
        history = gold.history(period="max")
        if not history.empty:
            history = history.reset_index()
            history['Date'] = history['Date'].dt.strftime('%Y-%m-%d')
            history = history.rename(columns={'Close': 'gold_price'})
            return history[['Date', 'gold_price']]
    except Exception as e:
        print(f"  [警告] 无法通过 yfinance 获取黄金历史数据: {e}")
    return pd.DataFrame(columns=['Date', 'gold_price'])

def fetch_cme_dec_futures_history(year):
    """从雅虎财经接口下载当年 12 月联邦基金期货合约的全部历史收盘价"""
    year_suffix = str(year)[2:]
    ticker = f"ZQZ{year_suffix}.CBT"
    print(f"  正在获取 12月期货合约历史数据: {ticker}...")
    
    end_time = int(datetime.datetime.now().timestamp())
    start_time = end_time - (730 * 24 * 3600)
    
    url = f"https://query1.finance.yahoo.com/v7/finance/download/{ticker}?period1={start_time}&period2={end_time}&interval=1d&events=history&includeAdjustedClose=true"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            lines = response.content.decode('utf-8').split('\n')
            data = []
            for line in lines[1:]:
                parts = line.strip().split(',')
                if len(parts) >= 6 and parts[4] != 'null' and parts[4] != '':
                    try:
                        date_str = parts[0]
                        close_price = float(parts[4])
                        data.append({'Date': date_str, 'futures_price': close_price})
                    except ValueError:
                        pass
            return pd.DataFrame(data)
        else:
            print(f"  [警告] 无法直接获取 {ticker} CSV 数据 (状态码: {response.status_code})。尝试使用 yfinance...")
    except Exception as e:
        print(f"  [警告] 请求期货数据异常: {e}。尝试使用 yfinance...")

    try:
        import yfinance as yf
        future = yf.Ticker(ticker)
        history = future.history(period="max")
        if not history.empty:
            history = history.reset_index()
            history['Date'] = history['Date'].dt.strftime('%Y-%m-%d')
            history = history.rename(columns={'Close': 'futures_price'})
            return history[['Date', 'futures_price']]
    except Exception as e:
        print(f"  [警告] yfinance 备份抓取失败: {e}")
        
    print("  [错误] 无法获取期货历史数据，年底降息预期曲线将无法呈现。")
    return pd.DataFrame(columns=['Date', 'futures_price'])

def calculate_derivatives(df):
    """计算各类利差、远期曲线指标、加权预期变化历史和滚动 Z-score，以及 DoD/WoW 变动幅度 (已换算为 bp)"""
    print("开始执行数学派生指标计算...")
    
    # 1. 计算核心期限利差 (Spreads)
    df['spread_10y_2y'] = df['dgs_10y'] - df['dgs_2y']
    df['spread_10y_3m'] = df['dgs_10y'] - df['dgs_3m']
    df['spread_30y_10y'] = df['dgs_30y'] - df['dgs_10y']
    
    # 2. 计算远期名义利率 (Forward Rates)
    df['fwd_1y1y'] = (((1 + df['dgs_2y']/100)**2 / (1 + df['dgs_1y']/100)) - 1) * 100
    df['fwd_2y1y'] = (((1 + df['dgs_3y']/100)**3 / (1 + df['dgs_2y']/100)**2) - 1) * 100
    df['fwd_5y5y_nom'] = ((((1 + df['dgs_10y']/100)**10 / (1 + df['dgs_5y']/100)**5)**(1/5)) - 1) * 100
    df['fwd_10y10y'] = ((((1 + df['dgs_20y']/100)**20 / (1 + df['dgs_10y']/100)**10)**(1/10)) - 1) * 100

    # 3. 计算 5Y5Y 远期实际利率锚 (5Y5Y Forward Real Yield)
    df['fwd_5y5y_real'] = ((((1 + df['real_10y']/100)**10 / (1 + df['real_5y']/100)**5)**(1/5)) - 1) * 100

    # 4. 计算 5Y5Y 远期通胀预期 (5Y5Y Forward Inflation)
    df['fwd_5y5y_infl'] = df['fwd_5y5y_nom'] - df['fwd_5y5y_real']

    # 5. 计算 10Y10Y 对 10Y 名义偏离程度的滚动 Z-score (与全历史样本 Z-score)
    df['fwd_deviation_spread'] = df['fwd_10y10y'] - df['dgs_10y']
    
    # 滚动 Z-score：由于 window=252 会自动适应最近一年中枢，用来显示短期“异动”
    rolling = df['fwd_deviation_spread'].rolling(window=252, min_periods=30)
    df['fwd_deviation_zscore'] = (df['fwd_deviation_spread'] - rolling.mean()) / rolling.std()
    
    # 新增：全历史样本 Z-score (以自2000年以来的全样本为常数均值和标准差，可准确刻画收益率创新高时的绝对历史极值偏离)
    full_mean = df['fwd_deviation_spread'].mean()
    full_std = df['fwd_deviation_spread'].std()
    df['fwd_deviation_zscore_full'] = (df['fwd_deviation_spread'] - full_mean) / full_std

    # 6. 计算 ACM 期望短期利率 (如果已经有 acm_exp_10y 列则进行校准对齐，否则使用减法)
    if 'acm_exp_10y' not in df.columns:
        df['acm_exp_10y'] = df['acm_nom_10y'] - df['acm_tp_10y']

    # 7. 计算当年 12 月期货降息预期的历史时间序列
    df['fed_target_midpoint'] = (df['fed_target_upper'] + df['fed_target_lower']) / 2
    if 'futures_price' in df.columns:
        df['futures_implied_rate'] = 100.0 - df['futures_price']
        df['weighted_change_history_bp'] = (df['futures_implied_rate'] - df['fed_target_midpoint']) * 100
    else:
        df['weighted_change_history_bp'] = 0.0

    # ==================== 计算 DoD / WoW 真实交易日变动 (单位：bp) ====================
    # 差分并乘以 100 转化为 bp，由于已在主函数剔除空交易日，此时的 shift 对应物理交易日
    df['dgs_2y_dod'] = (df['dgs_2y'] - df['dgs_2y'].shift(1)) * 100
    df['dgs_2y_wow'] = (df['dgs_2y'] - df['dgs_2y'].shift(5)) * 100
    
    df['dgs_10y_dod'] = (df['dgs_10y'] - df['dgs_10y'].shift(1)) * 100
    df['dgs_10y_wow'] = (df['dgs_10y'] - df['dgs_10y'].shift(5)) * 100
    
    df['spread_10y_2y_dod'] = (df['spread_10y_2y'] - df['spread_10y_2y'].shift(1)) * 100
    df['spread_10y_2y_wow'] = (df['spread_10y_2y'] - df['spread_10y_2y'].shift(5)) * 100
    
    df['real_10y_dod'] = (df['real_10y'] - df['real_10y'].shift(1)) * 100
    df['real_10y_wow'] = (df['real_10y'] - df['real_10y'].shift(5)) * 100
    
    df['bei_10y_dod'] = (df['bei_10y'] - df['bei_10y'].shift(1)) * 100
    df['bei_10y_wow'] = (df['bei_10y'] - df['bei_10y'].shift(5)) * 100
    
    df['fwd_5y5y_infl_dod'] = (df['fwd_5y5y_infl'] - df['fwd_5y5y_infl'].shift(1)) * 100
    df['fwd_5y5y_infl_wow'] = (df['fwd_5y5y_infl'] - df['fwd_5y5y_infl'].shift(5)) * 100
    
    df['acm_tp_10y_dod'] = (df['acm_tp_10y'] - df['acm_tp_10y'].shift(1)) * 100
    df['acm_tp_10y_wow'] = (df['acm_tp_10y'] - df['acm_tp_10y'].shift(5)) * 100
    
    df['acm_exp_10y_dod'] = (df['acm_exp_10y'] - df['acm_exp_10y'].shift(1)) * 100
    df['acm_exp_10y_wow'] = (df['acm_exp_10y'] - df['acm_exp_10y'].shift(5)) * 100
    
    df['acm_nom_10y_dod'] = (df['acm_nom_10y'] - df['acm_nom_10y'].shift(1)) * 100
    df['acm_nom_10y_wow'] = (df['acm_nom_10y'] - df['acm_nom_10y'].shift(5)) * 100

    # CME 期货预期变动已经是 bp 计价，不需要乘以 100
    df['weighted_change_history_bp_dod'] = df['weighted_change_history_bp'] - df['weighted_change_history_bp'].shift(1)
    df['weighted_change_history_bp_wow'] = df['weighted_change_history_bp'] - df['weighted_change_history_bp'].shift(5)

    return df

def main():
    print("===== 美债拆解数据抓取清洗脚本启动 =====")
    
    # 1. 循环抓取所有配置好的 FRED 数据列，并重命名为字典里的 key
    dfs = []
    for key, series_id in SERIES_IDS.items():
        df_series = get_fred_csv(series_id)
        if not df_series.empty:
            df_series = df_series.rename(columns={series_id: key})
            dfs.append(df_series)
            
    # 2. 从纽约联储获取每日更新的 ACM 期望与期限溢价数据
    df_acm = fetch_acm_term_premium_ny_fed()
    if not df_acm.empty:
        dfs.append(df_acm)
            
    # 3. 从雅虎财经拉取伦敦黄金历史价格 (GC=F) 绕过 FRED 版权限制
    df_gold = fetch_gold_price_history()
    if not df_gold.empty:
        dfs.append(df_gold)
        
    # 4. 从雅虎财经拉取 12 月联邦基金期货的历史价格数据
    current_year = datetime.datetime.now().year
    df_futures = fetch_cme_dec_futures_history(current_year)
    if not df_futures.empty:
        dfs.append(df_futures)
        
    if not dfs:
        print("错误: 无法下载任何核心数据系列。")
        sys.exit(1)
        
    # 5. 将所有下载的 DataFrame 根据 'Date' 字段进行外联接合并
    print("正在根据 Date 主键进行外联接合并...")
    merged_df = dfs[0]
    for df_next in dfs[1:]:
        merged_df = pd.merge(merged_df, df_next, on='Date', how='outer')
        
    merged_df['Date'] = pd.to_datetime(merged_df['Date'])
    merged_df = merged_df.sort_values('Date').reset_index(drop=True)
    
    # 修正问题 1：剔除非美债交易日的空值行 (消除周末带来的 0 变动)
    merged_df = merged_df.dropna(subset=['dgs_10y']).reset_index(drop=True)
    
    # 6. 执行缺失值的前向填充 (ffill) 和后向填充 (bfill) 补齐黄金和期货在美债交易日的开市错位
    cols_to_fill = [col for col in merged_df.columns if col != 'Date']
    merged_df[cols_to_fill] = merged_df[cols_to_fill].ffill().bfill()
    
    # 7. 调用核心函数进行衍生变量与数学公式指标计算 (包含 DoD / WoW 计算)
    merged_df = calculate_derivatives(merged_df)
    
    # 8. 数据截取：只保留 2000 年之后的数据，提高前端 ECharts 缩放图表渲染速度与历史可比性
    merged_df = merged_df[merged_df['Date'] >= '2000-01-01'].reset_index(drop=True)
    
    # 9. 计算最新的统计指标（用于首屏的四个卡片展示，使用 pd.notna 保证浮点转换健壮性）
    latest_row = merged_df.iloc[-1]
    
    latest_futures_price = float(latest_row['futures_price']) if 'futures_price' in latest_row and pd.notna(latest_row['futures_price']) else 95.5
    latest_target_midpoint = float(latest_row['fed_target_midpoint']) if 'fed_target_midpoint' in latest_row and pd.notna(latest_row['fed_target_midpoint']) else 5.375
    latest_implied_rate = 100.0 - latest_futures_price
    latest_weighted_change_bp = (latest_implied_rate - latest_target_midpoint) * 100
    
    print(f"最新基准中枢: {latest_target_midpoint}%, 期货隐含年底利率: {latest_implied_rate}%")
    print(f"最新年底加权预期变化: {latest_weighted_change_bp:.2f} bp")
    
    # 10. 格式化日期，生成字符串格式供前端 JS 处理
    merged_df['Date_str'] = merged_df['Date'].dt.strftime('%Y-%m-%d')
    # 修正：丢弃无法被 json 序列化的 pandas Timestamp 类型 'Date' 列
    merged_df = merged_df.drop(columns=['Date'])
    
    # 处理 pandas 计算中可能引入 of np.inf (正无穷) 和 -np.inf (负无穷)，均转换为 NaN
    merged_df = merged_df.replace([np.inf, -np.inf], np.nan)
    # 将所有的 NaN 转换为 Python 的 None
    merged_df = merged_df.where(pd.notnull(merged_df), None)
    
    # 将 DataFrame 转换为字典数组结构
    history_data = merged_df.to_dict(orient='records')
    
    # 构造最终写入本地 json 的总体层级结构
    final_output = {
        "update_time": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "target_midpoint": latest_target_midpoint,
        "futures_contract": f"ZQZ{str(current_year)[2:]}.CBT",
        "futures_price": latest_futures_price,
        "implied_rate": latest_implied_rate,
        "weighted_change_bp": round(latest_weighted_change_bp, 2),
        "history": history_data
    }
    
    # 11. 将结构化数据写入 data.js 缓存文件
    with open(DATA_JS_PATH, 'w', encoding='utf-8') as f:
        f.write("window.chartData = " + json.dumps(final_output, indent=4) + ";")
    print(f"  成功保存缓存至 {DATA_JS_PATH}")
    print("===== 数据源全自动化处理与注入流程运行成功！ =====")

if __name__ == "__main__":
    main()
