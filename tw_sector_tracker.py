#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股族群分類與每日盤後資料追蹤
==============================================================
功能：
  1. 以 yfinance 抓取個股日線資料（上市 .TW、上櫃 .TWO 自動判斷）
  2. 計算個股：收盤價、當日漲跌幅、近 5 日累積漲跌幅、成交金額
  3. 依自訂族群字典計算：
       - 族群當日漲幅（簡單平均 / 成交金額加權平均）
       - 族群近 5 日漲幅（簡單平均）
       - 資金占比（族群成交金額 / 所有追蹤個股總成交金額）
  4. 輸出 CSV（UTF-8 with BOM，Excel 可直接開啟）並回傳 DataFrame

安裝：
    pip install yfinance pandas

執行：
    python tw_sector_tracker.py

排程（每日盤後 14:30 之後執行，例如 15:00）：
    Linux/macOS crontab:  0 15 * * 1-5 /usr/bin/python3 /path/tw_sector_tracker.py
    Windows：工作排程器

注意：
  * yfinance 沒有「成交金額」欄位，本程式以「收盤價 × 成交股數」近似。
    若需要證交所官方成交金額，可改用 TWSE OpenAPI（見檔尾說明）。
  * 同一檔股票可同時屬於多個族群（例如同時是 PCB 與 CCL 概念）。
    各族群資金占比的分母是「不重複個股」的總成交金額，
    因此各族群占比加總可能超過 100%，這是正常現象。
"""

from __future__ import annotations

import datetime as dt
import os
from typing import Dict, List, Tuple

import pandas as pd
import yfinance as yf

# ----------------------------------------------------------------------
# 1. 設定區
# ----------------------------------------------------------------------

# 族群詳細字典：{ 族群: { 產業分類(供應鏈環節): [ (代號, 名稱, 說明), ... ] } }
# 同一檔股票可在不同族群有不同的產業分類與說明。
# 一檔股票可同時列在多個族群；同一族群內重複的代號會自動去重。
SECTOR_DETAIL: Dict[str, Dict[str, List[Tuple[str, str, str]]]] = {
    # ---------------- PCB 概念股 ----------------
    "PCB概念股": {
        "玻纖／銅箔": [
            ("1815", "富喬", "電子級玻纖紗／布"),
            ("1802", "台玻", "電子級玻纖布"),
            ("8358", "金居", "PCB 用電解銅箔"),
        ],
        "銅箔基板 CCL／FCCL": [
            ("2383", "台光電", "低損耗銅箔基板"),
            ("6274", "台燿", "高速低損耗銅箔基板"),
            ("6213", "聯茂", "銅箔基板／黏合片"),
            ("8039", "台虹", "軟性銅箔基板"),
        ],
        "硬板／HDI": [
            ("2368", "金像電", "伺服器／網通多層板"),
            ("3044", "健鼎", "多層板／HDI"),
            ("2313", "華通", "HDI／高層次板"),
            ("2367", "燿華", "HDI／軟硬結合板"),
        ],
        "軟板／軟硬結合板": [
            ("4958", "臻鼎-KY", "FPC／軟硬結合板"),
            ("6269", "台郡", "FPC／軟板模組"),
        ],
        "IC 載板": [
            ("3037", "欣興", "IC 載板／HDI"),
            ("8046", "南電", "ABF／BT 載板"),
            ("3189", "景碩", "FCBGA／CSP 載板"),
        ],
        "製程設備與耗材": [
            ("2467", "志聖", "曝光／壓膜／烘烤"),
            ("6664", "群翊", "塗佈／壓膜／乾燥"),
            ("6438", "迅得", "PCB 自動化設備"),
            ("3563", "牧德", "PCB 光學檢測"),
            ("8021", "尖點", "PCB 鑽針／銑刀"),
        ],
    },

    # ---------------- 設備與廠務概念股 ----------------
    "設備與廠務概念股": {
        "建廠與無塵室工程": [
            ("2404", "漢唐", "無塵室／廠務統包"),
            ("6139", "亞翔", "無塵室／機電工程"),
            ("5536", "聖暉", "無塵室／機電整合"),
            ("6196", "帆宣", "廠務統包／設備整合"),
            ("6691", "洋基工程", "無塵室／空調機電"),
        ],
        "廠務供應系統": [
            ("6613", "朋億", "化學品／氣體供應"),
            ("6667", "信紘科", "化學供應／廢液處理"),
            ("6944", "兆聯實業", "超純水／廢水回收"),
        ],
        "製程與封裝設備": [
            ("3131", "弘塑", "濕製程設備"),
            ("3583", "辛耘", "清洗／濕製程設備"),
            ("6187", "萬潤", "點膠／先進封裝設備"),
            ("6640", "均華", "晶粒挑揀／黏晶設備"),
        ],
        "測試設備與介面": [
            ("2360", "致茂", "半導體測試設備"),
            ("6223", "旺矽", "探針卡／測試設備"),
            ("6510", "精測", "探針卡／IC 測試板"),
        ],
    },

    # ---------------- 被動元件概念股 ----------------
    "被動元件概念股": {
        "MLCC 積層陶瓷電容": [
            ("2327", "國巨", "多元被動元件"),
            ("2492", "華新科", "MLCC／晶片電阻"),
            ("3026", "禾伸堂", "高壓、低損耗 MLCC"),
            ("6173", "信昌電", "特殊 MLCC／介電瓷粉"),
        ],
        "電阻": [
            ("2478", "大毅", "晶片電阻／電流感測電阻"),
            ("3624", "光頡", "薄膜精密／厚膜電阻"),
            ("6834", "天二科技", "厚膜／薄膜／金屬板電阻"),
        ],
        "鋁質電容": [
            ("2472", "立隆電", "液態／固態／混合電容"),
            ("8042", "金山電", "鋁電解／高分子固態"),
            ("6449", "鈺邦", "捲繞／疊層固態電容"),
            ("2375", "凱美", "鋁電解／固態／混合"),
        ],
        "電感與磁性元件": [
            ("3357", "臺慶科", "一體成型功率電感"),
            ("3236", "千如", "功率電感／濾波元件"),
            ("6432", "今展科", "功率電感／射頻電感"),
        ],
        "保護元件": [
            ("2428", "興勤", "熱敏電阻／壓敏電阻"),
            ("6224", "聚鼎", "PPTC 過電流保護"),
        ],
        "上游材料": [
            ("6127", "九豪", "晶片電阻用陶瓷基板"),
            ("4760", "勤凱", "被動元件導電漿"),
            ("6175", "立敦", "電容用化成箔／電蝕箔"),
        ],
    },

    # ---------------- 玻璃基板概念股 ----------------
    "玻璃基板概念股": {
        "基板開發與封裝整合": [
            ("2409", "友達", "玻璃核心基板開發"),
            ("3481", "群創", "玻璃核心／面板級封裝"),
            ("3037", "欣興", "玻璃核心載板技術開發"),
            ("3149", "正達", "玻璃薄化／精密加工、熱成型玻璃基板技術開發"),
            ("4768", "晶呈科技", "玻璃核心基板產品、玻璃通孔與金屬化製程"),
        ],
        "雷射與精密加工設備": [
            ("8027", "鈦昇", "玻璃雷射改質／切割"),
            ("8064", "東捷", "雷射改質／玻璃切割"),
            ("6207", "雷科", "TGV 飛秒雷射鑽孔"),
        ],
        "蝕刻、鍍膜與增層設備": [
            ("3583", "辛耘", "濕蝕刻／清洗設備"),
            ("3580", "友威科", "電漿處理／種子層鍍膜"),
            ("6664", "群翊", "壓膜／塗佈／烘烤"),
        ],
        "檢測與故障分析": [
            ("3055", "蔚華科", "雷射改質、孔洞與微裂紋檢測"),
            ("3289", "宜特", "玻璃基板故障分析服務"),
        ],
    },

    # ---------------- 石英元件概念股 ----------------
    "石英元件概念股": {
        "石英元件": [
            ("3042", "晶技", "石英晶體／振盪器；光通訊、伺服器與車用應用"),
            ("2484", "希華", "人工水晶／頻率元件；從晶棒到成品的垂直整合"),
            ("8289", "泰藝", "溫補／恆溫振盪器；精準時頻與高穩定度產品"),
            ("3221", "台嘉碩", "石英元件／SAW 濾波器；通訊與車用頻率元件"),
            ("8182", "加高", "石英晶體／振盪器；工業、車用與消費電子應用"),
            ("6174", "安碁", "石英晶體／振盪器；寬溫規格與客製化設計"),
        ],
    },

    # ---------------- 光通訊概念股 ----------------
    "光通訊概念股": {
        "磊晶材料": [
            ("2455", "全新", "雷射／光偵測器磊晶"),
            ("3081", "聯亞", "磷化銦磊晶／光源材料"),
            ("4971", "IET-KY", "磷化銦等化合物磊晶"),
        ],
        "晶片設計與晶圓製造": [
            ("2454", "聯發科", "高速傳輸／光互連晶片"),
            ("2330", "台積電", "矽光子／光電整合製程"),
            ("2303", "聯電", "矽光子晶圓代工"),
            ("3105", "穩懋", "光電元件／化合物半導體代工"),
            ("4991", "環宇-KY", "雷射／檢光元件晶圓代工"),
        ],
        "光被動元件、濾光片與光纖連接": [
            ("3363", "上詮", "光纖陣列／光連接"),
            ("3163", "波若威", "光被動元件／模組"),
            ("6442", "光聖", "光纖連接／光通訊元件"),
            ("6588", "東典光電", "光通訊薄膜濾光片"),
            ("6426", "統新", "光通訊濾光片／鍍膜"),
        ],
        "光電元件、封裝與光收發模組": [
            ("4979", "華星光", "雷射元件／高速光引擎"),
            ("3234", "光環", "雷射／檢光元件"),
            ("3450", "聯鈞", "光元件封裝測試"),
            ("6451", "訊芯-KY", "光收發模組／封裝測試"),
            ("4977", "眾達-KY", "光收發模組"),
            ("4908", "前鼎", "光收發模組／光次模組"),
        ],
        "耦光、自動化與測試設備": [
            ("4573", "高明鐵", "主動耦光／精密對位"),
            ("6187", "萬潤", "光耦合／點膠製程設備"),
            ("6223", "旺矽", "矽光子晶圓測試"),
            ("2360", "致茂", "光電量測／自動化測試"),
            ("7769", "鴻勁", "測試分選／CPO 測試佈局"),
        ],
    },

    # ---------------- 銅箔基板概念股 ----------------
    "銅箔基板概念股": {
        "銅箔基板": [
            ("2383", "台光電", "高頻高速／無鹵基材"),
            ("6274", "台燿", "高頻高速基板／黏合片"),
            ("6213", "聯茂", "高速低損耗基板"),
            ("1303", "南亞", "基板／銅箔／樹脂整合"),
        ],
        "軟性銅箔基板": [
            ("8039", "台虹", "軟性基板／高頻材料"),
            ("3354", "律勝", "軟性基板／保護膠膜"),
            ("4939", "亞電", "軟性基板／覆蓋膜"),
        ],
        "銅箔材料": [
            ("8358", "金居", "電解銅箔／PCB用材料"),
            ("4989", "榮科", "電解銅箔／低粗糙度銅箔"),
        ],
        "玻璃纖維紗與玻纖布": [
            ("1802", "台玻", "電子級玻纖紗／玻纖布"),
            ("1815", "富喬", "電子級玻纖紗／玻纖布"),
            ("5340", "建榮", "電子級／超薄型玻纖布"),
            ("5475", "德宏", "電子級／薄型玻纖布"),
        ],
        "樹脂、固化劑與電子材料": [
            ("4722", "國精化", "低介電損耗樹脂"),
            ("4764", "雙鍵", "CCL用電子材料"),
            ("7763", "崇舜", "特殊胺類固化劑"),
            ("1717", "長興", "合成樹脂／乾膜光阻"),
        ],
    },

    # ---------------- SiC 碳化矽概念股 ----------------
    "SiC碳化矽概念股": {
        "基板與磊晶": [
            ("6488", "環球晶", "SiC 基板"),
            ("6125", "廣運", "透過盛新材料布局 SiC 基板"),
            ("3016", "嘉晶", "SiC 磊晶"),
        ],
        "晶圓製造": [
            ("3707", "漢磊", "SiC 元件代工"),
            ("2342", "茂矽", "SiC 代工開發與合作"),
            ("3061", "台亞", "透過積亞製造 SiC 元件"),
        ],
        "功率元件": [
            ("2481", "強茂", "SiC 二極體／MOSFET"),
            ("5425", "台半", "SiC 蕭特基整流器"),
            ("3317", "尼克森", "SiC MOSFET／功率模組"),
            ("8261", "富鼎", "SiC 二極體／高壓 MOSFET"),
        ],
        "模組與封裝": [
            ("8255", "朋程", "SiC 功率模組"),
            ("6271", "同欣電", "SiC 模組封裝技術"),
        ],
        "製程設備、耗材與測試": [
            ("3583", "辛耘", "SiC 薄化貼合／剝離設備"),
            ("1560", "中砂", "SiC 晶圓研磨砂輪"),
            ("2360", "致茂", "SiC 功率元件測試設備"),
        ],
    },

    # ---------------- 封測族群概念股（原圖無個股說明，說明欄留空）----------------
    "封測族群概念股": {
        "綜合型封測": [
            ("3711", "日月光投控", ""),
        ],
        "記憶體封測": [
            ("6239", "力成", ""), ("2329", "華泰", ""), ("8110", "華東", ""), ("8131", "福懋科", ""),
        ],
        "驅動IC封測": [
            ("6147", "頎邦", ""), ("8150", "南茂", ""),
        ],
        "傳統／利基型封測": [
            ("2441", "超豐", ""), ("2369", "菱生", ""),
        ],
        "專業測試／後段服務": [
            ("2449", "京元電子", ""), ("6257", "矽格", ""), ("3264", "欣銓", ""), ("6261", "久元", ""),
        ],
        "晶圓級／特殊封裝": [
            ("3374", "精材", ""), ("6271", "同欣電", ""), ("6451", "訊芯-KY", ""), ("3265", "台星科", ""),
        ],
        "探針卡／測試介面": [
            ("6223", "旺矽", ""), ("6217", "中探針", ""), ("6510", "精測", ""),
            ("6515", "穎崴", ""), ("6683", "雍智科技", ""),
        ],
        # 久元 6261 原圖同時列在「專業測試」與此環節，明細表只保留第一筆（專業測試／後段服務）
        "測試／檢測設備": [
            ("2360", "致茂", ""), ("7769", "鴻勁", ""), ("6261", "久元", ""),
            ("3563", "牧德", ""), ("3455", "由田", ""),
        ],
    },

    # ---------------- AI 伺服器散熱概念股（說明欄＝原圖「核心能力」）----------------
    "AI伺服器散熱概念股": {
        "晶片級散熱元件": [
            ("3653", "健策", "均熱片、VC Lid、微流道蓋板"),
            ("2486", "一詮", "封裝均熱片、VC Lid、Metal TIM"),
        ],
        "散熱模組／液冷方案": [
            ("3017", "奇鋐", "散熱模組、水冷板、CDU"),
            ("3324", "雙鴻", "散熱模組、水冷板、液冷系統"),
            ("3483", "力致", "風扇、散熱模組、液冷方案"),
            ("6230", "尼得科超眾", "熱管、均熱板、液冷系統"),
            ("6831", "邁科", "3D VC、散熱模組、液冷模組"),
            ("2354", "鴻準", "液冷模組、分流器、金屬機構件"),
        ],
        "液冷關鍵零組件": [
            ("8996", "高力", "板式熱交換器、CDU、液冷系統"),
            ("6805", "富世達", "液冷快接頭、精密機構件"),
        ],
        "主動散熱": [
            ("2421", "建準", "風扇、水冷板、泵浦、CDU"),
        ],
        "系統整合": [
            ("2308", "台達電", "水冷板、CDU、資料中心整合"),
        ],
    },
}

# 由詳細字典衍生「族群 → {代號: 名稱}」，供下方計算流程使用（不需手動維護）
SECTORS: Dict[str, Dict[str, str]] = {
    sector: {code: name for items in segments.values() for code, name, _ in items}
    for sector, segments in SECTOR_DETAIL.items()
}

# 近幾日累積漲跌幅（5 個交易日）
N_DAYS = 5

# 抓取的歷史長度：N_DAYS 需要至少 N_DAYS+1 筆資料，抓一個月較保險（含假日）
HISTORY_PERIOD = "1mo"

# 輸出資料夾
OUTPUT_DIR = "output"

# 成交金額單位換算：1 億 = 1e8
YI = 1e8


# ----------------------------------------------------------------------
# 2. 資料抓取
# ----------------------------------------------------------------------

def _download_batch(codes: list[str], suffix: str) -> Dict[str, pd.DataFrame]:
    """
    一次下載多檔股票，回傳 {代號: DataFrame(Close, Volume)}。
    只回傳「資料筆數足夠」的股票；抓不到的由呼叫端改用其他後綴重試。
    """
    if not codes:
        return {}

    tickers = [f"{c}{suffix}" for c in codes]
    raw = yf.download(
        tickers,
        period=HISTORY_PERIOD,
        interval="1d",
        auto_adjust=False,   # 不還原股利，才會與盤面收盤價、漲跌幅一致
        group_by="ticker",   # 欄位為 (ticker, 欄位名) 的 MultiIndex
        progress=False,
        threads=True,
    )

    result: Dict[str, pd.DataFrame] = {}
    for code, ticker in zip(codes, tickers):
        try:
            # 單一股票時 yfinance 可能不回傳 MultiIndex，這裡一併處理
            df = raw[ticker] if isinstance(raw.columns, pd.MultiIndex) else raw
            df = df[["Close", "Volume"]].dropna()
        except (KeyError, TypeError):
            continue
        # 至少要有 N_DAYS + 1 筆收盤價才算得出 N 日漲幅
        if len(df) >= N_DAYS + 1:
            result[code] = df
    return result


def fetch_price_history(codes: list[str]) -> Dict[str, pd.DataFrame]:
    """
    先以上市(.TW)抓取，抓不到的再以上櫃(.TWO)重試。
    回傳 {代號: DataFrame}；兩邊都抓不到的代號會被略過並印出警告。
    """
    data = _download_batch(codes, ".TW")

    missing = [c for c in codes if c not in data]
    if missing:
        data.update(_download_batch(missing, ".TWO"))

    still_missing = [c for c in codes if c not in data]
    if still_missing:
        print(f"[警告] 以下代號抓不到資料（請確認代號/是否停牌）：{still_missing}")
    return data


# ----------------------------------------------------------------------
# 3. 個股指標計算
# ----------------------------------------------------------------------

def build_stock_table(history: Dict[str, pd.DataFrame],
                      names: Dict[str, str]) -> pd.DataFrame:
    """
    將每檔股票的歷史資料轉成「最新交易日」的指標一列。
    欄位：
      代號、名稱、日期、收盤價、當日漲跌幅(%)、近N日漲跌幅(%)、
      成交金額(元)、成交金額(億)
    """
    rows = []
    for code, df in history.items():
        close, volume = df["Close"], df["Volume"]

        last_close = close.iloc[-1]
        prev_close = close.iloc[-2]           # 前一交易日收盤
        base_close = close.iloc[-(N_DAYS + 1)]  # N 個交易日前的收盤

        pct_1d = (last_close / prev_close - 1) * 100
        pct_nd = (last_close / base_close - 1) * 100

        # 成交金額（近似）= 收盤價 × 成交股數
        turnover = float(last_close * volume.iloc[-1])

        rows.append({
            "代號": code,
            "名稱": names.get(code, ""),
            "日期": close.index[-1].date(),
            "收盤價": round(float(last_close), 2),
            "當日漲跌幅(%)": round(float(pct_1d), 2),
            f"近{N_DAYS}日漲跌幅(%)": round(float(pct_nd), 2),
            "成交金額(元)": turnover,
            "成交金額(億)": round(turnover / YI, 2),
        })

    stock_df = pd.DataFrame(rows)
    if stock_df.empty:
        return stock_df

    # 若部分股票最新資料日期落後（停牌等），以多數股票的最新日期為準，
    # 避免把不同日期的資料混在一起計算。
    trade_date = stock_df["日期"].max()
    stale = stock_df[stock_df["日期"] != trade_date]
    if not stale.empty:
        print(f"[警告] 以下股票最新資料不是 {trade_date}，已排除：{stale['代號'].tolist()}")
        stock_df = stock_df[stock_df["日期"] == trade_date]

    return stock_df.reset_index(drop=True)


# ----------------------------------------------------------------------
# 4. 族群指標計算
# ----------------------------------------------------------------------

def build_sector_table(stock_df: pd.DataFrame,
                       sectors: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    """
    族群摘要：
      - 檔數
      - 當日平均漲幅（簡單平均）
      - 當日加權漲幅（以成交金額加權）
      - 近 N 日平均漲幅（簡單平均）
      - 族群成交金額 / 資金占比
    """
    col_1d = "當日漲跌幅(%)"
    col_nd = f"近{N_DAYS}日漲跌幅(%)"

    # 族群 ←→ 個股 對應表（一檔股票可出現在多個族群）
    membership = pd.DataFrame(
        [(sector, code) for sector, stocks in sectors.items() for code in stocks],
        columns=["族群", "代號"],
    )
    merged = membership.merge(stock_df, on="代號", how="inner")

    # 資金占比分母：所有「不重複」追蹤個股的總成交金額
    total_turnover = stock_df["成交金額(元)"].sum()

    rows = []
    for sector, g in merged.groupby("族群", sort=False):
        sector_turnover = g["成交金額(元)"].sum()
        # 成交金額加權平均漲幅；若成交金額總和為 0 則退回簡單平均
        weighted_1d = (
            (g[col_1d] * g["成交金額(元)"]).sum() / sector_turnover
            if sector_turnover > 0 else g[col_1d].mean()
        )
        rows.append({
            "族群": sector,
            "有效檔數": len(g),
            "設定檔數": len(sectors[sector]),
            "當日平均漲幅(%)": round(g[col_1d].mean(), 2),
            "當日加權漲幅(%)": round(weighted_1d, 2),
            f"近{N_DAYS}日平均漲幅(%)": round(g[col_nd].mean(), 2),
            "成交金額(億)": round(sector_turnover / YI, 2),
            "資金占比(%)": round(sector_turnover / total_turnover * 100, 2),
        })

    sector_df = pd.DataFrame(rows)
    # 依當日平均漲幅由高到低排序
    return sector_df.sort_values("當日平均漲幅(%)", ascending=False).reset_index(drop=True)


# ----------------------------------------------------------------------
# 5. 主流程
# ----------------------------------------------------------------------

def build_detail_table(stock_df: pd.DataFrame) -> pd.DataFrame:
    """
    族群 × 個股 明細（長表）：同一檔股票依所屬族群各列一次，
    並帶上該族群內的「產業分類」與「說明」。
    """
    members = pd.DataFrame(
        [(sector, seg, code, name, desc)
         for sector, segments in SECTOR_DETAIL.items()
         for seg, items in segments.items()
         for code, name, desc in items],
        columns=["族群", "產業分類", "代號", "名稱", "說明"],
    )
    # 同一族群內若代號重複，只保留第一筆
    members = members.drop_duplicates(["族群", "代號"])
    merged = members.merge(stock_df.drop(columns=["名稱"]), on="代號", how="inner")
    cols = ["族群", "代號", "名稱", "產業分類", "說明", "日期", "收盤價",
            "當日漲跌幅(%)", f"近{N_DAYS}日漲跌幅(%)", "成交金額(億)", "成交金額(元)"]
    return merged[cols].reset_index(drop=True)


def run(output_dir: str = OUTPUT_DIR):
    """執行完整流程，回傳 (族群×個股明細 DataFrame, 族群摘要 DataFrame)。"""
    # 整理所有追蹤個股（去重）與名稱對照
    names: Dict[str, str] = {}
    for stocks in SECTORS.values():
        names.update(stocks)
    codes = list(names.keys())

    print(f"追蹤 {len(codes)} 檔個股、{len(SECTORS)} 個族群，下載資料中…")
    history = fetch_price_history(codes)

    stock_df = build_stock_table(history, names)
    if stock_df.empty:
        raise RuntimeError("沒有抓到任何資料，請檢查網路或股票代號。")

    sector_df = build_sector_table(stock_df, SECTORS)

    stock_df = stock_df.sort_values("當日漲跌幅(%)", ascending=False).reset_index(drop=True)

    # 族群×個股明細（含產業分類與說明）：儀表板與 CSV 都用這張表
    detail_df = build_detail_table(stock_df)

    # 輸出 CSV（utf-8-sig 讓 Excel 開啟中文不亂碼）
    os.makedirs(output_dir, exist_ok=True)
    trade_date = stock_df["日期"].iloc[0]
    stock_path = os.path.join(output_dir, f"stocks_{trade_date}.csv")
    sector_path = os.path.join(output_dir, f"sectors_{trade_date}.csv")
    # 成交金額(元) 轉成整數輸出，避免科學記號
    out_stock = detail_df.copy()
    out_stock["成交金額(元)"] = out_stock["成交金額(元)"].round(0).astype("int64")
    out_stock.to_csv(stock_path, index=False, encoding="utf-8-sig")
    sector_df.to_csv(sector_path, index=False, encoding="utf-8-sig")

    # 另存固定檔名（latest），讓 index.html 儀表板每天讀取同一個網址
    out_stock.to_csv(os.path.join(output_dir, "stocks_latest.csv"), index=False, encoding="utf-8-sig")
    sector_df.to_csv(os.path.join(output_dir, "sectors_latest.csv"), index=False, encoding="utf-8-sig")

    print(f"\n交易日：{trade_date}")
    print(f"已輸出：{stock_path}\n        {sector_path}\n        {output_dir}/stocks_latest.csv, sectors_latest.csv\n")
    print("=== 族群摘要 ===")
    print(sector_df.to_string(index=False))
    print("\n=== 族群×個股明細 ===")
    print(detail_df.drop(columns=["成交金額(元)"]).to_string(index=False))

    return detail_df, sector_df


if __name__ == "__main__":
    run()


# ----------------------------------------------------------------------
# 附註：改用證交所 OpenAPI 取得「官方成交金額」（選用）
# ----------------------------------------------------------------------
# 上市個股日資料（含 TradeValue 成交金額、ClosingPrice、Change 等）：
#   https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL
# 上櫃個股日資料：
#   https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes
#
# 範例：
#   import requests
#   r = requests.get("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", timeout=15)
#   df = pd.DataFrame(r.json())          # 欄位 Code, Name, TradeValue, ClosingPrice ...
#   df["TradeValue"] = pd.to_numeric(df["TradeValue"], errors="coerce")
#
# 這兩支 API 只提供「當日」資料，不含歷史；因此近 5 日漲幅仍需 yfinance 或
# 自行每天存檔累積。可把官方 TradeValue 取代 build_stock_table() 裡
# 「收盤價 × 成交股數」的近似值。
