"""讀取 output/sectors_日期.csv 歷史檔，產生 sectors_history.csv（趨勢圖）與 meta.json（更新時間）。放在每日更新腳本之後執行。"""
import csv, datetime, glob, json, re
rows = []
for f in sorted(glob.glob("output/sectors_????-??-??.csv"))[-60:]:
    d = re.search(r"\d{4}-\d{2}-\d{2}", f).group(0)
    with open(f, encoding="utf-8-sig") as fh:
        rows += [(d, r["族群"], r["資金占比(%)"]) for r in csv.DictReader(fh)]
with open("output/sectors_history.csv", "w", encoding="utf-8-sig", newline="") as fh:
    w = csv.writer(fh); w.writerow(["日期", "族群", "資金占比(%)"]); w.writerows(rows)
tw = datetime.timezone(datetime.timedelta(hours=8))
json.dump({"generated_at": datetime.datetime.now(tw).strftime("%Y-%m-%d %H:%M")}, open("output/meta.json", "w"))
