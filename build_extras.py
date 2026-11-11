"""讀取 output/sectors_日期.csv 歷史檔，產生 sectors_history.csv（趨勢圖）與 meta.json（更新時間）。
放在主程式之後執行；只有歷史資料真的變了才更新 meta.json，避免休市日也產生 commit。"""
import csv, datetime, glob, io, json, os, re
rows = []
for f in sorted(glob.glob("output/sectors_????-??-??.csv"))[-60:]:
    d = re.search(r"\d{4}-\d{2}-\d{2}", f).group(0)
    with open(f, encoding="utf-8-sig") as fh:
        rows += [(d, r["族群"], r["資金占比(%)"]) for r in csv.DictReader(fh)]
buf = io.StringIO(newline="")
w = csv.writer(buf); w.writerow(["日期", "族群", "資金占比(%)"]); w.writerows(rows)
new = "\ufeff" + buf.getvalue()
path, meta = "output/sectors_history.csv", "output/meta.json"
old = open(path, encoding="utf-8", newline="").read() if os.path.exists(path) else None
if new != old or not os.path.exists(meta):
    open(path, "w", encoding="utf-8", newline="").write(new)
    tw = datetime.timezone(datetime.timedelta(hours=8))
    json.dump({"generated_at": datetime.datetime.now(tw).strftime("%Y-%m-%d %H:%M")}, open(meta, "w"))
