#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""サンプル第5弾 介護・高齢者ダッシュボード: 生データ -> tidy CSV 変換
出典: 厚生労働省「介護保険事業状況報告」令和6年度年報 (e-Stat 政府統計コード 00450351)
      総務省統計局「人口推計」2024年10月1日現在 第11表 (政府統計コード 00200524)
すべて公開ダウンロードのみ。e-Stat API アプリID不使用。
"""
import csv
import os
import re

import openpyxl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
RAW = os.path.join(BASE, "data", "raw")
OUT = os.path.join(BASE, "data", "processed")
os.makedirs(OUT, exist_ok=True)

LEVELS = ["要支援１", "要支援２", "要介護１", "要介護２", "要介護３", "要介護４", "要介護５"]
LEVELS_N = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
SRC_LTCI = "厚生労働省 介護保険事業状況報告 令和6年度年報 (e-Stat 00450351)"


def norm(s):
    if isinstance(s, str):
        return s.replace("　", "").replace(" ", "").strip()
    return s


def find_header(ws, key="要支援１"):
    """レベル見出し行を探し、(行番号, レベル先頭列, 合計列) を返す"""
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [norm(c) for c in row]
        if key in cells:
            c0 = cells.index(key) + 1  # 1-based
            tot = None
            for j, c in enumerate(cells):
                if c in ("合計", "総数") and j + 1 > c0:
                    tot = j + 1
                    break
            return r, c0, tot
    raise ValueError(f"header not found in {ws.title}")


def read_level_table(path, name_cols):
    """都道府県別/保険者別の介護度別表を読む。name_cols=名前列の数(1 or 2)
    戻り値: list of dict {names:[...], levels:[...], total:...}"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    hr, c0, tot = find_header(ws)
    rows = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        names = [norm(row[i]) for i in range(name_cols)]
        if not names[0] or not isinstance(names[0], str):
            continue
        if names[0].startswith("※") or names[0].startswith("（"):
            continue
        levels = []
        ok = True
        for k in range(7):
            v = row[c0 - 1 + k]
            if v is None:
                v = 0
            if not isinstance(v, (int, float)):
                ok = False
                break
            levels.append(v)
        if not ok:
            continue
        total = row[tot - 1] if tot else sum(levels)
        rows.append({"names": names, "levels": levels, "total": total})
    wb.close()
    return rows


def read_insured(path, name_cols):
    """第1号被保険者数: 計, 65-74, 75-84, 85+"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    hr = None
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [norm(c) for c in row]
        if "都道府県" in cells and "計" in cells:
            hr = r
            c_total = cells.index("計") + 1
            break
    rows = []
    for row in ws.iter_rows(min_row=hr + 2, values_only=True):
        names = [norm(row[i]) for i in range(name_cols)]
        if not names[0] or not isinstance(names[0], str) or names[0].startswith("※"):
            continue
        vals = [row[c_total - 1 + k] for k in range(4)]
        if not isinstance(vals[0], (int, float)):
            continue
        rows.append({"names": names, "total": vals[0], "a65_74": vals[1], "a75_84": vals[2], "a85p": vals[3]})
    wb.close()
    return rows


def read_benefit_total(path, name_cols):
    """第8表 給付費: 総数ブロックの合計列だけを抜く (単位: 千円)"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    hr, c0, tot = find_header(ws)
    rows = {}
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        names = [norm(row[i]) for i in range(name_cols)]
        if not names[0] or not isinstance(names[0], str) or names[0].startswith("※"):
            continue
        v = row[tot - 1]
        if isinstance(v, (int, float)):
            key = tuple(names)
            if key not in rows:  # その1ブロックが総数
                rows[key] = v
    wb.close()
    return rows


def read_pop():
    """人口推計 第11表: 都道府県別 総人口・65歳以上 (千人)"""
    wb = openpyxl.load_workbook(os.path.join(RAW, "pop_est_pref_age3_2024.xlsx"), data_only=True, read_only=True)
    ws = wb.active
    out = {}
    # 列: 10=人口区分(総人口/日本人人口), 12=地域, 13=区分(人口/割合), 15=総数(男女計), 18=65歳以上(男女計) [千人]
    for row in ws.iter_rows(min_row=9, values_only=True):
        if len(row) < 18:
            continue
        if norm(row[9]) != "総人口" or norm(row[12]) != "人口":
            continue
        region = norm(row[11])
        total, o65 = row[14], row[17]
        if region and isinstance(total, (int, float)) and isinstance(o65, (int, float)):
            out[region] = {"total": total, "o65": o65}
    wb.close()
    return out


def zenkoku_totals(path):
    """全国計xlsxから 認定者総数(4-1表 最後の合計行) と 給付費総額(8-1表 エ合計行) を取る"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["4-1表"]
    cert = None
    for row in ws.iter_rows(values_only=True):
        vals = [c for c in row if c is not None]
        if vals and norm(vals[0]) == "合計" and isinstance(vals[-1], (int, float)):
            cert = vals[-1]
    benefit = None
    for sheet in ("8-1表", "8-1表 (2)"):
        if sheet not in wb.sheetnames:
            continue
        for row in wb[sheet].iter_rows(values_only=True):
            vals = [c for c in row if c is not None]
            if vals and norm(vals[0]) == "エ合計":
                nums = [v for v in vals if isinstance(v, (int, float))]
                benefit = nums[-1]  # 総数列
    wb.close()
    return cert, benefit


def main():
    # ---------- 都道府県レベル ----------
    insured = read_insured(os.path.join(RAW, "estat_r6_pref/p02_insured_cat1.xlsx"), 1)
    certified = read_level_table(os.path.join(RAW, "estat_r6_pref/p04_certified_total.xlsx"), 1)
    home = read_level_table(os.path.join(RAW, "estat_r6_pref/p05_home_recipients.xlsx"), 1)
    comm = read_level_table(os.path.join(RAW, "estat_r6_pref/p06_community_recipients.xlsx"), 1)
    faci = read_level_table(os.path.join(RAW, "estat_r6_pref/p07_facility_recipients.xlsx"), 1)
    benefit = read_benefit_total(os.path.join(RAW, "estat_r6_pref/p08_benefit_cost.xlsx"), 1)
    pop = read_pop()

    ins = {r["names"][0]: r for r in insured}
    cer = {r["names"][0]: r for r in certified}
    hom = {r["names"][0]: r["total"] for r in home}
    com = {r["names"][0]: r["total"] for r in comm}
    fac = {r["names"][0]: r["total"] for r in faci}
    ben = {k[0]: v for k, v in benefit.items()}

    # 公式認定率 (第1号被保険者ベース、令和6年度 概要 第3表、単位:%)
    official_rate = {
        "北海道": 21.4, "青森県": 18.2, "岩手県": 19.7, "宮城県": 19.4, "秋田県": 20.1, "山形県": 17.6,
        "福島県": 19.5, "茨城県": 16.5, "栃木県": 16.6, "群馬県": 18.2, "埼玉県": 17.9, "千葉県": 18.4,
        "東京都": 21.0, "神奈川県": 19.8, "新潟県": 18.9, "富山県": 20.1, "石川県": 18.7, "福井県": 18.0,
        "山梨県": 16.6, "長野県": 17.4, "岐阜県": 18.1, "静岡県": 17.4, "愛知県": 18.3, "三重県": 19.5,
        "滋賀県": 18.8, "京都府": 23.7, "大阪府": 24.3, "兵庫県": 21.6, "奈良県": 20.4, "和歌山県": 21.9,
        "鳥取県": 19.5, "島根県": 20.8, "岡山県": 21.6, "広島県": 20.1, "山口県": 19.3, "徳島県": 20.2,
        "香川県": 20.3, "愛媛県": 21.2, "高知県": 19.4, "福岡県": 19.8, "佐賀県": 18.2, "長崎県": 19.9,
        "熊本県": 19.6, "大分県": 19.1, "宮崎県": 16.3, "鹿児島県": 19.0, "沖縄県": 18.0,
    }
    prefs = [r["names"][0] for r in insured if r["names"][0] not in ("全国", "全国計")]
    rows = []
    for p in prefs:
        i, c = ins[p], cer[p]
        b_oku = ben[p] / 100000.0  # 千円 -> 億円
        pr = pop.get(p, {})
        rate = c["total"] / i["total"] * 100 if i["total"] else None
        rows.append({
            "都道府県": p,
            "第1号被保険者数": i["total"], "うち65_74歳": i["a65_74"], "うち75_84歳": i["a75_84"], "うち85歳以上": i["a85p"],
            **{f"認定者_{LEVELS_N[k]}": c["levels"][k] for k in range(7)},
            "認定者数_合計": c["total"],
            "認定率_pct_公式_第1号ベース": official_rate.get(p, ""),
            "認定者比率_pct_参考_総数ベース": round(rate, 2) if rate else "",
            "給付費_億円": round(b_oku, 1),
            "被保険者1人あたり給付費_円": round(ben[p] * 1000 / i["total"]) if i["total"] else "",
            "総人口_千人_人口推計": pr.get("total", ""),
            "65歳以上人口_千人_人口推計": pr.get("o65", ""),
            "高齢化率_pct_人口推計": round(pr["o65"] / pr["total"] * 100, 2) if pr.get("total") else "",
            "出典": SRC_LTCI + "; 総務省 人口推計2024年10月1日現在 第11表 (00200524)",
        })
    with open(os.path.join(OUT, "pref_insured_certified.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)

    rrows = []
    for p in prefs:
        h, cm, fc = hom.get(p, 0), com.get(p, 0), fac.get(p, 0)
        rrows.append({
            "都道府県": p,
            "居宅サービス受給者数_当年度累計_延人月": h,
            "地域密着型サービス受給者数_当年度累計_延人月": cm,
            "施設サービス受給者数_当年度累計_延人月": fc,
            "居宅_1ヶ月平均": round(h / 12), "地域密着_1ヶ月平均": round(cm / 12), "施設_1ヶ月平均": round(fc / 12),
            "注記": "サービス間の重複利用あり。1ヶ月平均=当年度累計(延人月)/12",
            "出典": SRC_LTCI,
        })
    with open(os.path.join(OUT, "pref_recipients.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rrows[0].keys()))
        w.writeheader(); w.writerows(rrows)

    # ---------- 保険者レベル ----------
    i_ins = read_insured(os.path.join(RAW, "estat_r6_insurer/i02_insured_cat1.xlsx"), 2)
    i_cer = read_level_table(os.path.join(RAW, "estat_r6_insurer/i04_certified_total.xlsx"), 2)
    i_hom = {tuple(r["names"]): r["total"] for r in read_level_table(os.path.join(RAW, "estat_r6_insurer/i05_home_recipients.xlsx"), 2)}
    i_com = {tuple(r["names"]): r["total"] for r in read_level_table(os.path.join(RAW, "estat_r6_insurer/i06_community_recipients.xlsx"), 2)}
    i_fac = {tuple(r["names"]): r["total"] for r in read_level_table(os.path.join(RAW, "estat_r6_insurer/i07_facility_recipients.xlsx"), 2)}
    i_ben = read_benefit_total(os.path.join(RAW, "estat_r6_insurer/i08_benefit_cost.xlsx"), 2)

    cer_map = {tuple(r["names"]): r for r in i_cer}
    ins_rows = []
    for r in i_ins:
        key = tuple(r["names"])
        if key[0] in ("全国", "全国計"):
            continue
        c = cer_map.get(key)
        b = i_ben.get(key)
        ins_rows.append({
            "都道府県": key[0], "保険者": key[1],
            "第1号被保険者数": r["total"],
            "認定者数_合計": c["total"] if c else "",
            **({f"認定者_{LEVELS_N[k]}": c["levels"][k] for k in range(7)} if c else {}),
            "認定者比率_pct_参考_総数ベース": round(c["total"] / r["total"] * 100, 2) if c and r["total"] else "",
            "給付費_億円": round(b / 100000.0, 2) if b else "",
            "被保険者1人あたり給付費_円": round(b * 1000 / r["total"]) if b and r["total"] else "",
            "居宅受給者数_累計_延人月": i_hom.get(key, ""), "地域密着受給者数_累計_延人月": i_com.get(key, ""),
            "施設受給者数_累計_延人月": i_fac.get(key, ""),
            "出典": SRC_LTCI,
        })
    with open(os.path.join(OUT, "insurer_detail.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(ins_rows[0].keys()))
        w.writeheader(); w.writerows(ins_rows)

    # ---------- 推移 (全国) ----------
    # 被保険者数 10年系列: 令和6年度 概要 第1表 (単位: 千人、四捨五入)
    years10 = ["2015年度", "2016年度", "2017年度", "2018年度", "2019年度", "2020年度", "2021年度", "2022年度", "2023年度", "2024年度"]
    ins10_total = [33816, 34405, 34879, 35252, 35548, 35788, 35887, 35845, 35890, 35839]
    ins10_65 = [17449, 17454, 17456, 17296, 17255, 17462, 17152, 16359, 15709, 15072]
    ins10_75 = [16366, 16951, 17423, 17955, 18292, 18326, 18735, 19486, 20181, 20767]
    trows = []
    for k, y in enumerate(years10):
        trows.append({"年度": y, "第1号被保険者数_千人_概要第1表": ins10_total[k],
                      "うち65_74歳_千人": ins10_65[k], "うち75歳以上_千人": ins10_75[k],
                      "認定者数_人_全国計": "", "給付費_介護給付予防給付_千円_全国計": "",
                      "出典": "厚生労働省 介護保険事業状況報告 令和6年度年報 概要 第1表"})
    exact = {"2021年度": "r03_zenkokukei.xlsx", "2022年度": "r04_zenkokukei.xlsx",
             "2023年度": "r05_zenkokukei.xlsx", "2024年度": "r06_zenkokukei.xlsx"}
    for y, fn in exact.items():
        cert, benf = zenkoku_totals(os.path.join(RAW, fn))
        for r in trows:
            if r["年度"] == y:
                r["認定者数_人_全国計"] = cert
                r["給付費_介護給付予防給付_千円_全国計"] = benf
                r["出典"] += "; 各年度 全国計 第4-1表・第8-1表"
    with open(os.path.join(OUT, "trend_national.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(trows[0].keys()))
        w.writeheader(); w.writerows(trows)

    # ---------- 検証 ----------
    nat_ins = sum(r["total"] for r in insured if r["names"][0] not in ("全国", "全国計"))
    nat_cer = sum(r["total"] for r in certified if r["names"][0] not in ("全国", "全国計"))
    nat_ben = sum(v for k, v in ben.items() if k not in ("全国", "全国計"))
    print(f" 参考: 給付費 全国計行 = {ben.get('全国計', 0):,.0f} 千円")
    cert_exact, ben_exact = zenkoku_totals(os.path.join(RAW, "r06_zenkokukei.xlsx"))
    print("検証:")
    print(f" 都道府県合計 被保険者数 = {nat_ins:,} (公表 35,839,278)")
    print(f" 都道府県合計 認定者数 = {nat_cer:,} (全国計xlsx 4-1表 = {cert_exact:,})")
    print(f" 都道府県合計 給付費(千円) = {nat_ben:,.0f} (全国計xlsx 8-1表 エ合計総数 = {ben_exact:,.0f})")
    print(f" 保険者ファイル 全国行 認定者数 = 7,207,487 / 被保険者数 = 35,839,278 と比較")
    print(f" 保険者レコード数 = {len(ins_rows)}")
    print(f" 人口推計 取得都道府県数 = {len(pop)}")


if __name__ == "__main__":
    main()
