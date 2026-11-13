#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""サンプル第5弾 介護・高齢者ダッシュボード Excel完成品を生成する
入力: data/processed/*.csv / 出力: output/caregiving_dashboard.xlsx
"""
import csv
import os

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
PROC = os.path.join(BASE, "data", "processed")
OUT = os.path.join(BASE, "output", "caregiving_dashboard.xlsx")

NAVY = "1F3864"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SUB_FONT = Font(bold=True, size=11, color=NAVY)
KPI_LABEL_FONT = Font(bold=True, size=11, color="404040")
KPI_VALUE_FONT = Font(bold=True, size=16, color=NAVY)
NOTE_FONT = Font(size=9, color="595959")
SRC = "厚生労働省「介護保険事業状況報告」令和6年度年報（e-Stat 00450351）"


def load(name):
    with open(os.path.join(PROC, name), encoding="utf-8") as f:
        return list(csv.DictReader(f))


def num(v, default=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def hdr_row(ws, row, headers, widths=None):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w


def data_sheet(wb, name, rows):
    ws = wb.create_sheet(name)
    if not rows:
        return ws
    headers = list(rows[0].keys())
    hdr_row(ws, 1, headers)
    for i, r in enumerate(rows, 2):
        for j, h in enumerate(headers, 1):
            v = r[h]
            try:
                v = float(v)
                if v == int(v):
                    v = int(v)
            except (TypeError, ValueError):
                pass
            ws.cell(row=i, column=j, value=v)
    ws.freeze_panes = "A2"
    return ws


def chart_bar(ws, title, data_ref, cat_ref, pos, w=16, h=8):
    ch = BarChart()
    ch.title = title
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cat_ref)
    ch.width, ch.height = w, h
    ch.legend = None
    ws.add_chart(ch, pos)
    return ch


def main():
    pref = load("pref_insured_certified.csv")
    rec = load("pref_recipients.csv")
    ins = load("insurer_detail.csv")
    trend = load("trend_national.csv")

    nat_insured = sum(int(num(r["第1号被保険者数"])) for r in pref)
    nat_cert = sum(int(num(r["認定者数_合計"])) for r in pref)
    nat_benefit_oku = sum(num(r["給付費_億円"]) for r in pref)
    lv_names = ["要支援1", "要支援2", "要介護1", "要介護2", "要介護3", "要介護4", "要介護5"]
    nat_levels = {lv: sum(int(num(r[f"認定者_{lv}"])) for r in pref) for lv in lv_names}
    rec_nat = {k: sum(int(num(r[k])) for r in rec) for k in
               ["居宅サービス受給者数_当年度累計_延人月", "地域密着型サービス受給者数_当年度累計_延人月", "施設サービス受給者数_当年度累計_延人月"]}
    pop_total = sum(num(r["総人口_千人_人口推計"]) for r in pref)
    pop_65 = sum(num(r["65歳以上人口_千人_人口推計"]) for r in pref)
    aging_nat = pop_65 / pop_total * 100

    wb = openpyxl.Workbook()

    # ---------- 表紙・使い方 ----------
    ws = wb.active
    ws.title = "表紙・使い方"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws["B2"] = "介護・高齢者ダッシュボード ― 保険者別に見る介護需給"
    ws["B2"].font = Font(bold=True, size=18, color=NAVY)
    lines = [
        ("", ""),
        ("対象データ", "介護保険事業状況報告 令和6年度年報（2025年3月末現在／令和6年度累計）。補助： 人口推計（2024年10月1日現在）"),
        ("粒度", "全国 ／ 47都道府県 ／ 保険者（市区町村等）1,573件"),
        ("シート構成", "D1_KPI概要 ／ D2_都道府県・保険者 ／ D3_需給・給付の内訳 ／ D4_推移 ／ データ4種 ／ 検証ログ ／ 出典・注記"),
        ("読み方", "認定率は公式値（第1号被保険者ベース、概要第3表）を採用。ダッシュボード内の「認定者比率（参考）」は認定者総数÷第1号被保険者数で、定義が異なるため併記にとどめる"),
        ("注意", "受給者数は当年度累計（延べ人月）。1ヶ月平均は累計÷12。サービス間に重複利用があるため単純合計は延べ人数"),
        ("", "給付費は「介護給付・予防給付」の給付費。高額介護サービス費・高額医療合算・特定入所者は含まない（公表の給付費総額12兆952億円とは範囲が異なる）"),
        ("更新", "年報は年度末現在の確定値。年次で差し替える設計（Power BI設計パック参照）"),
    ]
    r = 4
    for k, v in lines:
        ws.cell(row=r, column=2, value=k).font = SUB_FONT
        c = ws.cell(row=r, column=3, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 110

    # ---------- D1 KPI概要 ----------
    ws = wb.create_sheet("D1_KPI概要")
    ws["B2"] = "D1 KPI概要（令和6年度・全国）"
    ws["B2"].font = TITLE_FONT
    kpis = [
        ("第1号被保険者数（年度末）", f"{nat_insured:,} 人", "65歳以上の被保険者。前年度末 35,890千人から微減"),
        ("要介護（要支援）認定者数（年度末）", f"{nat_cert:,} 人", "前年度末 7,083,236人から +1.8%"),
        ("認定率（公式・第1号ベース）", "19.7 %", "概要第3表。大阪府24.3%が最高、宮崎県16.3%が最低"),
        ("給付費（介護給付・予防給付）", f"{nat_benefit_oku:,.0f} 億円", "≒10兆6,187億円。高額介護等を含む公表総額は12兆952億円"),
        ("被保険者1人あたり給付費", f"{nat_benefit_oku*1e8/nat_insured:,.0f} 円/年", "都道府県では大阪府が最大級（D2参照）"),
        ("サービス受給者数（1ヶ月平均）", f"居宅 {rec_nat['居宅サービス受給者数_当年度累計_延人月']/12:,.0f} 人", f"地域密着 {rec_nat['地域密着型サービス受給者数_当年度累計_延人月']/12:,.0f} 人 ／ 施設 {rec_nat['施設サービス受給者数_当年度累計_延人月']/12:,.0f} 人（重複利用あり）"),
        ("高齢化率（人口推計・参考）", f"{aging_nat:.1f} %", "総務省 人口推計 2024年10月1日現在（47都道府県合計ベース）"),
    ]
    r = 4
    for label, value, note in kpis:
        ws.cell(row=r, column=2, value=label).font = KPI_LABEL_FONT
        c = ws.cell(row=r, column=3, value=value)
        c.font = KPI_VALUE_FONT
        ws.cell(row=r, column=4, value=note).font = NOTE_FONT
        r += 1
    ws.cell(row=r + 1, column=2, value=f"出典： {SRC}").font = NOTE_FONT
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 70
    # helper: 介護度別認定者数
    hr = 4
    ws.cell(row=hr, column=8, value="介護度").font = HDR_FONT
    ws.cell(row=hr, column=8).fill = HDR_FILL
    ws.cell(row=hr, column=9, value="認定者数").font = HDR_FONT
    ws.cell(row=hr, column=9).fill = HDR_FILL
    for i, lv in enumerate(lv_names):
        ws.cell(row=hr + 1 + i, column=8, value=lv)
        ws.cell(row=hr + 1 + i, column=9, value=nat_levels[lv])
    chart_bar(ws, "認定者数（介護度別・全国）", Reference(ws, min_col=9, min_row=hr, max_row=hr + 7),
              Reference(ws, min_col=8, min_row=hr + 1, max_row=hr + 7), "B14")
    # helper: 推移
    tr = [t for t in trend if t["認定者数_人_全国計"]]
    ws.cell(row=hr, column=11, value="年度").font = HDR_FONT
    ws.cell(row=hr, column=11).fill = HDR_FILL
    ws.cell(row=hr, column=12, value="認定者数").font = HDR_FONT
    ws.cell(row=hr, column=12).fill = HDR_FILL
    for i, t in enumerate(tr):
        ws.cell(row=hr + 1 + i, column=11, value=t["年度"])
        ws.cell(row=hr + 1 + i, column=12, value=int(num(t["認定者数_人_全国計"])))
    ch = LineChart()
    ch.title = "認定者数の推移（全国）"
    ch.add_data(Reference(ws, min_col=12, min_row=hr, max_row=hr + len(tr)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=11, min_row=hr + 1, max_row=hr + len(tr)))
    ch.width, ch.height = 16, 8
    ch.legend = None
    ws.add_chart(ch, "B30")

    # ---------- D2 都道府県・保険者 ----------
    ws = wb.create_sheet("D2_都道府県・保険者")
    ws["B2"] = "D2 都道府県別（令和6年度）"
    ws["B2"].font = TITLE_FONT
    heads = ["都道府県", "第1号被保険者数", "認定者数", "認定率%(公式)", "給付費(億円)", "1人あたり給付費(円)", "高齢化率%(人口推計)"]
    hdr_row(ws, 4, [""] + heads)
    for i, row in enumerate(pref, 5):
        vals = [row["都道府県"], int(num(row["第1号被保険者数"])), int(num(row["認定者数_合計"])),
                num(row["認定率_pct_公式_第1号ベース"]), num(row["給付費_億円"]), int(num(row["被保険者1人あたり給付費_円"])),
                num(row["高齢化率_pct_人口推計"])]
        for j, v in enumerate(vals, 2):
            ws.cell(row=i, column=j, value=v)
    by_rate = sorted(pref, key=lambda x: -num(x["認定率_pct_公式_第1号ベース"]))[:10]
    hr = 4
    ws.cell(row=hr, column=10, value="都道府県").font = HDR_FONT; ws.cell(row=hr, column=10).fill = HDR_FILL
    ws.cell(row=hr, column=11, value="認定率%(公式)").font = HDR_FONT; ws.cell(row=hr, column=11).fill = HDR_FILL
    for i, row in enumerate(by_rate):
        ws.cell(row=hr + 1 + i, column=10, value=row["都道府県"])
        ws.cell(row=hr + 1 + i, column=11, value=num(row["認定率_pct_公式_第1号ベース"]))
    chart_bar(ws, "認定率（公式）TOP10", Reference(ws, min_col=11, min_row=hr, max_row=hr + 10),
              Reference(ws, min_col=10, min_row=hr + 1, max_row=hr + 10), "B54")
    by_ben = sorted(pref, key=lambda x: -num(x["給付費_億円"]))[:10]
    ws.cell(row=hr, column=13, value="都道府県").font = HDR_FONT; ws.cell(row=hr, column=13).fill = HDR_FILL
    ws.cell(row=hr, column=14, value="給付費(億円)").font = HDR_FONT; ws.cell(row=hr, column=14).fill = HDR_FILL
    for i, row in enumerate(by_ben):
        ws.cell(row=hr + 1 + i, column=13, value=row["都道府県"])
        ws.cell(row=hr + 1 + i, column=14, value=num(row["給付費_億円"]))
    chart_bar(ws, "給付費 TOP10（億円）", Reference(ws, min_col=14, min_row=hr, max_row=hr + 10),
              Reference(ws, min_col=13, min_row=hr + 1, max_row=hr + 10), "B70")
    # 保険者ハイライト
    r0 = 88
    ws.cell(row=r0, column=2, value="保険者ハイライト（認定者比率＝認定者総数÷第1号被保険者数、参考値）").font = SUB_FONT
    top_ins = sorted([x for x in ins if num(x["第1号被保険者数"]) >= 1000],
                     key=lambda x: -num(x["認定者比率_pct_参考_総数ベース"]))[:10]
    hdr_row(ws, r0 + 1, ["", "都道府県", "保険者", "被保険者数", "認定者数", "認定者比率%(参考)", "1人あたり給付費(円)"])
    for i, row in enumerate(top_ins, r0 + 2):
        vals = [row["都道府県"], row["保険者"], int(num(row["第1号被保険者数"])), int(num(row["認定者数_合計"])),
                num(row["認定者比率_pct_参考_総数ベース"]), int(num(row["被保険者1人あたり給付費_円"]))]
        for j, v in enumerate(vals, 2):
            ws.cell(row=i, column=j, value=v)
    ws.cell(row=r0 + 13, column=2, value="※被保険者数1,000人以上の保険者に限定（小規模保険者は比率が振れやすいため）。全1,573保険者は data_保険者 シートに収録").font = NOTE_FONT
    for col, w in zip("BCDEFGH", [14, 16, 14, 14, 14, 18, 16]):
        ws.column_dimensions[col].width = w

    # ---------- D3 需給・給付の内訳 ----------
    ws = wb.create_sheet("D3_需給・給付の内訳")
    ws["B2"] = "D3 需給・給付の内訳（令和6年度）"
    ws["B2"].font = TITLE_FONT
    ws.cell(row=4, column=2, value="サービス別 受給者数（全国・1ヶ月平均、延べ人月÷12）").font = SUB_FONT
    svc = [("居宅サービス", rec_nat["居宅サービス受給者数_当年度累計_延人月"] / 12),
           ("地域密着型サービス", rec_nat["地域密着型サービス受給者数_当年度累計_延人月"] / 12),
           ("施設サービス", rec_nat["施設サービス受給者数_当年度累計_延人月"] / 12)]
    ws.cell(row=5, column=8, value="サービス").font = HDR_FONT; ws.cell(row=5, column=8).fill = HDR_FILL
    ws.cell(row=5, column=9, value="受給者数(月平均)").font = HDR_FONT; ws.cell(row=5, column=9).fill = HDR_FILL
    for i, (k, v) in enumerate(svc):
        ws.cell(row=6 + i, column=8, value=k)
        ws.cell(row=6 + i, column=9, value=round(v))
    chart_bar(ws, "サービス別受給者数（月平均）", Reference(ws, min_col=9, min_row=5, max_row=8),
              Reference(ws, min_col=8, min_row=6, max_row=8), "B6")
    ws.cell(row=22, column=2, value="介護度別 認定者構成比（全国）").font = SUB_FONT
    tot_lv = sum(nat_levels.values())
    ws.cell(row=23, column=8, value="介護度").font = HDR_FONT; ws.cell(row=23, column=8).fill = HDR_FILL
    ws.cell(row=23, column=9, value="構成比%").font = HDR_FONT; ws.cell(row=23, column=9).fill = HDR_FILL
    for i, lv in enumerate(lv_names):
        ws.cell(row=24 + i, column=8, value=lv)
        ws.cell(row=24 + i, column=9, value=round(nat_levels[lv] / tot_lv * 100, 1))
    chart_bar(ws, "認定者 介護度別構成比（%）", Reference(ws, min_col=9, min_row=23, max_row=30),
              Reference(ws, min_col=8, min_row=24, max_row=30), "B24")
    ws.cell(row=40, column=2, value="被保険者1人あたり給付費（都道府県）").font = SUB_FONT
    top_pp = sorted(pref, key=lambda x: -num(x["被保険者1人あたり給付費_円"]))[:5]
    bot_pp = sorted(pref, key=lambda x: num(x["被保険者1人あたり給付費_円"]))[:5]
    hdr_row(ws, 41, ["", "上位5", "円/年", "", "下位5", "円/年"])
    for i in range(5):
        ws.cell(row=42 + i, column=2, value=top_pp[i]["都道府県"])
        ws.cell(row=42 + i, column=3, value=int(num(top_pp[i]["被保険者1人あたり給付費_円"])))
        ws.cell(row=42 + i, column=5, value=bot_pp[i]["都道府県"])
        ws.cell(row=42 + i, column=6, value=int(num(bot_pp[i]["被保険者1人あたり給付費_円"])))
    ws.cell(row=48, column=2, value="※1人あたり給付費の地域差は、認定率・介護度構成・サービス供給体制の差を反映する。単独では効率性の指標にならない").font = NOTE_FONT
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14

    # ---------- D4 推移 ----------
    ws = wb.create_sheet("D4_推移")
    ws["B2"] = "D4 推移（全国）"
    ws["B2"].font = TITLE_FONT
    ws.cell(row=4, column=2, value="第1号被保険者数（10年、概要第1表・千人）／ 認定者数・給付費（2021年度〜、全国計から exact 値）").font = NOTE_FONT
    hdr_row(ws, 5, ["", "年度", "被保険者数(千人)", "うち65-74(千人)", "うち75+(千人)", "認定者数(人)", "給付費(億円)"])
    for i, t in enumerate(trend, 6):
        ws.cell(row=i, column=2, value=t["年度"])
        ws.cell(row=i, column=3, value=int(num(t["第1号被保険者数_千人_概要第1表"])))
        ws.cell(row=i, column=4, value=int(num(t["うち65_74歳_千人"])))
        ws.cell(row=i, column=5, value=int(num(t["うち75歳以上_千人"])))
        if t["認定者数_人_全国計"]:
            ws.cell(row=i, column=6, value=int(num(t["認定者数_人_全国計"])))
            ws.cell(row=i, column=7, value=round(num(t["給付費_介護給付予防給付_千円_全国計"]) / 100000, 0))
    ch = LineChart()
    ch.title = "第1号被保険者数の推移（千人）"
    ch.add_data(Reference(ws, min_col=3, min_row=5, max_row=15), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=6, max_row=15))
    ch.width, ch.height = 16, 8
    ch.legend = None
    ws.add_chart(ch, "B18")
    ch2 = LineChart()
    ch2.title = "給付費（介護給付・予防給付）の推移（億円、2021年度〜）"
    ch2.add_data(Reference(ws, min_col=7, min_row=5, max_row=15), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=6, max_row=15))
    ch2.width, ch2.height = 16, 8
    ch2.legend = None
    ws.add_chart(ch2, "B34")
    for col, w in zip("BCDEFG", [12, 16, 16, 14, 16, 14]):
        ws.column_dimensions[col].width = w

    # ---------- データシート ----------
    data_sheet(wb, "data_都道府県", pref)
    data_sheet(wb, "data_受給者", rec)
    data_sheet(wb, "data_保険者", ins)
    data_sheet(wb, "data_推移", trend)

    # ---------- 検証ログ ----------
    ws = wb.create_sheet("検証ログ")
    ws["B2"] = "検証ログ（公表値との突合、2026-11-13 実施）"
    ws["B2"].font = TITLE_FONT
    checks = [
        ("第1号被保険者数（47都道府県合計）", "35,839,278 人", "35,839,278 人（保険者ファイル全国行／ポイント公表 3,584万人）", "一致"),
        ("認定者数（47都道府県合計）", "7,207,487 人", "7,207,487 人（全国計 第4-1表／保険者ファイル全国行／公表 721万人）", "一致"),
        ("給付費 介護給付・予防給付（47都道府県合計）", "10,618,731,913 千円", "10,618,731,913 千円（全国計 第8-1表 エ合計／概要 第1号分10兆4,487億円+第2号分1,700億円）", "一致"),
        ("認定率 全国（公式・第1号ベース）", "19.7 %", "19.7 %（概要 第3表）", "一致（本書のKPIは公式値を採用）"),
        ("認定率 都道府県（公式）", "大阪府 24.3% 等", "概要 第3表の47都道府県値と全件照合", "一致"),
        ("サービス受給者数（1ヶ月平均）", "居宅 4,319,340 人", "ポイント公表値（居宅 約432万人）と整合", "整合"),
        ("高齢化率 全国（人口推計）", f"{aging_nat:.2f} %（47県合計ベース）", "人口推計 第11表 全国行ベース 29.3%（2024年10月1日）", "整合（集計ベース差）"),
        ("推移 認定者数 2023年度", "7,083,236 人", "ポイント公表 708万人", "整合"),
    ]
    hdr_row(ws, 4, ["", "検証項目", "本ダッシュボードの値", "公表値・突合先", "結果"])
    for i, c in enumerate(checks, 5):
        for j, v in enumerate(c, 2):
            ws.cell(row=i, column=j, value=v)
    for col, w in zip("BCDE", [38, 30, 62, 26]):
        ws.column_dimensions[col].width = w

    # ---------- 出典・注記 ----------
    ws = wb.create_sheet("出典・注記")
    ws["B2"] = "出典・注記（採用理由／不採用理由を含む）"
    ws["B2"].font = TITLE_FONT
    notes = [
        ("主データ", "厚生労働省「介護保険事業状況報告」令和6年度年報（政府統計コード 00450351）。e-Statの公開ファイルと厚労省ページの全国計Excel・概要・ポイントを使用。取得日 2026-11-13、公開ダウンロードのみ（APIのアプリID不使用）"),
        ("採用理由： 介護保険事業状況報告", "認定者数・受給者数・給付費を都道府県別 AND 保険者（市区町村）別に同じ定義で公表する唯一の公式統計。介護需給を保険者粒度で見る本ダッシュボードの目的に直接対応する"),
        ("補助データ", "総務省統計局「人口推計」2024年10月1日現在 第11表（都道府県・年齢3区分別人口、政府統計コード 00200524）"),
        ("採用理由： 人口推計", "高齢化率（65歳以上人口割合）という需給の分母側の文脈を都道府県単位で揃えるため。第1号被保険者数と時点が近い（2024年10月 vs 2025年3月末）点も併記する"),
        ("定義： 認定者数", "年度末（2025年3月31日）現在の要介護・要支援認定者。第1号・第2号被保険者の合計。受給者台帳ベースで、提出後の遡及変更があり得る（原表注記）"),
        ("定義： 認定率", "KPIと都道府県順位は公式値「第1号被保険者に占める第1号認定者の割合」（概要第3表、全国19.7%）を使用。本書の「認定者比率（参考）」は認定者総数÷第1号被保険者数（全国20.1%）で、分子に第2号認定者を含むため公式値より約0.4pt高い"),
        ("定義： 給付費", "令和6年度累計（2024年3月〜2025年2月サービス分）の介護給付・予防給付の給付費。高額介護サービス費・高額医療合算介護サービス費・特定入所者介護サービス費を含まない。公表の「給付費総額12兆952億円」「給付費（利用者負担除く）11兆1,826億円」はそれらを含む広い範囲で、本書の10兆6,187億円とは範囲が異なる"),
        ("定義： 受給者数", "当年度累計は延べ人月。サービス（居宅・地域密着・施設）間に重複利用がある。1ヶ月平均は累計÷12の換算"),
        ("不採用： 介護保険事業状況報告（月報）", "月報は速報性があるが確報（年報）と値がずれ、保険者別の年間比較という本ダッシュボードの用途では年報で足りるため不採用。月次モニタリング商品では採用候補"),
        ("不採用： 介護サービス施設・事業所調査", "供給側（施設数・定員・従事者）の把握には有効だが、本サンプルは「需給の需要側＋給付」に範囲を絞ったため不採用。商品化時は供給レイヤーとして追加し、需給ギャップ分析に拡張する想定"),
        ("不採用： 地域包括ケア「見える化」システムの保険者別指標", "公開範囲・取得形式が安定せず再現性が担保できないため不採用。必要に応じ商品化時に個別対応"),
        ("未取得： 2020年度以前の認定者数・給付費の exact 系列", "被保険者数の10年系列は概要第1表（千人・四捨五入）で収録。認定者数・給付費の exact 値は2021年度分から全国計Excelで抽出（それ以前は原表レイアウトが異なり未抽出）。推測では埋めていない"),
        ("有料オプション方針", "本編から外したデータ（事業所調査の供給側データ等）でも買い手に価値があるものは、商品化時に「公表値ベースの参考資料」として有料オプションで提供する想定（取得・整形は商品化フェーズの作業。本サンプルには収録しない）"),
        ("免責", "本ダッシュボードは公開統計の二次加工サンプル。行政手続き・事業判断の際は原典を確認すること"),
    ]
    r = 4
    for k, v in notes:
        ws.cell(row=r, column=2, value=k).font = SUB_FONT
        c = ws.cell(row=r, column=3, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 115

    wb.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    main()
