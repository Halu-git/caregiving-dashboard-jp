#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""サンプル第5弾 データ取得の記録（再現用）
すべて公開ダウンロードのみ。e-Stat API のアプリIDは使用していない。取得日: 2026-11-13

厚生労働省「介護保険事業状況報告」令和6年度年報 (政府統計コード 00450351)
- 厚労省ページ https://www.mhlw.go.jp/topics/kaigo/osirase/jigyo/24/index.html
  - dl/r06_zenkokukei.xlsx  -> data/raw/r06_zenkokukei.xlsx
  - dl/r06_gaiyou.pdf       -> data/raw/r06_gaiyou.pdf （概要： 公式認定率の第3表を含む）
  - dl/r06_point.pdf        -> data/raw/r06_point.pdf （ポイント： 公表ヘッドライン値）
  - dl/r06_hihokensha.pdf   -> data/raw/r06_hihokensha.pdf
- 過去年度 全国計（推移用）: jigyo/23/dl/r05_zenkokukei.xlsx, /22/r04, /21/r03
- e-Stat ファイルダウンロード (stat-search/file-download?statInfId=<ID>&fileKind=0)
  都道府県別: 000040491595 第1号被保険者数 / 000040491597 認定者数(男女計・総数)
             000040491606 居宅受給者数 / 000040491609 地域密着受給者数 / 000040491612 施設受給者数
             000040491618 保険給付 給付費(総数)
  保険者別  : 000040491650 / 000040491652 / 000040491661 / 000040491664 / 000040491667 / 000040491673
             (同項目の保険者版) -> data/raw/estat_r6_pref/, data/raw/estat_r6_insurer/

総務省統計局「人口推計」2024年10月1日現在 (政府統計コード 00200524)
- 第11表 都道府県・年齢(3区分)・男女別人口: statInfId=000040268919 -> data/raw/pop_est_pref_age3_2024.xlsx

注記: 前任取得分 estat_r6/file-download は末尾欠損の破損ファイルだったため削除し、
      保険者別給付費 (000040491673) として再取得したものに置き換えた。
"""
print(__doc__)
