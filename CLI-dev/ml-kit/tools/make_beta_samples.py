# -*- coding: utf-8 -*-
"""실기 테스트용 샘플 데이터 3종을 만든다. 사람이 실행하는 개발 도구 (run_beta_check.py도 이 모듈의 main을 부른다).

    python make_beta_samples.py [출력 폴더]      # 출력 폴더 생략 시 현재 폴더 아래 실기테스트_샘플 폴더
    (저장소의 docs/opencode/실기테스트_샘플/은 이 스크립트로 만든 사본)

- 맵.xlsx           : A 0~100을 5 단위, B 증가(비선형), C 감소, D 올랐다 내려감 (D=64이면 A=20, 80)
- 공정_다중출력.xlsx : 보고서형(제목 1행, 머리글 3행, 단위 4행), 입력 6개(분진은 무관) x 출력 3개, 300행 (사용자 학습 규모, 2026-11-10: 800행이던 것).
                      생산효율 = 반응온도, 촉매비율 x 수분함량 / 불순물 = 반응온도, 압축압력 / 소비전력 = 압축압력 x 유량.
                      이상치 2행(생산효율), '제외' 열: 101행째 데이터 '센서오류', 241행째 데이터(이상치) '포함'
- 새조건.xlsx        : 예측용 조건 3행 (3번째는 반응온도 200, 학습 범위 밖)
난수 시드가 고정되어 있어 같은 파일이 만들어진다.
"""
import os
import sys

import numpy as np
import openpyxl
import pandas as pd


def main(out: str) -> None:
    os.makedirs(out, exist_ok=True)
    a = np.arange(0, 101, 5)
    pd.DataFrame({"A": a, "B": a * a / 10, "C": 100 - a, "D": 100 - (a - 50) ** 2 / 25}).to_excel(os.path.join(out, "맵.xlsx"), index=False)

    r = np.random.default_rng(31)
    n = 300
    T, P, C = r.uniform(170, 195, n), r.uniform(2, 4, n), r.uniform(0.5, 1.5, n)
    W, F, D = r.uniform(10, 30, n), r.uniform(5, 15, n), r.uniform(0, 1, n)
    eff = 60 + 0.8 * (T - 170) - 0.06 * (T - 182) ** 2 + 12 * C * (1 - W / 40) + r.normal(0, 0.4, n)
    imp = 50 + 2.5 * (T - 170) + 40 * (P - 3) ** 2 + r.normal(0, 1.5, n)
    pw = 20 + 6 * P * F / 10 + 0.5 * F + r.normal(0, 0.3, n)
    eff[[15, 240]] += [25, -30]
    mark = [""] * n
    mark[100], mark[240] = "센서오류", "포함"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "반응 공정 시험 기록 (실기 테스트용 샘플)"
    ws.append([])
    ws.append(["반응온도", "압축압력", "촉매비율", "수분함량", "유량", "분진", "생산효율", "불순물", "소비전력", "제외"])
    ws.append(["°C", "bar", "wt%", "%", "L/min", "mg", "%", "ppm", "kW", ""])
    for i in range(n):
        ws.append([round(T[i], 2), round(P[i], 3), round(C[i], 3), round(W[i], 2), round(F[i], 2), round(D[i], 3),
                   round(eff[i], 3), round(imp[i], 2), round(pw[i], 3), mark[i] or None])
    wb.save(os.path.join(out, "공정_다중출력.xlsx"))

    pd.DataFrame({"반응온도": [180, 190, 200], "압축압력": [3.0, 2.5, 3.0], "촉매비율": [1.0, 0.7, 1.0],
                  "수분함량": [20, 15, 20], "유량": [10, 12, 10], "분진": [0.5, 0.5, 0.5]}).to_excel(os.path.join(out, "새조건.xlsx"), index=False)
    print(f"[INFO] 샘플 3종 저장: {out}", file=sys.stderr)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), "실기테스트_샘플"))
