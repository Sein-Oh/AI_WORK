# -*- coding: utf-8 -*-
"""
ml-kit 성능예측표 내보내기 (ml_export.py, 2.9.0)

Turns a trained version into ONE Excel workbook that predicts with plain formulas (no Python, no macros):
  1. input selection : the model's raw numeric inputs (derived features follow their source inputs) are reduced
                       by backward elimination with real retraining (the same 5-model competition), never by
                       importance order (interacting inputs look unimportant alone). --keep inputs are never dropped.
                       Recommended set = fewest inputs whose validation R2 is within RECOMMEND_TOL of the best.
  2. grid table      : the chosen model evaluated on an evenly spaced grid over the training range
                       (points per axis = min(41, MAX_ROWS ** (1/k))), with the distance of every grid node to the
                       nearest training row (scaled inputs) for the "no measured data" warning.
  3. workbook        : 계산기 (input cells -> prediction, 90% interval, status), 일괄계산 (many rows), 사용법,
                       표_<target> (the grid, veryHidden), 정확도 (candidates), 정보. Multilinear interpolation with INDEX/MATCH only
                       (Excel 2016+), linear extension outside the table. Every sheet is protected without a password:
                       only the yellow input cells can be edited. Only the grid sheets are hidden.
The version folder is not changed except for <version>/export/ (the retrained model and the candidate table).
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ml_predict import (  # noqa: E402  shared runtime lives in ml_predict.py
    setup_console, run_cli, JsonArgumentParser, InputError, resolve_column_names, known_columns, coerce_numeric, load_data_file,
    strip_unit, avoid_overwrite, save_with_lock_fallback, apply_inference_evals, prepare_features, quote_arg,
    VERSION, SNAPSHOT_EXCL_PREFIX, _fmt, model_raw_inputs, check_time, retry_fs,
)
import ml_train as mt  # noqa: E402  same training competition as ml_train.py

setup_console()

import itertools  # noqa: E402
import time  # noqa: E402
import json  # noqa: E402
import re  # noqa: E402
from datetime import datetime  # noqa: E402
from typing import Any, Dict, List, Optional, Tuple  # noqa: E402

import joblib  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

MAX_ROWS = 200_000        # grid rows per target (Excel limit is 1,048,576; keeps the file a few MB)
MAX_POINTS = 41           # grid points per axis
MIN_POINTS_WARN = 6       # fewer points per axis than this: interpolation may lose accuracy (warned)
RECOMMEND_TOL = 0.005     # recommended set: fewest inputs with validation R2 >= best - this
UNCOVERED_DIST = 0.15     # nearest training row farther than this (inputs scaled to 0-1): "실측 없는 영역"
UNCOVERED_GRID = 0.5      # + this x the grid cell diagonal (sqrt(k)/(n-1)): the distance is interpolated between grid
                          #   nodes, so a coarse grid (many inputs) reads far even on a training row
MAX_INPUTS = 8            # inputs per table: 2^k corner terms per formula (9+ inputs exceed Excel's columns / memory)
SHOW_COMMANDS = 8         # next_commands shown; every candidate also carries its own command
BATCH_ROWS = 500          # pre-filled rows on the 일괄계산 sheet
CHECK_POINTS = 2000       # random points for the interpolation accuracy check
TITLE_MAX = 40


def _sig(v: Any, digits: int = 4) -> Any:
    """Rounds to significant digits (a tiny target, e.g. 2e-6, must not round to 0 like round(v, 4) does)."""
    if v is None or not np.isfinite(v):
        return v
    return float(f"{float(v):.{digits}g}")


def value_format(values: np.ndarray) -> str:
    """Excel number format (also the TEXT() pattern) with about 4 significant digits of the target's magnitude:
    at least 2 decimals as before, more for small targets (0.005 -> 0.000000), scientific below 1e-5."""
    v = np.asarray(values, dtype=float)
    v = v[np.isfinite(v)]
    m = float(np.max(np.abs(v))) if v.size else 0.0
    if m == 0:
        return "0.00"
    dec = max(2, 3 - int(np.floor(np.log10(m))))
    return "0.000E+00" if dec > 8 else "0." + "0" * dec


def progress(msg: str) -> None:
    sys.stderr.write(f"[PROGRESS] {msg}\n")
    sys.stderr.flush()


# ==============================================================================
# 1. Version folder
# ==============================================================================

def load_version(vdir: str, targets_arg: Optional[str]) -> Dict[str, Any]:
    vdir = os.path.abspath(vdir)
    if not os.path.isdir(vdir):
        raise FileNotFoundError(f"버전 폴더가 없습니다: {vdir}")
    snap = os.path.join(vdir, "model_dataset.csv")
    if not os.path.isfile(snap):
        raise InputError(f"버전 폴더에 학습 데이터 사본(model_dataset.csv)이 없습니다: {vdir}. 다시 학습한 버전으로 실행하십시오.")
    metas = {}
    for name in sorted(os.listdir(vdir)):
        if name.startswith("model_") and name.endswith(".json"):
            with open(os.path.join(vdir, name), encoding="utf-8") as f:
                m = json.load(f)
            metas[m.get("target", name[6:-5])] = m
    if not metas:
        raise InputError(f"버전 폴더에 모델 정보(model_*.json)가 없습니다: {vdir}")
    # training order (leaderboard.json keeps the -y order); file-name order only as a fallback
    try:
        with open(os.path.join(vdir, "leaderboard.json"), encoding="utf-8") as f:
            order = list(json.load(f).get("targets", {}))
    except Exception:
        order = []
    names = [t for t in order if t in metas] + [t for t in metas if t not in order]
    if targets_arg:
        names, _ = resolve_column_names([t.strip() for t in targets_arg.split(",") if t.strip()], names, "타겟(-y)")
    skipped = [t for t in names if metas[t].get("task", "regression") != "regression"]
    names = [t for t in names if t not in skipped]
    if not names:
        raise InputError("성능예측표는 숫자 타겟(회귀)만 만들 수 있습니다. 이 버전에는 숫자 타겟이 없습니다.")
    return {"dir": vdir, "metas": metas, "targets": names, "skipped": skipped, "data": load_data_file(snap)}


def raw_inputs(meta: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    """(raw numeric inputs, derived feature expressions) of a model; text/date raw inputs are not supported."""
    schema = meta.get("feature_schema") or {}
    evals = list(meta.get("derived_evals") or [])
    raw = model_raw_inputs(meta)
    bad = [f for f in raw if f in schema.get("categorical", {}) or f in schema.get("datetime", {})]
    if bad or not schema:
        raise InputError(f"글자·날짜 입력 {bad}이(가) 있는 모델은 성능예측표로 만들 수 없습니다 (숫자 입력만 가능).")
    return raw, evals


def _refers(rhs: str, name: str) -> bool:
    """Whole-name match ('온도' is not found inside '반응온도'); expressions are stored with full column names."""
    return re.search(r"(?<!\w)" + re.escape(name) + r"(?!\w)", rhs) is not None


def evals_for(subset: List[str], evals: List[str], raw: List[str]) -> List[str]:
    """Derived features whose sources are all kept: raw inputs in the subset and earlier derived features kept
    (a derived feature built from another one is dropped with it)."""
    out, lhs_all = [], [e.split("=", 1)[0].strip() for e in evals]
    kept_lhs: List[str] = []
    for e, lhs in zip(evals, lhs_all):
        rhs = e.split("=", 1)[1]
        if all(r in subset for r in raw if _refers(rhs, r)) and all(d in kept_lhs for d in lhs_all if d != lhs and _refers(rhs, d)):
            out.append(e)
            kept_lhs.append(lhs)
    return out


# ==============================================================================
# 2. Training a subset (same competition as ml_train.py)
# ==============================================================================

# 2.9.40: every retrained candidate is saved in <version>/export/_fits (key = data, target, inputs, formulas, seed,
# versions). A command stopped at the time limit (check_time) or killed by the agent's 120 s limit continues where it
# stopped when it is run again, and a later export with other --keep reuses the same candidates.
FIT_CACHE: Optional[str] = None
_FIT_SECONDS: List[float] = []
_FIT_STATS = {"reused": 0, "trained": 0}


def _fit_key(data: pd.DataFrame, target: str, subset: List[str], kept: List[str], seed: int) -> str:
    import hashlib
    import sklearn
    h = hashlib.sha1()
    for part in (VERSION, sklearn.__version__, target, "\x1f".join(subset), "\x1f".join(kept), seed):
        h.update(str(part).encode("utf-8") + b"\x1e")
    h.update(pd.util.hash_pandas_object(data[list(subset) + [target]], index=False).values.tobytes())
    return h.hexdigest()[:32]


def _cached_fit(key: str) -> Optional[Dict[str, Any]]:
    if not FIT_CACHE:
        return None
    path = os.path.join(FIT_CACHE, key + ".joblib")
    if not os.path.isfile(path):
        return None
    try:
        res = joblib.load(path)
        if isinstance(res, dict) and "pipe" in res and "val_r2" in res:
            return res
    except Exception:
        pass
    try:                                      # damaged cache entry: drop it and train again (JIT: cache rebuilt)
        os.remove(path)
    except OSError:
        pass
    return None


def _save_fit(key: str, res: Dict[str, Any]) -> None:
    if not FIT_CACHE:
        return
    try:
        os.makedirs(FIT_CACHE, exist_ok=True)
        path = os.path.join(FIT_CACHE, key + ".joblib")
        tmp = path + ".tmp"
        joblib.dump(res, tmp)
        retry_fs(lambda: os.replace(tmp, path))
    except Exception as e:                    # no cache only means a slower next run
        sys.stderr.write(f"[WARN] 학습 결과를 저장하지 못했습니다 (다음 실행이 느려질 뿐): {type(e).__name__}: {e}\n")


def fit_subset(data: pd.DataFrame, target: str, subset: List[str], evals: List[str], raw: List[str],
               seed: int = 42) -> Dict[str, Any]:
    kept = evals_for(subset, evals, raw)
    key = _fit_key(data, target, subset, kept, seed)
    hit = _cached_fit(key)
    if hit is not None:
        _FIT_STATS["reused"] += 1
        return hit
    done = _FIT_STATS["reused"] + _FIT_STATS["trained"]
    if _FIT_STATS["trained"]:                 # at least one new candidate per run: running again always moves on
        check_time(max(_FIT_SECONDS, default=0.0) * 1.2,
                   f"성능예측표 입력 고르기: 후보 {done}개 학습 저장됨 ('{strip_unit(target)}' 입력 {len(subset)}개 시험 전에 멈춤)")
    t0 = time.monotonic()
    df = data[subset].apply(coerce_numeric)
    if kept:
        df = apply_inference_evals(df, kept)
    derived = [e.split("=", 1)[0].strip() for e in kept]
    cols = subset + derived
    schema = {"numeric": cols, "categorical": {}, "datetime": {}, "input_columns": cols}
    X = prepare_features(df, schema)
    quiet, mt.log_progress = mt.log_progress, (lambda p, m: None)
    try:
        _, pipe, info, _ = mt.train_and_benchmark_models(X=X, y_series=coerce_numeric(data[target]),
                                                         target_name=target, schema=schema, seed=seed)
    finally:
        mt.log_progress = quiet
    res = {"inputs": list(subset), "evals": kept, "schema": schema, "pipe": pipe,
           "val_r2": float(info["val_r2"]), "rmse": float(info["rmse"]), "model": info["champion_model"],
           "interval": info.get("prediction_interval")}
    _FIT_SECONDS.append(time.monotonic() - t0)
    _FIT_STATS["trained"] += 1
    _save_fit(key, res)
    return dict(res)


def search(data: pd.DataFrame, target: str, inputs: List[str], keep: List[str], evals: List[str],
           raw: List[str], stop_at: int = 1) -> List[Dict[str, Any]]:
    """Backward elimination with retraining: one candidate per input count, from all inputs down to the keep set /
    stop_at (1, or MAX_INPUTS when only a set that fits the workbook is needed)."""
    cur = list(inputs)
    progress(f"'{target}': 입력 {len(cur)}개 학습")
    best = fit_subset(data, target, cur, evals, raw)
    cands = [best]
    while len(cur) > max(stop_at, len(keep)):
        trials = []
        for drop in [c for c in cur if c not in keep]:
            sub = [c for c in cur if c != drop]
            progress(f"'{target}': 입력 {len(sub)}개 시험 ({strip_unit(drop)} 빼기)")
            trials.append((drop, fit_subset(data, target, sub, evals, raw)))
        if not trials:
            break
        drop, res = max(trials, key=lambda tr: tr[1]["val_r2"])
        res["dropped"] = drop
        cands.append(res)
        cur = res["inputs"]
    return cands


def recommend(cands: List[Dict[str, Any]]) -> Dict[str, Any]:
    best = max(c["val_r2"] for c in cands)
    ok = [c for c in cands if c["val_r2"] >= best - RECOMMEND_TOL]
    return min(ok, key=lambda c: (len(c["inputs"]), -c["val_r2"]))


# ==============================================================================
# 3. Grid table
# ==============================================================================

def build_grid(cand: Dict[str, Any], data: pd.DataFrame) -> Dict[str, Any]:
    inputs = cand["inputs"]
    k = len(inputs)
    n = int(max(2, min(MAX_POINTS, np.floor(MAX_ROWS ** (1.0 / k) + 1e-9))))
    lo = np.array([float(coerce_numeric(data[c]).min()) for c in inputs])
    hi = np.array([float(coerce_numeric(data[c]).max()) for c in inputs])
    axes = [np.linspace(lo[j], hi[j], n) for j in range(k)]
    G = pd.DataFrame(list(itertools.product(*axes)), columns=inputs)
    # the values written to the workbook (12 significant digits): interpolation checks and the agreement use them
    pred = np.array([float(f"{v:.12g}") for v in predict_frame(cand, G)])
    # distance of every node to the nearest training row (inputs scaled to 0-1)
    from scipy.spatial import cKDTree
    span = np.where(hi - lo > 0, hi - lo, 1.0)
    T = (np.column_stack([coerce_numeric(data[c]).to_numpy(dtype=float) for c in inputs]) - lo) / span
    T = T[np.isfinite(T).all(axis=1)]
    dist, _ = cKDTree(T).query((G.to_numpy(dtype=float) - lo) / span)
    # interpolation accuracy: table vs model at random points inside the range
    from scipy.interpolate import RegularGridInterpolator
    rng = np.random.default_rng(0)
    P = pd.DataFrame({c: rng.uniform(lo[j], hi[j], CHECK_POINTS) for j, c in enumerate(inputs)})
    interp = RegularGridInterpolator(axes, pred.reshape([n] * k))(P.to_numpy(dtype=float))
    model = predict_frame(cand, P)
    err = interp - model
    return {"n": n, "axes": axes, "grid": G, "pred": pred, "dist": dist, "lo": lo, "hi": hi,
            "interp_rmse": float(np.sqrt(np.mean(err ** 2))), "interp_max": float(np.max(np.abs(err))),
            "agreement": agreement(interp, model, cand.get("interval")),
            "uncovered_dist": UNCOVERED_DIST + UNCOVERED_GRID * float(np.sqrt(k)) / (n - 1), "fmt": value_format(pred)}


def agreement(table: np.ndarray, model: np.ndarray, interval: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """How closely the table follows the model over the operating range (random points), in words people use:
    'matches the model by 99.7% on average, within 0.9% in 95% of conditions'. Relative to the model value;
    relative to the output range instead when the output passes near zero (a % of a tiny value means nothing)."""
    d = np.abs(table - model)
    span = float(np.ptp(model)) or 1.0
    near_zero = bool(np.min(np.abs(model)) < 0.05 * span)
    rel = d / span * 100 if near_zero else d / np.abs(model) * 100
    half = (float(interval["residual_upper"]) - float(interval["residual_lower"])) / 2 if interval else None
    return {"agreement_pct": round(100 - float(rel.mean()), 2), "p95_diff_pct": round(float(np.quantile(rel, 0.95)), 2),
            "max_diff_pct": round(float(rel.max()), 2), "basis": "출력 범위 대비" if near_zero else "예측값 대비",
            "within_interval_pct": round(float(np.mean(d <= half)) * 100, 1) if half is not None else None,
            "points": int(len(d))}


def agreement_text(a: Dict[str, Any]) -> str:
    return (f"모델 예측과 평균 {a['agreement_pct']:.1f}% 일치, 95% 조건에서 차이 {a['p95_diff_pct']:.1f}% 이내"
            + (f" ({a['basis']})" if a["basis"] != "예측값 대비" else ""))


def predict_frame(cand: Dict[str, Any], df: pd.DataFrame) -> np.ndarray:
    d = df[cand["inputs"]].astype(float)
    if cand["evals"]:
        d = apply_inference_evals(d, cand["evals"])
    return np.asarray(cand["pipe"].predict(prepare_features(d, cand["schema"])), dtype=float)


def excel_interp(table: Dict[str, Any], x: np.ndarray, col: np.ndarray) -> float:
    """Python mirror of the workbook formulas (same MATCH clamp, same corner order and strides)."""
    axes, n = table["axes"], table["n"]
    k = len(axes)
    pos, t = [], []
    for j in range(k):
        p = int(np.clip(np.searchsorted(axes[j], x[j], side="right"), 1, n - 1))   # MATCH(...,1) clamped to [1, n-1]
        a0, a1 = axes[j][p - 1], axes[j][p]
        pos.append(p)
        t.append((x[j] - a0) / (a1 - a0))
    strides = [n ** (k - 1 - j) for j in range(k)]
    val = 0.0
    for bits in itertools.product((0, 1), repeat=k):
        w = float(np.prod([t[j] if b else 1 - t[j] for j, b in enumerate(bits)]))
        row = 1 + sum((pos[j] - 1 + b) * strides[j] for j, b in enumerate(bits))
        val += w * col[row - 1]
    return val


# ==============================================================================
# 4. Workbook
# ==============================================================================

def _col(i: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def _label(name: str) -> Tuple[str, str]:
    """'반응온도(°C)' -> ('반응온도', '°C')."""
    m = re.match(r"^(.*?)\s*[\(\[]([^\(\)\[\]]*)[\)\]]\s*$", str(name))
    return (m.group(1).strip(), m.group(2).strip()) if m else (str(name), "")


def _nice(v: float, span: float) -> float:
    """Rounds a value for display to ~3 significant digits of the input's range (170.0 ~ 195.0, 0.001 ~ 1.000)."""
    dec = int(max(0, 3 - np.floor(np.log10(span)))) if span > 0 else 3
    return round(float(v), min(dec, 6))


def _nice_txt(v: float, span: float) -> str:
    dec = int(max(0, 3 - np.floor(np.log10(span)))) if span > 0 else 3
    return f"{float(v):.{min(dec, 6)}f}"


def target_labels(targets: List[str]) -> Dict[str, str]:
    """Each target's short name for headers, sheet names and notes: the name without unit ('수율' for '수율(%)'),
    or the full name when another target has the same name without unit ('수율(1차)', '수율(최종)')."""
    count: Dict[str, int] = {}
    for t in targets:
        count[strip_unit(t)] = count.get(strip_unit(t), 0) + 1
    return {t: (str(t) if count[strip_unit(t)] > 1 else strip_unit(t)) for t in targets}


def _sheet_name(prefix: str, name: str, used: set) -> str:
    base = re.sub(r"[\\/\?\*\[\]:']", "_", f"{prefix}{name}")[:28]
    taken = {u.lower() for u in used}           # Excel compares sheet names without case
    out, i = base, 2
    while out.lower() in taken:
        out = f"{base[:25]}_{i}"
        i += 1
    used.add(out)
    return out


def default_title(meta: Dict[str, Any]) -> str:
    desc = (meta.get("description") or "").strip().splitlines()[0] if meta.get("description") else ""
    desc = re.split(r"(?<=[.!?。])\s", desc)[0] if desc else ""
    if not desc:
        return "성능 예측"
    return desc if len(desc) <= TITLE_MAX else desc[:TITLE_MAX - 3].rstrip() + "..."


def write_workbook(path: str, title: str, vinfo: Dict[str, Any], plans: List[Dict[str, Any]]) -> str:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.workbook.protection import WorkbookProtection

    FONT = "맑은 고딕"
    YELLOW = PatternFill("solid", fgColor="FFF2CC")
    HEAD = PatternFill("solid", fgColor="E7ECF5")
    # conditional formats: Excel reads bgColor for a solid fill here (fgColor alone shows no color)
    GOOD = {"fill": PatternFill("solid", fgColor="DCFCE7", bgColor="DCFCE7"), "font": Font(color="166534", bold=True)}
    WARN = {"fill": PatternFill("solid", fgColor="FCE4D6", bgColor="FCE4D6"), "font": Font(color="C2410C", bold=True)}
    NAVY = "1E3A8A"
    thin = Side(style="thin", color="C9D1E0")
    BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
    UNLOCK = Protection(locked=False)

    def style(c, size=10, bold=False, color="1F2937", fill=None, align="left", border=True, fmt=None):
        c.font = Font(name=FONT, size=size, bold=bold, color=color)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
        if fill is not None:
            c.fill = fill
        if border:
            c.border = BOX
        if fmt:
            c.number_format = fmt

    def status_colors(ws, cells, first):
        """Green for 정상, orange for 범위 밖 / 참고용. first = top-left cell of cells (relative reference)."""
        ws.conditional_formatting.add(cells, FormulaRule(formula=[f'{first}="정상"'], **GOOD))
        ws.conditional_formatting.add(cells, FormulaRule(formula=[f'OR({first}="범위 밖",LEFT({first},3)="참고용")'], **WARN))

    wb = Workbook()
    calc = wb.active
    calc.title = "계산기"
    batch = wb.create_sheet("일괄계산")
    howto = wb.create_sheet("사용법")
    used = {"계산기", "일괄계산", "사용법", "정확도", "정보"}
    short = target_labels([p["target"] for p in plans])   # '수율(1차)' / '수율(최종)': full names, never '수율' twice
    for p in plans:
        p["sheet"] = _sheet_name("표_", short[p["target"]], used)
        p["ws"] = wb.create_sheet(p["sheet"])
        p["ws"].sheet_state = "veryHidden"   # 2.9.3: the grid is not shown (not even in "unhide"); formulas still read it
    acc = wb.create_sheet("정확도")
    info = wb.create_sheet("정보")

    # union of inputs, in model order of the first target, then the others
    all_inputs: List[str] = []
    for p in plans:
        for c in p["cand"]["inputs"]:
            if c not in all_inputs:
                all_inputs.append(c)
    # training range per input = the span every table using it covers (rows left out of training excluded),
    # so "범위 밖" in the calculator means "outside the table", never a silent extrapolation
    rng = {}
    for c in all_inputs:
        users = [p for p in plans if c in p["cand"]["inputs"]]
        lo = max(float(p["table"]["lo"][p["cand"]["inputs"].index(c)]) for p in users)
        hi = min(float(p["table"]["hi"][p["cand"]["inputs"].index(c)]) for p in users)
        med = float(pd.concat([coerce_numeric(p["data"][c]) for p in users]).median())
        rng[c] = (lo, hi, min(max(med, lo), hi))

    # -- table sheets ------------------------------------------------------------------------
    for p in plans:
        ws, tab, cand = p["ws"], p["table"], p["cand"]
        k, n = len(cand["inputs"]), tab["n"]
        heads = [_label(c)[0] + (f" ({_label(c)[1]})" if _label(c)[1] else "") for c in cand["inputs"]]
        heads += [f"예측 {short[p['target']]}", "실측까지 거리"]
        for j, h in enumerate(heads, 1):
            style(ws.cell(row=1, column=j, value=h), bold=True, fill=HEAD, align="center")
        G = tab["grid"].to_numpy(dtype=float)
        for i in range(len(G)):
            r = i + 2
            for j in range(k):
                ws.cell(row=r, column=j + 1, value=float(G[i, j]))
            ws.cell(row=r, column=k + 1, value=float(tab["pred"][i])).number_format = tab["fmt"]
            ws.cell(row=r, column=k + 2, value=round(float(tab["dist"][i]), 4))
        p["pred_rng"] = f"'{p['sheet']}'!${_col(k + 1)}$2:${_col(k + 1)}${len(G) + 1}"
        p["dist_rng"] = f"'{p['sheet']}'!${_col(k + 2)}$2:${_col(k + 2)}${len(G) + 1}"
        ax0 = k + 4
        style(ws.cell(row=1, column=ax0 - 1, value="축 값"), bold=True, fill=HEAD)
        p["axis_rng"] = []
        for j in range(k):
            c = ax0 + j
            style(ws.cell(row=1, column=c, value=heads[j]), bold=True, fill=HEAD, align="center")
            for i, v in enumerate(tab["axes"][j]):
                ws.cell(row=i + 2, column=c, value=float(v))
            p["axis_rng"].append(f"'{p['sheet']}'!${_col(c)}$2:${_col(c)}${n + 1}")
            ws.column_dimensions[_col(c)].width = 14
        for j in range(1, k + 3):
            ws.column_dimensions[_col(j)].width = 15
        ws.freeze_panes = "A2"
        p["strides"] = [n ** (k - 1 - j) for j in range(k)]

    def interp_cells(ws, p, x_refs: Dict[str, str], row: int, col0: int) -> Tuple[str, str, int]:
        """Writes the helper cells of one interpolation (axis positions/fractions, corner terms) starting at col0.
        Returns (prediction formula, distance formula, next free column)."""
        cand = p["cand"]
        k, n = len(cand["inputs"]), p["table"]["n"]
        pos, frac = [], []
        c = col0
        for j, name in enumerate(cand["inputs"]):
            x, ax = x_refs[name], p["axis_rng"][j]
            ws.cell(row=row, column=c, value=f"=MIN(MAX(IFERROR(MATCH({x},{ax},1),1),1),{n - 1})")
            pos.append(f"{_col(c)}{row}")
            ws.cell(row=row, column=c + 1, value=f"=({x}-INDEX({ax},{pos[-1]}))/(INDEX({ax},{pos[-1]}+1)-INDEX({ax},{pos[-1]}))")
            frac.append(f"{_col(c + 1)}{row}")
            c += 2
        # corner terms in two contiguous blocks (prediction, distance) summed as ranges: 2^k terms would exceed
        # Excel's 255-argument limit as a list from 8 inputs on
        corners = list(itertools.product((0, 1), repeat=k))
        p0, d0 = c, c + len(corners)
        for i, bits in enumerate(corners):
            w = "*".join(frac[j] if b else f"(1-{frac[j]})" for j, b in enumerate(bits))
            idx = "1+" + "+".join(f"({pos[j]}-1{'+1' if b else ''})*{p['strides'][j]}" for j, b in enumerate(bits))
            ws.cell(row=row, column=p0 + i, value=f"={w}*INDEX({p['pred_rng']},{idx})")
            ws.cell(row=row, column=d0 + i, value=f"={w}*INDEX({p['dist_rng']},{idx})")
        end = d0 + len(corners)
        return (f"SUM({_col(p0)}{row}:{_col(d0 - 1)}{row})", f"SUM({_col(d0)}{row}:{_col(end - 1)}{row})", end)

    # -- 계산기 ------------------------------------------------------------------------------
    calc.sheet_view.showGridLines = False
    calc.column_dimensions["A"].width = 3
    for col, w in zip("BCDEF", (18, 14, 10, 22, 26)):
        calc.column_dimensions[col].width = w
    calc.merge_cells("B2:F2")
    style(calc["B2"], size=18, bold=True, color=NAVY, border=False)
    calc["B2"].value = title
    calc.merge_cells("B3:F3")
    style(calc["B3"], size=9, color="6B7280", border=False)
    calc["B3"].value = f"모델 {os.path.basename(vinfo['dir'])}  |  생성 {datetime.now():%Y-%m-%d}  |  ml-kit {VERSION}"
    calc.merge_cells("B4:F4")
    style(calc["B4"], size=10, color="374151", border=False)
    calc["B4"].value = "노란 칸에 값을 넣으면 아래에 예측값이 바로 계산됩니다."
    calc.row_dimensions[2].height = 30
    r = 5
    for j, h in enumerate(("입력", "값", "단위", "학습 범위", "상태"), 2):
        style(calc.cell(row=r, column=j, value=h), bold=True, fill=HEAD, align="center")
    x_refs: Dict[str, str] = {}
    dv = DataValidation(type="decimal", operator="between", formula1="-1E+307", formula2="1E+307", allow_blank=True,
                        showErrorMessage=True, errorTitle="숫자만 입력", error="숫자만 입력할 수 있습니다.")
    calc.add_data_validation(dv)
    first_input = None
    for i, c in enumerate(all_inputs):
        row = r + 1 + i
        lo, hi, med = rng[c]
        name, unit = _label(c)
        style(calc.cell(row=row, column=2, value=name), bold=True)
        v = calc.cell(row=row, column=3, value=_nice(med, hi - lo))
        style(v, size=11, bold=True, fill=YELLOW, align="right")
        v.protection = UNLOCK
        dv.add(v.coordinate)
        first_input = first_input or v.coordinate
        style(calc.cell(row=row, column=4, value=unit), align="center")
        style(calc.cell(row=row, column=5, value=f"{_nice_txt(lo, hi - lo)} ~ {_nice_txt(hi, hi - lo)}"), align="center")
        x = f"$C${row}"
        x_refs[c] = x
        chk = f"OR({x}<{lo!r},{x}>{hi!r})"
        style(calc.cell(row=row, column=6, value=f'=IF(ISNUMBER({x}),IF({chk},"범위 밖","정상"),"값을 넣으세요")'), align="center")
    in_end = r + len(all_inputs)
    status_colors(calc, f"F{r + 1}:F{in_end}", f"F{r + 1}")

    r2 = in_end + 2
    for j, h in enumerate(("결과", "예측값", "단위", "90% 범위", "판단"), 2):
        style(calc.cell(row=r2, column=j, value=h), bold=True, fill=HEAD, align="center")
    helper_col = 9
    style(calc.cell(row=5, column=helper_col, value="계산용 셀 (자동 계산, 수정 불가)"), size=9, color="6B7280", border=False)
    for ti, p in enumerate(plans):
        row = r2 + 1 + ti
        name, unit = _label(p["target"])
        name = name if short[p["target"]] == strip_unit(p["target"]) else short[p["target"]]
        own = [x_refs[c] for c in p["cand"]["inputs"]]
        pf, df_, nxt = interp_cells(calc, p, x_refs, row, helper_col)
        helper_col_end = nxt
        iv = p["cand"]["interval"] or {"residual_lower": 0.0, "residual_upper": 0.0}
        ready = f"COUNT({','.join(own)})={len(own)}"
        style(calc.cell(row=row, column=2, value=name), bold=True)
        fmt = p["table"]["fmt"]
        style(calc.cell(row=row, column=3, value=f'=IF({ready},{pf},"")'), size=14, bold=True, color=NAVY, align="right", fmt=fmt)
        style(calc.cell(row=row, column=4, value=unit), align="center")
        style(calc.cell(row=row, column=5, value=(f'=IF({ready},TEXT(C{row}+({iv["residual_lower"]!r}),"{fmt}")&" ~ "&'
                                                   f'TEXT(C{row}+({iv["residual_upper"]!r}),"{fmt}"),"")')), align="center")
        out_rng = "OR(" + ",".join(f"OR({x}<{rng[c][0]!r},{x}>{rng[c][1]!r})" for c, x in zip(p["cand"]["inputs"], own)) + ")"
        style(calc.cell(row=row, column=6, value=(f'=IF(NOT({ready}),"",IF({out_rng},"참고용 (학습 범위 밖)",'
                                                   f'IF({df_}>{p["table"]["uncovered_dist"]!r},"참고용 (실측 없는 영역)","정상")))')), align="center")
        status_colors(calc, f"F{row}", f"F{row}")
        for c in range(helper_col, helper_col_end):
            calc.column_dimensions[_col(c)].width = 9
            calc.cell(row=row, column=c).font = Font(name=FONT, size=8, color="9CA3AF")
    note_row = r2 + len(plans) + 2
    agree_lines = [f"- {short[p['target']]}: {agreement_text(p['table']['agreement'])}" for p in plans]
    calc.merge_cells(f"B{note_row}:F{note_row + 3 + len(plans)}")
    style(calc[f"B{note_row}"], size=9, color="374151", border=False)
    calc[f"B{note_row}"].alignment = Alignment(wrap_text=True, vertical="top")
    calc[f"B{note_row}"].value = ("90% 범위: 실제 값이 이 범위에 들 가능성이 약 90%입니다. "
                                  "학습 범위 밖이거나 실측 데이터가 없는 영역의 값은 참고용입니다. "
                                  "자세한 내용은 '사용법' 시트를 보세요.\n"
                                  "이 계산기와 원래 모델 예측의 일치도 (학습 범위 안 무작위 조건 2,000개):\n" + "\n".join(agree_lines))
    note_row += 1 + len(plans)
    calc.print_area = f"B2:F{note_row + 2}"
    calc.sheet_properties.pageSetUpPr.fitToPage = True
    calc.page_setup.fitToWidth = 1
    calc.page_setup.fitToHeight = 0

    # -- 일괄계산 ------------------------------------------------------------------------------
    batch.freeze_panes = "A3"
    batch.merge_cells("A1:F1")
    style(batch["A1"], size=11, bold=True, color=NAVY, border=False)
    batch["A1"].value = f"노란 열에 조건을 붙여 넣으면 결과가 계산됩니다 (최대 {BATCH_ROWS}건)."
    heads = [(_label(c)[0] + (f" ({_label(c)[1]})" if _label(c)[1] else "")) for c in all_inputs]
    col = 1
    for h in heads:
        style(batch.cell(row=2, column=col, value=h), bold=True, fill=YELLOW, align="center")
        batch.column_dimensions[_col(col)].width = 14
        col += 1
    res_cols = []
    for p in plans:
        tn = short[p["target"]]
        for h in (f"예측 {tn}", f"{tn} 하한(90%)", f"{tn} 상한(90%)", f"{tn} 판단"):
            style(batch.cell(row=2, column=col, value=h), bold=True, fill=HEAD, align="center")
            batch.column_dimensions[_col(col)].width = 16
            res_cols.append(col)
            col += 1
    helper0 = col + 1
    style(batch.cell(row=2, column=helper0, value="계산용 (자동)"), size=8, color="9CA3AF", border=False)
    bdv = DataValidation(type="decimal", operator="between", formula1="-1E+307", formula2="1E+307", allow_blank=True,
                         showErrorMessage=True, errorTitle="숫자만 입력", error="숫자만 입력할 수 있습니다.")
    batch.add_data_validation(bdv)
    bdv.add(f"A3:{_col(len(all_inputs))}{BATCH_ROWS + 2}")
    for i in range(BATCH_ROWS):
        row = i + 3
        refs = {c: f"${_col(j + 1)}{row}" for j, c in enumerate(all_inputs)}
        for j in range(len(all_inputs)):
            cell = batch.cell(row=row, column=j + 1)
            cell.fill = YELLOW
            cell.protection = UNLOCK
            cell.border = BOX
        hc = helper0 + 1
        rc = 0
        for p in plans:
            own = [refs[c] for c in p["cand"]["inputs"]]
            pf, df_, hc = interp_cells(batch, p, refs, row, hc)
            iv = p["cand"]["interval"] or {"residual_lower": 0.0, "residual_upper": 0.0}
            ready = f"COUNT({','.join(own)})={len(own)}"
            pc = res_cols[rc]
            fmt = p["table"]["fmt"]
            batch.cell(row=row, column=pc, value=f'=IF({ready},{pf},"")').number_format = fmt
            batch.cell(row=row, column=pc + 1, value=f'=IF({ready},{_col(pc)}{row}+({iv["residual_lower"]!r}),"")').number_format = fmt
            batch.cell(row=row, column=pc + 2, value=f'=IF({ready},{_col(pc)}{row}+({iv["residual_upper"]!r}),"")').number_format = fmt
            out_rng = "OR(" + ",".join(f"OR({x}<{rng[c][0]!r},{x}>{rng[c][1]!r})" for c, x in zip(p["cand"]["inputs"], own)) + ")"
            batch.cell(row=row, column=pc + 3, value=(f'=IF(NOT({ready}),"",IF({out_rng},"참고용 (학습 범위 밖)",'
                                                      f'IF({df_}>{p["table"]["uncovered_dist"]!r},"참고용 (실측 없는 영역)","정상")))'))
            rc += 4
    for c in res_cols[3::4]:
        status_colors(batch, f"{_col(c)}3:{_col(c)}{BATCH_ROWS + 2}", f"{_col(c)}3")

    # -- 사용법 ----------------------------------------------------------------------------------
    lines = [
        ("사용법", True),
        ("1. '계산기' 시트의 노란 칸에 조건을 넣으면 예측값과 90% 범위가 바로 계산됩니다.", False),
        ("2. 여러 조건은 '일괄계산' 시트의 노란 열에 붙여 넣으세요 (한 줄에 한 조건).", False),
        ("3. 노란 칸 말고는 수정할 수 없습니다 (계산이 깨지지 않도록 잠겨 있습니다).", False),
        ("", False),
        ("결과 읽는 법", True),
        ("- 예측값: 모델로 미리 계산해 파일 안에 넣어 둔 표에서, 입력값 주변 값을 비율로 섞어 구한 값입니다.", False),
        ("- 90% 범위: 실제 값이 이 범위 안에 들 가능성이 약 90%입니다. 좁을수록 정밀합니다.", False),
        ("- 판단 '참고용 (학습 범위 밖)': 입력이 모델이 배운 범위를 벗어났습니다. 값이 부정확할 수 있습니다.", False),
        ("- 판단 '참고용 (실측 없는 영역)': 각 입력은 범위 안이지만, 이런 조합의 실제 데이터가 근처에 없었습니다.", False),
        ("", False),
        ("이 파일에 대해", True),
        ("- 입력 항목과 모델 정보는 '정보' 시트, 입력 수에 따른 정확도는 '정확도' 시트에 있습니다.", False),
        ("- 모델이 예측한 값이므로 중요한 결정 전에는 실제 시험으로 확인하세요.", False),
    ]
    howto.column_dimensions["B"].width = 110
    howto.sheet_view.showGridLines = False
    for i, (t, bold) in enumerate(lines, 2):
        style(howto.cell(row=i, column=2, value=t), size=12 if bold else 10, bold=bold, color=NAVY if bold else "1F2937", border=False)

    # -- 정확도 / 정보 ---------------------------------------------------------------------------
    acc.sheet_view.showGridLines = False
    row = 2
    for p in plans:
        style(acc.cell(row=row, column=2, value=f"{p['target']}: 입력 수별 후보 (선택: 입력 {len(p['cand']['inputs'])}개)"),
              size=12, bold=True, color=NAVY, border=False)
        row += 1
        hdr = ("입력 수", "입력", "검증 R²", "RMSE", "모델", "표 행 수", "표 보간 오차(RMSE)", "예상 최종 오차", "선택")
        for j, h in enumerate(hdr, 2):
            style(acc.cell(row=row, column=j, value=h), bold=True, fill=HEAD, align="center")
        for c in p["cands"]:
            row += 1
            vals = (len(c["inputs"]), ", ".join(strip_unit(x) for x in c["inputs"]), round(c["val_r2"], 4), _sig(c["rmse"]),
                    c["model"], c.get("rows"), c.get("interp_rmse"), c.get("total_rmse"), "선택" if c is p["cand"] else "")
            for j, v in enumerate(vals, 2):
                style(acc.cell(row=row, column=j, value=v), align="left" if j == 3 else "right")
        a = p["table"]["agreement"]
        row += 2
        style(acc.cell(row=row, column=2, value="선택한 안의 모델 대비 일치도 (학습 범위 안 무작위 조건 2,000개, 성능예측표 값 vs 모델 예측값)"),
              bold=True, border=False)
        for label, v in (("평균 일치도", f"{a['agreement_pct']:.2f}%"), ("95% 조건의 차이", f"{a['p95_diff_pct']:.2f}% 이내"),
                         ("최대 차이", f"{a['max_diff_pct']:.2f}%"), ("기준", a["basis"]),
                         ("차이가 모델 90% 범위 안인 조건", "-" if a["within_interval_pct"] is None else f"{a['within_interval_pct']:.1f}%")):
            row += 1
            style(acc.cell(row=row, column=2, value=label), border=False)
            style(acc.cell(row=row, column=3, value=v), border=False)
        row += 3
    for col, w in zip("BCDEFGHIJ", (8, 46, 10, 10, 22, 11, 18, 15, 7)):
        acc.column_dimensions[col].width = w

    info.sheet_view.showGridLines = False
    rows = [("제목", title), ("만든 날짜", f"{datetime.now():%Y-%m-%d %H:%M}"), ("ml-kit", VERSION),
            ("원래 모델", short_version(vinfo["dir"]))]
    for p in plans:
        c, m = p["cand"], vinfo["metas"][p["target"]]
        dropped = [x for x in p["all_inputs"] if x not in c["inputs"]]
        rows += [("", ""), (f"[{p['target']}]", ""),
                 ("원래 모델", f"{m.get('champion_model')} / 입력 {len(p['all_inputs'])}개 / 검증 R² {m.get('val_r2')}"),
                 ("이 표의 모델", f"{c['model']} / 입력 {len(c['inputs'])}개 / 검증 R² {round(c['val_r2'], 4)} / RMSE {_sig(c['rmse'])}"),
                 ("입력", ", ".join(c["inputs"])),
                 ("뺀 입력 (학습 범위)", ", ".join(f"{x} ({_fmt(rng_of(p['data'], x)[0])}~{_fmt(rng_of(p['data'], x)[1])})" for x in dropped) or "없음"),
                 ("파생변수 (자동 계산)", ", ".join(c["evals"]) or "없음"),
                 ("표", f"축마다 {p['table']['n']}점, {len(p['table']['grid']):,}행, 보간 오차 RMSE {_sig(p['table']['interp_rmse'])}"),
                 ("실측 없는 영역 기준", f"가까운 학습 행까지 거리(입력 0~1 환산) {p['table']['uncovered_dist']:.3f} 초과 "
                                    f"({UNCOVERED_DIST} + 표 칸 대각선의 {UNCOVERED_GRID}배)"),
                 ("모델 대비 일치도", agreement_text(p["table"]["agreement"])),
                 ("90% 범위 (예측값 기준)", f"{(c['interval'] or {}).get('residual_lower')} ~ +{(c['interval'] or {}).get('residual_upper')}")]
    for i, (a, b) in enumerate(rows, 2):
        style(info.cell(row=i, column=2, value=a), bold=True, border=False)
        style(info.cell(row=i, column=3, value=b), border=False)
    info.column_dimensions["B"].width = 24
    info.column_dimensions["C"].width = 100

    # -- protection (no password): only the yellow cells are editable; structure fixed --------------
    for ws in wb.worksheets:
        ws.protection.sheet = True
        ws.protection.formatCells = True
        ws.protection.formatColumns = False   # widths can still be adjusted
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
    wb.security = WorkbookProtection(lockStructure=True)
    wb.active = 0
    calc.sheet_view.selection[0].activeCell = first_input
    calc.sheet_view.selection[0].sqref = first_input

    def _save(pth: str) -> None:
        wb.save(pth)
    return save_with_lock_fallback(path, _save, what="성능예측표")


def short_version(vdir: str) -> str:
    """'<data>_ml_data\\v3' only: the full path (user name, folders) is personal information and is not distributed."""
    return os.path.join(os.path.basename(os.path.dirname(vdir)), os.path.basename(vdir))


def rng_of(data: pd.DataFrame, col: str) -> Tuple[float, float]:
    s = coerce_numeric(data[col])
    return float(s.min()), float(s.max())


# ==============================================================================
# 5. Verification
# ==============================================================================

def verify_workbook(path: str, plans: List[Dict[str, Any]]) -> List[str]:
    """Re-reads the file: sheets, protection, unlocked inputs, table size; checks the formula math in Python;
    recalculates in Excel when it is installed (Windows)."""
    from openpyxl import load_workbook
    problems = []
    try:
        wb = load_workbook(path)
    except Exception as e:
        return [f"저장한 파일을 다시 열지 못함: {e}"]
    for name in ["계산기", "일괄계산", "사용법", "정확도", "정보"] + [p["sheet"] for p in plans]:
        if name not in wb.sheetnames:
            problems.append(f"'{name}' 시트 없음")
        elif not wb[name].protection.sheet:
            problems.append(f"'{name}' 시트 잠금 안 됨")
    if not (wb.security and wb.security.lockStructure):
        problems.append("통합문서 구조 잠금 안 됨")
    for p in plans:
        ws = wb[p["sheet"]] if p["sheet"] in wb.sheetnames else None
        if ws is not None and ws.max_row - 1 < len(p["table"]["grid"]):
            problems.append(f"'{p['sheet']}' 표 행 수 부족 ({ws.max_row - 1} < {len(p['table']['grid'])})")
        # formula math (same corners / strides as the workbook) against an independent interpolator
        from scipy.interpolate import RegularGridInterpolator
        tab = p["table"]
        rgi = RegularGridInterpolator(tab["axes"], tab["pred"].reshape([tab["n"]] * len(tab["axes"])))
        r = np.random.default_rng(1)
        for _ in range(20):
            x = np.array([r.uniform(a[0], a[-1]) for a in tab["axes"]])
            if abs(excel_interp(tab, x, tab["pred"]) - float(rgi(x)[0])) > 1e-6 * max(1.0, abs(float(rgi(x)[0]))):
                problems.append(f"'{p['target']}' 수식 계산 확인 실패")
                break
    calc = wb["계산기"] if "계산기" in wb.sheetnames else None
    if calc is not None:
        unlocked = [c.coordinate for row in calc.iter_rows() for c in row if c.protection is not None and not c.protection.locked]
        if not unlocked:
            problems.append("계산기에 입력 가능한 칸이 없음")
    return problems


def excel_check(path: str, plans: List[Dict[str, Any]]) -> Optional[str]:
    """Windows with Excel: open, recalculate, compare the calculator result with Python. None = checked OK,
    'skip' = Excel not available, otherwise a problem."""
    import importlib.util
    if importlib.util.find_spec("win32com") is None:
        return "skip"
    try:
        import win32com.client as w32
        xl = w32.DispatchEx("Excel.Application")
    except Exception as e:
        return f"엑셀로 계산 확인 실패: {type(e).__name__}: {e}"
    wb = None
    try:   # Excel is always closed, also when opening the file fails (no hidden EXCEL.EXE left behind)
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
        ws = wb.Worksheets("계산기")
        xl.CalculateFull()
        problems = []
        for p in plans:
            got = ws.Range(p["calc_cell"]).Value
            x = np.array([float(ws.Range(p["input_cells"][c]).Value) for c in p["cand"]["inputs"]])
            want = excel_interp(p["table"], x, p["table"]["pred"])
            if got is None or abs(float(got) - want) > 1e-6 * max(1.0, abs(want)):
                problems.append(f"'{p['target']}' 엑셀 계산 {got} != {want:.6g}")
        return "; ".join(problems) or None
    except Exception as e:
        return f"엑셀로 계산 확인 실패: {type(e).__name__}: {e}"
    finally:
        for close in ((lambda: wb.Close(SaveChanges=False)) if wb is not None else None, xl.Quit):
            try:
                if close:
                    close()
            except Exception:
                pass


# ==============================================================================
# 6. Command
# ==============================================================================

def _names_in(names: List[str], inputs: List[str], what: str) -> List[str]:
    """The given names that are inputs of this target (unit-less names allowed); names of other targets are skipped."""
    got = []
    for nm in names:
        try:
            got.append(resolve_column_names([nm], inputs, what)[0][0])
        except InputError:
            pass
    return list(dict.fromkeys(got))


def _check_known(names: List[str], all_inputs: Dict[str, List[str]], what: str) -> None:
    """A name no target has is an input error (with the [후보] of resolve_column_names; the other names of the same
    option are never listed: '압룍' next to '압력' in --keep "압력,압룍")."""
    union = list(dict.fromkeys(c for ins in all_inputs.values() for c in ins))
    for nm in names:
        if not any(_names_in([nm], ins, what) for ins in all_inputs.values()):
            # raises with a suggestion, never an input another given name already is ('압룍' next to '압력')
            resolve_column_names([nm], union, what, exclude=known_columns([n for n in names if n != nm], union))
            raise InputError(f"{what} '{nm}'은(는) 어느 타겟의 입력도 아닙니다. [후보] {', '.join(union[:10])}")


def parse_inputs_spec(text: Optional[str], targets: List[str], all_inputs: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """--inputs: 'a,b' (every target that has those inputs) or '<타겟>=a,b;<타겟2>=c' (per target).
    Returns {target: inputs}; a target not in the result is optimized automatically."""
    if not text:
        return {}
    what = "입력(--inputs)"
    if "=" in text:
        out: Dict[str, List[str]] = {}
        for part in [x for x in text.split(";") if x.strip()]:
            if "=" not in part:
                raise InputError(f"--inputs 타겟별 형식은 '<타겟>=입력1,입력2;<타겟2>=입력3'입니다: '{part.strip()}'")
            t, names = part.split("=", 1)
            tgt = resolve_column_names([t.strip()], targets, "--inputs의 타겟")[0][0]
            out[tgt] = resolve_column_names([s.strip() for s in names.split(",") if s.strip()], all_inputs[tgt], what)[0]
        return out
    names = [s.strip() for s in text.split(",") if s.strip()]
    _check_known(names, all_inputs, what)
    return {t: got for t in targets if (got := _names_in(names, all_inputs[t], what))}


SPEC_CHARS = ",;="     # separators of --inputs: a target or input name with one of them cannot be written there


def spec_safe(choice: Dict[str, List[str]], single: bool = False) -> bool:
    """True when every name written in --inputs can be written there: per target ('<타겟>=a,b;...') no ',', ';' or
    '=' in a target or input name; one target ('a,b', no target name) only the input names, without ',' or '='."""
    if single:
        return not any(ch in name for v in choice.values() for name in v for ch in ",=")
    return not any(ch in name for t, v in choice.items() for name in [t] + list(v) for ch in SPEC_CHARS)


def inputs_arg(choice: Dict[str, List[str]], single: bool) -> str:
    if single:
        return ",".join(next(iter(choice.values())))
    return ";".join(f"{t}={','.join(v)}" for t, v in choice.items())


def execute_export(args) -> Dict[str, Any]:
    global FIT_CACHE
    vinfo = load_version(args.model, args.targets)
    FIT_CACHE = os.path.join(vinfo["dir"], "export", "_fits")
    data_all = vinfo["data"]
    warnings: List[str] = []
    if vinfo["skipped"]:
        warnings.append(f"분류 타겟 {vinfo['skipped']}은(는) 성능예측표에서 뺐습니다 (숫자 타겟만 가능)")
    target_inputs = {t: raw_inputs(vinfo["metas"][t]) for t in vinfo["targets"]}
    in_lists = {t: v[0] for t, v in target_inputs.items()}
    keep_names = [s.strip() for s in (args.keep or "").split(",") if s.strip()]
    _check_known(keep_names, in_lists, "꼭 포함할 입력(--keep)")
    spec = parse_inputs_spec(args.inputs, vinfo["targets"], in_lists)
    if args.inputs:
        auto = [t for t in vinfo["targets"] if t not in spec]
        if auto:
            warnings.append(f"--inputs의 입력이 없는 타겟 {auto}은(는) 입력을 자동으로 골랐습니다")
    plans = []
    over: Dict[str, Dict[str, Any]] = {}     # --inputs over MAX_INPUTS: target -> best candidate with <= MAX_INPUTS
    over_n: Dict[str, int] = {}
    for target in vinfo["targets"]:
        inputs, evals = target_inputs[target]
        missing = [c for c in inputs + [target] if c not in data_all.columns]
        if missing:
            raise InputError(f"학습 데이터 사본에 {missing} 열이 없습니다. 다시 학습한 버전으로 실행하십시오.")
        data = data_all.copy()
        excl = f"{SNAPSHOT_EXCL_PREFIX}{target}"
        if excl in data.columns:
            data = data[coerce_numeric(data[excl]).fillna(0) != 1]
        data = data[coerce_numeric(data[target]).notna()].reset_index(drop=True)
        keep = _names_in(keep_names, inputs, "꼭 포함할 입력(--keep)")
        if len(keep) > MAX_INPUTS:
            raise InputError(f"'{target}': --keep 입력이 {len(keep)}개라 성능예측표 상한 {MAX_INPUTS}개를 넘습니다. "
                             f"꼭 넣을 입력을 {MAX_INPUTS}개 이하로 줄여 주십시오.")
        # an input with a single value carries no information and cannot be a table axis
        const = [c for c in inputs if coerce_numeric(data[c]).nunique(dropna=True) <= 1]
        if const:
            warnings.append(f"'{target}': 값이 하나뿐인 입력 {const}은(는) 성능예측표 입력에서 뺐습니다 (학습 데이터에서 변하지 않음)")
        usable = [c for c in inputs if c not in const]
        if not usable:
            raise InputError(f"'{target}': 학습 데이터에서 값이 변하는 입력이 없어 성능예측표를 만들 수 없습니다.")
        keep = [c for c in keep if c not in const]
        if target in spec:
            chosen_inputs = [c for c in spec[target] if c not in const]
            if not chosen_inputs:
                # the user chose these inputs: another set is the user's choice too ([후보], never an automatic retry
                # that would also drop the --inputs of the other targets)
                raise InputError(f"'{target}': --inputs의 입력 {spec[target]}이(가) 모두 학습 데이터에서 값이 하나뿐이라 표를 "
                                 f"만들 수 없습니다. 값이 변하는 입력: [후보] {', '.join(usable)}. 어느 입력으로 만들지 사용자에게 "
                                 f"물으십시오 (--inputs 없이 실행하면 입력을 자동으로 고릅니다).")
            missing_keep = [k for k in keep if k not in chosen_inputs]
            if missing_keep:
                raise InputError(f"--inputs에 --keep 입력 {missing_keep}이(가) 빠져 있습니다.")
            if len(chosen_inputs) > MAX_INPUTS:
                progress(f"'{target}': --inputs가 {len(chosen_inputs)}개라 {MAX_INPUTS}개 이하 추천 안을 찾는 중")
                cands = search(data, target, chosen_inputs, keep, evals, inputs, stop_at=MAX_INPUTS)
                over[target] = recommend([c for c in cands if len(c["inputs"]) <= MAX_INPUTS])
                over_n[target] = len(chosen_inputs)
                continue          # every target over the limit is collected first: one [추천] for all of them
            if over:
                continue          # an error follows anyway: skip the slow work
            progress(f"'{target}': 입력 {len(chosen_inputs)}개로 학습 (최적화 없음)")
            cands = [fit_subset(data, target, chosen_inputs, evals, inputs)]
            cand, reason = cands[0], "지정한 입력 (--inputs)"
        else:
            if over:
                continue
            cands = search(data, target, usable, keep, evals, inputs)
            eligible = [c for c in cands if len(c["inputs"]) <= MAX_INPUTS]
            cand = recommend(eligible)
            best = max(c["val_r2"] for c in eligible)
            reason = (f"최고 검증 R² {best:.4f}와 {RECOMMEND_TOL} 이내에서 입력 수가 가장 적은 안"
                      + (f" (꼭 포함: {', '.join(keep)})" if keep else "")
                      + (f" (입력 {MAX_INPUTS}개 이하 안 중에서)" if len(usable) > MAX_INPUTS else ""))
            if len(usable) > MAX_INPUTS:
                warnings.append(f"'{target}': 입력이 {len(usable)}개라 성능예측표 상한 {MAX_INPUTS}개 이하 안 중에서 골랐습니다")
        progress(f"'{target}': 후보별 표 정확도 계산")
        for c in cands:
            if len(c["inputs"]) > MAX_INPUTS:
                continue          # no table: the workbook cannot hold it (candidate listed with its score only)
            tab = build_grid(c, data)
            c["table"] = tab
            c["rows"] = len(tab["grid"])
            c["interp_rmse"] = _sig(tab["interp_rmse"])
            c["total_rmse"] = _sig(float(np.sqrt(c["rmse"] ** 2 + tab["interp_rmse"] ** 2)))
        if cand["table"]["n"] < MIN_POINTS_WARN:
            warnings.append(f"'{target}': 입력이 {len(cand['inputs'])}개라 표가 축마다 {cand['table']['n']}점뿐입니다. "
                            f"보간 오차 {cand['interp_rmse']} (모델 RMSE {_sig(cand['rmse'])})")
        plans.append({"target": target, "cand": cand, "cands": cands, "table": cand["table"], "all_inputs": inputs,
                      "reason": reason, "keep": keep, "data": data})

    progress(f"후보 학습: 새로 {_FIT_STATS['trained']}개, 저장된 것 재사용 {_FIT_STATS['reused']}개")
    if over:
        # the suggestion keeps the --inputs of every other target, so running it does not change them
        choice = {**spec, **{t: c["inputs"] for t, c in over.items()}}
        what = ", ".join(f"'{t}' {over_n[t]}개 -> {len(c['inputs'])}개 (검증 R² {c['val_r2']:.4f})" for t, c in over.items())
        head = (f"--inputs가 성능예측표 상한 {MAX_INPUTS}개를 넘습니다 (입력이 많으면 표와 수식이 엑셀 한도를 넘음). "
                f"입력을 줄여 다시 학습한 안: {what}.")
        if not spec_safe(choice, len(vinfo["targets"]) == 1):
            raise InputError(f"{head} 줄인 입력: {choice}. 타겟이나 입력 이름에 ',' ';' '='가 있어 --inputs로 쓸 수 없습니다. "
                             f"--inputs 없이 실행하면 입력을 자동으로 고릅니다 (--keep으로 꼭 넣을 입력 지정).")
        raise InputError(f"{head} [추천] {inputs_arg(choice, len(vinfo['targets']) == 1)}")

    # retrained models and the candidate table next to the version (never inside the version files)
    exp_dir = os.path.join(vinfo["dir"], "export")
    os.makedirs(exp_dir, exist_ok=True)
    for p in plans:
        c = p["cand"]
        safe = re.sub(r'[\\/*?:"<>|]', "_", p["target"])
        joblib.dump(c["pipe"], os.path.join(exp_dir, f"model_{safe}.joblib"), compress=3)
        with open(os.path.join(exp_dir, f"model_{safe}.json"), "w", encoding="utf-8") as f:
            json.dump({"target": p["target"], "task": "regression", "input_columns_required": c["inputs"],
                       "features_required": c["inputs"] + [e.split("=", 1)[0].strip() for e in c["evals"]],
                       "derived_evals": c["evals"], "feature_schema": c["schema"], "champion_model": c["model"],
                       "val_r2": round(c["val_r2"], 4), "rmse": _sig(c["rmse"]), "prediction_interval": c["interval"],
                       "source_version": short_version(vinfo["dir"]), "exported_at": datetime.now().isoformat(timespec="seconds"),
                       "ml_kit": VERSION}, f, ensure_ascii=False, indent=2)
    pd.DataFrame([{"타겟": p["target"], "입력 수": len(c["inputs"]), "입력": ", ".join(c["inputs"]), "검증 R2": round(c["val_r2"], 4),
                   "RMSE": _sig(c["rmse"]), "모델": c["model"], "표 행 수": c.get("rows"), "표 보간 오차": c.get("interp_rmse"),
                   "예상 최종 오차": c.get("total_rmse"), "선택": "선택" if c is p["cand"] else ""}
                  for p in plans for c in p["cands"]]).to_csv(os.path.join(exp_dir, "candidates.csv"), index=False, encoding="utf-8-sig")

    title = (args.title or "").strip() or default_title(vinfo["metas"][plans[0]["target"]])
    if len(title) > TITLE_MAX:
        warnings.append(f"제목이 {len(title)}자로 길어 {TITLE_MAX}자로 줄였습니다")
        title = title[:TITLE_MAX - 3].rstrip() + "..."
    # same folder as the report and history.jsonl (<data>_ml_data), name paired with ML결과보고서_<data>_<version>.docx
    data_root = os.path.dirname(vinfo["dir"])
    stem = re.sub(r"_ml_data$", "", os.path.basename(data_root)) or "모델"
    default = os.path.join(data_root, f"성능예측표_{stem}_{os.path.basename(vinfo['dir'])}.xlsx")
    path = args.save if args.save else default
    if args.save and not os.path.isabs(path):
        path = os.path.join(data_root, path)
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    path = avoid_overwrite(os.path.abspath(path), getattr(args, "overwrite", False), what="성능예측표")
    progress("엑셀 파일 작성")
    real = write_workbook(path, title, vinfo, plans)
    # remember where the calculator puts each result (for the Excel check)
    all_inputs = list(dict.fromkeys(c for p in plans for c in p["cand"]["inputs"]))
    for ti, p in enumerate(plans):
        p["input_cells"] = {c: f"C{6 + all_inputs.index(c)}" for c in all_inputs}
        p["calc_cell"] = f"C{5 + len(all_inputs) + 3 + ti}"
    progress("저장한 파일 확인")
    problems = verify_workbook(real, plans)
    xl = excel_check(real, plans)
    unverified = list(problems)
    if xl == "skip":
        unverified.append("엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함 (파일 구조, 잠금, 수식 계산 방식은 확인함)")
    elif xl:
        unverified.append(xl)

    # every candidate carries the command that rebuilds the workbook with it; the other targets keep their chosen inputs
    here = os.path.dirname(os.path.abspath(__file__))
    single = len(plans) == 1

    def rebuild_command(p: Dict[str, Any], x: Dict[str, Any]) -> Optional[str]:
        if len(x["inputs"]) > MAX_INPUTS:
            return None
        choice = {q["target"]: (x["inputs"] if q is p else q["cand"]["inputs"]) for q in plans}
        if not spec_safe(choice, single):
            return None   # a name with ',', ';' or '=' cannot be written in --inputs: no command rather than a broken one
        return (f"python {quote_arg(os.path.join(here, 'ml_export.py'))} -m {quote_arg(vinfo['dir'])}"
                + (f" -y {quote_arg(args.targets)}" if args.targets else "")
                + f" --inputs {quote_arg(inputs_arg(choice, single))} --title {quote_arg(title)}")

    out_targets = {}
    next_commands = []
    for p in plans:
        c = p["cand"]
        cands_out = []
        for x in p["cands"]:
            cmd = None if x is c else rebuild_command(p, x)
            cands_out.append({"count": len(x["inputs"]), "inputs": x["inputs"], "val_r2": round(x["val_r2"], 4), "rmse": _sig(x["rmse"]),
                              "table_rows": x.get("rows"), "interp_rmse": x.get("interp_rmse"), "total_rmse": x.get("total_rmse"),
                              **({"command": cmd} if cmd else {}),
                              **({"note": f"입력이 {MAX_INPUTS}개를 넘어 성능예측표로 만들 수 없음"} if len(x["inputs"]) > MAX_INPUTS else
                                 {"note": "이름에 ',' ';' '='가 있어 --inputs 명령을 만들 수 없음"} if x is not c and not cmd else {})})
            if cmd:
                next_commands.append({"purpose": f"'{p['target']}' 입력 {len(x['inputs'])}개 안으로 다시 만들기 (검증 R² {round(x['val_r2'], 4)}"
                                                 + (", 다른 타겟은 지금 입력 유지)" if not single else ")"),
                                      "command": cmd, "needs_input": False})
        out_targets[p["target"]] = {
            "inputs": c["inputs"], "dropped": [x for x in p["all_inputs"] if x not in c["inputs"]], "keep": p["keep"],
            "val_r2": round(c["val_r2"], 4), "rmse": _sig(c["rmse"]), "model": c["model"], "reason": p["reason"],
            "table_rows": c["rows"], "points_per_axis": c["table"]["n"], "interp_rmse": c["interp_rmse"], "total_rmse": c["total_rmse"],
            "uncovered_dist": round(c["table"]["uncovered_dist"], 4), "agreement": c["table"]["agreement"],
            "candidates": cands_out,
        }
    brief = "; ".join(f"{t}: 입력 {len(v['inputs'])}개 ({', '.join(strip_unit(i) for i in v['inputs'])}), 검증 R² {v['val_r2']}, "
                      f"표 {v['table_rows']:,}행, 모델과 평균 {v['agreement']['agreement_pct']:.1f}% 일치"
                      for t, v in out_targets.items())
    res = {"status": "ok", "output_file": real, "title": title, "targets": out_targets, "export_model_dir": exp_dir,
           "verified": not unverified, "preview": f"성능예측표 생성: {brief}", "next_commands": next_commands[:SHOW_COMMANDS],
           "next_commands_count": len(next_commands), "truncated": len(next_commands) > SHOW_COMMANDS}
    if warnings:
        res["warnings"] = warnings
    if unverified:
        res["unverified"] = unverified
    res["agent_directive"] = (
        ("VERIFIED: " if not unverified else "UNVERIFIED: " + "; ".join(unverified) + ". 재시도하지 말고 그대로 알리세요. ")
        + "output_file 경로, 타겟별 입력(inputs)과 뺀 입력(dropped), 검증 R², 모델 대비 일치도(agreement.agreement_pct, p95_diff_pct)를 알리세요. "
        + "다른 입력 조합을 원하면 targets.<타겟>.candidates를 보여 주고, 사용자가 고른 안의 command를 그대로 실행하세요"
        + (f" (next_commands는 {SHOW_COMMANDS}개만 담았습니다. 모든 안의 명령은 candidates[].command)." if len(next_commands) > SHOW_COMMANDS else ".")
        + (" warnings도 알리세요." if warnings else ""))
    return res


def main():
    setup_console()
    parser = JsonArgumentParser(description="ml-kit 성능예측표 (엑셀, 수식만, 잠금) 내보내기")
    parser.add_argument("--model", "-m", required=True, help="학습 결과 버전 폴더 (예: 공정_ml_data/v3)")
    parser.add_argument("--targets", "-y", help="표로 만들 숫자 타겟 (쉼표 구분, 생략 시 전부)")
    parser.add_argument("--keep", help="입력을 줄일 때 꼭 포함할 입력 (쉼표 구분)")
    parser.add_argument("--inputs", help=f"최적화 없이 이 입력 조합으로 바로 만들기 (쉼표 구분, 최대 {MAX_INPUTS}개). "
                                         "타겟마다 다르게: '<타겟>=입력1,입력2;<타겟2>=입력3'")
    parser.add_argument("--title", help="계산기 제목 (데이터 설명을 짧게 정리한 글, 40자 이내)")
    parser.add_argument("--save", help="저장 경로 (.xlsx, 생략 시 보고서와 같은 <데이터>_ml_data 폴더에 성능예측표_<데이터>_<버전>.xlsx)")
    parser.add_argument("--overwrite", action="store_true",
                        help="같은 이름의 성능예측표가 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    args = parser.parse_args()
    run_cli(lambda: execute_export(args), hint="Check the version folder and input names in SKILL.md.")


if __name__ == "__main__":
    main()
