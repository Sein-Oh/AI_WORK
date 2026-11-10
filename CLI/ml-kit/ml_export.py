# -*- coding: utf-8 -*-
"""
ml-kit 성능예측표 내보내기 (ml_export.py)

Turns a trained version into ONE Excel workbook that computes the version's own model with plain formulas
(no Python, no macros, Excel 2016+), 2.9.84:
  1. model         : the version's model as it is (never retrained): every raw numeric input the user trained with,
                       derived features (--eval) as cell formulas, the same median for a formula error.
  2. trees         : RandomForest / ExtraTrees / HistGradientBoosting are moved leaf by leaf: every leaf is the box of
                       input values that reaches it (bounds in the user's units, the scaler and the float32 rounding
                       of scikit-learn folded into the bounds), and the prediction is one SUMPRODUCT over the boxes.
                       Ridge / ElasticNet: the linear formula. The same computation as the model, not an approximation.
  3. workbook      : 계산기 (input cells -> prediction, 90% interval, status), 일괄계산 (many rows), 사용법,
                       모델_<target> (leaves and training points, veryHidden), 정확도, 정보. Every sheet is protected
                       without a password: only the yellow input cells can be edited.
  4. check         : the formulas are mirrored in Python from the saved file and compared with the model on random
                       conditions and every training row; on Windows with Excel the calculator itself is recalculated.
The version folder is not changed.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ml_runtime import (  # noqa: E402  shared runtime of ml-kit
    check_core_libraries, exit_missing_library,
    setup_console, run_cli, JsonArgumentParser, InputError, resolve_column_names, coerce_numeric, load_data_file,
    strip_unit, avoid_overwrite, save_with_lock_fallback, apply_inference_evals, prepare_features, compile_formula,
    VERSION, SNAPSHOT_EXCL_PREFIX, model_raw_inputs, mark_committed, model_dir_from_data, version_hint, saved_as,
    command_text, quote_arg, data_folder_of, resolve_save,
)
check_core_libraries("성능예측표")
try:  # R5: a missing library answers with 1 JSON line (missing_library + install_command); the saved models need it
    __import__("sklearn")
except ImportError as _import_exc:
    exit_missing_library(_import_exc, "성능예측표")
setup_console()

import ast  # noqa: E402
import json  # noqa: E402
import re  # noqa: E402
from datetime import datetime  # noqa: E402
from typing import Any, Dict, List, Optional, Tuple  # noqa: E402

import joblib  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

MAX_LEAVES = 400_000      # leaves in one workbook (all targets): about 25 MB and under a minute to write and check
COVER_ROWS = 20_000       # training rows kept for the "no measured data nearby" check (random sample above this)
UNCOVERED_DIST = 0.15     # nearest training row farther than this (inputs scaled to 0-1): "실측 없는 영역", or farther
UNCOVERED_Q = 0.95        #   than this share of training rows are from their own nearest row (sparse data, many inputs)
BATCH_ROWS = 500          # pre-filled rows on the 일괄계산 sheet
CHECK_POINTS = 2000       # random conditions of the model check (plus every training row)
EXCEL_CHECKS = 50         # random conditions recalculated in Excel itself (Windows, one recalculation for all)
TITLE_MAX = 40
LOW_R2 = 0.70             # 검증 R²가 이보다 낮은 모델(보고서 판정 '부적합', ml_report CONDITIONAL_SCORE와 같은 값)은 계산기 판단이 '참고용'
FORMULA_MAX = 8192        # 엑셀 수식 한 칸의 글자 수 한도 (넘으면 엑셀이 파일을 복구 대상으로 열거나 수식을 지움)
FUNC_ARGS_MAX = 255       # 엑셀 함수 인자 수 한도 (COUNT, OR에 입력이 하나씩 들어감)
NAMES_SHOWN = 8           # preview와 보고서에 적는 입력 이름 수 (나머지는 '외 N개')
OPEN = 1e300              # bound of a leaf box that has no limit on a feature
TREE_MODELS = ("RandomForestRegressor", "ExtraTreesRegressor")
LINEAR_MODELS = ("RidgeCV", "ElasticNetCV", "Ridge", "ElasticNet", "LinearRegression")


def _sig(v: Any, digits: int = 4) -> Any:
    """Rounds to significant digits (a tiny target, e.g. 2e-6, must not round to 0 like round(v, 4) does)."""
    if v is None or not np.isfinite(v):
        return v
    return float(f"{float(v):.{digits}g}")


def value_format(values: np.ndarray) -> str:
    """Excel number format (also the TEXT() pattern) with about 4 significant digits of the target's magnitude:
    at least 2 decimals, more for small targets (0.005 -> 0.000000), scientific below 1e-5."""
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


def num(v: float) -> str:
    """A number inside a formula: every digit of the double (repr), negative numbers in parentheses."""
    s = repr(float(v)).replace("e", "E")
    return f"({s})" if s.startswith("-") else s


# ==============================================================================
# 1. Version folder
# ==============================================================================

def load_version(vdir: str, targets_arg: Optional[str]) -> Dict[str, Any]:
    vdir = os.path.abspath(vdir)
    if not os.path.isdir(vdir):
        raise FileNotFoundError(f"버전 폴더가 없습니다: {vdir}." + version_hint(vdir))
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
    # training order (leaderboard.json keeps the --targets order); file-name order only as a fallback
    try:
        with open(os.path.join(vdir, "leaderboard.json"), encoding="utf-8") as f:
            order = list(json.load(f).get("targets", {}))
    except Exception:
        order = []
    names = [t for t in order if t in metas] + [t for t in metas if t not in order]
    if targets_arg:
        names, _ = resolve_column_names([t.strip() for t in targets_arg.split(",") if t.strip()], names, "타겟(--targets)")
    skipped = [t for t in names if metas[t].get("task", "regression") != "regression"]
    names = [t for t in names if t not in skipped]
    if not names:
        raise InputError("성능예측표는 숫자 타겟(회귀)만 만들 수 있습니다. 이 버전에는 숫자 타겟이 없습니다.")
    return {"dir": vdir, "metas": metas, "targets": names, "skipped": skipped, "data": load_data_file(snap)}


def text_date_inputs(meta: Dict[str, Any]) -> List[str]:
    """Raw inputs of a model that are text (categories) or dates: the calculator takes numbers only."""
    schema = meta.get("feature_schema") or {}
    return [f for f in model_raw_inputs(meta) if f in schema.get("categorical", {}) or f in schema.get("datetime", {})]


def raw_inputs(meta: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    """(raw numeric inputs, derived feature expressions) of a model; text/date raw inputs are not supported."""
    schema = meta.get("feature_schema") or {}
    evals = list(meta.get("derived_evals") or [])
    raw = model_raw_inputs(meta)
    bad = text_date_inputs(meta)
    if bad or not schema:
        raise InputError(f"글자·날짜 입력 {bad}이(가) 있는 모델은 성능예측표로 만들 수 없습니다 (엑셀 계산기는 숫자 입력만 가능). "
                         f"사용자에게 알리십시오.")
    return raw, evals


def load_pipe(vdir: str, meta: Dict[str, Any]):
    safe = meta.get("target_safe_name") or re.sub(r'[\\/*?:"<>|]', "_", str(meta.get("target")))
    pipe = joblib.load(os.path.join(vdir, f"model_{safe}.joblib"))
    return pipe["pipeline"] if isinstance(pipe, dict) and "pipeline" in pipe else pipe


# ==============================================================================
# 2. The model as formulas
# ==============================================================================

def _unsupported(what: str) -> InputError:
    return InputError(f"이 모델은 성능예측표(엑셀 계산기)로 옮길 수 없습니다: {what}. 사용자에게 알리십시오.")


def scaler_parts(pipe) -> Tuple[List[str], np.ndarray, np.ndarray, np.ndarray, Any]:
    """(model feature names, imputer medians, scaler mean, scaler scale, regressor) of the ml-kit pipeline
    ColumnTransformer[num: SimpleImputer(median) -> StandardScaler] -> regressor."""
    steps = getattr(pipe, "steps", None)
    if not steps or len(steps) != 2:
        raise _unsupported("전처리와 모델 두 단계가 아님")
    ct, reg = steps[0][1], steps[1][1]
    parts = [(n, t, c) for n, t, c in getattr(ct, "transformers_", []) if not (isinstance(t, str) and t == "drop")]
    if len(parts) != 1 or parts[0][0] != "num":
        raise _unsupported("숫자 입력 전처리 하나가 아님")
    sub = parts[0][1]
    named = getattr(sub, "named_steps", {})
    imp, sc = named.get("imputer"), named.get("scaler")
    if imp is None or sc is None or len(sub.steps) != 2 or getattr(imp, "strategy", "") != "median":
        raise _unsupported("빈 값 채우기·표준화 전처리가 아님")
    feats = [str(c) for c in parts[0][2]]
    mean = np.zeros(len(feats)) if getattr(sc, "mean_", None) is None else np.asarray(sc.mean_, dtype=float)
    scale = np.ones(len(feats)) if getattr(sc, "scale_", None) is None else np.asarray(sc.scale_, dtype=float)
    return feats, np.asarray(imp.statistics_, dtype=float), mean, scale, reg


def raw_bounds(feat: np.ndarray, thr: np.ndarray, mean: np.ndarray, scale: np.ndarray, f32: bool) -> np.ndarray:
    """Largest double r per split with  cast((r - mean) / scale) <= thr  (cast = float32 for RandomForest /
    ExtraTrees, which compare float32 inputs; none for HistGradientBoosting): the split in the user's units, so
    Excel compares the input itself and the comparison is the model's own, bit for bit (x <= r for every double x)."""
    mu, sg, t = mean[feat], scale[feat], thr.astype(float)

    def left(r: np.ndarray) -> np.ndarray:
        z = (r - mu) / sg
        return (z.astype(np.float32) <= t) if f32 else (z <= t)

    c = mu + t * sg
    w = (np.abs(t) + 1.0) * np.abs(sg) * 1e-6 + 1e-300
    lo, hi = c - w, c + w
    for _ in range(80):                      # bracket: left(lo) and not left(hi)
        bad = ~left(lo) | left(hi)
        if not bad.any():
            break
        w = np.where(bad, w * 4, w)
        lo, hi = np.where(bad, c - w, lo), np.where(bad, c + w, hi)
    else:
        raise _unsupported("갈림 기준값을 입력 단위로 바꾸지 못함")
    for _ in range(200):                     # bisection down to neighbouring doubles
        m = lo + (hi - lo) / 2
        done = (m <= lo) | (m >= hi)
        if done.all():
            break
        ok = left(m)
        lo = np.where(~done & ok, m, lo)
        hi = np.where(~done & ~ok, m, hi)
    return lo


def _boxes(left: np.ndarray, right: np.ndarray, feat: np.ndarray, bound: np.ndarray, value: np.ndarray,
           leaf: np.ndarray, k: int, lows: List[np.ndarray], highs: List[np.ndarray], vals: List[float], scale: float):
    """Walks one tree: every leaf gives the box (low < x <= high per feature) that reaches it."""
    stack = [(0, np.full(k, -OPEN), np.full(k, OPEN))]
    while stack:
        n, lo, hi = stack.pop()
        if leaf[n]:
            lows.append(lo)
            highs.append(hi)
            vals.append(float(value[n]) * scale)
            continue
        f, b = int(feat[n]), float(bound[n])
        h2 = hi.copy()
        h2[f] = min(hi[f], b)
        stack.append((int(left[n]), lo, h2))
        l2 = lo.copy()
        l2[f] = max(lo[f], b)
        stack.append((int(right[n]), l2, hi))


def model_formula(pipe) -> Dict[str, Any]:
    """The model as data the workbook computes: {"kind": "trees", leaves (lo, hi, value), base} or
    {"kind": "linear", coef, intercept}, with the features, imputer medians and scaler."""
    feats, med, mean, scale, reg = scaler_parts(pipe)
    k = len(feats)
    name = type(reg).__name__
    out: Dict[str, Any] = {"features": feats, "median": med, "mean": mean, "scale": scale, "model_class": name}
    if name in LINEAR_MODELS:
        coef = np.asarray(reg.coef_, dtype=float).reshape(-1)
        if coef.size != k:
            raise _unsupported("선형 모델 계수 수가 입력 수와 다름")
        out.update(kind="linear", coef=coef, intercept=float(np.asarray(reg.intercept_, dtype=float).reshape(-1)[0]))
        return out
    lows: List[np.ndarray] = []
    highs: List[np.ndarray] = []
    vals: List[float] = []
    if name in TREE_MODELS:
        trees = reg.estimators_
        for est in trees:
            tr = est.tree_
            if tr.n_outputs != 1 or tr.value.shape[2] != 1:
                raise _unsupported("타겟이 여럿인 나무")
            leaf = tr.children_left == -1
            bound = np.zeros(tr.node_count)
            inner = ~leaf
            bound[inner] = raw_bounds(tr.feature[inner], tr.threshold[inner], mean, scale, f32=True)
            _boxes(tr.children_left, tr.children_right, tr.feature, bound, tr.value[:, 0, 0], leaf, k,
                   lows, highs, vals, 1.0 / len(trees))
        base = 0.0
    elif name == "HistGradientBoostingRegressor":
        if getattr(reg, "loss", "squared_error") != "squared_error" or getattr(reg, "_preprocessor", None) is not None:
            raise _unsupported("제곱 오차가 아니거나 범주 입력이 있는 HistGradientBoosting")
        for preds in reg._predictors:
            if len(preds) != 1:
                raise _unsupported("타겟이 여럿인 HistGradientBoosting")
            nodes = preds[0].nodes
            leaf = nodes["is_leaf"].astype(bool)
            if np.any(nodes["is_categorical"][~leaf]):
                raise _unsupported("범주 갈림이 있는 HistGradientBoosting")
            bound = np.zeros(len(nodes))
            inner = ~leaf
            bound[inner] = raw_bounds(nodes["feature_idx"][inner].astype(int), nodes["num_threshold"][inner],
                                      mean, scale, f32=False)
            _boxes(nodes["left"], nodes["right"], nodes["feature_idx"], bound, nodes["value"], leaf, k,
                   lows, highs, vals, 1.0)
        base = float(np.asarray(reg._baseline_prediction, dtype=float).reshape(-1)[0])
    else:
        raise _unsupported(f"모델 종류 {name}")
    lo, hi = np.array(lows), np.array(highs)
    used = [j for j in range(k) if np.any(lo[:, j] > -OPEN) or np.any(hi[:, j] < OPEN)]
    out.update(kind="trees", lo=lo[:, used], hi=hi[:, used], value=np.array(vals), used=used, base=base,
               trees=len(reg.estimators_) if name in TREE_MODELS else len(reg._predictors))
    return out


def mirror(form: Dict[str, Any], F: np.ndarray) -> np.ndarray:
    """Python copy of the workbook formulas on model features F (rows x features, the user's units)."""
    F = np.where(np.isfinite(F), F, form["median"])
    if form["kind"] == "linear":
        return form["intercept"] + ((F - form["mean"]) / form["scale"]) @ form["coef"]
    X = F[:, form["used"]]
    lo, hi, val = form["lo"], form["hi"], form["value"]
    out = np.full(len(F), float(form["base"]))
    step = max(1, 4_000_000 // max(1, lo.size))           # points per block (memory: block x leaves x features)
    for s in range(0, len(X), step):
        x = X[s:s + step, None, :]
        inside = np.all((x > lo[None]) & (x <= hi[None]), axis=2)
        out[s:s + step] += inside @ val
    return out


def model_features(cand: Dict[str, Any], df: pd.DataFrame) -> np.ndarray:
    """Model feature values in the user's units (raw inputs, derived formulas) before the imputer and scaler."""
    d = df[cand["inputs"]].astype(float)
    if cand["evals"]:
        d = apply_inference_evals(d, cand["evals"])
    return prepare_features(d, cand["schema"])[cand["form"]["features"]].to_numpy(dtype=float)


def predict_frame(cand: Dict[str, Any], df: pd.DataFrame) -> np.ndarray:
    d = df[cand["inputs"]].astype(float)
    if cand["evals"]:
        d = apply_inference_evals(d, cand["evals"])
    return np.asarray(cand["pipe"].predict(prepare_features(d, cand["schema"])), dtype=float)


# -- derived features (--eval) as Excel formulas ----------------------------------------------------

_XL_FUNCS = {
    "log": lambda a: f"LN(MAX(ABS({a[0]}),1E-9))",
    "log1p": lambda a: f"LN(1+ABS({a[0]}))",
    "log10": lambda a: f"LOG10(MAX(ABS({a[0]}),1E-9))",
    "sqrt": lambda a: f"SQRT(ABS({a[0]}))",
    "exp": lambda a: f"EXP(MIN(MAX({a[0]},-50),50))",
    "abs": lambda a: f"ABS({a[0]})",
    "sin": lambda a: f"SIN({a[0]})",
    "cos": lambda a: f"COS({a[0]})",
    "tan": lambda a: f"TAN({a[0]})",
    "square": lambda a: f"(({a[0]})^2)",
    "clip": lambda a: f"MIN(MAX({a[0]},{a[1]}),{a[2]})",
    "radians": lambda a: f"RADIANS({a[0]})",
    "degrees": lambda a: f"DEGREES({a[0]})",
}


def excel_expr(node: ast.AST, refs: Dict[str, str]) -> str:
    """One --eval expression (ast of compile_formula's tokens) as an Excel formula; every operation in parentheses
    (Excel's -2^2 is 4, Python's is -4)."""
    if isinstance(node, ast.Expression):
        return excel_expr(node.body, refs)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return num(node.value)
    if isinstance(node, ast.Name):
        if node.id in refs:
            return refs[node.id]
        if node.id == "pi":
            return "PI()"
        if node.id == "e":
            return "EXP(1)"
    if isinstance(node, ast.UnaryOp):
        a = excel_expr(node.operand, refs)
        return f"(-({a}))" if isinstance(node.op, ast.USub) else a
    if isinstance(node, ast.BinOp):
        a, b = excel_expr(node.left, refs), excel_expr(node.right, refs)
        op = type(node.op)
        if op is ast.Pow:
            # the model's (Python) power: 0 ** 0 = 1 (Excel #NUM!), a negative base with a fractional exponent is no
            # number (-> the imputer's median) where Excel gives the odd root ((-8)^(1/3) = -2) (결정 1-8)
            return f"IF(AND(({a})=0,({b})=0),1,IF(AND(({a})<0,({b})<>INT({b})),NA(),(({a})^({b}))))"
        if op in (ast.Add, ast.Sub, ast.Mult, ast.Div):
            sym = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/"}[op]
            return f"(({a}){sym}({b}))"
        if op is ast.FloorDiv:
            return f"INT(({a})/({b}))"
        if op is ast.Mod:
            return f"MOD({a},{b})"
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in _XL_FUNCS:
        return _XL_FUNCS[node.func.id]([excel_expr(x, refs) for x in node.args])
    raise _unsupported(f"파생변수 수식을 엑셀 수식으로 바꾸지 못함 ({type(node).__name__})")


def derived_formulas(cand: Dict[str, Any]) -> List[Tuple[str, Any, float]]:
    """[(name, function(refs by name) -> Excel formula, median)] of the derived features in order (later ones may use
    earlier ones). The formula is the value before filling: an error (division by 0, a negative base with a fractional
    exponent) stays an error, so a later formula that reads it is an error too, like the model's empty value. The model
    reads the filled cell IFERROR(value, median): the imputer's median of that feature (결정 1-8)."""
    known = list(cand["inputs"])
    med = dict(zip(cand["form"]["features"], cand["form"]["median"]))
    out = []
    for e in cand["evals"]:
        lhs, rhs = (x.strip() for x in e.split("=", 1))
        tokens, tok2col, _ = compile_formula(rhs, known)
        tree = ast.parse(tokens, mode="eval")

        def make(refs: Dict[str, str], tree=tree, tok2col=tok2col) -> str:
            return excel_expr(tree, {t: refs[c] for t, c in tok2col.items()})
        out.append((lhs, make, float(med.get(lhs, 0.0))))
        known.append(lhs)
    return out


# ==============================================================================
# 3. Workbook
# ==============================================================================

def _col(i: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def _label(name: str) -> Tuple[str, str]:
    """'반응온도(°C)' -> ('반응온도', '°C')."""
    m = re.match(r"^(.*?)\s*[\(\[]([^\(\)\[\]]*)[\)\]]\s*$", str(name))
    return (m.group(1).strip(), m.group(2).strip()) if m else (str(name), "")


def _decimals(span: float) -> int:
    return int(max(0, 3 - np.floor(np.log10(span)))) if span > 0 else 3


def _nice(v: float, span: float) -> float:
    """Rounds a value for display to ~3 significant digits of the input's range (170.0 ~ 195.0, 0.001 ~ 1.000,
    1.23e-07 for an input of 1e-7 size: never rounded to 0, 2.9.88 D1)."""
    return round(float(v), _decimals(span))


def _nice_txt(v: float, span: float) -> str:
    """The range text: decimals as _nice, exponent form when more than 6 decimals would be needed ('1.230E-07')."""
    dec = _decimals(span)
    if dec > 6:
        return f"{float(v):.3E}"
    return f"{float(v):.{dec}f}"


def names_brief(names: List[str], limit: int = NAMES_SHOWN) -> str:
    """Input names for preview and the report: without unit (the full name when two would read the same, 2.9.88 D2),
    the first `limit` and '외 N개' (B7: 140 names made a 2,000-letter preview)."""
    short = target_labels(list(names))
    shown = ", ".join(short[n] for n in list(names)[:limit])
    return shown + (f" 외 {len(names) - limit}개" if len(names) > limit else "")


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


def method_text(form: Dict[str, Any]) -> str:
    if form["kind"] == "linear":
        return f"{form['model_class']} 선형식 그대로"
    return f"{form['model_class']} 나무 {form['trees']}개의 잎 {len(form['value']):,}개 그대로"


def write_workbook(path: str, title: str, vinfo: Dict[str, Any], plans: List[Dict[str, Any]]) -> str:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.formula import ArrayFormula
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
        p["sheet"] = _sheet_name("모델_", short[p["target"]], used)
        p["ws"] = wb.create_sheet(p["sheet"])
        p["ws"].sheet_state = "veryHidden"   # not shown (not even in "unhide"); formulas still read it
    acc = wb.create_sheet("정확도")
    info = wb.create_sheet("정보")

    # union of inputs, in model order of the first target, then the others
    all_inputs: List[str] = []
    for p in plans:
        for c in p["cand"]["inputs"]:
            if c not in all_inputs:
                all_inputs.append(c)
    # training range per input = the range every target using it was trained on: "범위 밖" is outside all of them
    rng = {}
    for c in all_inputs:
        users = [p for p in plans if c in p["cand"]["inputs"]]
        lo = max(rng_of(p["data"], c)[0] for p in users)
        hi = min(rng_of(p["data"], c)[1] for p in users)
        med = float(pd.concat([coerce_numeric(p["data"][c]) for p in users]).median())
        rng[c] = (lo, hi, min(max(med, lo), hi))

    # -- model sheets: leaf boxes, then the training points --------------------------------------------
    for p in plans:
        ws, form, cand = p["ws"], p["cand"]["form"], p["cand"]
        head: List[Any] = []
        blocks: List[np.ndarray] = []
        p["leaf_rng"] = None
        if form["kind"] == "trees":
            n = len(form["value"])
            lo_r, hi_r = [], []
            for j, f in enumerate(form["used"]):
                fname = form["features"][f]
                head += [f"{fname} 하한(초과)", f"{fname} 상한(이하)"]
                lo_r.append(f"'{p['sheet']}'!${_col(len(head) - 1)}$2:${_col(len(head) - 1)}${n + 1}")
                hi_r.append(f"'{p['sheet']}'!${_col(len(head))}$2:${_col(len(head))}${n + 1}")
            head.append("잎 값")
            vcol = len(head)
            p["leaf_rng"] = {"lo": lo_r, "hi": hi_r, "value": f"'{p['sheet']}'!${_col(vcol)}$2:${_col(vcol)}${n + 1}"}
            p["leaf_cols"] = {"first": 1, "value": vcol, "rows": n}
            leafs = np.empty((n, vcol))
            leafs[:, 0:-1:2], leafs[:, 1:-1:2], leafs[:, -1] = form["lo"], form["hi"], form["value"]
            blocks.append(leafs)
            head.append(None)                  # one empty column between the leaves and the training points
            blocks.append(np.full((n, 1), np.nan))
        T = p["cover"]
        c0 = len(head) + 1
        head += [f"학습 {c}" for c in cand["inputs"]]
        p["cover_rng"] = [f"'{p['sheet']}'!${_col(c0 + j)}$2:${_col(c0 + j)}${len(T) + 1}" for j in range(len(cand["inputs"]))]
        blocks.append(T)
        ws.append(head)                        # the numbers go in after saving (sheet_numbers): much faster
        p["numbers"] = blocks

    def model_cells(ws, p, x_refs: Dict[str, str], row: int, col0: int, ready: str) -> Tuple[str, str, int]:
        """Writes the helper cells of one prediction (derived features, distance to the nearest training row)
        starting at col0. Returns (prediction formula, distance cell, next free column)."""
        cand, form = p["cand"], p["cand"]["form"]
        refs = dict(x_refs)          # what a formula reads: inputs, the derived values before filling
        filled = dict(x_refs)        # what the model reads: inputs, the derived values filled with the median
        c = col0
        for name, make, fill in p["derived"]:
            ws.cell(row=row, column=c, value=f"={make(refs)}")
            ws.cell(row=row, column=c + 1, value=f"=IFERROR({_col(c)}{row},{num(fill)})")
            refs[name] = f"${_col(c)}{row}"
            filled[name] = f"${_col(c + 1)}{row}"
            c += 2
        feats = [filled[f] for f in form["features"]]
        if form["kind"] == "linear":
            pred = num(form["intercept"]) + "".join(
                f"+{num(w)}*(({x})-{num(m)})/{num(s)}" for x, w, m, s in zip(feats, form["coef"], form["mean"], form["scale"]))
        else:
            lr = p["leaf_rng"]
            terms = "*".join(f"({feats[f]}>{lo})*({feats[f]}<={hi})" for f, lo, hi in zip(form["used"], lr["lo"], lr["hi"]))
            pred = (f"{num(form['base'])}+" if form["base"] else "") + (f"SUMPRODUCT({terms}*{lr['value']})" if terms
                                                                       else f"SUM({lr['value']})")
        spans = [(x_refs[ci], r, rng_of(p["data"], ci)) for ci, r in zip(cand["inputs"], p["cover_rng"])]
        parts = [f"(({x}-{r})/{num(b - a)})^2" for x, r, (a, b) in spans if b > a]
        dist = f"${_col(c)}{row}"
        ws[dist.replace("$", "")] = ArrayFormula(dist.replace("$", ""), f"=IF({ready},MIN(SQRT({'+'.join(parts)})),\"\")"
                                                 if parts else "=0")
        return pred, dist, c + 1

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
    low = [p for p in plans if p["cand"]["val_r2"] < LOW_R2]
    style(calc["B4"], size=10, color="C2410C" if low else "374151", bold=bool(low), border=False)
    calc["B4"].value = "노란 칸에 값을 넣으면 아래에 예측값이 바로 계산됩니다." + "".join(
        f" 주의: '{short[p['target']]}' 모델은 검증 R² {p['cand']['val_r2']:.2f}로 정확도가 낮아 예측값은 참고용입니다."
        for p in low)
    calc.row_dimensions[2].height = 30
    r = 5
    for j, h in enumerate(("입력", "값", "단위", "학습 범위", "상태"), 2):
        style(calc.cell(row=r, column=j, value=h), bold=True, fill=HEAD, align="center")
    x_refs: Dict[str, str] = {}
    dv = DataValidation(type="decimal", operator="between", formula1="-1E+307", formula2="1E+307", allow_blank=True,
                        showErrorMessage=True, errorTitle="숫자만 입력", error="숫자만 입력할 수 있습니다.")
    calc.add_data_validation(dv)
    first_input = None
    in_short = target_labels(all_inputs)       # '온도(입구)', '온도(출구)': full names, never '온도' twice (D2)
    for i, c in enumerate(all_inputs):
        row = r + 1 + i
        lo, hi, med = rng[c]
        name, unit = _label(c)
        if in_short[c] != strip_unit(c):
            name, unit = str(c), ""
        style(calc.cell(row=row, column=2, value=name), bold=True)
        v = calc.cell(row=row, column=3, value=_nice(med, hi - lo) if hi > lo else float(med))
        style(v, size=11, bold=True, fill=YELLOW, align="right")
        v.protection = UNLOCK
        dv.add(v.coordinate)
        first_input = first_input or v.coordinate
        style(calc.cell(row=row, column=4, value=unit), align="center")
        style(calc.cell(row=row, column=5, value=f"{_nice_txt(lo, hi - lo)} ~ {_nice_txt(hi, hi - lo)}"), align="center")
        x = f"$C${row}"
        x_refs[c] = x
        chk = f"OR({x}<{num(lo)},{x}>{num(hi)})"
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
        ready = f"COUNT({','.join(own)})={len(own)}"
        pf, dist, nxt = model_cells(calc, p, x_refs, row, helper_col, ready)
        iv = p["cand"]["interval"] or {"residual_lower": 0.0, "residual_upper": 0.0}
        style(calc.cell(row=row, column=2, value=name), bold=True)
        fmt = p["fmt"]
        style(calc.cell(row=row, column=3, value=f'=IF({ready},{pf},"")'), size=14, bold=True, color=NAVY, align="right", fmt=fmt)
        style(calc.cell(row=row, column=4, value=unit), align="center")
        style(calc.cell(row=row, column=5, value=(f'=IF({ready},TEXT(C{row}+{num(iv["residual_lower"])},"{fmt}")&" ~ "&'
                                                   f'TEXT(C{row}+{num(iv["residual_upper"])},"{fmt}"),"")')), align="center")
        out_rng = "OR(" + ",".join(f"OR({x}<{num(rng[c][0])},{x}>{num(rng[c][1])})" for c, x in zip(p["cand"]["inputs"], own)) + ")"
        style(calc.cell(row=row, column=6, value=(f'=IF(NOT({ready}),"",IF({out_rng},"참고용 (학습 범위 밖)",'
                                                   f'IF({dist}>{num(p["uncovered"])},"참고용 (실측 없는 영역)","{p["ok_text"]}")))')),
              align="center")
        status_colors(calc, f"F{row}", f"F{row}")
        p["calc_cell"] = f"C{row}"
        for c in range(helper_col, nxt):
            calc.column_dimensions[_col(c)].width = 9
            calc.cell(row=row, column=c).font = Font(name=FONT, size=8, color="9CA3AF")
    note_row = r2 + len(plans) + 2
    calc.merge_cells(f"B{note_row}:F{note_row + 3}")
    style(calc[f"B{note_row}"], size=9, color="374151", border=False)
    calc[f"B{note_row}"].alignment = Alignment(wrap_text=True, vertical="top")
    calc[f"B{note_row}"].value = ("예측값은 학습한 모델과 같은 계산입니다 (모델을 수식으로 그대로 옮김). "
                                  "90% 범위: 실제 값이 이 범위에 들 가능성이 약 90%입니다. "
                                  "학습 범위 밖이거나 실측 데이터가 없는 영역의 값은 참고용입니다.")
    calc.print_area = f"B2:F{note_row + 3}"
    calc.sheet_properties.pageSetUpPr.fitToPage = True
    calc.page_setup.fitToWidth = 1
    calc.page_setup.fitToHeight = 0
    for p in plans:
        p["input_cells"] = {c: x_refs[c].replace("$", "") for c in p["cand"]["inputs"]}

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
        p["batch_inputs"] = {c: all_inputs.index(c) + 1 for c in p["cand"]["inputs"]}     # for the Excel check
        p["batch_pred_col"] = col
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
            ready = f"COUNT({','.join(own)})={len(own)}"
            pf, dist, hc = model_cells(batch, p, refs, row, hc, ready)
            iv = p["cand"]["interval"] or {"residual_lower": 0.0, "residual_upper": 0.0}
            pc = res_cols[rc]
            fmt = p["fmt"]
            batch.cell(row=row, column=pc, value=f'=IF({ready},{pf},"")').number_format = fmt
            batch.cell(row=row, column=pc + 1, value=f'=IF({ready},{_col(pc)}{row}+{num(iv["residual_lower"])},"")').number_format = fmt
            batch.cell(row=row, column=pc + 2, value=f'=IF({ready},{_col(pc)}{row}+{num(iv["residual_upper"])},"")').number_format = fmt
            out_rng = "OR(" + ",".join(f"OR({x}<{num(rng[c][0])},{x}>{num(rng[c][1])})" for c, x in zip(p["cand"]["inputs"], own)) + ")"
            batch.cell(row=row, column=pc + 3, value=(f'=IF(NOT({ready}),"",IF({out_rng},"참고용 (학습 범위 밖)",'
                                                      f'IF({dist}>{num(p["uncovered"])},"참고용 (실측 없는 영역)","{p["ok_text"]}")))'))
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
        ("- 예측값: 학습한 모델을 엑셀 수식으로 그대로 옮겨 계산한 값입니다 (모델 예측과 같습니다).", False),
        ("- 90% 범위: 실제 값이 이 범위 안에 들 가능성이 약 90%입니다. 좁을수록 정밀합니다.", False),
        ("- 판단 '참고용 (학습 범위 밖)': 입력이 모델이 배운 범위를 벗어났습니다. 값이 부정확할 수 있습니다.", False),
        ("- 판단 '참고용 (실측 없는 영역)': 각 입력은 범위 안이지만, 이런 조합의 실제 데이터가 근처에 없었습니다.", False),
        (f"- 판단 '참고용 (모델 정확도 낮음)': 모델의 검증 R²가 {LOW_R2} 미만이라 예측값을 그대로 쓰기 어렵습니다.", False),
        ("- 입력 칸이 하나라도 비어 있으면 예측값을 내지 않습니다 (모델은 빈 값을 학습 데이터의 중앙값으로 채워 예측하지만, "
         "계산기는 값을 모두 넣어야 계산합니다).", False),
        ("", False),
        ("이 파일에 대해", True),
        ("- 입력 항목과 모델 정보는 '정보' 시트, 모델 정확도와 계산 확인 결과는 '정확도' 시트에 있습니다.", False),
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
        c, chk = p["cand"], p["check"]
        style(acc.cell(row=row, column=2, value=f"{p['target']}"), size=12, bold=True, color=NAVY, border=False)
        for label, v in (("모델", c["model"]), ("입력", f"{len(c['inputs'])}개"), ("검증 R²", round(c["val_r2"], 4)),
                         ("RMSE", _sig(c["rmse"])), ("계산 방식", method_text(c["form"])),
                         ("모델과 같은지 확인", f"조건 {chk['points']:,}개 (무작위 {CHECK_POINTS:,}개 + 학습 데이터 {chk['points'] - CHECK_POINTS:,}행), "
                                         f"최대 차이 {chk['max_diff']:.3g} (예측값의 {chk['max_rel_pct']:.2g}%)")):
            row += 1
            style(acc.cell(row=row, column=2, value=label), border=False)
            style(acc.cell(row=row, column=3, value=v), border=False)
        row += 2
    acc.column_dimensions["B"].width = 20
    acc.column_dimensions["C"].width = 90

    info.sheet_view.showGridLines = False
    rows = [("제목", title), ("만든 날짜", f"{datetime.now():%Y-%m-%d %H:%M}"), ("ml-kit", VERSION),
            ("원래 모델", short_version(vinfo["dir"]))]
    for p in plans:
        c = p["cand"]
        rows += [("", ""), (f"[{p['target']}]", ""),
                 ("모델", f"{c['model']} / 입력 {len(c['inputs'])}개 / 검증 R² {round(c['val_r2'], 4)} / RMSE {_sig(c['rmse'])}"),
                 ("입력", ", ".join(c["inputs"])),
                 ("파생변수 (자동 계산)", ", ".join(c["evals"]) or "없음"),
                 ("계산 방식", method_text(c["form"])),
                 ("실측 없는 영역 기준", f"가까운 학습 행까지 거리(입력 0~1 환산) {p['uncovered']:.3f} 초과 (학습 행끼리 가장 "
                                    f"가까운 거리의 {UNCOVERED_Q:.0%}가 이 안, 최소 {UNCOVERED_DIST})"
                                    + (f" (학습 데이터 {COVER_ROWS:,}행 표본)" if p["cover_sampled"] else "")),
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

    check_formula_limits(calc, batch, all_inputs)
    sheet_files = {f"xl/worksheets/sheet{wb.sheetnames.index(p['sheet']) + 1}.xml": p for p in plans}

    def _save(pth: str) -> None:
        wb.save(pth)
        sheet_numbers(pth, {name: p["numbers"] for name, p in sheet_files.items()})
    return save_with_lock_fallback(path, _save, what="성능예측표")


def check_formula_limits(calc, batch, inputs: List[str]) -> None:
    """Excel's limits (2.9.88 B2): a formula of at most FORMULA_MAX letters and a function with at most FUNC_ARGS_MAX
    arguments (COUNT and OR take every input). A longer formula makes Excel repair the file or drop the formula, so the
    workbook is not saved: STOP_RETRY with the size and the way out (fewer inputs)."""
    longest, where = 0, ""
    rows = list(calc.iter_rows()) + list(batch.iter_rows(min_row=3, max_row=3))
    for row in rows:
        for c in row:
            v = c.value
            text = getattr(v, "text", v)          # ArrayFormula keeps its formula in .text
            if isinstance(text, str) and text.startswith("=") and len(text) > longest:
                longest, where = len(text), f"{c.parent.title}!{c.coordinate}"
    if len(inputs) > FUNC_ARGS_MAX or longest > FORMULA_MAX:
        raise InputError(f"입력이 {len(inputs)}개라 엑셀 계산기의 수식이 엑셀 한도를 넘어 성능예측표를 만들지 않았습니다 "
                         f"(가장 긴 수식 {longest:,}자 ({where}), 한도 {FORMULA_MAX:,}자, 함수 인자 한도 {FUNC_ARGS_MAX}개). "
                         f"입력을 줄여 다시 학습하면 만들 수 있습니다. 사용자에게 알리십시오.")


def _rows_xml(blocks: List[np.ndarray]) -> bytes:
    """Rows 2.. of a model sheet: the blocks side by side, every number with all its digits (repr: the cell holds
    the exact double; openpyxl would round to 16 digits and move a split bound), an empty cell for NaN or a
    shorter block."""
    from openpyxl.utils import get_column_letter
    width = sum(b.shape[1] for b in blocks)
    letters = [get_column_letter(i + 1) for i in range(width)]
    rows = max(len(b) for b in blocks)
    out = []
    for i in range(rows):
        r = i + 2
        cells = []
        c = 0
        for b in blocks:
            if i < len(b):
                for v in b[i]:
                    if v == v:
                        cells.append(f'<c r="{letters[c]}{r}" t="n"><v>{repr(float(v))}</v></c>')
                    c += 1
            else:
                c += b.shape[1]
        out.append(f'<row r="{r}">{"".join(cells)}</row>')
    return "".join(out).encode("ascii")


def sheet_numbers(path: str, sheets: Dict[str, List[np.ndarray]]) -> None:
    """Puts the leaf and training numbers into the saved model sheets (rows after the header). openpyxl takes
    about a minute for a few million cells; this writes the same XML in seconds."""
    import zipfile
    tmp = path + ".part"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename in sheets:
                blocks = sheets[item.filename]
                rows = _rows_xml(blocks)
                end = data.index(b"</sheetData>")
                data = data[:end] + rows + data[end:]
                last = f"{_col(sum(x.shape[1] for x in blocks))}{max(len(x) for x in blocks) + 1}"
                data = re.sub(rb'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{last}"/>'.encode("ascii"), data, count=1)
            dst.writestr(item, data)
    os.replace(tmp, path)


def short_version(vdir: str) -> str:
    """'<data>_ml_data\\v3' only: the full path (user name, folders) is personal information and is not distributed."""
    return os.path.join(os.path.basename(os.path.dirname(vdir)), os.path.basename(vdir))


def rng_of(data: pd.DataFrame, col: str) -> Tuple[float, float]:
    s = coerce_numeric(data[col])
    return float(s.min()), float(s.max())


# ==============================================================================
# 4. Verification
# ==============================================================================

def uncovered_distance(T: np.ndarray, data: pd.DataFrame, inputs: List[str]) -> float:
    """The "no measured data nearby" distance: farther from every training row than UNCOVERED_Q of the training rows
    are from their own nearest one (inputs scaled to 0-1, like the workbook), at least UNCOVERED_DIST."""
    from scipy.spatial import cKDTree
    span = np.array([b - a for a, b in (rng_of(data, c) for c in inputs)])
    keep = span > 0
    if len(T) < 3 or not keep.any():
        return UNCOVERED_DIST
    lo = np.array([rng_of(data, c)[0] for c in inputs])
    Z = ((T - lo) / np.where(keep, span, 1.0))[:, keep]
    d, _ = cKDTree(Z).query(Z, k=2)
    return float(max(UNCOVERED_DIST, np.quantile(d[:, 1], UNCOVERED_Q)))


def check_points(cand: Dict[str, Any], data: pd.DataFrame) -> pd.DataFrame:
    """Random conditions inside the training range plus every training row (the rows the splits were made from)."""
    rng = np.random.default_rng(0)
    P = pd.DataFrame({c: rng.uniform(*rng_of(data, c), CHECK_POINTS) for c in cand["inputs"]})
    T = data[cand["inputs"]].apply(coerce_numeric)
    T = T[np.isfinite(T.to_numpy(dtype=float)).all(axis=1)]
    return pd.concat([P, T], ignore_index=True)


def model_check(cand: Dict[str, Any], form: Dict[str, Any], P: pd.DataFrame) -> Dict[str, Any]:
    """The workbook formulas (mirror) against the model's own prediction."""
    want = predict_frame(cand, P)
    got = mirror(form, model_features(cand, P))
    d = np.abs(got - want)
    scale = np.maximum(np.abs(want), 1e-12)
    return {"points": int(len(P)), "max_diff": float(d.max()), "max_rel_pct": float((d / scale).max() * 100),
            "same": bool(np.all(d <= 1e-9 * (np.abs(want) + float(np.ptp(want)) + 1e-12)))}


def verify_workbook(path: str, plans: List[Dict[str, Any]]) -> List[str]:
    """Re-reads the saved file (read-only, fast): sheets, sheet and structure protection, unlocked inputs, the
    prediction formulas, and the leaves in the file bit for bit the leaves checked against the model."""
    import zipfile
    from openpyxl import load_workbook
    problems = []
    try:
        wb = load_workbook(path, read_only=True)
    except Exception as e:
        return [f"저장한 파일을 다시 열지 못함: {e}"]
    try:
        names = ["계산기", "일괄계산", "사용법", "정확도", "정보"] + [p["sheet"] for p in plans]
        problems += [f"'{n}' 시트 없음" for n in names if n not in wb.sheetnames]
        with zipfile.ZipFile(path) as z:        # openpyxl writes the i-th sheet as sheet<i>.xml
            for i, n in enumerate(wb.sheetnames, 1):
                if b"<sheetProtection" not in z.read(f"xl/worksheets/sheet{i}.xml"):
                    problems.append(f"'{n}' 시트 잠금 안 됨")
        if not (wb.security and wb.security.lockStructure):
            problems.append("통합문서 구조 잠금 안 됨")
        if "계산기" in wb.sheetnames:
            calc = wb["계산기"]
            for p in plans:
                if any(calc[c].protection.locked for c in p["input_cells"].values()):
                    problems.append(f"'{p['target']}' 계산기 입력 칸이 잠김")
                if not str(calc[p["calc_cell"]].value or "").startswith("=IF("):
                    problems.append(f"'{p['target']}' 계산기 예측 수식 없음")
        for p in plans:
            form = p["cand"]["form"]
            if form["kind"] != "trees" or p["sheet"] not in wb.sheetnames:
                continue
            lc = p["leaf_cols"]
            vals = np.array(list(wb[p["sheet"]].iter_rows(min_row=2, max_row=lc["rows"] + 1, min_col=lc["first"],
                                                          max_col=lc["value"], values_only=True)), dtype=float)
            if not (vals.shape == (len(form["value"]), 2 * len(form["used"]) + 1)
                    and np.array_equal(vals[:, 0:-1:2], form["lo"]) and np.array_equal(vals[:, 1:-1:2], form["hi"])
                    and np.array_equal(vals[:, -1], form["value"])):
                problems.append(f"'{p['target']}' 파일에 저장한 잎 값이 모델에서 옮긴 값과 다름")
    finally:
        wb.close()
    return problems


EXCEL_TOL = 0.01            # 엑셀 다시 계산이 모델과 이만큼(예측 범위의 1%, 99% 일치)까지 달라도 통과 (사용자 결정 2026-11-10)
EXCEL_NOTES: List[str] = []  # 허용 안에서 다른 곳 (결과 warnings로 알림)
EXCEL_EPS = 1e-13           # Excel compares numbers that agree to about 15 significant digits as equal (=0.1+0.2=0.3 is TRUE)


def above_for_excel(v: float) -> float:
    """The smallest step above a split bound that Excel still tells apart from the bound (2026-11-10 company PC F7:
    the next double above a bound compared as equal in Excel, so the row landed in the left leaf; LibreOffice was exact)."""
    return float(v + max(abs(v) * EXCEL_EPS, EXCEL_EPS))


def edge_points(p: Dict[str, Any]) -> pd.DataFrame:
    """Conditions exactly on a split bound of a raw input and just above it: Excel must read the bounds with all their
    digits for these to land on the model's side. 'Just above' is the smallest step Excel tells apart (above_for_excel):
    numbers closer to a bound than about 1e-15 of its size are the bound itself for Excel."""
    form, cand = p["cand"]["form"], p["cand"]
    if form["kind"] != "trees":
        return pd.DataFrame(columns=cand["inputs"])
    med = {c: float(coerce_numeric(p["data"][c]).median()) for c in cand["inputs"]}
    rows = []
    for j, f in enumerate(form["used"]):
        name = form["features"][f]
        if name not in cand["inputs"]:
            continue
        hi = form["hi"][:, j]
        b = hi[hi < OPEN]
        if b.size:
            v = float(b[len(b) // 2])
            rows += [dict(med, **{name: v}), dict(med, **{name: above_for_excel(v)})]
        if len(rows) >= 6:
            break
    return pd.DataFrame(rows, columns=cand["inputs"])


DERIVED_POINTS = 12          # conditions for the derived formulas in the Excel check (0 and negative inputs)


def derived_points(p: Dict[str, Any]) -> pd.DataFrame:
    """Conditions where the derived formulas meet 0 and negative values (0 ** 0, a negative base with a fractional
    exponent, division by 0): all inputs 0, all -1.5, and each input 0 or negative with the others at the median
    (결정 1-8: Excel and the model's Python differ there unless the formula says what Python does)."""
    cand = p["cand"]
    if not cand["evals"]:
        return pd.DataFrame(columns=cand["inputs"])
    med = {c: float(coerce_numeric(p["data"][c]).median()) for c in cand["inputs"]}
    rows = [{c: 0.0 for c in cand["inputs"]}, {c: -1.5 for c in cand["inputs"]}]
    for c in cand["inputs"]:
        rows += [dict(med, **{c: 0.0}), dict(med, **{c: -abs(med[c]) - 1.0})]
    return pd.DataFrame(rows[:DERIVED_POINTS], columns=cand["inputs"])


def compare_excel(target: str, got: List[Any], want: List[float], wheres: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """엑셀이 다시 계산한 값과 모델 예측 비교 -> (문제, 알림). 차이가 확인한 조건들의 예측 범위(최대 - 최소)의
    EXCEL_TOL(1%, 99% 일치)을 넘으면 문제, 0보다 크고 그 안이면 알림 (사용자 결정 2026-11-10: 99%쯤 일치하면 통과,
    결정 1-5: 기준은 예측값 크기가 아니라 타겟이 움직이는 폭. 모든 조건의 예측이 같으면(범위 0) 예측값 크기)."""
    for g, where in zip(got, wheres):
        if not isinstance(g, (int, float)) or isinstance(g, bool):
            return f"'{target}' 엑셀 계산({where})이 숫자가 아님: {g}", None
    w_arr = np.asarray(want, dtype=float)
    span = float(np.ptp(w_arr)) if w_arr.size else 0.0
    size = float(np.max(np.abs(w_arr))) if w_arr.size else 0.0
    base, what = (span, f"예측 범위 {span:.6g}") if span > 1e-12 * (size + 1.0) else (max(size, 1e-12), f"예측값 {size:.6g}")
    worst = (0.0, "", 0.0, 0.0)
    for g, w, where in zip(got, want, wheres):
        diff = abs(float(g) - w)
        rel = diff / base if diff > 1e-9 * (abs(w) + span + 1.0) else 0.0
        if rel > worst[0]:
            worst = (rel, where, float(g), float(w))
    if worst[0] > EXCEL_TOL:
        return (f"'{target}' 엑셀 계산({worst[1]}) {worst[2]:.10g}이 모델 {worst[3]:.10g}과 {what}의 {worst[0] * 100:.2f}% 다름 "
                f"(허용 {EXCEL_TOL * 100:g}%)"), None
    if worst[0] > 0:
        return None, f"'{target}' 엑셀 계산이 모델과 최대 {what}의 {worst[0] * 100:.3g}% 다름({worst[1]}, 허용 {EXCEL_TOL * 100:g}% 안)"
    return None, None


def excel_check(path: str, plans: List[Dict[str, Any]]) -> Optional[str]:
    """Windows with Excel: opens the file (not saved), puts the check conditions (random ones and the ones on a split
    bound) in the 일괄계산 rows at once and one condition in the calculator, recalculates once per target and compares
    with the model (2.9.85: was one full recalculation per condition, minutes for a big workbook).
    None = checked OK, 'skip' = Excel not available, 'off' = turned off (ML_KIT_SKIP_EXCEL_CHECK=1, tests), otherwise a problem."""
    import importlib.util
    if os.environ.get("ML_KIT_SKIP_EXCEL_CHECK") == "1":
        return "off"
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
        calc, batch = wb.Worksheets("계산기"), wb.Worksheets("일괄계산")
        problems = []
        for p in plans:
            P = pd.concat([p["points"].sample(EXCEL_CHECKS, random_state=1), edge_points(p), derived_points(p)],
                          ignore_index=True)
            want = predict_frame(p["cand"], P)
            n = len(P)
            for c, col in p["batch_inputs"].items():
                batch.Range(batch.Cells(3, col), batch.Cells(2 + n, col)).Value = tuple((float(v),) for v in P[c])
            for c in p["cand"]["inputs"]:
                calc.Range(p["input_cells"][c]).Value = float(P.at[0, c])
            xl.Calculate()
            got = [r[0] for r in batch.Range(batch.Cells(3, p["batch_pred_col"]), batch.Cells(2 + n, p["batch_pred_col"])).Value]
            got.append(calc.Range(p["calc_cell"]).Value)
            problem, note = compare_excel(p["target"], got, list(want) + [want[0]],
                                          [f"일괄계산 {i + 3}행" for i in range(n)] + ["계산기"])
            if problem:
                problems.append(problem)
            elif note:
                EXCEL_NOTES.append(note)
        return "; ".join(problems) or None
    except Exception as e:
        return f"엑셀로 계산 확인 실패: {type(e).__name__}: {e}"
    finally:
        # xl.Quit도 lambda 안에서 찾음: 엑셀 연결이 끊기면 속성 찾기(xl.Quit)부터 AttributeError가 나는데, 여기서 나면 try 밖이라
        # 확인 결과 대신 그 오류가 올라가 성능예측표 전체가 실패로 보고됨 (회사 PC 2026-11-10 beta F1)
        for close in ((lambda: wb.Close(SaveChanges=False)) if wb is not None else None, lambda: xl.Quit()):
            try:
                if close:
                    close()
            except Exception:
                pass


# ==============================================================================
# 5. Command
# ==============================================================================

def execute_export(args) -> Dict[str, Any]:
    args.model = model_dir_from_data(args.model)
    vinfo = load_version(args.model, args.targets)
    data_all = vinfo["data"]
    warnings: List[str] = []
    if vinfo["skipped"]:
        warnings.append(f"분류 타겟 {vinfo['skipped']}은(는) 성능예측표에서 뺐습니다 (숫자 타겟만 가능)")
    plans = []
    for target in vinfo["targets"]:
        meta = vinfo["metas"][target]
        inputs, evals = raw_inputs(meta)
        missing = [c for c in inputs + [target] if c not in data_all.columns]
        if missing:
            raise InputError(f"학습 데이터 사본에 {missing} 열이 없습니다. 다시 학습한 버전으로 실행하십시오.")
        data = data_all.copy()
        excl = f"{SNAPSHOT_EXCL_PREFIX}{target}"
        if excl in data.columns:
            data = data[coerce_numeric(data[excl]).fillna(0) != 1]
        data = data[coerce_numeric(data[target]).notna()].reset_index(drop=True)
        const = [c for c in inputs if coerce_numeric(data[c]).nunique(dropna=True) <= 1]
        if const:
            warnings.append(f"'{target}': 입력 {const}은(는) 학습 데이터에서 값이 하나뿐입니다 (계산기 기본값이 그 값이고, "
                            f"다른 값은 '범위 밖')")
        progress(f"'{target}': 모델을 엑셀 수식으로 옮기는 중")
        pipe = load_pipe(vinfo["dir"], meta)
        form = model_formula(pipe)
        cand = {"inputs": list(inputs), "evals": list(evals), "schema": meta["feature_schema"], "pipe": pipe,
                "val_r2": float(meta["val_r2"]), "rmse": float(meta["rmse"]), "model": meta.get("champion_model"),
                "interval": meta.get("prediction_interval"), "form": form}
        known = set(inputs) | {e.split("=", 1)[0].strip() for e in evals}
        if not set(form["features"]) <= known:
            raise _unsupported(f"모델 입력 {sorted(set(form['features']) - known)}을(를) 입력·파생변수에서 찾지 못함")
        T = data[inputs].apply(coerce_numeric)
        T = T[np.isfinite(T.to_numpy(dtype=float)).all(axis=1)].to_numpy(dtype=float)
        sampled = len(T) > COVER_ROWS
        if sampled:
            T = T[np.sort(np.random.default_rng(0).choice(len(T), COVER_ROWS, replace=False))]
        uncovered = uncovered_distance(T, data, inputs)
        P = check_points(cand, data)
        chk = model_check(cand, form, P)
        if not chk["same"]:
            raise _unsupported(f"옮긴 수식이 모델과 다름 (조건 {chk['points']:,}개 중 최대 차이 {chk['max_diff']:.3g})")
        if cand["val_r2"] < LOW_R2:      # C2: the report calls it '부적합'; the calculator says so in its own cells
            warnings.append(f"'{target}': 검증 R² {cand['val_r2']:.2f}로 {LOW_R2} 미만이라 예측값은 참고용입니다 (계산기 판단 칸에 "
                            f"'참고용 (모델 정확도 낮음)'으로 표시)")
        plans.append({"target": target, "cand": cand, "data": data, "cover": T, "cover_sampled": sampled, "uncovered": uncovered,
                      "derived": derived_formulas(cand), "check": chk, "points": P,
                      "fmt": value_format(predict_frame(cand, P)),
                      "ok_text": "정상" if cand["val_r2"] >= LOW_R2 else "참고용 (모델 정확도 낮음)"})
    leaves = sum(len(p["cand"]["form"].get("value", [])) for p in plans)
    if leaves > MAX_LEAVES:
        each = ", ".join(f"'{p['target']}' {len(p['cand']['form'].get('value', [])):,}개" for p in plans)
        raise InputError(f"모델이 커서 성능예측표를 만들지 않았습니다: 나무 잎이 모두 {leaves:,}개 ({each}), 상한 {MAX_LEAVES:,}개 "
                         f"(엑셀 파일이 너무 커짐). 데이터 행이 많을수록 RandomForest·ExtraTrees의 잎이 많아집니다. "
                         f"--targets로 타겟을 나눠 만들 수 있는지 사용자에게 알리십시오.")

    title = (args.title or "").strip() or default_title(vinfo["metas"][plans[0]["target"]])
    if len(title) > TITLE_MAX:
        warnings.append(f"제목이 {len(title)}자로 길어 {TITLE_MAX}자로 줄였습니다")
        title = title[:TITLE_MAX - 3].rstrip() + "..."
    # same folder as the report and history.jsonl (<data>_ml_data), name paired with ML결과보고서_<data>_<version>.docx
    data_root = os.path.dirname(vinfo["dir"])
    stem = re.sub(r"_ml_data$", "", os.path.basename(data_root)) or "모델"
    default = os.path.join(data_root, f"성능예측표_{stem}_{os.path.basename(vinfo['dir'])}.xlsx")
    path = save_path(args.save, data_folder_of(vinfo["dir"])) if args.save else default
    # the 'recovered' note of a new name waits for the saved file (B6: a run stopped before saving reported a file
    # that was never written)
    path = avoid_overwrite(os.path.abspath(path), getattr(args, "overwrite", False), what="성능예측표", defer=True)
    progress("엑셀 파일 작성")
    real = write_workbook(path, title, vinfo, plans)
    saved_as(path)
    mark_committed(f"성능예측표 {os.path.basename(real)}", output_file=real)   # a stop now: saved, not CONTINUE
    progress("저장한 파일 확인")
    problems = verify_workbook(real, plans)
    EXCEL_NOTES.clear()
    xl = excel_check(real, plans)
    warnings += EXCEL_NOTES
    unverified = list(problems)
    if xl == "off":
        unverified.append("엑셀 계산 확인을 껐음 (ML_KIT_SKIP_EXCEL_CHECK=1, 파일 구조, 잠금, 저장한 값으로 한 계산은 확인함)")
    elif xl == "skip":
        unverified.append("엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함 (파일 구조, 잠금, 저장한 값으로 한 계산은 확인함)")
    elif xl:
        unverified.append(xl)

    out_targets = {}
    for p in plans:
        c, chk = p["cand"], p["check"]
        out_targets[p["target"]] = {
            "inputs": c["inputs"], "model": c["model"], "val_r2": round(c["val_r2"], 4), "rmse": _sig(c["rmse"]),
            "method": method_text(c["form"]), "leaves": len(c["form"]["value"]) if c["form"]["kind"] == "trees" else None,
            "check": {"points": chk["points"], "max_diff": _sig(chk["max_diff"], 3), "max_rel_pct": _sig(chk["max_rel_pct"], 2)},
        }
        if c["evals"]:     # D5: the check above computes derived features in Python; their Excel formulas only in Excel
            out_targets[p["target"]]["derived_note"] = derived_note(len(c["evals"]), xl is None)
    res = {"status": "ok", "output_file": real, "title": title, "targets": out_targets,
           "verified": not unverified, "preview": f"성능예측표 생성: {export_brief(out_targets)}"}
    if warnings:
        res["warnings"] = warnings
    if unverified:
        res["unverified"] = unverified
    res["agent_directive"] = (
        ("VERIFIED: " if not unverified else "UNVERIFIED: 성능예측표 확인 실패 (" + "; ".join(unverified) + "). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + "preview와 output_file 경로를 그대로 알리세요."
        + (" warnings도 그대로 알리세요." if warnings else ""))
    return res


def derived_note(n: int, excel_ok: bool) -> str:
    """What was checked of the derived features' Excel formulas (결정 1-8, C영역 D5: the model check computes them in
    Python, so only Excel itself checks the formulas)."""
    if excel_ok:
        return f"파생변수 수식 {n}개는 엑셀에서 다시 계산해 확인"
    return f"파생변수 수식 {n}개는 엑셀이 있는 PC에서만 확인됨 (이번에는 엑셀로 확인 못 함)"


def export_brief(out_targets: Dict[str, Dict[str, Any]]) -> str:
    """The preview text: the inputs once when every target has the same ones (B7), names shortened (names_brief)."""
    items = list(out_targets.items())
    same = len(items) > 1 and all(v["inputs"] == items[0][1]["inputs"] for _, v in items)

    def one(t: str, v: Dict[str, Any]) -> str:
        low = " (정확도 낮음: 참고용)" if isinstance(v.get("val_r2"), (int, float)) and v["val_r2"] < LOW_R2 else ""
        return ((f"{t}: " if same else f"{t}: 입력 {len(v['inputs'])}개 ({names_brief(v['inputs'])}), ")
                + f"{v['model']}, 검증 R² {v['val_r2']}{low}, 모델과 같은 계산 (조건 {v['check']['points']:,}개 확인"
                + (f", {v['derived_note']})" if v.get("derived_note") else ")"))
    head = f"입력 {len(items[0][1]['inputs'])}개 ({names_brief(items[0][1]['inputs'])}), 모든 타겟 공통; " if same else ""
    return head + "; ".join(one(t, v) for t, v in items)


def save_path(save: str, data_folder: str) -> str:
    """--save -> the workbook path. A relative path is next to the original data file (결정 1-6: like the report, the
    dataset and the [추천] commands); no extension -> .xlsx; another extension -> STOP_RETRY + [추천] the same command
    with .xlsx (2.9.88 B3: 'my.xls' was saved as 'my.xls.xlsx'). A folder that does not exist is made; one that cannot
    be made is STOP_RETRY."""
    path = resolve_save(save, data_folder)
    ext = os.path.splitext(path)[1]
    if not ext:
        path += ".xlsx"
    elif ext.lower() != ".xlsx":
        fixed = os.path.splitext(path)[0] + ".xlsx"
        cmd = command_text(drop={"--save": "save"})
        rec = f"{cmd} --save {quote_arg(fixed)}" if cmd else f"--save {quote_arg(fixed)}"
        raise InputError(f"--save는 엑셀 파일(.xlsx)만 됩니다: '{save}'. [추천] {rec}")
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
    except OSError as e:
        raise InputError(f"성능예측표를 저장할 폴더를 만들 수 없습니다: {os.path.dirname(path)} ({e}). "
                         "--save 경로를 사용자에게 확인하십시오.")
    return path


def main():
    setup_console()
    parser = JsonArgumentParser(description="ml-kit 성능예측표 (엑셀, 모델을 수식으로 그대로 옮김, 잠금) 내보내기")
    parser.add_argument("--model", required=True, help="학습한 데이터 파일(도구가 최종 버전, 곧 마지막 학습을 씀) 또는 버전 폴더 (예: 공정_ml_data/v3)")
    parser.add_argument("--targets", help="표로 만들 숫자 타겟 (쉼표 구분, 생략 시 전부)")
    parser.add_argument("--title", help="계산기 제목 (데이터 설명을 짧게 정리한 글, 40자 이내)")
    parser.add_argument("--save", help="저장 경로 (.xlsx, 상대 경로와 이름만은 원본 데이터 파일 옆, 없는 폴더는 만듦. "
                                       "생략 시 보고서와 같은 <데이터>_ml_data 폴더에 성능예측표_<데이터>_<버전>.xlsx)")
    parser.add_argument("--overwrite", action="store_true",
                        help="같은 이름의 성능예측표가 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    args = parser.parse_args()
    run_cli(lambda: execute_export(args), hint="확인할 것: 버전 폴더 경로.")


if __name__ == "__main__":
    main()
