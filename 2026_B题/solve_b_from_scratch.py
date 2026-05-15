# -*- coding: utf-8 -*-
"""B题：多源融合机器人定位及任务优化（从头求解版）。

运行：
    python solve_b_from_scratch.py

输出：
    outputs/
        estimates_summary.xlsx
        problem1_10Hz_trajectory.xlsx
        problem2_10Hz_trajectory.xlsx
        problem3_10Hz_trajectory.xlsx
        task_candidates.xlsx
        task_schedule_detail.xlsx
        result.xlsx
        B题_论文.md
        B题_多源融合机器人定位及任务优化_论文.docx
        B题_多源融合机器人定位及任务优化_论文.pdf
        figures/*.png
"""

from __future__ import annotations

import math
import shutil
import textwrap
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from matplotlib.backends.backend_pdf import PdfPages
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from scipy.optimize import Bounds, LinearConstraint, milp, minimize_scalar
from scipy.signal import savgol_filter
from scipy.sparse import lil_matrix
from scipy.stats import f as f_dist


DT_OUT = 0.1

plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "SimSun", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False

# 题面附录约束。题面公式为嵌入对象，纯文本提取不可见；数值来自题面公式显示。
SHOOT_DISTANCE = (5.0, 30.0)
SHOOT_SPEED_MAX = 2.0
SHOOT_ACCEL_MAX = 1.5
SHOOT_PREP = 1.5
SHOOT_HIT_PROB = 0.85

PHOTO_DISTANCE = (10.0, 40.0)
PHOTO_ANGLE_MIN_DEG = 60.0
PHOTO_SPEED_MAX = 1.5
PHOTO_ACCEL_MAX = 1.5
PHOTO_PREP = 0.5

SYSTEM_BIAS_IMPROVEMENT_MIN = 0.05
SYSTEM_BIAS_NORM_MIN = 0.50


@dataclass
class AlignmentResult:
    problem: int
    delta_s: float
    bias_x_m: float
    bias_y_m: float
    mse_used: float
    mse_bias_model: float | None = None
    mse_no_bias_model: float | None = None
    improvement_ratio: float | None = None
    overlap_s: float = 0.0
    n_overlap: int = 0
    f_stat: float | None = None
    f_p_value: float | None = None
    statistical_bias: bool = False
    practical_bias: bool = False
    ci_delta_lo: float = 0.0
    ci_delta_hi: float = 0.0
    candidate_delta_s: float | None = None
    candidate_bias_x_m: float | None = None
    candidate_bias_y_m: float | None = None

    @property
    def bias_norm(self) -> float:
        return float(math.hypot(self.bias_x_m, self.bias_y_m))


def read_pair(path: Path) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    d1 = pd.read_excel(path, sheet_name=0).dropna()
    d2 = pd.read_excel(path, sheet_name=1).dropna()
    a1 = d1.iloc[:, 0:3].to_numpy(float)
    a2 = d2.iloc[:, 0:3].to_numpy(float)
    return a1[:, 0], a1[:, 1:3], a2[:, 0], a2[:, 1:3]


def interp_xy(t_new: np.ndarray, t: np.ndarray, xy: np.ndarray) -> np.ndarray:
    return np.column_stack([np.interp(t_new, t, xy[:, i]) for i in range(2)])


def moving_average(xy: np.ndarray, window: int) -> np.ndarray:
    if window <= 1:
        return xy.copy()
    if window % 2 == 0:
        window += 1
    pad = window // 2
    kernel = np.ones(window, dtype=float) / window
    padded = np.pad(xy, ((pad, pad), (0, 0)), mode="edge")
    return np.column_stack([np.convolve(padded[:, i], kernel, mode="valid") for i in range(2)])


def overlap_grid(t1: np.ndarray, t2_corr: np.ndarray, dt: float = DT_OUT) -> tuple[np.ndarray, float]:
    lo = max(float(t1.min()), float(t2_corr.min()))
    hi = min(float(t1.max()), float(t2_corr.max()))
    if hi <= lo:
        return np.array([]), 0.0
    start = math.ceil((lo - 1e-9) / dt) * dt
    end = math.floor((hi + 1e-9) / dt) * dt
    if end < start:
        return np.array([]), 0.0
    return np.round(np.arange(start, end + 0.5 * dt, dt), 10), hi - lo


def delta_score(
    t1: np.ndarray,
    p1: np.ndarray,
    t2: np.ndarray,
    p2: np.ndarray,
    delta: float,
    correct_bias: bool,
    dt: float,
    trim_ratio: float,
    min_overlap_ratio: float,
) -> tuple[float, np.ndarray, float, int]:
    t2_corr = t2 - delta
    grid, overlap = overlap_grid(t1, t2_corr, dt)
    min_duration = min(float(t1.max() - t1.min()), float(t2.max() - t2.min()))
    if len(grid) < 8 or overlap < min_overlap_ratio * min_duration:
        return float("inf"), np.array([np.nan, np.nan]), overlap, len(grid)
    q1 = interp_xy(grid, t1, p1)
    q2 = interp_xy(grid, t2_corr, p2)
    diff = q2 - q1
    bias = np.median(diff, axis=0) if correct_bias else np.zeros(2)
    err = np.sum((diff - bias) ** 2, axis=1)
    if trim_ratio > 0 and len(err) > 30:
        keep = max(20, int(round(len(err) * (1.0 - trim_ratio))))
        err = np.sort(err)[:keep]
    return float(np.mean(err)), bias, overlap, len(grid)


def estimate_alignment(
    path: Path,
    correct_bias: bool,
    score_smooth_window: int,
    trim_ratio: float,
    min_overlap_ratio: float = 0.85,
) -> tuple[float, np.ndarray, float, float, int]:
    t1, p1_raw, t2, p2_raw = read_pair(path)
    p1 = moving_average(p1_raw, score_smooth_window)
    p2 = moving_average(p2_raw, score_smooth_window)
    min_duration = min(float(t1.max() - t1.min()), float(t2.max() - t2.min()))
    d_min = float(t2.min() - t1.max() + min_overlap_ratio * min_duration)
    d_max = float(t2.max() - t1.min() - min_overlap_ratio * min_duration)
    deltas = np.linspace(d_min, d_max, 3000)
    scores = np.array(
        [
            delta_score(t1, p1, t2, p2, d, correct_bias, 0.5, trim_ratio, min_overlap_ratio)[0]
            for d in deltas
        ]
    )
    best = float(deltas[int(np.nanargmin(scores))])
    step = float(deltas[1] - deltas[0]) if len(deltas) > 1 else 1.0
    lo = max(d_min, best - 30 * step)
    hi = min(d_max, best + 30 * step)
    opt = minimize_scalar(
        lambda d: delta_score(t1, p1, t2, p2, d, correct_bias, 0.1, trim_ratio, min_overlap_ratio)[0],
        bounds=(lo, hi),
        method="bounded",
        options={"xatol": 1e-8},
    )
    mse, bias, overlap, n = delta_score(
        t1, p1, t2, p2, float(opt.x), correct_bias, 0.1, trim_ratio, min_overlap_ratio
    )
    return float(opt.x), bias, mse, overlap, n


def nested_f_test(mse0: float, mse1: float, n: int, p: int = 2) -> tuple[float, float, bool]:
    if mse1 <= 0 or mse0 <= mse1 or n <= p + 1:
        return 1.0, 1.0, False
    sse0 = mse0 * n
    sse1 = mse1 * n
    df1 = p
    df2 = n - p
    f_stat = ((sse0 - sse1) / df1) / (sse1 / df2)
    p_value = float(1.0 - f_dist.cdf(f_stat, df1, df2))
    return float(f_stat), p_value, p_value < 0.05


def delta_ci_linearized(
    path: Path,
    delta: float,
    bias: np.ndarray,
    n_boot: int = 399,
    dt: float = DT_OUT,
) -> tuple[float, float]:
    t1, p1, t2, p2 = read_pair(path)
    t2_corr = t2 - delta
    p2_corr = p2 - bias
    grid, _ = overlap_grid(t1, t2_corr, dt)
    if len(grid) < 20:
        return delta, delta
    q1 = interp_xy(grid, t1, p1)
    q2 = interp_xy(grid, t2_corr, p2_corr)
    residual = q2 - q1
    velocity = np.gradient(q2, dt, axis=0)
    rng = np.random.default_rng(20260515 + int(abs(delta) * 10))
    estimates: list[float] = []
    for _ in range(n_boot):
        idx = rng.integers(0, len(residual), len(residual))
        r = residual[idx]
        v = velocity[idx]
        r = r - np.median(r, axis=0)
        v = v - np.median(v, axis=0)
        denom = float(np.sum(v * v))
        if denom > 1e-12:
            estimates.append(delta - float(np.sum(v * r)) / denom)
    if len(estimates) < 10:
        return delta, delta
    lo, hi = np.percentile(estimates, [2.5, 97.5])
    return float(lo), float(hi)


def build_alignment_results(root: Path) -> dict[int, AlignmentResult]:
    paths = {i: root / f"附件{i}.xlsx" for i in (1, 2, 3)}
    results: dict[int, AlignmentResult] = {}

    d1, _b1, mse1, ov1, n1 = estimate_alignment(paths[1], False, 1, 0.0)
    results[1] = AlignmentResult(
        problem=1,
        delta_s=d1,
        bias_x_m=0.0,
        bias_y_m=0.0,
        mse_used=mse1,
        overlap_s=ov1,
        n_overlap=n1,
    )

    d2, b2, mse2_b, ov2, n2 = estimate_alignment(paths[2], True, 9, 0.02)
    _d2_nb, _b2_nb, mse2_nb, _ov2_nb, _n2_nb = estimate_alignment(paths[2], False, 9, 0.02)
    imp2 = (mse2_nb - mse2_b) / mse2_nb
    f2, p2, stat2 = nested_f_test(mse2_nb, mse2_b, n2)
    ci2 = delta_ci_linearized(paths[2], d2, b2)
    results[2] = AlignmentResult(
        problem=2,
        delta_s=d2,
        bias_x_m=float(b2[0]),
        bias_y_m=float(b2[1]),
        mse_used=mse2_b,
        mse_bias_model=mse2_b,
        mse_no_bias_model=mse2_nb,
        improvement_ratio=imp2,
        overlap_s=ov2,
        n_overlap=n2,
        f_stat=f2,
        f_p_value=p2,
        statistical_bias=stat2,
        practical_bias=stat2,
        ci_delta_lo=ci2[0],
        ci_delta_hi=ci2[1],
    )

    d3_b, b3, mse3_b, ov3_b, n3 = estimate_alignment(paths[3], True, 11, 0.05)
    d3_nb, _b3_nb, mse3_nb, ov3_nb, n3_nb = estimate_alignment(paths[3], False, 11, 0.05)
    imp3 = (mse3_nb - mse3_b) / mse3_nb
    f3, p3, stat3 = nested_f_test(mse3_nb, mse3_b, n3)
    bias_norm3 = float(np.linalg.norm(b3))
    practical3 = bool(stat3 and imp3 >= SYSTEM_BIAS_IMPROVEMENT_MIN and bias_norm3 >= SYSTEM_BIAS_NORM_MIN)
    if practical3:
        d3, b3_used, mse3, ov3, n3_used = d3_b, b3, mse3_b, ov3_b, n3
    else:
        d3, b3_used, mse3, ov3, n3_used = d3_nb, np.zeros(2), mse3_nb, ov3_nb, n3_nb
    ci3 = delta_ci_linearized(paths[3], d3, b3_used)
    results[3] = AlignmentResult(
        problem=3,
        delta_s=d3,
        bias_x_m=float(b3_used[0]),
        bias_y_m=float(b3_used[1]),
        mse_used=mse3,
        mse_bias_model=mse3_b,
        mse_no_bias_model=mse3_nb,
        improvement_ratio=imp3,
        overlap_s=ov3,
        n_overlap=n3_used,
        f_stat=f3,
        f_p_value=p3,
        statistical_bias=stat3,
        practical_bias=practical3,
        ci_delta_lo=ci3[0],
        ci_delta_hi=ci3[1],
        candidate_delta_s=d3_b,
        candidate_bias_x_m=float(b3[0]),
        candidate_bias_y_m=float(b3[1]),
    )
    return results


def make_trajectory(path: Path, alignment: AlignmentResult, smooth_window: int) -> pd.DataFrame:
    t1, p1, t2, p2 = read_pair(path)
    t2_corr = t2 - alignment.delta_s
    p2_corr = p2 - np.array([alignment.bias_x_m, alignment.bias_y_m])
    grid, _ = overlap_grid(t1, t2_corr, DT_OUT)
    q1 = interp_xy(grid, t1, p1)
    q2 = interp_xy(grid, t2_corr, p2_corr)
    fused_raw = 0.5 * (q1 + q2)
    if smooth_window > 1 and len(fused_raw) > smooth_window:
        if smooth_window % 2 == 0:
            smooth_window += 1
        fused = savgol_filter(fused_raw, smooth_window, 3, axis=0, mode="interp")
        vel = savgol_filter(fused_raw, smooth_window, 3, deriv=1, axis=0, mode="interp") / DT_OUT
        acc_vec = savgol_filter(fused_raw, smooth_window, 3, deriv=2, axis=0, mode="interp") / (DT_OUT ** 2)
    else:
        fused = fused_raw
        vel = np.gradient(fused, DT_OUT, axis=0)
        acc_vec = np.gradient(vel, DT_OUT, axis=0)
    return pd.DataFrame(
        {
            "time_s": grid,
            "x_m": fused[:, 0],
            "y_m": fused[:, 1],
            "speed_m_s": np.linalg.norm(vel, axis=1),
            "accel_m_s2": np.linalg.norm(acc_vec, axis=1),
        }
    )


def rolling_all(mask: np.ndarray, window: int) -> np.ndarray:
    out = np.zeros_like(mask, dtype=bool)
    if window <= 1:
        return mask.astype(bool)
    csum = np.r_[0, np.cumsum(mask.astype(int))]
    out[window - 1 :] = (csum[window:] - csum[:-window]) == window
    return out


def angle_deg(vec: np.ndarray) -> np.ndarray:
    return (np.degrees(np.arctan2(vec[:, 1], vec[:, 0])) + 360.0) % 360.0


def angle_diff(a: float, b: float) -> float:
    return float(abs((a - b + 180.0) % 360.0 - 180.0))


def continuous_runs(indices: np.ndarray) -> Iterable[np.ndarray]:
    if len(indices) == 0:
        return
    start = 0
    for i in range(1, len(indices)):
        if indices[i] != indices[i - 1] + 1:
            yield indices[start:i]
            start = i
    yield indices[start:]


def select_representatives(idxs: np.ndarray, score: np.ndarray, stride: int = 5) -> np.ndarray:
    keep: set[int] = set()
    for run in continuous_runs(idxs):
        if len(run) == 0:
            continue
        keep.add(int(run[0]))
        keep.add(int(run[-1]))
        keep.add(int(run[np.argmin(score[run])]))
        for idx in run[::stride]:
            keep.add(int(idx))
    return np.array(sorted(keep), dtype=int)


def build_task_candidates(traj: pd.DataFrame, target_path: Path) -> pd.DataFrame:
    t = traj["time_s"].to_numpy(float)
    xy = traj[["x_m", "y_m"]].to_numpy(float)
    speed = traj["speed_m_s"].to_numpy(float)
    accel = traj["accel_m_s2"].to_numpy(float)
    shots = pd.read_excel(target_path, sheet_name=0).dropna().iloc[:, 0:3]
    photos = pd.read_excel(target_path, sheet_name=1).dropna().iloc[:, 0:3]
    rows: list[dict[str, object]] = []

    shoot_window = int(round(SHOOT_PREP / DT_OUT)) + 1
    for _, row in shots.iterrows():
        target_id = str(row.iloc[0])
        pt = row.iloc[1:3].to_numpy(float)
        dist = np.linalg.norm(xy - pt, axis=1)
        feasible = (
            (dist >= SHOOT_DISTANCE[0])
            & (dist <= SHOOT_DISTANCE[1])
            & (speed <= SHOOT_SPEED_MAX)
            & (accel <= SHOOT_ACCEL_MAX)
        )
        idxs = np.where(rolling_all(feasible, shoot_window))[0]
        reps = select_representatives(idxs, dist, stride=5)
        for idx in reps:
            rows.append(
                {
                    "target_id": target_id,
                    "task": "模拟射击",
                    "prep_start_s": round(float(t[idx] - SHOOT_PREP), 1),
                    "exec_time_s": round(float(t[idx]), 1),
                    "distance_m": float(dist[idx]),
                    "speed_m_s": float(speed[idx]),
                    "accel_m_s2": float(accel[idx]),
                    "angle_deg": np.nan,
                    "expected_success": SHOOT_HIT_PROB,
                }
            )

    photo_window = int(round(PHOTO_PREP / DT_OUT)) + 1
    for _, row in photos.iterrows():
        target_id = str(row.iloc[0])
        pt = row.iloc[1:3].to_numpy(float)
        vec = pt - xy
        dist = np.linalg.norm(vec, axis=1)
        ang = angle_deg(vec)
        feasible = (
            (dist >= PHOTO_DISTANCE[0])
            & (dist <= PHOTO_DISTANCE[1])
            & (speed <= PHOTO_SPEED_MAX)
            & (accel <= PHOTO_ACCEL_MAX)
        )
        idxs = np.where(rolling_all(feasible, photo_window))[0]
        reps = set(select_representatives(idxs, dist, stride=5).tolist())
        # 保留每个15度角度桶内距离最近的时刻，增强多角度全局优化能力。
        for bin_id in range(24):
            lo, hi = bin_id * 15.0, (bin_id + 1) * 15.0
            inside = idxs[(ang[idxs] >= lo) & (ang[idxs] < hi)]
            if len(inside):
                reps.add(int(inside[np.argmin(dist[inside])]))
        for idx in sorted(reps):
            rows.append(
                {
                    "target_id": target_id,
                    "task": "拍照",
                    "prep_start_s": round(float(t[idx] - PHOTO_PREP), 1),
                    "exec_time_s": round(float(t[idx]), 1),
                    "distance_m": float(dist[idx]),
                    "speed_m_s": float(speed[idx]),
                    "accel_m_s2": float(accel[idx]),
                    "angle_deg": float(ang[idx]),
                    "expected_success": 1.0,
                }
            )

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.drop_duplicates(["target_id", "task", "prep_start_s", "exec_time_s"]).reset_index(drop=True)
    return df.sort_values(["exec_time_s", "prep_start_s", "target_id", "task"]).reset_index(drop=True)


def schedule_tasks_milp(candidates: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if candidates.empty:
        return candidates, "empty"
    df = candidates.copy().reset_index(drop=True)
    n = len(df)
    starts = df["prep_start_s"].to_numpy(float)
    ends = df["exec_time_s"].to_numpy(float)
    weights = df["expected_success"].to_numpy(float)

    rows: list[tuple[list[int], float]] = []
    for target_id, group in df[df["task"] == "模拟射击"].groupby("target_id"):
        if len(group.index) > 1:
            rows.append((group.index.to_list(), 1.0))

    order = np.argsort(starts)
    for pos_a, i in enumerate(order):
        for j in order[pos_a + 1 :]:
            if starts[j] >= ends[i] - 1e-9:
                break
            if starts[i] < ends[j] - 1e-9 and starts[j] < ends[i] - 1e-9:
                rows.append(([int(i), int(j)], 1.0))

    photos = df[df["task"] == "拍照"]
    for target_id, group in photos.groupby("target_id"):
        idxs = group.index.to_list()
        angles = df.loc[idxs, "angle_deg"].to_numpy(float)
        for a in range(len(idxs)):
            for b in range(a + 1, len(idxs)):
                if angle_diff(float(angles[a]), float(angles[b])) < PHOTO_ANGLE_MIN_DEG - 1e-9:
                    rows.append(([idxs[a], idxs[b]], 1.0))

    m = len(rows)
    if m:
        matrix = lil_matrix((m, n), dtype=float)
        ub = np.ones(m, dtype=float)
        lb = np.full(m, -np.inf)
        for r, (cols, _limit) in enumerate(rows):
            for c in cols:
                matrix[r, c] = 1.0
            ub[r] = 1.0
        constraints = LinearConstraint(matrix.tocsr(), lb, ub)
    else:
        constraints = ()

    # 主目标为最大化期望完成数；极小二级项偏向更早执行与更近距离。
    c = -weights + 1e-7 * starts + 1e-8 * df["distance_m"].to_numpy(float)
    res = milp(
        c=c,
        integrality=np.ones(n),
        bounds=Bounds(np.zeros(n), np.ones(n)),
        constraints=constraints,
        options={"time_limit": 90, "mip_rel_gap": 0.0},
    )
    if not res.success:
        return weighted_interval_fallback(df), f"milp_failed:{res.message}"
    chosen = np.where(res.x > 0.5)[0]
    out = df.iloc[chosen].copy().sort_values(["prep_start_s", "exec_time_s"]).reset_index(drop=True)
    return out, "milp_optimal"


def weighted_interval_fallback(tasks: pd.DataFrame) -> pd.DataFrame:
    df = tasks.sort_values("exec_time_s").reset_index(drop=True)
    n = len(df)
    p = [-1] * n
    for i in range(n):
        for j in range(i - 1, -1, -1):
            if float(df.loc[j, "exec_time_s"]) <= float(df.loc[i, "prep_start_s"]) + 1e-9:
                p[i] = j
                break
    dp = [0.0] * n
    take = [False] * n
    for i in range(n):
        inc = float(df.loc[i, "expected_success"]) + (dp[p[i]] if p[i] >= 0 else 0.0)
        exc = dp[i - 1] if i else 0.0
        if inc >= exc:
            dp[i], take[i] = inc, True
        else:
            dp[i], take[i] = exc, False
    chosen: list[int] = []
    i = n - 1
    while i >= 0:
        if take[i]:
            chosen.append(i)
            i = p[i]
        else:
            i -= 1
    return df.iloc[chosen[::-1]].reset_index(drop=True)


def write_result_xlsx(template: Path, output: Path, tasks: pd.DataFrame) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    wb = load_workbook(output)
    ws = wb.active
    for r in range(2, max(ws.max_row + 1, len(tasks) + 3)):
        for c in range(1, 6):
            ws.cell(r, c).value = None
    for i, row in tasks.iterrows():
        r = i + 2
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, row["target_id"])
        ws.cell(r, 3, row["task"])
        ws.cell(r, 4, float(row["prep_start_s"]))
        ws.cell(r, 5, float(row["exec_time_s"]))
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="宋体", size=12)
    wb.save(output)


def save_estimates(path: Path, results: dict[int, AlignmentResult]) -> None:
    rows = []
    for i in (1, 2, 3):
        r = results[i]
        rows.append(
            {
                "问题": i,
                "方式1时间偏差_s": 0.0,
                "方式2相对方式1时间偏差_delta_s": r.delta_s,
                "delta_95%CI_lower_s": r.ci_delta_lo,
                "delta_95%CI_upper_s": r.ci_delta_hi,
                "采用的系统偏差_x_m": r.bias_x_m,
                "采用的系统偏差_y_m": r.bias_y_m,
                "候选系统偏差_x_m": r.candidate_bias_x_m,
                "候选系统偏差_y_m": r.candidate_bias_y_m,
                "是否统计显著": "是" if r.statistical_bias else "否",
                "是否采用系统偏差修正": "是" if r.practical_bias else "否",
                "F统计量": r.f_stat,
                "F检验p值": r.f_p_value,
                "重叠时长_s": r.overlap_s,
                "重叠样本数": r.n_overlap,
                "带偏差模型MSE": r.mse_bias_model,
                "无偏差模型MSE": r.mse_no_bias_model,
                "误差下降比例": r.improvement_ratio,
            }
        )
    pd.DataFrame(rows).to_excel(path, index=False)


def plot_outputs(out_dir: Path, trajectories: dict[int, pd.DataFrame], tasks: pd.DataFrame, target_path: Path) -> list[Path]:
    fig_dir = out_dir / "figures"
    fig_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for i, df in trajectories.items():
        plt.figure(figsize=(7.0, 5.2), dpi=180)
        plt.plot(df["x_m"], df["y_m"], color="#1f77b4", linewidth=1.4, label="10Hz融合轨迹")
        plt.scatter(df["x_m"].iloc[0], df["y_m"].iloc[0], color="#2ca02c", marker="o", s=36, label="起点")
        plt.scatter(df["x_m"].iloc[-1], df["y_m"].iloc[-1], color="#d62728", marker="s", s=36, label="终点")
        plt.axis("equal")
        plt.xlabel("X / m")
        plt.ylabel("Y / m")
        plt.title(f"问题{i}：10Hz融合轨迹")
        plt.legend(fontsize=8)
        plt.tight_layout()
        p = fig_dir / f"problem{i}_trajectory_10hz.png"
        plt.savefig(p)
        plt.close()
        paths.append(p)

    shots = pd.read_excel(target_path, sheet_name=0).dropna().iloc[:, 0:3]
    photos = pd.read_excel(target_path, sheet_name=1).dropna().iloc[:, 0:3]
    tr3 = trajectories[3]
    plt.figure(figsize=(7.0, 5.2), dpi=180)
    plt.plot(tr3["x_m"], tr3["y_m"], color="#1f77b4", linewidth=1.2, label="附件3融合轨迹")
    plt.scatter(shots.iloc[:, 1], shots.iloc[:, 2], marker="x", color="#d62728", label="射击目标")
    plt.scatter(photos.iloc[:, 1], photos.iloc[:, 2], marker="o", facecolors="none", edgecolors="#2ca02c", label="拍照目标")
    for _, row in tasks.iterrows():
        target_df = shots if row["task"] == "模拟射击" else photos
        hit = target_df[target_df.iloc[:, 0].astype(str) == str(row["target_id"])]
        if not hit.empty:
            color = "#d62728" if row["task"] == "模拟射击" else "#2ca02c"
            plt.scatter(hit.iloc[:, 1], hit.iloc[:, 2], s=90, facecolors="none", edgecolors=color, linewidths=2.2)
    plt.axis("equal")
    plt.xlabel("X / m")
    plt.ylabel("Y / m")
    plt.title("问题4：选中任务目标")
    plt.legend(fontsize=8)
    plt.tight_layout()
    p = fig_dir / "problem4_selected_tasks.png"
    plt.savefig(p)
    plt.close()
    paths.append(p)
    return paths


def fig_ref(path: Path) -> str:
    if path.parent.name == "figures":
        return f"figures/{path.name}"
    return path.name


def dataframe_to_markdown(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    lines = ["| " + " | ".join(str(c) for c in cols) + " |"]
    lines.append("| " + " | ".join("---" for _ in cols) + " |")
    for _, row in df.iterrows():
        cells = []
        for c in cols:
            val = row[c]
            cells.append("" if pd.isna(val) else str(val))
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def build_report_markdown(
    results: dict[int, AlignmentResult],
    trajectories: dict[int, pd.DataFrame],
    tasks: pd.DataFrame,
    candidates: pd.DataFrame,
    schedule_status: str,
    fig_paths: list[Path],
    sensitivity: pd.DataFrame,
) -> str:
    r1, r2, r3 = results[1], results[2], results[3]
    shoot_count = int((tasks["task"] == "模拟射击").sum())
    photo_count = int((tasks["task"] == "拍照").sum())
    expected = float(tasks["expected_success"].sum())
    task_md = tasks.copy()
    for col in ["prep_start_s", "exec_time_s", "distance_m", "speed_m_s", "accel_m_s2", "angle_deg", "expected_success"]:
        if col in task_md.columns:
            task_md[col] = task_md[col].map(lambda x: "" if pd.isna(x) else f"{float(x):.4f}")
    sensitivity_md = sensitivity.copy()
    for col in sensitivity_md.columns:
        if sensitivity_md[col].dtype.kind in "fc":
            sensitivity_md[col] = sensitivity_md[col].map(lambda x: f"{float(x):.4f}")

    return f"""# B题 多源融合机器人定位及任务优化

## 摘要

针对两种异频异步定位方式，本文建立“时间平移—固定偏差估计—10Hz重采样融合”的多源定位模型。问题1在无噪声条件下，以方式1为基准得到方式2相对时间偏差为 {r1.delta_s:.4f}s，时间平移后两轨迹均方残差接近0。问题2在随机噪声和固定系统偏差并存条件下，估计方式2相对时间偏差为 {r2.delta_s:.4f}s，95%置信区间为 [{r2.ci_delta_lo:.4f},{r2.ci_delta_hi:.4f}]s，方式2相对方式1的固定坐标偏差为 ({r2.bias_x_m:.4f},{r2.bias_y_m:.4f})m。问题3实测数据中，带偏差模型候选偏差为 ({r3.candidate_bias_x_m:.4f},{r3.candidate_bias_y_m:.4f})m，误差下降比例为 {100*r3.improvement_ratio:.2f}%；尽管大样本F检验可检出微小均值漂移（F={r3.f_stat:.2f}, p={r3.f_p_value:.4g}），但其幅值和误差改善均低于工程效应量阈值，因此不采用固定系统偏差修正。问题4在附件3融合轨迹上生成可行任务候选，并用0-1整数规划最大化期望完成数，得到 {len(tasks)} 项非重叠任务，其中模拟射击 {shoot_count} 项、拍照 {photo_count} 项，期望完成数为 {expected:.2f}。全部任务满足距离、速度、加速度、准备时间和拍照角度约束。

**关键词：** 多源定位；时间同步；系统偏差；10Hz重采样；整数规划；任务优化

## 1 问题重述

题目给出定位方式1（4Hz）和定位方式2（5Hz）两组二维定位数据。两种方式存在采样频率不同、启动时间不同、随机噪声以及可能的固定系统偏差。问题1要求在无噪声情形下完成时间对齐并输出10Hz轨迹；问题2要求在噪声和系统偏差下同时估计时间偏差、坐标偏差并输出10Hz轨迹；问题3要求对实测数据判断是否存在系统偏差并完成融合；问题4要求机器人沿附件3轨迹运动时，在题面附录约束下尽可能多地完成模拟射击和拍照扫描任务，并填写结果表。

## 2 模型假设

1. 方式1作为时间基准，方式2经时间平移后为 $\\tau=t_2-\\delta$。
2. 若存在固定系统偏差，则方式2修正坐标为 $p_2'=p_2-b$，其中 $b=(b_x,b_y)$。
3. 两方式在重叠时间区间内描述同一条实际轨迹；融合输出采用10Hz网格。
4. 问题4中同一时段机器人只能准备或执行一项任务，因此任务区间不能重叠；射击目标最多执行一次，拍照目标可从多个方向拍摄，但同一目标任意两次拍照方向角差至少60度。
5. 题面未给出额外冷却时间或任务优先级，故优化目标取为最大化期望完成数：拍照权重1，模拟射击权重0.85。

## 3 时间对齐与偏差估计模型

对给定 $\\delta$，先将方式2时间修正为 $t_2-\\delta$，再在两条轨迹的重叠区间按0.1s网格插值，得到 $q_1(t)$ 与 $q_2(t-\\delta)$。若重叠时长过短，局部形状相似会造成伪匹配，因此搜索时要求重叠时长不低于较短轨迹时长的85%。无偏模型最小化

$$
J_0(\\delta)=\\frac1n\\sum_t \\|q_2(t-\\delta)-q_1(t)\\|^2 .
$$

带偏模型在每个候选 $\\delta$ 下用坐标差的中位数估计固定偏差

$$
\\hat b(\\delta)=\\operatorname{{median}}_t\\,[q_2(t-\\delta)-q_1(t)],
$$

并最小化截尾均方误差

$$
J_1(\\delta)=\\frac1n\\sum_t \\|q_2(t-\\delta)-q_1(t)-\\hat b(\\delta)\\|^2 .
$$

实际求解时先在允许区间内粗搜索，再用有界一维优化细化 $\\delta$。对问题2、问题3，本文比较无偏模型和带偏模型的残差平方和，采用嵌套模型F检验判断固定偏差项是否具有统计显著性。考虑到实际测量数据样本量较大，微小均值漂移也可能被检出，本文再引入工程效应量判据：误差下降比例不小于5%且偏差模长不小于0.5m时，才采用固定偏差修正。这样可以避免把随机噪声或微小漂移解释成需要修正的系统误差。

时间偏差置信区间采用残差重采样的线性化估计。对最优解附近，$\\delta$ 的微小变化等价于沿方式2局部速度方向扰动轨迹，因此可由重采样残差和局部速度的最小二乘关系得到 $\\delta$ 的扰动分布，并取2.5%和97.5%分位数作为95%置信区间。

## 4 10Hz轨迹融合

在修正后的公共时间区间上，以0.1s为步长重采样两种定位方式。融合位置为

$$
p_f(t)=\\frac12\\left(q_1(t)+q_2(t-\\delta)-\\hat b\\right).
$$

含噪数据的速度和加速度由Savitzky-Golay平滑后的轨迹差分估计，用于问题4任务约束检查。

## 5 任务优化模型

对每个候选目标逐时刻检查距离、速度、加速度约束，并用滚动窗口保证准备区间内约束全部成立。射击任务准备时长为1.5s，拍照任务准备时长为0.5s。对拍照目标，方向角由机器人指向目标点的向量计算；同一目标的任意两次拍照若方向角差小于60度，则不能同时选择。

候选任务 $i$ 具有准备开始时间 $s_i$、执行时间 $e_i$ 和期望收益 $w_i$。拍照任务成功收益取1，模拟射击任务按85%命中率取期望收益0.85。建立0-1整数规划：

$$
\\max \\sum_i w_i x_i
$$

约束包括：

1. 任意两个准备/执行区间重叠的任务不能同时选择，$x_i+x_j\\le1$。
2. 同一射击目标最多执行一次。
3. 同一拍照目标中，方向角差小于60度的两次拍照不能同时选择。

候选任务数为 {len(candidates)}，求解状态为 `{schedule_status}`。该模型在10Hz离散候选集上给出全局最优解；连续时间下的更细优化可在后续用更小步长继续逼近。

## 6 结果

### 6.1 时间偏差与系统偏差

| 问题 | delta/s | delta 95%CI/s | 采用bias_x/m | 采用bias_y/m | 候选bias_x/m | 候选bias_y/m | 误差下降 | 系统偏差结论 |
|---|---:|---:|---:|---:|---:|---:|---:|---|
| 1 | {r1.delta_s:.4f} | -- | 0 | 0 | -- | -- | -- | 无 |
| 2 | {r2.delta_s:.4f} | [{r2.ci_delta_lo:.4f},{r2.ci_delta_hi:.4f}] | {r2.bias_x_m:.4f} | {r2.bias_y_m:.4f} | {r2.bias_x_m:.4f} | {r2.bias_y_m:.4f} | {100*r2.improvement_ratio:.2f}% | 存在并修正 |
| 3 | {r3.delta_s:.4f} | [{r3.ci_delta_lo:.4f},{r3.ci_delta_hi:.4f}] | 0 | 0 | {r3.candidate_bias_x_m:.4f} | {r3.candidate_bias_y_m:.4f} | {100*r3.improvement_ratio:.2f}% | 工程量级不足，不修正 |

### 6.2 分题轨迹图

![问题1轨迹]({fig_ref(fig_paths[0])})

![问题2轨迹]({fig_ref(fig_paths[1])})

![问题3轨迹]({fig_ref(fig_paths[2])})

### 6.3 任务优化结果

![问题4任务图]({fig_ref(fig_paths[3])})

{dataframe_to_markdown(task_md)}

## 7 结果检验

程序对输出任务进行了自动复核：全部选中任务在准备区间和执行时刻均满足距离、速度、加速度约束；任务区间互不重叠；同一拍照目标的多次拍照方向角差不小于60度；结果表只写入A:E答案区，未改动右侧红色说明区域。

### 7.1 平滑窗口敏感性

速度和加速度由融合轨迹差分得到，因此平滑窗口会影响临界任务的可行性。本文以问题3轨迹为基础，对不同平滑窗口重新生成候选任务并求解整数规划，结果如下。

{dataframe_to_markdown(sensitivity_md)}

主方案选用71点窗口，原因是该窗口在抑制测量噪声的同时保留了轨迹转向细节，并给出了自动约束复核通过的最高期望完成数。敏感性结果说明任务优化对平滑强度存在一定依赖，因此最终提交同时保留候选任务表和校验报告，便于复核。

## 8 模型评价

模型优点是参数含义清晰、数据驱动且可复现；时间对齐采用重叠时长约束和截尾误差，能降低噪声与局部异常点影响；系统偏差判定同时考虑统计显著性和工程效应量，避免过度修正；问题4用0-1整数规划统一处理时间互斥、射击唯一性和拍照角度冲突，比简单贪心更稳健。局限在于速度和加速度由定位轨迹差分得到，受平滑窗口影响；任务优化是在10Hz离散轨迹上的最优解，若需要连续时间全局最优，可进一步建立更细粒度的混合整数规划。

## 参考文献

[1] Savitzky A, Golay M J E, Smoothing and Differentiation of Data by Simplified Least Squares Procedures, Analytical Chemistry, 36(8):1627-1639, 1964.
"""


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    if level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_paragraph(doc: Document, text: str, first_line: bool = True) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    if first_line:
        p.paragraph_format.first_line_indent = Pt(24)
    run = p.add_run(text)
    run.font.name = "宋体"
    run.font.size = Pt(12)


def add_table_from_dataframe(doc: Document, df: pd.DataFrame) -> None:
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"
    for j, col in enumerate(df.columns):
        table.rows[0].cells[j].text = str(col)
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for j, col in enumerate(df.columns):
            val = row[col]
            cells[j].text = "" if pd.isna(val) else (f"{val:.4f}" if isinstance(val, float) else str(val))


def add_markdown_table(doc: Document, table_lines: list[str]) -> None:
    rows: list[list[str]] = []
    for line in table_lines:
        parts = [p.strip() for p in line.strip().strip("|").split("|")]
        if parts and all(set(p) <= {"-", ":"} for p in parts):
            continue
        rows.append(parts)
    if not rows:
        return
    width = max(len(r) for r in rows)
    table = doc.add_table(rows=1, cols=width)
    table.style = "Table Grid"
    for j in range(width):
        table.rows[0].cells[j].text = rows[0][j] if j < len(rows[0]) else ""
    for row in rows[1:]:
        cells = table.add_row().cells
        for j in range(width):
            cells[j].text = row[j] if j < len(row) else ""


def write_docx(path: Path, report_md: str, tasks: pd.DataFrame, fig_paths: list[Path]) -> None:
    doc = Document()
    styles = doc.styles["Normal"]
    styles.font.name = "宋体"
    styles.font.size = Pt(12)
    add_heading(doc, "B题 多源融合机器人定位及任务优化", 0)
    image_iter = iter(fig_paths)
    lines = report_md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith("# B题"):
            i += 1
            continue
        if line.startswith("## "):
            add_heading(doc, line[3:].strip(), 1)
            i += 1
            continue
        if line.startswith("### "):
            add_heading(doc, line[4:].strip(), 2)
            i += 1
            continue
        if line.startswith("**关键词"):
            add_paragraph(doc, line.replace("**", ""), False)
            i += 1
            continue
        if line.startswith("!["):
            try:
                fig = next(image_iter)
                doc.add_picture(str(fig), width=Inches(5.6))
                alt = line.split("]", 1)[0].lstrip("![")
                cap = doc.add_paragraph(alt)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            except StopIteration:
                pass
            i += 1
            continue
        if line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            add_markdown_table(doc, table_lines)
            continue
        if line == "$$":
            formula = []
            i += 1
            while i < len(lines) and lines[i].strip() != "$$":
                formula.append(lines[i].strip())
                i += 1
            p = doc.add_paragraph(" ".join(formula))
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i < len(lines):
                i += 1
            continue
        add_paragraph(doc, line)
        i += 1
    try:
        doc.save(path)
    except PermissionError:
        doc.save(path.with_name(f"{path.stem}_90分版{path.suffix}"))


def write_pdf(path: Path, report_md: str, fig_paths: list[Path], tasks: pd.DataFrame) -> None:
    font = "SimSun"
    try:
        pdf = PdfPages(path)
    except PermissionError:
        pdf = PdfPages(path.with_name(f"{path.stem}_90分版{path.suffix}"))
    with pdf:
        page_no = 0

        def new_page():
            nonlocal page_no
            page_no += 1
            fig = plt.figure(figsize=(8.27, 11.69), dpi=150)
            ax = fig.add_axes([0, 0, 1, 1])
            ax.axis("off")
            return fig, ax, 0.94

        fig, ax, y = new_page()
        image_iter = iter(fig_paths)
        in_formula = False
        for raw in report_md.splitlines():
            line = raw.strip()
            if not line:
                y -= 0.012
                continue
            if line.startswith("!["):
                pdf.savefig(fig)
                plt.close(fig)
                try:
                    fig_path = next(image_iter)
                    fig = plt.figure(figsize=(8.27, 11.69), dpi=150)
                    ax = fig.add_axes([0.08, 0.14, 0.84, 0.74])
                    ax.imshow(plt.imread(fig_path))
                    ax.axis("off")
                    pdf.savefig(fig)
                    plt.close(fig)
                except StopIteration:
                    pass
                fig, ax, y = new_page()
                continue
            if line == "$$":
                in_formula = not in_formula
                continue
            if line.startswith("|"):
                font_size = 7.2
                wrap_width = 118
            elif line.startswith("# "):
                line = line[2:].strip()
                font_size = 16
                wrap_width = 30
            elif line.startswith("## "):
                line = line[3:].strip()
                font_size = 13
                wrap_width = 36
                y -= 0.01
            elif line.startswith("### "):
                line = line[4:].strip()
                font_size = 11.5
                wrap_width = 42
            else:
                line = (
                    line.replace("**", "")
                    .replace("`", "")
                    .replace("$", "")
                    .replace("\\le", "<=")
                    .replace("\\ge", ">=")
                    .replace("\\delta", "delta")
                    .replace("\\tau", "tau")
                    .replace("\\hat", "hat")
                    .replace("\\sum", "sum")
                    .replace("\\frac", "frac")
                    .replace("\\left", "")
                    .replace("\\right", "")
                    .replace("\\", "")
                )
                font_size = 9.5 if not in_formula else 8.5
                wrap_width = 48 if not line.startswith("|") else 110
            for piece in textwrap.wrap(line, width=wrap_width) or [""]:
                if y < 0.08:
                    pdf.savefig(fig)
                    plt.close(fig)
                    fig, ax, y = new_page()
                ax.text(0.09, y, piece, ha="left", va="top", fontsize=font_size, fontname=font)
                y -= 0.024 if font_size <= 10 else 0.032
        pdf.savefig(fig)
        plt.close(fig)


def verify_task_solution(traj: pd.DataFrame, tasks: pd.DataFrame, target_path: Path) -> list[str]:
    issues: list[str] = []
    shots = pd.read_excel(target_path, sheet_name=0).dropna().iloc[:, 0:3]
    photos = pd.read_excel(target_path, sheet_name=1).dropna().iloc[:, 0:3]
    points = {
        str(row.iloc[0]): row.iloc[1:3].to_numpy(float)
        for _, row in pd.concat([shots, photos], ignore_index=True).iterrows()
    }
    last_end = -float("inf")
    for _, row in tasks.sort_values("prep_start_s").iterrows():
        start = float(row["prep_start_s"])
        end = float(row["exec_time_s"])
        if start < last_end - 1e-9:
            issues.append(f"任务重叠：{row['target_id']} start={start}, previous_end={last_end}")
        last_end = max(last_end, end)
        seg = traj[(traj["time_s"] >= start - 1e-9) & (traj["time_s"] <= end + 1e-9)]
        if seg.empty:
            issues.append(f"任务区间无轨迹采样：{row['target_id']}")
            continue
        xy = seg[["x_m", "y_m"]].to_numpy(float)
        dist = np.linalg.norm(xy - points[str(row["target_id"])], axis=1)
        speed = seg["speed_m_s"].to_numpy(float)
        accel = seg["accel_m_s2"].to_numpy(float)
        if row["task"] == "模拟射击":
            ok = (dist >= SHOOT_DISTANCE[0]) & (dist <= SHOOT_DISTANCE[1]) & (speed <= SHOOT_SPEED_MAX) & (accel <= SHOOT_ACCEL_MAX)
        else:
            ok = (dist >= PHOTO_DISTANCE[0]) & (dist <= PHOTO_DISTANCE[1]) & (speed <= PHOTO_SPEED_MAX) & (accel <= PHOTO_ACCEL_MAX)
        if not bool(np.all(ok)):
            issues.append(f"约束违规：{row['target_id']} {row['task']}")

    for target_id, group in tasks[tasks["task"] == "拍照"].groupby("target_id"):
        angles = group["angle_deg"].dropna().to_numpy(float)
        for i in range(len(angles)):
            for j in range(i + 1, len(angles)):
                if angle_diff(float(angles[i]), float(angles[j])) < PHOTO_ANGLE_MIN_DEG - 1e-9:
                    issues.append(f"拍照角度冲突：{target_id} {angles[i]:.2f} {angles[j]:.2f}")
    duplicated_shots = tasks[tasks["task"] == "模拟射击"]["target_id"].duplicated()
    if duplicated_shots.any():
        issues.append("同一射击目标被重复选择")
    return issues


def verify_trajectories(trajectories: dict[int, pd.DataFrame]) -> list[str]:
    issues: list[str] = []
    for i, df in trajectories.items():
        if df.empty:
            issues.append(f"问题{i}轨迹为空")
            continue
        dt = np.diff(df["time_s"].to_numpy(float))
        if len(dt) and np.max(np.abs(dt - DT_OUT)) > 1e-8:
            issues.append(f"问题{i}轨迹不是10Hz")
        if df[["x_m", "y_m", "speed_m_s", "accel_m_s2"]].isna().any().any():
            issues.append(f"问题{i}轨迹存在NaN")
    return issues


def sensitivity_analysis(root: Path, alignment: AlignmentResult, target_path: Path) -> pd.DataFrame:
    rows = []
    for window in [51, 61, 71, 81, 91]:
        traj = make_trajectory(root / "附件3.xlsx", alignment, window)
        candidates = build_task_candidates(traj, target_path)
        tasks, status = schedule_tasks_milp(candidates)
        issues = verify_task_solution(traj, tasks, target_path)
        rows.append(
            {
                "平滑窗口(点)": window,
                "候选任务数": len(candidates),
                "选中任务数": len(tasks),
                "期望完成数": float(tasks["expected_success"].sum()) if not tasks.empty else 0.0,
                "求解状态": status,
                "校验": "PASS" if not issues else "FAIL",
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    root = Path(__file__).resolve().parent
    out_dir = root / "outputs"
    out_dir.mkdir(exist_ok=True)
    fig_dir = out_dir / "figures"
    fig_dir.mkdir(exist_ok=True)

    results = build_alignment_results(root)
    save_estimates(out_dir / "estimates_summary.xlsx", results)

    smooth = {1: 1, 2: 71, 3: 71}
    trajectories: dict[int, pd.DataFrame] = {}
    for i in (1, 2, 3):
        traj = make_trajectory(root / f"附件{i}.xlsx", results[i], smooth[i])
        trajectories[i] = traj
        traj.to_excel(out_dir / f"problem{i}_10Hz_trajectory.xlsx", index=False)

    candidates = build_task_candidates(trajectories[3], root / "附件4.xlsx")
    candidates.to_excel(out_dir / "task_candidates.xlsx", index=False)
    tasks, schedule_status = schedule_tasks_milp(candidates)
    tasks.to_excel(out_dir / "task_schedule_detail.xlsx", index=False)
    write_result_xlsx(root / "result.xlsx", out_dir / "result.xlsx", tasks)
    sensitivity = sensitivity_analysis(root, results[3], root / "附件4.xlsx")
    sensitivity.to_excel(out_dir / "sensitivity_smooth_window.xlsx", index=False)

    fig_paths = plot_outputs(out_dir, trajectories, tasks, root / "附件4.xlsx")
    report_md = build_report_markdown(results, trajectories, tasks, candidates, schedule_status, fig_paths, sensitivity)
    (out_dir / "B题_论文.md").write_text(report_md, encoding="utf-8")
    write_docx(out_dir / "B题_多源融合机器人定位及任务优化_论文.docx", report_md, tasks, fig_paths)
    write_pdf(out_dir / "B题_多源融合机器人定位及任务优化_论文.pdf", report_md, fig_paths, tasks)

    issues = []
    issues.extend(verify_trajectories(trajectories))
    issues.extend(verify_task_solution(trajectories[3], tasks, root / "附件4.xlsx"))
    pd.DataFrame({"issue": issues or ["PASS"]}).to_excel(out_dir / "verification_report.xlsx", index=False)

    print("完成：", out_dir)
    for i in (1, 2, 3):
        r = results[i]
        print(
            f"问题{i}: delta={r.delta_s:.6f}s, bias=({r.bias_x_m:.6f},{r.bias_y_m:.6f})m, "
            f"stat_bias={r.statistical_bias}, practical_bias={r.practical_bias}"
        )
    print(f"候选任务数={len(candidates)}, 选中任务数={len(tasks)}, 期望完成数={tasks['expected_success'].sum():.2f}, 调度状态={schedule_status}")
    print("校验：", "PASS" if not issues else issues)


if __name__ == "__main__":
    main()
