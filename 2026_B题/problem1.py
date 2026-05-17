# -*- coding: utf-8 -*-
"""问题1及前三问共用的轨迹配准工具。"""
from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
from scipy.signal import savgol_filter
from scipy.stats import f as f_dist


DT = 0.1
MIN_OVERLAP_RATIO = 0.85
BIAS_ALPHA = 0.05


@dataclass
class AlignmentResult:
    """保存一次时间同步和偏差估计的结果。"""

    delta_s: float
    bias_x_m: float = 0.0
    bias_y_m: float = 0.0
    mse: float = 0.0
    overlap_s: float = 0.0
    n_overlap: int = 0
    f_stat: float | None = None
    p_value: float | None = None
    ci_low_s: float | None = None
    ci_high_s: float | None = None
    adopt_bias: bool = False


def project_root() -> Path:
    """返回本题代码所在目录。"""
    return Path(__file__).resolve().parent


def output_dir(root: Path | None = None) -> Path:
    """创建并返回输出目录。"""
    out = (root or project_root()) / "outputs"
    out.mkdir(exist_ok=True)
    return out


def read_pair(path: Path) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """读取一个附件中的两种定位方式数据。"""
    d1 = pd.read_excel(path, sheet_name=0).dropna()
    d2 = pd.read_excel(path, sheet_name=1).dropna()
    a1 = d1.iloc[:, 0:3].to_numpy(float)
    a2 = d2.iloc[:, 0:3].to_numpy(float)
    return a1[:, 0], a1[:, 1:3], a2[:, 0], a2[:, 1:3]


def interp_xy(t_new: np.ndarray, t: np.ndarray, xy: np.ndarray) -> np.ndarray:
    """把二维坐标线性插值到指定时间网格。"""
    return np.column_stack([np.interp(t_new, t, xy[:, i]) for i in range(2)])


def moving_average(xy: np.ndarray, window: int) -> np.ndarray:
    """用于配准目标函数的小窗口平滑。"""
    if window <= 1:
        return xy.copy()
    if window % 2 == 0:
        window += 1
    pad = window // 2
    kernel = np.ones(window) / window
    padded = np.pad(xy, ((pad, pad), (0, 0)), mode="edge")
    return np.column_stack([np.convolve(padded[:, i], kernel, mode="valid") for i in range(2)])


def overlap_grid(t1: np.ndarray, t2_corr: np.ndarray, dt: float = DT) -> tuple[np.ndarray, float]:
    """计算两条轨迹在当前时间偏差下的公共10Hz时间网格。"""
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
    min_overlap_ratio: float = MIN_OVERLAP_RATIO,
) -> tuple[float, np.ndarray, float, int]:
    """计算给定时间偏差下的匹配误差和系统偏差估计。"""
    t2_corr = t2 - delta
    grid, overlap = overlap_grid(t1, t2_corr, dt)
    min_duration = min(float(t1.max() - t1.min()), float(t2.max() - t2.min()))
    if len(grid) < 8 or overlap < min_overlap_ratio * min_duration:
        return float("inf"), np.zeros(2), overlap, len(grid)

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
) -> AlignmentResult:
    """先粗搜索再一维优化，估计方式2相对方式1的时间偏差。"""
    t1, p1_raw, t2, p2_raw = read_pair(path)
    p1 = moving_average(p1_raw, score_smooth_window)
    p2 = moving_average(p2_raw, score_smooth_window)

    min_duration = min(float(t1.max() - t1.min()), float(t2.max() - t2.min()))
    d_min = float(t2.min() - t1.max() + MIN_OVERLAP_RATIO * min_duration)
    d_max = float(t2.max() - t1.min() - MIN_OVERLAP_RATIO * min_duration)
    deltas = np.linspace(d_min, d_max, 3000)
    scores = np.array(
        [delta_score(t1, p1, t2, p2, d, correct_bias, 0.5, trim_ratio)[0] for d in deltas]
    )
    best = float(deltas[int(np.nanargmin(scores))])
    step = float(deltas[1] - deltas[0]) if len(deltas) > 1 else 1.0
    lo = max(d_min, best - 30 * step)
    hi = min(d_max, best + 30 * step)

    opt = minimize_scalar(
        lambda d: delta_score(t1, p1, t2, p2, d, correct_bias, DT, trim_ratio)[0],
        bounds=(lo, hi),
        method="bounded",
        options={"xatol": 1e-8},
    )
    mse, bias, overlap, n = delta_score(t1, p1, t2, p2, float(opt.x), correct_bias, DT, trim_ratio)
    return AlignmentResult(float(opt.x), float(bias[0]), float(bias[1]), mse, overlap, n)


def nested_f_test(mse0: float, mse1: float, n: int) -> tuple[float, float, bool]:
    """用嵌套F检验判断系统偏差项是否显著降低残差。"""
    if mse1 <= 0 or mse0 <= mse1 or n <= 3:
        return 1.0, 1.0, False
    sse0 = mse0 * n
    sse1 = mse1 * n
    f_stat = ((sse0 - sse1) / 2.0) / (sse1 / (n - 2))
    p_value = float(1.0 - f_dist.cdf(f_stat, 2, n - 2))
    return float(f_stat), p_value, p_value < BIAS_ALPHA


def delta_ci_linearized(path: Path, delta: float, bias: np.ndarray, n_boot: int = 399) -> tuple[float, float]:
    """用残差重采样给出时间偏差的近似95%置信区间。"""
    t1, p1, t2, p2 = read_pair(path)
    t2_corr = t2 - delta
    p2_corr = p2 - bias
    grid, _ = overlap_grid(t1, t2_corr, DT)
    if len(grid) < 20:
        return delta, delta
    q1 = interp_xy(grid, t1, p1)
    q2 = interp_xy(grid, t2_corr, p2_corr)
    residual = q2 - q1
    velocity = np.gradient(q2, DT, axis=0)
    rng = np.random.default_rng(20260515 + int(abs(delta) * 10))
    estimates: list[float] = []
    for _ in range(n_boot):
        idx = rng.integers(0, len(residual), len(residual))
        r = residual[idx] - np.median(residual[idx], axis=0)
        v = velocity[idx] - np.median(velocity[idx], axis=0)
        denom = float(np.sum(v * v))
        if denom > 1e-12:
            estimates.append(delta - float(np.sum(v * r)) / denom)
    if len(estimates) < 10:
        return delta, delta
    lo, hi = np.percentile(estimates, [2.5, 97.5])
    return float(lo), float(hi)


def make_trajectory(path: Path, alignment: AlignmentResult, smooth_window: int) -> pd.DataFrame:
    """按0.1s网格输出融合轨迹，并估计速度和加速度。"""
    t1, p1, t2, p2 = read_pair(path)
    t2_corr = t2 - alignment.delta_s
    p2_corr = p2 - np.array([alignment.bias_x_m, alignment.bias_y_m])
    grid, _ = overlap_grid(t1, t2_corr, DT)
    q1 = interp_xy(grid, t1, p1)
    q2 = interp_xy(grid, t2_corr, p2_corr)
    fused_raw = 0.5 * (q1 + q2)

    if smooth_window > 1 and len(fused_raw) > smooth_window:
        if smooth_window % 2 == 0:
            smooth_window += 1
        fused = savgol_filter(fused_raw, smooth_window, 3, axis=0, mode="interp")
        vel = savgol_filter(fused_raw, smooth_window, 3, deriv=1, axis=0, mode="interp") / DT
        acc = savgol_filter(fused_raw, smooth_window, 3, deriv=2, axis=0, mode="interp") / (DT**2)
    else:
        fused = fused_raw
        vel = np.gradient(fused, DT, axis=0)
        acc = np.gradient(vel, DT, axis=0)

    return pd.DataFrame(
        {
            "time_s": grid,
            "x_m": fused[:, 0],
            "y_m": fused[:, 1],
            "speed_m_s": np.linalg.norm(vel, axis=1),
            "accel_m_s2": np.linalg.norm(acc, axis=1),
        }
    )


def solve_problem1(root: Path | None = None, write_outputs: bool = True) -> tuple[AlignmentResult, pd.DataFrame]:
    """求解问题1：无噪声条件下的时间同步和10Hz轨迹输出。"""
    root = root or project_root()
    alignment = estimate_alignment(root / "附件1.xlsx", correct_bias=False, score_smooth_window=1, trim_ratio=0.0)
    trajectory = make_trajectory(root / "附件1.xlsx", alignment, smooth_window=1)
    if write_outputs:
        out = output_dir(root)
        trajectory.to_excel(out / "problem1_10Hz_trajectory.xlsx", index=False)
        pd.DataFrame(
            [
                {
                    "problem": 1,
                    "delta_s": alignment.delta_s,
                    "overlap_s": alignment.overlap_s,
                    "n_overlap": alignment.n_overlap,
                    "bias_x_m": 0.0,
                    "bias_y_m": 0.0,
                    "mse": alignment.mse,
                }
            ]
        ).to_excel(out / "problem1_estimate.xlsx", index=False)
    return alignment, trajectory


def main() -> None:
    alignment, _trajectory = solve_problem1()
    print(f"problem1 delta = {alignment.delta_s:.6f} s")
    print(f"overlap = {alignment.overlap_s:.3f} s, samples = {alignment.n_overlap}")


if __name__ == "__main__":
    main()
