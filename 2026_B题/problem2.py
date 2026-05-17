# -*- coding: utf-8 -*-
"""问题2：含随机噪声和固定系统偏差的数据融合。"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from problem1 import (
    AlignmentResult,
    delta_ci_linearized,
    estimate_alignment,
    make_trajectory,
    nested_f_test,
    output_dir,
    project_root,
)


def solve_problem2(root: Path | None = None, write_outputs: bool = True) -> tuple[AlignmentResult, pd.DataFrame]:
    """估计时间偏差和固定系统偏差，输出修正后的10Hz融合轨迹。"""
    root = root or project_root()
    path = root / "附件2.xlsx"

    bias_model = estimate_alignment(path, correct_bias=True, score_smooth_window=9, trim_ratio=0.02)
    no_bias = estimate_alignment(path, correct_bias=False, score_smooth_window=9, trim_ratio=0.02)
    f_stat, p_value, significant = nested_f_test(no_bias.mse, bias_model.mse, bias_model.n_overlap)
    ci_low, ci_high = delta_ci_linearized(
        path, bias_model.delta_s, np.array([bias_model.bias_x_m, bias_model.bias_y_m])
    )

    alignment = AlignmentResult(
        delta_s=bias_model.delta_s,
        bias_x_m=bias_model.bias_x_m,
        bias_y_m=bias_model.bias_y_m,
        mse=bias_model.mse,
        overlap_s=bias_model.overlap_s,
        n_overlap=bias_model.n_overlap,
        f_stat=f_stat,
        p_value=p_value,
        ci_low_s=ci_low,
        ci_high_s=ci_high,
        adopt_bias=significant,
    )
    trajectory = make_trajectory(path, alignment, smooth_window=71)

    if write_outputs:
        out = output_dir(root)
        trajectory.to_excel(out / "problem2_10Hz_trajectory.xlsx", index=False)
        pd.DataFrame(
            [
                {
                    "problem": 2,
                    "delta_s": alignment.delta_s,
                    "delta_ci_low_s": ci_low,
                    "delta_ci_high_s": ci_high,
                    "bias_x_m": alignment.bias_x_m,
                    "bias_y_m": alignment.bias_y_m,
                    "overlap_s": alignment.overlap_s,
                    "n_overlap": alignment.n_overlap,
                    "mse_bias_model": bias_model.mse,
                    "mse_no_bias_model": no_bias.mse,
                    "improvement_ratio": (no_bias.mse - bias_model.mse) / no_bias.mse,
                    "f_stat": f_stat,
                    "p_value": p_value,
                    "adopt_bias": significant,
                }
            ]
        ).to_excel(out / "problem2_estimate.xlsx", index=False)
    return alignment, trajectory


def main() -> None:
    alignment, _trajectory = solve_problem2()
    print(f"problem2 delta = {alignment.delta_s:.6f} s")
    print(f"bias = ({alignment.bias_x_m:.6f}, {alignment.bias_y_m:.6f}) m")
    print(f"F p-value = {alignment.p_value:.6g}")


if __name__ == "__main__":
    main()
