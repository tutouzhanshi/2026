# B题从头求解说明

## 运行方式

```powershell
cd 2026_B题
python solve_b_from_scratch.py
```

或双击/运行：

```powershell
run_b_solution.bat
```

## 主要输出

输出目录为 `outputs/`：

- `estimates_summary.xlsx`：问题1-3时间偏差、系统偏差与检验结果。
- `problem1_10Hz_trajectory.xlsx`、`problem2_10Hz_trajectory.xlsx`、`problem3_10Hz_trajectory.xlsx`：三问10Hz融合轨迹。
- `task_candidates.xlsx`：问题4候选任务。
- `task_schedule_detail.xlsx`：整数规划选中的任务明细。
- `result.xlsx`：按题目模板填好的结果表。
- `B题_论文.md`、`B题_多源融合机器人定位及任务优化_论文.docx`、`B题_多源融合机器人定位及任务优化_论文.pdf`：论文输出。
- `verification_report.xlsx`：自动校验报告，`PASS` 表示轨迹和任务约束检查通过。

## 方法摘要

方式1作为时间基准，方式2通过一维时间平移对齐。带偏模型在每个候选时间偏差下用坐标差中位数估计固定偏差，并用截尾均方误差评价匹配质量。问题3采用“统计显著性 + 工程效应量”综合判定是否采用系统偏差修正。问题4先在10Hz融合轨迹上生成满足准备时间约束的候选任务，再用0-1整数规划最大化期望完成数。
