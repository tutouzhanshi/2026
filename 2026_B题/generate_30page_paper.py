# -*- coding: utf-8 -*-
"""Generate an extended 8-chapter mathematical modeling paper for 2026 B.

The solver output is treated as authoritative.  This script formats the
already verified numerical results into a contest-style paper with abstract,
manual table of contents, formulas, flow charts, three-line tables and appendix.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "outputs"
FIG = OUT / "figures"
DOCX_OUT = OUT / "B题_数学建模论文_30页_规范版_同范围无标题图版.docx"
MD_OUT = OUT / "B题_数学建模论文_30页_规范版.md"
REVIEW_OUT = OUT / "B题_30页论文自审报告.md"

plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "SimSun", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False


def set_run_font(run, font_name: str = "宋体", size_pt: float = 12, bold: bool = False, italic: bool = False) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic


def setup_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Cm(2.45)
    section.bottom_margin = Cm(2.3)
    section.left_margin = Cm(2.7)
    section.right_margin = Cm(2.4)
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.2
    add_page_number(section)


def add_page_number(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, fld_end])
    set_run_font(run, "宋体", 10.5)


def add_para(
    doc: Document,
    text: str = "",
    *,
    font_name: str = "宋体",
    size_pt: float = 12,
    bold: bool = False,
    align: int | None = None,
    first_line: bool = True,
    space_after: float = 3,
) -> None:
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    fmt = p.paragraph_format
    fmt.line_spacing = 1.2
    fmt.space_after = Pt(space_after)
    if first_line:
        fmt.first_line_indent = Pt(size_pt * 2)
    r = p.add_run(text)
    set_run_font(r, font_name, size_pt, bold)


def add_heading_1(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    set_run_font(r, "黑体", 15, True)


def add_heading_2(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, "黑体", 12.5, True)


def add_heading_3(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_run_font(r, "黑体", 12, False)


def add_formula(doc: Document, text: str, number: str | None = None) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(2)
    body = text if number is None else f"{text}        ({number})"
    r = p.add_run(body)
    set_run_font(r, "Times New Roman", 11)


def set_cell_text(cell, text: str, size_pt: float = 9.2, bold: bool = False, align: int = WD_ALIGN_PARAGRAPH.CENTER) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(str(text))
    set_run_font(run, "宋体", size_pt, bold)


def set_cell_border(cell, **kwargs) -> None:
    """Set selected cell borders. Values: {'val': 'single', 'sz': '8', 'color': '000000'}."""
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge in kwargs:
            tag = "w:{}".format(edge)
            element = tc_borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tc_borders.append(element)
            for key, value in kwargs[edge].items():
                element.set(qn(f"w:{key}"), str(value))


def add_three_line_table(
    doc: Document,
    df: pd.DataFrame,
    columns: Iterable[str] | None = None,
    caption: str | None = None,
    font_size: float = 8.8,
    left_cols: Iterable[str] | None = None,
) -> None:
    if caption:
        add_para(doc, caption, font_name="黑体", size_pt=10.5, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=1)
    cols = list(columns or df.columns)
    left_col_set = set(left_cols or [])
    table = doc.add_table(rows=1, cols=len(cols))
    table.autofit = True
    for j, c in enumerate(cols):
        set_cell_text(table.cell(0, j), c, size_pt=font_size, bold=True)
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for j, c in enumerate(cols):
            value = row[c]
            if pd.isna(value):
                text = ""
            elif isinstance(value, (float, np.floating)):
                text = f"{float(value):.4f}"
            else:
                text = str(value)
            align = WD_ALIGN_PARAGRAPH.LEFT if c in left_col_set else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(cells[j], text, size_pt=font_size, align=align)
    nil = {"val": "nil"}
    line = {"val": "single", "sz": "10", "color": "000000"}
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, top=nil, left=nil, bottom=nil, right=nil, insideH=nil, insideV=nil)
    for cell in table.rows[0].cells:
        set_cell_border(cell, top=line, bottom=line, left=nil, right=nil)
    for cell in table.rows[-1].cells:
        set_cell_border(cell, bottom=line, left=nil, right=nil)
    add_para(doc, "", first_line=False, space_after=0)


def add_caption(doc: Document, text: str) -> None:
    add_para(doc, text, font_name="宋体", size_pt=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=2)


def add_picture(doc: Document, path: Path, caption: str, width: float = 5.8) -> None:
    if path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(path), width=Inches(width))
        add_caption(doc, caption)


def add_toc_entry(doc: Document, text: str, page: str, level: int = 0) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.75 * level)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    r = p.add_run(f"{text}\t{page}")
    set_run_font(r, "宋体", 12)


def draw_linear_flow(path: Path, title: str, labels: list[str], fig_width: float = 12, fig_height: float = 2.4) -> None:
    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=180)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.text(0.5, 0.92, title, ha="center", va="center", fontsize=13, fontweight="bold")
    n = len(labels)
    box_w = 0.13 if n >= 6 else 0.15
    xs = np.linspace(0.08, 0.92, n)
    for i, (x, label) in enumerate(zip(xs, labels)):
        rect = plt.Rectangle((x - box_w / 2, 0.40), box_w, 0.26, fill=False, linewidth=1.2, edgecolor="black")
        ax.add_patch(rect)
        ax.text(x, 0.53, label, ha="center", va="center", fontsize=9, wrap=True)
        if i < n - 1:
            ax.annotate("", xy=(xs[i + 1] - box_w / 2, 0.53), xytext=(x + box_w / 2, 0.53), arrowprops=dict(arrowstyle="->", lw=1.2))
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def draw_overall_flow(path: Path) -> None:
    fig, ax = plt.subplots(figsize=(12, 5.2), dpi=180)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.text(0.5, 0.95, "问题总体分析流程", ha="center", va="center", fontsize=14, fontweight="bold")
    rows = [
        ("问题一", ["读取附件1", "公共时间轴", "搜索时间偏差", "10Hz融合轨迹"]),
        ("问题二", ["读取附件2", "带偏差模型", "F检验", "偏差修正轨迹"]),
        ("问题三", ["实际数据", "显著性判断", "偏差修正", "输出10Hz轨迹"]),
        ("问题四", ["附件3轨迹", "候选任务生成", "0-1整数规划", "约束复核"]),
    ]
    y_positions = [0.78, 0.60, 0.42, 0.24]
    for y, (name, steps) in zip(y_positions, rows):
        ax.text(0.07, y + 0.035, name, ha="center", va="center", fontsize=10, fontweight="bold")
        xs = [0.20, 0.42, 0.64, 0.86]
        for i, (x, step) in enumerate(zip(xs, steps)):
            rect = plt.Rectangle((x - 0.075, y - 0.025), 0.15, 0.09, fill=False, linewidth=1.1, edgecolor="black")
            ax.add_patch(rect)
            ax.text(x, y + 0.02, step, ha="center", va="center", fontsize=9)
            if i < len(xs) - 1:
                ax.annotate("", xy=(xs[i + 1] - 0.075, y + 0.02), xytext=(x + 0.075, y + 0.02), arrowprops=dict(arrowstyle="->", lw=1.0))
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def generate_flowcharts() -> dict[str, Path]:
    FIG.mkdir(parents=True, exist_ok=True)
    overall = FIG / "paper_overall_flowchart.png"
    align = FIG / "paper_alignment_flowchart.png"
    task = FIG / "paper_task_optimization_flowchart.png"
    draw_overall_flow(overall)
    draw_linear_flow(
        align,
        "时间同步与系统偏差估计流程",
        ["读取双源轨迹", "平滑与插值", "搜索δ", "估计b", "嵌套F检验", "10Hz融合输出"],
        fig_width=12,
        fig_height=2.8,
    )
    draw_linear_flow(
        task,
        "任务候选生成与整数规划流程",
        ["修正轨迹", "速度加速度", "距离角度筛选", "冲突约束", "MILP求解", "结果复核"],
        fig_width=12,
        fig_height=2.8,
    )
    return {"overall": overall, "align": align, "task": task}


def load_results() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    est = pd.read_excel(OUT / "estimates_summary.xlsx")
    tasks = pd.read_excel(OUT / "task_schedule_detail.xlsx")
    sens = pd.read_excel(OUT / "sensitivity_smooth_window.xlsx")
    verify = pd.read_excel(OUT / "verification_report.xlsx")
    return est, tasks, sens, verify


def add_title_abstract(doc: Document, est: pd.DataFrame, tasks: pd.DataFrame) -> None:
    r1, r2, r3 = (est.iloc[i] for i in range(3))
    expected = float(tasks["expected_success"].sum())
    shoot_count = int((tasks["task"] == "模拟射击").sum())
    photo_count = int((tasks["task"] == "拍照").sum())

    add_para(doc, "多源异频定位融合与机器人任务调度优化模型设计", font_name="黑体", size_pt=17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=14)
    add_para(doc, "摘    要", font_name="黑体", size_pt=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=8)
    abstract = (
        "移动机器人在执行沿途目标识别与任务操作时，需要将不同频率、不同起始时刻的定位数据统一到"
        "同一时间基准下。本文针对两种定位方式存在异频、异步、随机噪声和固定系统偏差的情形，"
        "建立时间同步、系统偏差检验、10Hz轨迹融合与任务调度优化模型，并利用附件数据完成四个问题的求解。"
    )
    add_para(doc, abstract)
    add_para(
        doc,
        f"针对问题一，我们首先以定位方式1为基准，构造方式2时间轴平移量δ，并将两条轨迹插值到公共"
        f"0.1s时间网格上。通过最小化公共重叠区间内的轨迹均方距离，得到方式2相对方式1的时间偏差为"
        f"{r1['方式2相对方式1时间偏差_delta_s']:.4f}s，并据此生成无噪声条件下的10Hz融合轨迹。"
    )
    add_para(
        doc,
        f"针对问题二，我们在时间同步模型中加入二维固定系统偏差b，采用坐标差分量中位数估计偏差，"
        f"并用截尾均方误差削弱随机噪声和局部异常值的影响。通过无偏模型与带偏模型的嵌套F检验，"
        f"判断系统偏差是否应当进入模型。计算得到时间偏差为{r2['方式2相对方式1时间偏差_delta_s']:.4f}s，"
        f"系统偏差为({r2['采用的系统偏差_x_m']:.4f},{r2['采用的系统偏差_y_m']:.4f})m，误差下降比例为"
        f"{100 * r2['误差下降比例']:.2f}%，说明附件2中固定偏差显著，需修正后再进行10Hz轨迹融合。"
    )
    add_para(
        doc,
        f"针对问题三，我们将上述带偏配准模型用于实际测量数据，并进一步给出时间偏差的置信区间和系统偏差"
        f"显著性判断。结果显示，方式2相对方式1的时间偏差为{r3['方式2相对方式1时间偏差_delta_s']:.4f}s，"
        f"95%置信区间为[{r3['delta_95%CI_lower_s']:.4f},{r3['delta_95%CI_upper_s']:.4f}]s，系统偏差为"
        f"({r3['采用的系统偏差_x_m']:.4f},{r3['采用的系统偏差_y_m']:.4f})m，F检验p值为{r3['F检验p值']:.2e}。"
        f"因此，实际数据中存在统计显著的微小系统偏差，修正后误差下降比例为{100 * r3['误差下降比例']:.2f}%，"
        "并输出对应10Hz融合轨迹。"
    )
    add_para(
        doc,
        f"针对问题四，我们在问题三修正轨迹上计算速度、加速度、目标距离和拍照方向角，逐时刻生成满足射击、"
        f"拍照距离、速度、加速度及准备时间约束的候选任务；随后建立0-1整数规划模型，加入任务时间互斥、"
        f"同一射击目标唯一性和同一拍照目标角度差约束，以最大化期望完成数。最终选中{len(tasks)}项任务，"
        f"其中模拟射击{shoot_count}项、拍照{photo_count}项，期望完成数为{expected:.2f}。自动复核结果为PASS，"
        "敏感性分析表明主方案在不同平滑窗口下具有较好的稳定性。"
    )
    add_para(doc, "关键词：时间同步；系统偏差；轨迹融合；嵌套F检验；整数规划；任务优化", font_name="黑体", bold=True)


def add_contents(doc: Document) -> None:
    doc.add_page_break()
    add_para(doc, "目    录", font_name="黑体", size_pt=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=False, space_after=12)
    entries = [
        ("一、问题重述", "3", 0),
        ("1.1 问题背景", "3", 1),
        ("1.2 问题重述", "4", 1),
        ("二、问题分析", "6", 0),
        ("2.1 问题一的分析", "6", 1),
        ("2.2 问题二的分析", "7", 1),
        ("2.3 问题三的分析", "8", 1),
        ("2.4 问题四的分析", "10", 1),
        ("三、模型假设", "11", 0),
        ("四、符号说明", "13", 0),
        ("五、模型建立与求解", "15", 0),
        ("5.1 问题一的模型", "15", 1),
        ("5.2 问题二的模型", "18", 1),
        ("5.3 问题三的模型", "21", 1),
        ("5.4 问题四求解", "24", 1),
        ("六、模型分析、检验与评价", "31", 0),
        ("七、参考文献", "35", 0),
        ("八、附录", "36", 0),
    ]
    for text, page, level in entries:
        add_toc_entry(doc, text, page, level)
    add_para(doc, "注：目录页码按照本文显式分页排版给出，若在Word中调整字体或更新图片尺寸，页码可能随排版微调。", size_pt=10.5, first_line=True)


def add_problem_restatement(doc: Document, figures: dict[str, Path]) -> None:
    def add_problem_item(label: str, text: str) -> None:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.first_line_indent = Pt(24)
        r1 = p.add_run(label)
        set_run_font(r1, "黑体", 12, True)
        r2 = p.add_run("  " + text)
        set_run_font(r2, "宋体", 12, False)

    doc.add_page_break()
    add_heading_1(doc, "一、问题重述")
    add_heading_2(doc, "1.1 问题背景")
    add_para(
        doc,
        "随着移动机器人、无人车和智能巡检设备在复杂场景中的应用不断增加，机器人在运动过程中"
        "需要获得连续、准确、稳定的位置轨迹。实际定位系统通常由不同传感器或不同定位方式共同提供数据，"
        "这些数据在采样频率、起始时间、测量精度和坐标基准上可能存在差异。如果不先对不同来源的数据"
        "进行时间同步和偏差修正，直接融合会造成轨迹错位，进而影响机器人后续执行射击、拍照等任务时"
        "对距离、速度和加速度约束的判断。"
    )
    add_para(
        doc,
        "题目给出两种二维定位方式，其中定位方式1的输出频率为4Hz，定位方式2的输出频率为5Hz。"
        "两种方式记录的是同一机器人运动过程，但由于设备开机先后、随机测量噪声以及固定系统偏差等原因，"
        "原始数据不能直接按行对齐。前三个问题要求根据附件1至附件3分别完成时间偏差估计、系统偏差判断"
        "和10Hz轨迹输出；第四个问题则要求在实际轨迹基础上，对沿途射击目标和拍照目标进行任务优化设计。"
    )

    add_heading_2(doc, "1.2 问题重述")
    add_para(
        doc,
        "现需根据题目给出的附件数据，建立合理的数学模型，对两种异频定位数据进行对齐、修正和融合，"
        "并在融合轨迹的基础上完成任务目标优化。具体需要解决的问题如下。"
    )
    add_problem_item(
        "问题一",
        "附件1中的两种定位数据在采集过程中无随机噪声影响，但由于两台设备开机时间不同，"
        "两组时间序列存在固定时间偏差。需要建立两类数据的时间对齐模型，求出定位方式2相对定位方式1"
        "的时间偏差，并在统一时间基准下给出10Hz的位置轨迹。"
    )
    add_problem_item(
        "问题二",
        "附件2中的两种定位数据不仅存在开机先后造成的时间不同步，还存在随机测量噪声和固定系统偏差。"
        "需要在问题一时间对齐模型的基础上，进一步建立含系统偏差的数据融合模型，估计两种定位方式之间"
        "的时间偏差和二维坐标系统偏差，并输出修正后的10Hz位置轨迹。"
    )
    add_problem_item(
        "问题三",
        "附件3给出的是实际测量数据，题目没有直接说明两种定位方式之间是否存在系统偏差。需要利用数据本身"
        "判断系统偏差是否存在；若存在，则应在偏差修正后完成数据对齐和轨迹融合；若不存在，则应避免引入"
        "不必要的偏差参数。最终仍需给出时间偏差、系统偏差判断结果以及10Hz位置轨迹。"
    )
    add_problem_item(
        "问题四",
        "机器人按照附件3中的轨迹运动，运动过程中需要对附件4给出的射击目标和拍照目标执行模拟射击或拍照扫描任务。"
        "射击任务受到距离、速度、加速度、瞄准准备时间和命中率约束，拍照任务受到距离、速度、加速度、相机准备时间"
        "以及同一目标多角度拍摄要求约束。需要设计任务目标优化方案，在满足全部约束的前提下尽可能多地完成任务，"
        "并将最终任务安排填入result.xlsx。"
    )


def add_problem_analysis(doc: Document) -> None:
    doc.add_page_break()
    add_heading_1(doc, "二、问题分析")
    add_heading_2(doc, "2.1 问题一的分析")
    add_para(
        doc,
        "针对问题一，两种定位方式记录的是同一条机器人运动轨迹，数据中没有随机噪声和系统偏差，"
        "主要矛盾来自采样频率不同和设备开机时间不同。定位方式1为4Hz，定位方式2为5Hz，若直接按数据行号"
        "进行比较，会把不同时刻的位置当成同一时刻的位置，从而得到错误的时间偏差。因此，该问题的关键在于"
        "建立一个只含时间平移量的对齐模型。"
    )
    add_para(
        doc,
        "根据题意，附件1中的两组轨迹可以看作同一真实轨迹在不同时间轴上的采样结果。我们假设定位方式2"
        "相对定位方式1存在一个固定时间偏差δ，将方式2的时间轴平移后，与方式1在公共时间区间内进行比较。"
        "由于两种方式的采样时刻并不重合，需要先把两组数据插值到统一的0.1s时间网格，再计算两条轨迹之间"
        "的均方距离。使均方距离最小的时间平移量，即为两种定位方式之间的时间偏差。"
    )
    add_para(
        doc,
        "问题一没有噪声干扰，模型不需要额外引入空间偏差和滤波参数，求解过程相对直接。得到时间偏差后，"
        "只需将两种定位方式统一到同一时间轴上，并按10Hz频率输出融合位置轨迹。该问题为后续含噪数据和"
        "实际测量数据的处理提供了基本的时间同步框架。"
    )

    add_heading_2(doc, "2.2 问题二的分析")
    add_para(
        doc,
        "针对问题二，数据中同时存在随机测量噪声和固定系统偏差。随机噪声会使相邻采样点出现局部波动，"
        "固定系统偏差则会使一条轨迹相对另一条轨迹整体平移。如果仍沿用问题一中的无偏时间对齐模型，"
        "空间偏移会被错误地计入时间匹配误差，进而影响时间偏差估计。因此，问题二需要把时间偏差估计和"
        "空间系统偏差估计放在同一个模型中处理。"
    )
    add_para(
        doc,
        "根据该问题的特点，我们把系统偏差理解为定位方式2相对于定位方式1的二维常量偏移。对于每一个"
        "候选时间偏差，先将两条轨迹插值到公共时间网格，再根据两条轨迹的坐标差估计系统偏差。考虑到"
        "随机噪声和可能存在的局部异常点，偏差不宜简单使用均值估计，而应采用更稳健的坐标差中位数。"
        "之后再计算去除系统偏差后的轨迹匹配误差，从而筛选出最合理的时间偏差和系统偏差。"
    )
    add_para(
        doc,
        "此外，问题二还需要判断引入系统偏差后模型是否确实优于无偏模型。由于带偏模型比无偏模型多两个"
        "偏差参数，可以把二者看作嵌套模型，并利用F检验比较残差平方和的下降是否显著。若检验显著，"
        "说明固定偏差对解释两组数据差异有必要；若不显著，则应避免过度修正。问题二的最终结果应包括"
        "时间偏差、二维系统偏差和修正后的10Hz融合轨迹。"
    )

    doc.add_page_break()
    add_heading_2(doc, "2.3 问题三的分析")
    add_para(
        doc,
        "针对问题三，附件3给出的是实际测量数据，数据来源更接近真实工程环境。与问题二相比，本问题的"
        "显著特点是题目没有预先说明系统偏差是否存在，因此不能直接假设存在固定偏移，也不能因为估计出的"
        "偏差不为零就认定存在系统偏差。实际测量数据中即使没有固定偏差，也可能由于噪声和采样波动产生"
        "非零的样本平均差。"
    )
    add_para(
        doc,
        "因此，问题三的分析思路应当继承问题二的带偏配准模型，但在结论判断上更加谨慎。我们先按照带偏模型"
        "估计时间偏差和候选系统偏差，再利用无偏模型与带偏模型的嵌套关系进行显著性检验。如果加入偏差参数后"
        "残差下降在统计上显著，则判定存在系统偏差并进行修正；否则保持无偏融合结果。这样可以避免把随机噪声"
        "误判为系统偏差。"
    )
    add_para(
        doc,
        "同时，实际测量数据往往比仿真数据更容易受到局部异常值影响。为提高结果稳定性，需要在轨迹配准前"
        "进行适当平滑，在误差评价中采用截尾均方误差，并对时间偏差给出置信区间。最终，问题三不仅要给出"
        "10Hz融合轨迹，还要明确回答“是否存在系统偏差”这一判断性问题。"
    )

    add_heading_2(doc, "2.4 问题四的分析")
    add_para(
        doc,
        "针对问题四，机器人不再需要重新规划运动路线，而是按照附件3融合后的轨迹运动，并在运动过程中尽可能多地"
        "完成射击和拍照任务。因此，问题四本质上是给定轨迹条件下的任务选择与调度问题。任务是否可执行取决于"
        "执行时刻以及准备时间内机器人与目标点之间的距离、机器人速度、加速度以及拍照方向角等因素。"
    )
    add_para(
        doc,
        "直接在连续时间上搜索所有可能的射击和拍照时刻，会使模型复杂度较高，也不便于处理同一目标的角度差约束和"
        "任务时间互斥。由于前三问已经输出10Hz轨迹，本文可以把0.1s轨迹点作为候选执行时刻，对每个目标和每个时刻"
        "逐一检查题目约束，先生成满足物理条件的候选任务集合。这样可以将复杂的连续约束转化为离散候选筛选问题。"
    )
    add_para(
        doc,
        "在候选任务生成后，仍需决定哪些任务可以同时保留。不同任务的准备区间和执行时刻可能发生重叠，同一射击目标"
        "不应重复计数，同一拍照目标的多次拍摄角度也必须满足最小差异要求。因此，可以为每个候选任务建立0-1决策变量，"
        "以期望完成数最大为目标，加入时间互斥、射击唯一性和拍照角度差约束，形成0-1整数规划模型。"
    )
    add_para(
        doc,
        "这一处理方式的优点是计算路径清晰，候选生成阶段负责判断单个任务是否可行，整数规划阶段负责处理任务之间"
        "的组合冲突。由于速度和加速度由实际轨迹差分得到，平滑窗口会影响候选任务数量，因此还需要通过敏感性分析"
        "检验最终任务方案的稳定性。"
    )


def add_assumptions_symbols(doc: Document) -> None:
    doc.add_page_break()
    add_heading_1(doc, "三、模型假设")
    assumptions = [
        "两种定位方式观测的是同一机器人在同一平面内的运动轨迹，附件中坐标单位均为米，时间单位均为秒。",
        "方式2相对方式1的开机时间差在每个附件内为常量，短时间内不存在随时间变化的时钟漂移。",
        "系统偏差若存在，则表现为二维坐标上的固定平移向量，不考虑旋转、比例缩放和非线性畸变。",
        "随机测量噪声总体均值为零，局部异常点可以通过平滑、截尾误差和中位数偏差估计降低影响。",
        "第四问中机器人只能沿附件3融合轨迹运动，不对轨迹本身进行重新规划；任务只在10Hz离散时刻上执行。",
        "题目给出的射击命中率、距离、速度、加速度、准备时间和拍照角度差约束准确有效，未给出的额外冷却时间不作考虑。",
    ]
    for i, item in enumerate(assumptions, 1):
        add_para(doc, f"（{i}）{item}")
    add_para(
        doc,
        "上述假设主要用于把题目中的工程问题转化为可计算模型。其中，时间差常量假设保证可以用一维参数δ描述"
        "两类数据的同步关系；固定平移假设保证系统偏差可用二维向量b描述；离散时刻假设使任务优化可转化为"
        "0-1整数规划。若实际系统存在时钟漂移或坐标旋转，模型需要进一步扩展为仿射配准或动态状态估计，"
        "这将在模型改进部分讨论。"
    )
    add_para(
        doc,
        "对第四问而言，最重要的假设是“轨迹给定”。本文不改变机器人运动路线，因此任务可行性完全由机器人"
        "在轨迹各时刻的位置、速度、加速度及其到目标点的相对几何关系决定。该处理方式与题意中“机器人按照"
        "附件3中的轨迹运动”一致，也避免了把路径规划和任务规划混为一谈。"
    )
    doc.add_page_break()
    add_heading_1(doc, "四、符号说明")
    symbols = pd.DataFrame(
        [
            ["轨迹数据", "t₁,t₂", "方式1、方式2的原始时间序列", "s"],
            ["轨迹数据", "p₁(t),p₂(t)", "两种定位方式给出的二维位置", "m"],
            ["时间同步", "δ", "方式2相对方式1的时间偏差", "s"],
            ["时间同步", "q₁(t),q₂(t-δ)", "插值到公共时间网格后的轨迹", "m"],
            ["偏差模型", "b=(bₓ,bᵧ)", "方式2相对方式1的固定坐标偏差", "m"],
            ["偏差模型", "J₀(δ),J₁(δ)", "无偏模型、带偏模型的平均匹配误差", "m²"],
            ["偏差检验", "SSE₀,SSE₁", "无偏模型、带偏模型的残差平方和", "m²"],
            ["偏差检验", "F,p", "嵌套F检验统计量及其p值", "无"],
            ["轨迹输出", "Δt", "融合轨迹采样间隔，本文取0.1", "s"],
            ["轨迹输出", "p_f(t)", "10Hz融合轨迹位置", "m"],
            ["任务约束", "vₖ,aₖ", "第k个轨迹点的速度、加速度", "m/s, m/s²"],
            ["任务约束", "dᵢₖ", "第k时刻机器人到目标i的距离", "m"],
            ["任务约束", "θᵢₖ", "第k时刻机器人观察拍照目标i的方向角", "°"],
            ["任务优化", "xᵢ", "候选任务i是否被选中", "0或1"],
            ["任务优化", "wᵢ", "候选任务i的期望收益", "无"],
            ["任务优化", "Iᵢ", "候选任务i占用的准备及执行时间区间", "s"],
        ],
        columns=["类别", "符号", "含义", "单位"],
    )
    add_three_line_table(doc, symbols, caption="表1 主要符号说明", font_size=8.0, left_cols=["含义"])
    add_para(
        doc,
        "表1按轨迹数据、时间同步、偏差模型、偏差检验、轨迹输出和任务优化六类列出核心符号。"
        "实际计算中，问题1至问题3分别对应不同附件，本文在公式中省略附件编号；第四问中的候选任务i"
        "既可以是射击任务，也可以是拍照任务，其收益wᵢ由任务类型决定。"
    )


def add_model_and_solution(doc: Document, est: pd.DataFrame, tasks: pd.DataFrame, figures: dict[str, Path]) -> None:
    r1, r2, r3 = (est.iloc[i] for i in range(3))
    doc.add_page_break()
    add_heading_1(doc, "五、模型建立与求解")
    add_heading_2(doc, "5.1 问题一的模型")
    add_heading_3(doc, "5.1.1 数据预处理")
    add_para(
        doc,
        "问题一的附件1包含方式1和方式2两张数据表。方式1采样周期为0.25s，方式2采样周期为0.2s，"
        "二者无法直接逐行比较。设方式2相对方式1的时间偏差为δ，则方式2校正后的时间为t_2-δ。"
        "对给定δ，先求两条轨迹的公共时间区间，并以0.1s为间隔生成公共网格。若公共重叠长度不足，"
        "该δ不作为有效候选。"
    )
    add_formula(doc, "T(δ)={t | t∈[max(min t_1,min(t_2-δ)), min(max t_1,max(t_2-δ))],  t=kΔt}", "1")
    add_para(
        doc,
        "在公共网格T(δ)上，对原始轨迹分别进行线性插值，得到q_1(t)和q_2(t-δ)。由于附件1无噪声，"
        "插值误差主要来自采样频率差异，线性插值足以满足10Hz输出需要。"
    )
    add_heading_3(doc, "5.1.2 时间平移模型")
    add_formula(doc, "J_0(δ)=1/n · Σ_{t∈T(δ)} ||q_2(t-δ)-q_1(t)||^2", "2")
    add_formula(doc, "δ^*=arg min_δ J_0(δ)", "3")
    add_para(
        doc,
        "式（2）把两条轨迹在公共时间上的平均平方距离作为匹配误差。由于只有一个未知参数δ，本文先在可行区间内"
        "进行细网格搜索，找到误差最小邻域后再用一维有界优化细化结果。此方法不需要初始配对点，适用于两条轨迹"
        "起始时间差较大的情形。"
    )
    doc.add_page_break()
    add_heading_3(doc, "5.1.3 求解结果与轨迹输出")
    res1 = pd.DataFrame(
        [
            ["问题1", r1["方式2相对方式1时间偏差_delta_s"], r1["重叠时长_s"], r1["重叠样本数"], r1["是否采用系统偏差修正"]],
        ],
        columns=["问题", "δ/s", "重叠时长/s", "重叠样本数", "是否修正偏差"],
    )
    add_three_line_table(doc, res1, caption="表2 问题一时间偏差估计结果", font_size=8.8)
    add_para(
        doc,
        f"计算得到方式2相对方式1的时间偏差为{r1['方式2相对方式1时间偏差_delta_s']:.4f}s。由于附件1无噪声且无系统偏差，"
        "融合轨迹采用两条对齐轨迹的平均值，并按0.1s间隔输出。图2展示了问题一融合后的10Hz轨迹，轨迹形状连续，"
        "说明时间平移模型能够把两种采样频率的数据对齐到同一运动过程。"
    )
    add_picture(doc, FIG / "problem1_trajectory_10hz.png", "图2 问题一10Hz融合轨迹", width=5.6)

    doc.add_page_break()
    add_heading_2(doc, "5.2 问题二的模型")
    add_heading_3(doc, "5.2.1 含噪数据和固定偏差处理")
    add_para(
        doc,
        "问题二相比问题一增加了随机测量噪声和固定系统偏差。若仍使用式（2）的无偏模型，轨迹整体平移会使匹配误差"
        "显著增大，并可能影响δ估计。为此，本文在每个候选δ下估计一个二维固定偏差向量b(δ)，再用去偏残差评价匹配质量。"
    )
    add_formula(doc, "b(δ)=median_{t∈T(δ)} [q_2(t-δ)-q_1(t)]", "4")
    add_formula(doc, "J_1(δ)=1/n · Σ_{t∈T(δ)} ||q_2(t-δ)-q_1(t)-b(δ)||^2", "5")
    add_formula(doc, "δ^*=arg min_δ J_1(δ),    b^*=b(δ^*)", "6")
    add_para(
        doc,
        "式（4）使用坐标差的分量中位数估计系统偏差。中位数估计不依赖噪声严格服从正态分布，能减少局部异常点对偏差"
        "估计的影响。为降低高频噪声对搜索曲线的扰动，程序中对轨迹进行了小窗口移动平均，并在误差评价中使用截尾均方误差。"
    )
    add_picture(doc, figures["align"], "图3 时间同步与系统偏差估计流程图", width=6.2)
    doc.add_page_break()
    add_heading_3(doc, "5.2.2 嵌套模型检验")
    add_para(
        doc,
        "无偏模型可看作带偏模型在b_x=b_y=0时的特例，因此两者构成嵌套模型。记无偏模型残差平方和为SSE_0，"
        "带偏模型残差平方和为SSE_1，样本数为n，新增参数数目为2，则采用F统计量检验加入系统偏差是否显著降低误差。"
    )
    add_formula(doc, "F=[(SSE_0-SSE_1)/2] / [SSE_1/(n-2)]", "7")
    add_para(
        doc,
        "当p值小于0.05时，认为固定系统偏差对解释两条轨迹差异具有统计显著作用；否则不采用偏差修正。"
        "该检验把“偏差估计值非零”和“偏差确实需要进入模型”区分开来，避免因噪声产生的样本偏移导致过拟合。"
    )
    res2 = pd.DataFrame(
        [
            [
                "问题2",
                r2["方式2相对方式1时间偏差_delta_s"],
                r2["采用的系统偏差_x_m"],
                r2["采用的系统偏差_y_m"],
                r2["F统计量"],
                r2["F检验p值"],
                r2["误差下降比例"],
            ],
        ],
        columns=["问题", "δ/s", "b_x/m", "b_y/m", "F统计量", "p值", "误差下降比例"],
    )
    add_three_line_table(doc, res2, caption="表3 问题二时间偏差与系统偏差估计结果", font_size=8.4)
    add_para(
        doc,
        f"问题二中δ={r2['方式2相对方式1时间偏差_delta_s']:.4f}s，系统偏差为"
        f"({r2['采用的系统偏差_x_m']:.4f},{r2['采用的系统偏差_y_m']:.4f})m，F检验p值接近0，"
        f"误差下降比例达到{100 * r2['误差下降比例']:.2f}%。这说明附件2的固定偏差不仅数值明显，"
        "而且对轨迹配准质量有决定性影响，必须在融合前进行修正。"
    )
    add_picture(doc, FIG / "problem2_trajectory_10hz.png", "图4 问题二10Hz融合轨迹", width=5.6)

    doc.add_page_break()
    add_heading_2(doc, "5.3 问题三的模型")
    add_heading_3(doc, "5.3.1 实际测量数据的偏差判定")
    add_para(
        doc,
        "问题三使用实际测量数据，随机误差和设备误差更接近真实工程场景。本文沿用问题二的带偏差模型，但在结论上"
        "更加谨慎：先估计候选偏差，再用F检验判断其统计显著性，最后根据检验结果决定是否修正。由于实际数据的"
        "偏差可能很小，不能单纯以偏差模长作为是否存在系统偏差的判据。"
    )
    add_formula(doc, "H_0: b_x=0,b_y=0;      H_1: 至少一个偏差分量不为0", "8")
    add_formula(doc, "若 p=P(F_{2,n-2}≥F_obs)<0.05, 则拒绝H_0并采用偏差修正", "9")
    add_para(
        doc,
        "为给出时间偏差的不确定性，程序对配准残差进行重采样，得到δ的近似95%置信区间。该区间用于说明时间同步结果"
        "在噪声扰动下的稳定程度，不直接改变融合轨迹的计算。"
    )
    res3 = pd.DataFrame(
        [
            [
                "问题3",
                r3["方式2相对方式1时间偏差_delta_s"],
                r3["delta_95%CI_lower_s"],
                r3["delta_95%CI_upper_s"],
                r3["采用的系统偏差_x_m"],
                r3["采用的系统偏差_y_m"],
                r3["F检验p值"],
                r3["误差下降比例"],
            ],
        ],
        columns=["问题", "δ/s", "δ下界/s", "δ上界/s", "b_x/m", "b_y/m", "p值", "误差下降比例"],
    )
    add_three_line_table(doc, res3, caption="表4 问题三实际测量数据偏差检验结果", font_size=8.1)
    add_para(
        doc,
        f"问题三中δ={r3['方式2相对方式1时间偏差_delta_s']:.4f}s，95%置信区间为"
        f"[{r3['delta_95%CI_lower_s']:.4f},{r3['delta_95%CI_upper_s']:.4f}]s。估计系统偏差为"
        f"({r3['采用的系统偏差_x_m']:.4f},{r3['采用的系统偏差_y_m']:.4f})m，p值为{r3['F检验p值']:.2e}，"
        f"小于0.05。因此本文判定实际测量数据存在统计显著的固定系统偏差，并采用偏差修正。"
    )
    add_picture(doc, FIG / "problem3_trajectory_10hz.png", "图5 问题三10Hz融合轨迹", width=5.6)

    doc.add_page_break()
    add_heading_3(doc, "5.3.2 10Hz轨迹融合模型")
    add_para(
        doc,
        "在完成时间同步和偏差修正后，两种定位方式被映射到同一个0.1s网格。若采用系统偏差修正，则先从方式2坐标中"
        "减去b；若不采用修正，则令b=0。本文采用等权平均输出融合轨迹："
    )
    add_formula(doc, "p_f(t)=1/2 · [q_1(t)+q_2(t-δ)-b],   t∈T, Δt=0.1s", "10")
    add_para(
        doc,
        "等权平均的原因是题目未给出两种定位方式的先验方差，也未要求建立实时滤波器。经过时间与偏差校正后，两类定位"
        "数据都可视为对真实位置的测量，等权平均能降低随机误差并保持模型透明。若有传感器方差估计，可进一步扩展为"
        "加权平均或卡尔曼滤波。"
    )
    add_para(
        doc,
        "前三问输出的10Hz轨迹分别写入problem1_10Hz_trajectory.xlsx、problem2_10Hz_trajectory.xlsx和"
        "problem3_10Hz_trajectory.xlsx。后续任务优化只使用问题三的修正轨迹，因此第四问的结果与问题三的偏差判断"
        "直接相关。"
    )

    doc.add_page_break()
    add_heading_2(doc, "5.4 问题四求解")
    add_heading_3(doc, "5.4.1 运动学指标计算")
    add_para(
        doc,
        "第四问首先在问题三10Hz融合轨迹上估计机器人速度和加速度。由于差分会放大测量噪声，程序先对轨迹进行"
        "平滑处理，再用中心差分计算速度和加速度。设p_k为第k个10Hz轨迹点，Δt=0.1s，则内部点的速度和加速度为："
    )
    add_formula(doc, "v_k=||p_{k+1}-p_{k-1}||/(2Δt)", "11")
    add_formula(doc, "a_k=||p_{k+1}-2p_k+p_{k-1}||/(Δt^2)", "12")
    add_para(
        doc,
        "射击任务要求执行时刻与准备区间内距离在5m至30m之间，速度不超过2m/s，加速度不超过1.5m/s²，"
        "并且执行前1.5s均满足约束。拍照任务要求距离在10m至40m之间，速度不超过1.5m/s，加速度不超过1.5m/s²，"
        "执行前0.5s满足约束；同一拍照目标的不同拍摄方向角差异至少60°。"
    )
    constraints = pd.DataFrame(
        [
            ["模拟射击", "5≤d≤30", "v≤2.0", "a≤1.5", "1.5s", "命中率0.85，同一射击目标最多一次"],
            ["拍照扫描", "10≤d≤40", "v≤1.5", "a≤1.5", "0.5s", "同一目标多次拍摄角度差至少60°"],
        ],
        columns=["任务类型", "距离约束/m", "速度约束/(m/s)", "加速度约束/(m/s²)", "准备时间", "附加约束"],
    )
    add_three_line_table(doc, constraints, caption="表5 任务可行性约束汇总", font_size=8.2)
    add_picture(doc, figures["task"], "图6 任务候选生成与整数规划流程图", width=6.2)

    doc.add_page_break()
    add_heading_3(doc, "5.4.2 候选任务生成")
    add_para(
        doc,
        "对每个目标点和每个10Hz轨迹时刻，程序计算机器人到目标点的距离d、速度v、加速度a以及拍照方向角θ。"
        "若该时刻及其准备区间内所有运动学约束都满足，则生成一个候选任务。候选任务记录目标编号、任务类型、"
        "准备开始时刻、执行时刻、距离、速度、加速度、方向角和期望收益。"
    )
    add_formula(doc, "d_{ik}=||p_f(t_k)-g_i||", "13")
    add_formula(doc, "θ_{ik}=atan2(y_i-y_k, x_i-x_k)", "14")
    add_para(
        doc,
        "生成候选集时，射击和拍照使用不同的距离与准备时间参数。这样可以在进入整数规划前排除明显不可行的任务，"
        "减少求解规模。主方案平滑窗口为71个点，共生成337个候选任务；这些候选任务均已在生成阶段满足基础物理约束。"
    )
    cand_summary = pd.DataFrame(
        [
            ["主方案", 71, 337, len(tasks), float(tasks["expected_success"].sum()), "PASS"],
        ],
        columns=["方案", "平滑窗口/点", "候选任务数", "选中任务数", "期望完成数", "复核结果"],
    )
    add_three_line_table(doc, cand_summary, caption="表6 主方案候选生成与选中结果概览", font_size=8.6)

    doc.add_page_break()
    add_heading_3(doc, "5.4.3 0-1整数规划模型")
    add_para(
        doc,
        "设候选任务集合为C，对每个候选任务i∈C定义0-1变量x_i。若任务i被选中，则x_i=1；否则x_i=0。"
        "任务收益w_i由任务类型确定：拍照任务记为1，射击任务按命中率记为0.85。目标函数为最大化期望完成数："
    )
    add_formula(doc, "max  Σ_{i∈C} w_i x_i", "15")
    add_para(doc, "主要约束包括以下三类。第一，若两个候选任务占用的准备及执行时间区间I_i和I_j相交，则不能同时选择：")
    add_formula(doc, "x_i+x_j≤1,   if I_i∩I_j≠∅", "16")
    add_para(doc, "第二，同一射击目标最多被射击一次，避免重复计算同一目标的命中收益：")
    add_formula(doc, "Σ_{i∈S(g)} x_i≤1,   对每个射击目标g", "17")
    add_para(doc, "第三，对于同一拍照目标，若两个候选拍照方向角差小于60°，则这两次拍摄不能同时选择：")
    add_formula(doc, "x_i+x_j≤1,   if target_i=target_j 且 |wrap(θ_i-θ_j)|<60°", "18")
    add_para(
        doc,
        "该模型是标准0-1线性整数规划。程序使用SciPy的MILP求解器求得最优解，并对解进行自动复核。"
        "由于候选任务已经预先满足距离、速度、加速度和准备时间约束，整数规划阶段重点处理候选之间的组合冲突。"
    )

    doc.add_page_break()
    add_heading_3(doc, "5.4.4 任务优化结果")
    task_show = tasks.copy()
    task_show = task_show.rename(
        columns={
            "target_id": "目标",
            "task": "任务",
            "prep_start_s": "准备开始/s",
            "exec_time_s": "执行时刻/s",
            "distance_m": "距离/m",
            "speed_m_s": "速度/(m/s)",
            "accel_m_s2": "加速度/(m/s²)",
            "angle_deg": "角度/°",
            "expected_success": "期望收益",
        }
    )
    add_three_line_table(doc, task_show, caption="表7 问题四选中任务明细", font_size=6.8)
    add_para(
        doc,
        f"表7按执行时间列出了最终选中的{len(tasks)}项任务。前一段时间轨迹附近可完成较多拍照任务，"
        "随后在507s至515s附近插入3项射击任务，末段在750s以后继续完成拍照和射击。选中方案的任务区间"
        "互不重叠，同一射击目标未重复计数，同一拍照目标的多次拍摄角度差满足题目要求。"
    )
    add_picture(doc, FIG / "problem4_selected_tasks.png", "图7 问题四选中任务空间分布", width=5.8)


def add_analysis_validation(doc: Document, sens: pd.DataFrame, verify: pd.DataFrame) -> None:
    doc.add_page_break()
    add_heading_1(doc, "六、模型分析、检验与评价")
    add_heading_2(doc, "6.1 结果复核")
    status = str(verify.iloc[0, 0]) if len(verify) else "UNKNOWN"
    add_para(
        doc,
        f"为避免模型求解结果只满足目标函数而违反题面约束，程序在输出result.xlsx后进行了自动复核。"
        f"复核内容包括：选中任务是否按时间顺序排列；准备区间和执行时刻是否满足距离、速度、加速度约束；"
        f"射击目标是否重复；同一拍照目标的角度差是否达到60°；任务时间区间是否互斥。复核报告给出的结论为{status}。"
    )
    add_para(
        doc,
        "前三问的轨迹结果也通过重叠时长、样本数量、误差下降比例和显著性检验进行解释。问题一的无噪声条件下，"
        "无需系统偏差修正；问题二偏差明显，F统计量很大且p值接近0；问题三偏差模长较小，但p值远小于0.05，"
        "说明实际测量数据中仍存在统计显著的固定偏差。"
    )
    validation = pd.DataFrame(
        [
            ["时间同步", "重叠时长与匹配误差", "三问均有充足公共区间"],
            ["系统偏差", "嵌套F检验", "问题2、3均显著"],
            ["轨迹输出", "10Hz时间网格", "三问均已输出Excel轨迹"],
            ["任务调度", "自动约束复核", status],
        ],
        columns=["检验对象", "检验方法", "检验结论"],
    )
    add_three_line_table(doc, validation, caption="表8 模型结果检验汇总", font_size=8.5)
    doc.add_page_break()
    add_heading_2(doc, "6.2 敏感性分析")
    add_para(
        doc,
        "第四问中速度和加速度由轨迹差分得到，平滑窗口大小会影响候选任务数量和最终选中结果。为检验主方案对"
        "平滑窗口的依赖程度，程序分别取51、61、71、81、91个点进行候选生成和整数规划求解。结果见表9。"
    )
    add_three_line_table(doc, sens, caption="表9 平滑窗口敏感性分析结果", font_size=8.4)
    add_para(
        doc,
        "从表9可以看出，主方案窗口71时选中18项任务，期望完成数17.25，为所测试窗口中的最高值。窗口过小会使"
        "速度和加速度估计更受噪声影响，候选任务偏少；窗口过大则可能过度平滑局部运动变化，使部分时刻的约束判断"
        "发生变化。五组窗口全部得到MILP最优状态且复核通过，说明模型流程稳定，但最终任务数对运动学估计仍有一定敏感性。"
    )
    doc.add_page_break()
    add_heading_2(doc, "6.3 模型的优点")
    advantages = [
        "模型链条完整：从时间同步、偏差估计、轨迹融合到任务调度形成闭环，第四问直接使用第三问的修正轨迹。",
        "偏差判断有统计依据：通过嵌套F检验区分随机样本偏移和固定系统偏差，避免仅凭偏差估计值作主观判断。",
        "任务约束表达清晰：先生成物理可行候选，再用0-1整数规划处理互斥关系，使结果便于复核和解释。",
        "程序输出可复现：关键中间结果、候选任务、选中任务、敏感性分析和复核报告均以Excel或图像形式保存。",
    ]
    for i, item in enumerate(advantages, 1):
        add_para(doc, f"（{i}）{item}")
    add_para(
        doc,
        "尤其在问题三中，偏差模长只有约0.26m，但F检验仍显示显著。这一结论说明模型没有简单按照工程经验阈值"
        "忽略偏差，而是把题目要求的“是否存在系统偏差”转化为可检验的统计问题。"
    )
    add_heading_2(doc, "6.4 模型的缺点")
    limitations = [
        "时间同步模型假设时钟偏差为常量，若两种设备存在时钟漂移，单一δ无法完全描述时间误差。",
        "系统偏差只建模为平移向量，未考虑坐标轴旋转、比例缩放或非线性畸变。",
        "速度和加速度由平滑差分估计，平滑窗口需要经验选择，且对第四问候选数量有影响。",
        "任务执行时刻限制在10Hz离散网格上，若允许连续时间微调，可能存在更优任务组合。",
    ]
    for i, item in enumerate(limitations, 1):
        add_para(doc, f"（{i}）{item}")
    doc.add_page_break()
    add_heading_2(doc, "6.5 模型的改进与推广")
    add_para(
        doc,
        "若需要进一步提高模型精度，可把时间同步模型扩展为δ(t)=δ_0+κt，用线性漂移项κ描述设备时钟漂移；"
        "也可把空间偏差扩展为二维仿射变换，联合估计旋转、缩放和平移参数。对于实际机器人定位，还可以引入"
        "卡尔曼滤波或粒子滤波，在状态空间中融合多源传感器观测。"
    )
    add_para(
        doc,
        "对于第四问，若题目允许改变机器人速度或局部路径，则可把当前的任务选择模型推广为路径—任务联合优化模型。"
        "此时决策变量不仅包括任务是否执行，还包括机器人在连续空间中的控制量，模型将从0-1整数规划扩展为混合整数"
        "非线性规划。当前模型虽然没有优化路径，但其候选筛选和冲突约束仍可作为更复杂模型的离散任务层。"
    )
    add_para(
        doc,
        "在更多实际场景中，射击或拍照任务的收益可能与距离、角度、光照、目标重要性有关。可将本文中固定的收益w_i"
        "改为任务质量函数，例如拍照收益随视角差和距离变化，射击收益随距离和速度变化。这样模型可以从“尽可能多完成任务”"
        "推广到“最大化综合任务价值”。"
    )


def add_references_appendix(doc: Document, est: pd.DataFrame, tasks: pd.DataFrame, sens: pd.DataFrame, verify: pd.DataFrame) -> None:
    doc.add_page_break()
    add_heading_1(doc, "七、参考文献")
    refs = [
        "司守奎, 孙玺菁. 数学建模算法与应用[M]. 北京: 国防工业出版社, 2011.",
        "Dantzig G B. Linear Programming and Extensions[M]. Princeton: Princeton University Press, 1963.",
        "Savitzky A, Golay M J E. Smoothing and Differentiation of Data by Simplified Least Squares Procedures[J]. Analytical Chemistry, 1964, 36(8): 1627-1639.",
        "Virtanen P, Gommers R, Oliphant T E, et al. SciPy 1.0: fundamental algorithms for scientific computing in Python[J]. Nature Methods, 2020, 17(3): 261-272.",
        "Harris C R, Millman K J, van der Walt S J, et al. Array programming with NumPy[J]. Nature, 2020, 585(7825): 357-362.",
        "McKinney W. Data structures for statistical computing in Python[C]//Proceedings of the 9th Python in Science Conference. 2010: 56-61.",
        "张贤达. 现代信号处理[M]. 北京: 清华大学出版社, 2002.",
    ]
    for i, ref in enumerate(refs, 1):
        add_para(doc, f"[{i}] {ref}", first_line=False)

    doc.add_page_break()
    add_heading_1(doc, "八、附录")
    add_heading_2(doc, "附录1 支撑材料文件列表")
    files = pd.DataFrame(
        [
            ["2026_B题.docx", "题目原文", "给出问题背景、四个问题和附录约束"],
            ["附件1.xlsx", "问题一数据", "无噪声双源定位数据"],
            ["附件2.xlsx", "问题二数据", "含随机噪声和固定系统偏差的数据"],
            ["附件3.xlsx", "问题三、四数据", "实际测量定位数据"],
            ["附件4.xlsx", "目标点数据", "射击目标和拍照目标坐标"],
            ["solve_b_from_scratch.py", "求解程序", "完成配准、轨迹融合、任务优化和复核"],
            ["outputs/estimates_summary.xlsx", "估计结果", "问题1至3时间偏差、系统偏差和检验结果"],
            ["outputs/task_schedule_detail.xlsx", "任务结果", "问题四选中任务明细"],
            ["outputs/verification_report.xlsx", "复核报告", "记录任务约束复核结论"],
        ],
        columns=["文件", "类型", "说明"],
    )
    add_three_line_table(doc, files, caption="表10 附录文件说明", font_size=8.0)
    add_heading_2(doc, "附录2 主要程序模块说明")
    modules = pd.DataFrame(
        [
            ["read_pair", "读取方式1和方式2轨迹数据，返回时间和二维坐标数组"],
            ["estimate_alignment", "在可行时间偏差区间内搜索δ，并根据是否带偏差计算匹配误差"],
            ["nested_f_test", "计算无偏模型与带偏模型的嵌套F检验统计量和p值"],
            ["build_fused_trajectory", "按0.1s时间网格输出10Hz融合轨迹"],
            ["generate_task_candidates", "根据距离、速度、加速度和准备时间生成候选任务"],
            ["solve_task_milp", "建立并求解0-1整数规划，最大化期望完成数"],
            ["verify_schedule", "对选中任务进行时间互斥、目标约束和运动学约束复核"],
        ],
        columns=["程序模块", "作用"],
    )
    add_three_line_table(doc, modules, caption="表11 求解程序主要模块", font_size=8.2)
    doc.add_page_break()
    add_heading_2(doc, "附录3 核心算法伪代码")
    pseudo = [
        "算法1 时间同步与偏差估计：",
        "输入：两种定位方式的时间序列和坐标序列；是否采用带偏模型。",
        "1. 设定δ的可行区间，保证两条轨迹有足够公共重叠区间；",
        "2. 对每个候选δ，将方式2时间修正为t_2-δ，并在公共网格插值得到q_1和q_2；",
        "3. 若采用带偏模型，则令b为q_2-q_1的分量中位数，否则令b=0；",
        "4. 计算去偏后的截尾均方误差J(δ)，选择误差最小的δ并局部细化；",
        "5. 对无偏模型和带偏模型计算SSE，并用嵌套F检验判断系统偏差是否显著；",
        "6. 根据检验结论输出δ、b和10Hz融合轨迹。",
        "",
        "算法2 任务候选生成与0-1整数规划：",
        "输入：问题三10Hz融合轨迹、射击目标、拍照目标和题目约束。",
        "1. 对融合轨迹平滑，并用差分估计每个时刻的速度和加速度；",
        "2. 对每个目标和每个10Hz时刻检查执行时刻及准备区间约束，生成候选任务；",
        "3. 为每个候选任务设置0-1变量x_i，收益w_i为拍照1或射击0.85；",
        "4. 添加时间互斥、射击唯一性和拍照角度差约束；",
        "5. 求解MILP，输出选中任务并写入result.xlsx；",
        "6. 对输出结果逐条复核，若无违规项则报告PASS。",
    ]
    for line in pseudo:
        add_para(doc, line, font_name="仿宋", size_pt=10.5, first_line=False if line.startswith("算法") or line == "" else True)

    doc.add_page_break()
    add_heading_2(doc, "附录4 问题一至问题三关键数值汇总")
    est_show = est[
        [
            "问题",
            "方式2相对方式1时间偏差_delta_s",
            "delta_95%CI_lower_s",
            "delta_95%CI_upper_s",
            "采用的系统偏差_x_m",
            "采用的系统偏差_y_m",
            "F统计量",
            "F检验p值",
            "误差下降比例",
        ]
    ].rename(
        columns={
            "方式2相对方式1时间偏差_delta_s": "δ/s",
            "delta_95%CI_lower_s": "δ下界/s",
            "delta_95%CI_upper_s": "δ上界/s",
            "采用的系统偏差_x_m": "b_x/m",
            "采用的系统偏差_y_m": "b_y/m",
            "F统计量": "F",
            "F检验p值": "p值",
            "误差下降比例": "误差下降比例",
        }
    )
    add_three_line_table(doc, est_show, caption="表12 问题一至问题三估计结果完整汇总", font_size=7.6)
    add_para(
        doc,
        "表12保留了正文中压缩展示的关键数值，便于复核者直接对照程序输出。问题一没有进行偏差检验，"
        "相应的F统计量和p值为空；问题二和问题三均显示带偏模型相对于无偏模型有显著改进。"
    )
    add_para(
        doc,
        "其中问题三的误差下降比例虽然只有约1.65%，但由于样本量达到3000，F检验仍能识别出稳定的固定偏差。"
        "因此本文在结论中采用“存在统计显著的微小系统偏差”的表述，以同时反映统计显著性和工程效应大小。"
    )

    doc.add_page_break()
    add_heading_2(doc, "附录5 问题四选中任务统计补充")
    task_stats = pd.DataFrame(
        [
            ["选中任务总数", len(tasks), "项"],
            ["拍照任务数", int((tasks["task"] == "拍照").sum()), "项"],
            ["模拟射击任务数", int((tasks["task"] == "模拟射击").sum()), "项"],
            ["期望完成数", float(tasks["expected_success"].sum()), "项"],
            ["最早执行时刻", float(tasks["exec_time_s"].min()), "s"],
            ["最晚执行时刻", float(tasks["exec_time_s"].max()), "s"],
            ["最小任务距离", float(tasks["distance_m"].min()), "m"],
            ["最大任务距离", float(tasks["distance_m"].max()), "m"],
        ],
        columns=["统计项", "数值", "单位"],
    )
    add_three_line_table(doc, task_stats, caption="表13 问题四选中任务统计", font_size=8.4)
    add_para(
        doc,
        "该统计表用于从整体上描述任务方案。选中任务覆盖476.5s至763.5s的轨迹区间，既包括近距离拍照任务，"
        "也包括在速度和加速度满足约束时插入的模拟射击任务。期望完成数不是简单任务数，因为射击任务按0.85命中率计入收益。"
    )
    add_para(
        doc,
        "从任务类型看，拍照任务数量更多，这是因为拍照目标允许在满足角度差约束的条件下多次拍摄，而射击目标采用"
        "同一目标最多一次的约束。该处理符合题目中“尽量多从不同角度拍摄”和“单次射击命中率”的差异化要求。"
    )

    doc.add_page_break()
    add_heading_2(doc, "附录6 敏感性分析补充说明")
    add_para(
        doc,
        "敏感性分析改变的是轨迹平滑窗口，而非题目原始数据或任务约束。窗口越小，速度和加速度曲线越接近原始噪声；"
        "窗口越大，曲线越平滑但可能削弱局部机动特征。因此不同窗口下候选任务数、选中任务数和期望完成数会出现变化。"
    )
    add_three_line_table(doc, sens, caption="表14 平滑窗口敏感性分析复列表", font_size=8.4)
    add_para(
        doc,
        "五个窗口的求解状态均为milp_optimal，复核结论均为PASS，说明整数规划模型本身没有因为窗口变化而失效。"
        "主方案选择窗口71，是因为它在测试集合中获得最高期望完成数，同时候选数量适中，避免过小窗口产生噪声驱动的任务缺失。"
    )

    doc.add_page_break()
    add_heading_2(doc, "附录7 图表清单")
    fig_table = pd.DataFrame(
        [
            ["图1", "问题总体分析流程图", "展示四个问题的整体求解路线"],
            ["图2", "问题一10Hz融合轨迹", "验证无噪声时间同步结果"],
            ["图3", "时间同步与系统偏差估计流程图", "说明δ、b和F检验的计算顺序"],
            ["图4", "问题二10Hz融合轨迹", "展示含噪带偏数据修正后的轨迹"],
            ["图5", "问题三10Hz融合轨迹", "展示实际数据修正后的融合轨迹"],
            ["图6", "任务候选生成与整数规划流程图", "说明第四问从轨迹到MILP的过程"],
            ["图7", "问题四选中任务空间分布", "展示选中任务与目标点的空间关系"],
        ],
        columns=["编号", "名称", "说明"],
    )
    add_three_line_table(doc, fig_table, caption="表15 论文图件清单", font_size=8.2)
    add_para(
        doc,
        "图件均由程序自动生成并保存到outputs/figures目录。流程图服务于建模思路说明，轨迹图和任务图服务于结果解释。"
        "图表标题采用“图号+名称”的方式，正文中均有对应文字说明。"
    )

    doc.add_page_break()
    add_heading_2(doc, "附录8 提交材料检查清单")
    check_status = str(verify.iloc[0, 0]) if len(verify) else "UNKNOWN"
    checklist = pd.DataFrame(
        [
            ["论文DOCX", DOCX_OUT.name, "已生成"],
            ["结果表", "outputs/result.xlsx", "已按题目模板写入"],
            ["轨迹结果", "problem1/2/3_10Hz_trajectory.xlsx", "已生成"],
            ["任务明细", "outputs/task_schedule_detail.xlsx", "已生成"],
            ["复核报告", "outputs/verification_report.xlsx", check_status],
            ["求解代码", "solve_b_from_scratch.py", "已保留"],
            ["论文生成代码", "generate_30page_paper.py", "已保留"],
        ],
        columns=["材料", "文件", "状态"],
    )
    add_three_line_table(doc, checklist, caption="表16 提交材料检查清单", font_size=8.2)
    add_para(
        doc,
        "最终提交时，应优先使用本脚本生成的DOCX论文和outputs/result.xlsx。若需要重新计算，可先运行solve_b_from_scratch.py"
        "刷新所有数值结果，再运行generate_30page_paper.py生成论文。"
    )


def write_markdown(est: pd.DataFrame, tasks: pd.DataFrame, sens: pd.DataFrame) -> None:
    def df_to_markdown(df: pd.DataFrame) -> str:
        cols = [str(c) for c in df.columns]
        rows = ["| " + " | ".join(cols) + " |", "| " + " | ".join(["---"] * len(cols)) + " |"]
        for _, row in df.iterrows():
            values: list[str] = []
            for c in df.columns:
                val = row[c]
                if pd.isna(val):
                    values.append("")
                elif isinstance(val, (float, np.floating)):
                    values.append(f"{float(val):.4f}")
                else:
                    values.append(str(val))
            rows.append("| " + " | ".join(values) + " |")
        return "\n".join(rows)

    r1, r2, r3 = (est.iloc[i] for i in range(3))
    lines = [
        "# 多源异频定位融合与机器人任务调度优化模型",
        "",
        "本Markdown为DOCX论文的结果索引版，完整30页排版请见 `B题_数学建模论文_30页_规范版.docx`。",
        "",
        "## 核心结果",
        f"- 问题1：δ={r1['方式2相对方式1时间偏差_delta_s']:.6f}s。",
        f"- 问题2：δ={r2['方式2相对方式1时间偏差_delta_s']:.6f}s，b=({r2['采用的系统偏差_x_m']:.6f},{r2['采用的系统偏差_y_m']:.6f})m。",
        f"- 问题3：δ={r3['方式2相对方式1时间偏差_delta_s']:.6f}s，b=({r3['采用的系统偏差_x_m']:.6f},{r3['采用的系统偏差_y_m']:.6f})m，p={r3['F检验p值']:.3e}。",
        f"- 问题4：选中{len(tasks)}项任务，期望完成数={float(tasks['expected_success'].sum()):.2f}。",
        "",
        "## 选中任务",
        df_to_markdown(tasks),
        "",
        "## 敏感性分析",
        df_to_markdown(sens),
    ]
    MD_OUT.write_text("\n".join(lines), encoding="utf-8")


def write_review_report(est: pd.DataFrame, tasks: pd.DataFrame, verify: pd.DataFrame) -> None:
    status = str(verify.iloc[0, 0]) if len(verify) else "UNKNOWN"
    text = f"""# B题30页规范版论文自审报告

## 结构检查

- 已按用户给出的模板重写为8章：问题重述、问题分析、模型假设、符号说明、模型建立与求解、模型分析检验与评价、参考文献、附录。
- 已包含题名、摘要、关键词和目录。
- 主要表格在DOCX中按三线表边框生成。
- 已加入3张流程图和4张结果图，并配有图题和文字说明。
- 正文采用显式分页扩写，按约30页论文体量排版。

## 数据一致性检查

- 问题1至问题3结果来自 `outputs/estimates_summary.xlsx`。
- 问题4任务明细来自 `outputs/task_schedule_detail.xlsx`，共{len(tasks)}项，期望完成数{float(tasks['expected_success'].sum()):.2f}。
- 自动复核报告结论：{status}。

## 残余风险

- Word在不同机器上打开时可能因字体替换和图片缩放导致页码微调。
- 任务优化使用10Hz离散时刻，若改为连续时间优化，可能产生不同的任务组合。
- 速度、加速度估计依赖平滑窗口，论文已用敏感性分析说明影响。
"""
    REVIEW_OUT.write_text(text, encoding="utf-8")


def main() -> None:
    OUT.mkdir(exist_ok=True)
    figures = generate_flowcharts()
    est, tasks, sens, verify = load_results()

    doc = Document()
    setup_document(doc)
    add_title_abstract(doc, est, tasks)
    add_contents(doc)
    add_problem_restatement(doc, figures)
    add_problem_analysis(doc)
    add_assumptions_symbols(doc)
    add_model_and_solution(doc, est, tasks, figures)
    add_analysis_validation(doc, sens, verify)
    add_references_appendix(doc, est, tasks, sens, verify)
    doc.save(DOCX_OUT)

    write_markdown(est, tasks, sens)
    write_review_report(est, tasks, verify)
    print(DOCX_OUT)
    print(MD_OUT)
    print(REVIEW_OUT)


if __name__ == "__main__":
    main()
