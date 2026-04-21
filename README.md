# 工程经济学案例作业仓库说明

本仓库用于完成《工程经济学》案例研究作业，案例对象为宁波怡丰汇商业项目。

## 这份作业最终交什么

按照作业要求，最终提交物是：

- 一份 PDF 报告
- 一份包含完整计算过程的 Excel 文件
- 将以上两个文件打包成一个 ZIP 提交

仓库中作为“主版本报告 PDF”保留的是：

- `docs/analysis/final-report.pdf`

实际提交使用的生成文件路径是：

- PDF 报告：`outputs/submission/EE2025_Case_Study_Report.pdf`
- Excel 计算书：`outputs/submission/EE2025_Case_Study_Calculations.xlsx`
- 提交压缩包：`outputs/submission/EE2025_Case_Study_Submission.zip`

## 仓库目录说明

- `CaseStudy作业要求/`
  原始作业要求目录，保留压缩包、题目文档、原始 Excel 和中文提取说明。

- `docs/assignment/`
  题目原始文档备份目录。

- `docs/analysis/`
  分析过程和报告正文目录。
  这里面最重要的是：
  `final-report.tex`：LaTeX 报告源码
  `final-report.pdf`：仓库内保留的最终报告 PDF
  `final-report.md`：较早期的 Markdown 草稿
  `assignment-checklist.md`：把题目要求拆成可执行清单

- `docs/project/`
  项目管理和分工说明目录。
  这里面主要是：
  `branching-and-roles.md`：分支和分工规则
  `work-plan.md`：阶段任务计划

- `data/raw/`
  原始数据与原始附件备份，不建议直接修改。

- `data/processed/`
  预留给后续清洗数据或中间结果，目前基本空置。

- `src/`
  脚本目录，已按职责分层。
  `analysis/`：现金流模型、估值逻辑、参数与结果导出
  `excel/`：分步骤 Excel 工作簿生成
  `reporting/`：图表生成与报告相关输出
  顶层兼容入口：
  `calculate_case_study.py`
  `generate_report_figures.py`
  `build_submission.py`

- `outputs/`
  自动生成结果目录。
  `outputs/submission/` 下放最终提交件和图表输出。

## 仓库内各个 Markdown 文件是干什么的

- `CaseStudy作业要求/作业要求.md`
  从原始作业说明里提取出的中文要求，方便快速查看题目。

- `docs/analysis/assignment-checklist.md`
  把老师要求拆成“必须完成哪些分析项”的清单。

- `docs/analysis/final-report.md`
  报告的早期 Markdown 文本草稿，现在主要作参考，不是最终排版源文件。

- `docs/project/branching-and-roles.md`
  记录 `codex / cc / main` 等分支和协作角色分工。

- `docs/project/work-plan.md`
  记录从建模、敏感性分析到最终写作的推进计划。

- `src/README.md`
  简要说明 `src/` 目录结构和脚本职责。

## 当前推荐使用方式

如果要重新生成整套提交件，优先运行：

- `python3 src/build_submission.py`

这条命令会依次完成：

- 生成现金流与估值结果
- 生成报告图表
- 生成分步骤 Excel 工作簿
- 编译 LaTeX 报告
- 打包最终 ZIP

## 分支说明

- `codex`
  当前主要工作分支。

- `main`
  预留给最终整合后的稳定版本。

- `cc`
  预留给并行修改或另一位成员的工作分支。
