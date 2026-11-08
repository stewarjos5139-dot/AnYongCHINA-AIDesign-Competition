# 多源运营数据「相似度」模糊匹配工具

面向审计对账场景的中文企业名模糊匹配工具。输入 A 系统（ERP 客户交易明细）与 B 系统（银行流水明细），自动完成名称清洗、相似度计算、领域规则判定、一对一分配与五档分档，输出符合赛题模板的四 Sheet Excel 报表。

## 核心功能

- **无损读取**：表头自动探测；原始名称列原样保留，清洗只新增辅助列。
- **文本清洗**：剔除隐形字符（Cc / Zs / Zl / Zp / Cf 全类、NBSP、零宽、BOM、TAG）、全角转半角、去空格、英文大写。
- **多算法加权**：`ratio` 0.45 + `partial_ratio` 0.20 + `WRatio` 0.15 + `core_ratio` 0.20；另支持 6 种基础算法切换。
- **企业名领域规则**：
  - 字号差异惩罚：剥离公共前后缀，只比「字号」片段；
  - 通用词折叠 + 关键字号保护：R2 同一字号抬分、R1 关键字号冲突封顶、R3 同字号异业务只调候选；
  - 包含关系抬分：简称 / 全称按覆盖率抬回高度匹配；
  - 单字错别字、括号附注两类豁免。
- **一对一分配**：高分先挑、低分让位，同一条 B 记录只落到一条 A 记录。
- **同分裁决**：仅在名称分完全相等时，用交易日期与金额差打破平局。
- **可解释输出**：24 个受控标签写入「备注」列，可在 Excel 中按标签筛选复核。
- **双入口**：命令行 `main.py` 与 PyQt6 桌面端 `gui_app.py` 共用同一套 `src/pipeline.py` 编排逻辑。
- **可交付**：PyInstaller 单文件 exe，换数据无需重新打包。

## 处理流程

```
读取 → 清洗 → 基础打分 → 字号惩罚（含两类豁免）
     → 通用词折叠 + 关键字号保护 → 包含关系抬分
     → 精确兜底 + 同分裁决 + 一对一
     → 五档分档 → 四 Sheet Excel 导出
```

## 实测结果（赛题数据）

| 指标 | 结果 |
|---|---:|
| A 系统 / B 系统 | 100 条 / 97 条 |
| 全量比对 | 9,700 对 |
| 完全匹配 | 45 |
| 高度匹配 | 40 |
| 中低匹配 / 低置信度 | 0 / 0 |
| A 系统独有 | 15 |
| B 系统独有 | 12 |
| 匹配成功合计 | 85 |
| 端到端耗时 | 约 0.14 秒 |

账目对平：`85 + 15 = 100`（A 侧）、`85 + 12 = 97`（B 侧）。

赛题点名陷阱对均正确拦截：石油天然气 / 石油化工 45.4；建设银行 / 中国银行 19.9；安永华明 / 大华 44.5；顺丰控股 / 顺丰快递 46.0。腾讯简称场景判为 92.0 高度匹配。27 条独有记录单列为「待复核」，并附最佳候选、最高相似度与处理建议。

## 目录结构

```
.
├── main.py                     # 命令行入口
├── gui_app.py                  # PyQt6 桌面端
├── build_exe.py / build.bat    # PyInstaller 打包
├── requirements.txt
├── src/
│   ├── config.py               # 路径、Sheet、字段配置
│   ├── data_loader.py          # 无损读取、表头探测、Unicode 诊断
│   ├── preprocessor.py         # 清洗流水线、字号分析、词表
│   ├── matcher.py              # 打分矩阵、规则、选优、LLM 仲裁扩展位
│   ├── explainer.py            # 可解释标签
│   ├── exporter.py             # 四 Sheet 报表与条件格式
│   └── pipeline.py             # CLI / GUI 共用编排层
└── 团体赛赛道考题/              # 默认 A/B 数据（xlsx）
```

## 快速开始

```bash
# 1. 安装依赖
python -m pip install -r requirements.txt

# 2. 命令行全流程
python main.py

# 3. 桌面端
python gui_app.py

# 4. 离屏自检
python gui_app.py --selftest

# 5. 打包 exe
python build_exe.py
```

命令行常用参数：

```bash
python main.py --high 90 --low 70 --floor 60   # 分档阈值
python main.py --algo weighted                # 基础算法
python main.py --list-algo                    # 列出全部算法
python main.py -A path/to/A.xlsx -B path/to/B.xlsx
python main.py --no-export                    # 只匹配不导出
```

## 环境

- Python 3.10+（开发环境 3.13）
- pandas / numpy / openpyxl / rapidfuzz / tqdm
- GUI：PyQt6 / matplotlib
- 打包：PyInstaller

## 说明

- 金额不作为匹配依据；仅在名称分完全相等时用于同分流水裁决。
- 默认一对一匹配；抢不到 B 的记录显式判为 A 系统独有。
- 分档阈值默认 100 / 90 / 70 / 60，可在 GUI 或命令行调整。
- 规则参数与边界说明见项目内技术文档（不随仓库提供）。
