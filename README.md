# 商业模拟竞赛结果分析方法论工具集 v3.0

本仓库包含企业模拟经营战报结果分析的完整工具集，基于**方法论3.0版本**。适用于商业模拟竞赛（如Cesim）的多回合数据分析。

---

## 🔄 分析方法

本工具集采用**三步分析法**，从结果数据到决策方案的完整流程：

### 步骤一：结果数据分析

**工具**：`scripts/analyze_comprehensive_v3.py`

**输入**：回合结果Excel文件（`results-ir00.xls`、`results-pr01.xls`、`results-pr02.xls`、`results-pr03.xls`等）

**输出**：方法论3.0完整分析报告（包含财务健康度、竞争分析、趋势分析等）

**使用方法**：
```bash
python scripts/analyze_comprehensive_v3.py --input-dir ./data --output-dir ./reports
```

### 步骤二：页面数据分析

**工具**：`docs/prompts/MHTML页面解析提示词.md`（AI提示词）

**输入**：Cesim决策页面的MHTML格式文件

**输出**：各页面分析报告（需求预测、生产、营销、研发、物流、税收、财务等页面）

**使用方法**：
1. 将Cesim决策页面保存为MHTML格式
2. 使用大模型（如GPT、Claude等）加载`MHTML页面解析提示词.md`
3. 将MHTML文件内容输入大模型，获取结构化分析报告

### 步骤三：决策方案制定

**工具**：`docs/prompts/决策制定提示词.md`（AI提示词）

**输入**：
- 步骤一生成的方法论3.0完整分析报告
- 步骤二生成的所有页面分析报告

**输出**：决策项具体方案与输入建议（包含69项决策的完整方案）

**使用方法**：
1. 使用大模型加载`决策制定提示词.md`
2. 将步骤一和步骤二的分析报告输入大模型
3. 获取完整的决策方案文档

---

## 📋 目录结构

```
cesim18th/
├── scripts/                     # 分析脚本目录
│   └── analyze_comprehensive_v3.py  # 核心分析脚本（v3.0）
├── utils/                       # 工具模块目录
│   └── utils_data_analysis.py   # 数据分析工具模块
├── docs/                        # 文档目录
│   ├── methodology/             # 方法论文档
│   │   └── 结果分析方法.md      # 方法论文档（v3.0）
│   └── prompts/                 # 提示词文档
│       ├── 决策制定提示词.md    # 决策制定相关提示词文档
│       └── MHTML页面解析提示词.md # MHTML页面解析相关提示词文档
├── README.md                    # 本说明文件
├── .gitignore                   # Git忽略配置
└── __pycache__/                 # Python缓存目录（自动生成）
```

---

## 🔧 核心文件说明

### 1. analyze_comprehensive_v3.py

**核心分析脚本** - 严格按照方法论3.0版本进行完整分析

**功能特性**：
- ✅ 数据基础建设（数据读取、指标映射、验证清洗）
- ✅ 自身诊断分析（财务健康度红绿灯、现金流分析、区域市场）
- ✅ 竞争分析解码（三维度对标矩阵、策略突变检测、意图预测）
- ✅ 多回合趋势分析（关键指标趋势和环比增长率）

**使用方法**：
```bash
# 进入项目目录
cd cesim18th

# 方式1：使用默认路径（数据在项目上级目录的 '结果/' 文件夹，报告输出到 '分析/' 文件夹）
python scripts/analyze_comprehensive_v3.py

# 方式2：指定数据输入目录
python scripts/analyze_comprehensive_v3.py --input-dir /path/to/data

# 方式3：指定数据输入目录和输出目录
python scripts/analyze_comprehensive_v3.py --input-dir /path/to/data --output-dir /path/to/output

# 方式4：使用短参数
python scripts/analyze_comprehensive_v3.py -i ./data -o ./reports

# 查看帮助信息
python scripts/analyze_comprehensive_v3.py --help
```

**命令行参数**：
- `--input-dir, -i`: 数据输入目录路径（包含 results-ir00.xls 等文件）
- `--output-dir, -o`: 分析报告输出目录路径
- `--help, -h`: 显示帮助信息

**输出**：
- 分析报告会生成在指定的输出目录中（默认在项目上级目录的 `分析/` 文件夹）
- 主要报告：`方法论3.0完整分析报告.md`

---

### 2. utils_data_analysis.py

**数据分析工具模块** - 提供通用数据读取和处理功能

**主要函数**：
```python
from utils.utils_data_analysis import (
    read_excel_data,      # 读取Excel文件并解析数据
    find_metric,          # 根据关键词查找指标
    get_metric_value,     # 获取特定队伍和指标的数值
    check_excel_structure, # 检查Excel文件结构
    diagnose_missing_data  # 诊断缺失数据
)
```

**使用示例**：
```python
from utils.utils_data_analysis import read_excel_data, get_metric_value

# 读取数据
metrics_dict, teams = read_excel_data('results-pr03.xls')

# 获取指标值
cash = get_metric_value(metrics_dict, '现金', '做大做强')
```

---

### 3. 结果分析方法.md

**方法论文档** - 完整的方法论3.0版本说明文档

**文件位置**：`docs/methodology/结果分析方法.md`

**包含内容**：
- 方法论概述（定位、流程、时间管理）
- 数据基础建设（指标映射、验证清洗）
- 自身诊断分析（财务健康度、现金流、市场表现）
- 竞争分析解码（对标矩阵、策略突变、意图预测）
- 决策支持体系（复盘分析、策略建议）
- 标准输出模板
- 工具与方法库（代码示例、公式库）
- 迭代与优化机制

---

### 4. 决策制定提示词.md

**决策制定提示词文档** - AI执行版的决策制定提示词

**文件位置**：`docs/prompts/决策制定提示词.md`

**用途**：为AI提供决策制定的完整流程和提示词模板

---

### 5. MHTML页面解析提示词.md

**MHTML页面解析提示词文档** - AI执行版的页面解析提示词

**文件位置**：`docs/prompts/MHTML页面解析提示词.md`

**用途**：为AI提供从MHTML格式页面中提取结构化数据的完整流程

---

## 📂 数据文件配置

### 数据文件要求

分析脚本需要以下回合的Excel结果文件：
- `results-ir00.xls` - 初始回合（回合0）
- `results-pr01.xls` - 第一回合
- `results-pr02.xls` - 第二回合
- `results-pr03.xls` - 第三回合

### 配置数据路径

**推荐方式：使用命令行参数**（无需修改代码）：
```bash
# 指定数据输入目录和输出目录
python scripts/analyze_comprehensive_v3.py --input-dir /path/to/data --output-dir /path/to/output
```

**默认配置**：
如果不指定参数，脚本默认从项目上级目录的 `结果/` 文件夹读取数据，输出到 `分析/` 文件夹。

**通过代码配置**（高级用法）：
如果需要修改默认路径，可以编辑 `scripts/analyze_comprehensive_v3.py` 中的默认配置：
```python
# 默认输入目录
DEFAULT_INPUT_DIR = Path(__file__).parent.parent.parent / '结果'

# 默认输出目录
DEFAULT_OUTPUT_DIR = Path(__file__).parent.parent.parent / '分析'
```

**配置队伍名称映射**：
如果队伍名称需要映射，编辑 `TEAM_NAME_MAPPING` 字典：
```python
TEAM_NAME_MAPPING = {
    '原始队伍名1': '映射后的名称1',
    '原始队伍名2': '映射后的名称2',
}
```

---

## 📊 输出报告

分析脚本会生成以下报告（输出位置在脚本中配置）：

1. **方法论3.0完整分析报告.md**
   - 内容：完整的分析报告，包含执行摘要、数据验证、财务诊断、竞争分析、趋势分析等
   - 包含章节：
     - 执行摘要
     - 数据基础建设验证
     - 自身诊断分析（财务健康度、现金流、市场表现）
     - 竞争分析解码（对标矩阵、策略突变、意图预测）
     - 多回合趋势分析

2. **分析执行总结.md**
   - 内容：分析执行情况总结和关键发现

3. **数据异常原因分析报告.md**
   - 内容：数据异常问题的诊断和解决方案

**注意**：输出目录路径在脚本中配置，默认在项目上级目录的 `分析/` 文件夹中。

---

## 🚀 快速开始

### 1. 环境要求

- Python 3.7+
- 必需的Python包

### 2. 安装依赖

```bash
pip install pandas numpy xlrd openpyxl
```

**说明**：
- `pandas` - 数据处理和分析
- `numpy` - 数值计算
- `xlrd` - 读取旧版Excel文件（.xls格式）
- `openpyxl` - 读取新版Excel文件（.xlsx格式，如需要）

### 3. 准备数据文件

将回合结果Excel文件放置在配置的数据目录中：
- `results-ir00.xls` - 初始回合
- `results-pr01.xls` - 第一回合
- `results-pr02.xls` - 第二回合
- `results-pr03.xls` - 第三回合

### 4. 配置脚本（可选）

根据需要修改 `scripts/analyze_comprehensive_v3.py` 中的配置：
- `BASE_DIR` - 数据文件路径
- `TEAM_NAME_MAPPING` - 队伍名称映射
- `THRESHOLDS` - 分析阈值配置

### 5. 运行分析

```bash
# 进入项目目录
cd cesim18th

# 方式1：使用默认路径
python scripts/analyze_comprehensive_v3.py

# 方式2：指定自定义路径
python scripts/analyze_comprehensive_v3.py --input-dir ./data --output-dir ./reports
```

**注意**：如果使用自定义路径，确保数据目录中包含以下文件：
- `results-ir00.xls` - 初始回合
- `results-pr01.xls` - 第一回合
- `results-pr02.xls` - 第二回合
- `results-pr03.xls` - 第三回合

### 6. 查看报告

分析完成后，在配置的输出目录中查看生成的报告文件。

---