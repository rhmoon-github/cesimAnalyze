# 商业模拟竞赛结果分析方法论工具集 v3.0

本仓库包含企业模拟经营战报结果分析的完整工具集，基于**方法论3.0版本**。适用于商业模拟竞赛（如Cesim）的多回合数据分析。

---

## 📋 目录结构

```
cesim18th/
├── analyze_comprehensive_v3.py  # 核心分析脚本（v3.0）
├── utils_data_analysis.py       # 数据分析工具模块
├── 结果分析方法.md                # 方法论文档（v3.0）
├── README.md                    # 本说明文件
└── .gitignore                   # Git忽略配置
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
python analyze_comprehensive_v3.py

# 方式2：指定数据输入目录
python analyze_comprehensive_v3.py --input-dir /path/to/data

# 方式3：指定数据输入目录和输出目录
python analyze_comprehensive_v3.py --input-dir /path/to/data --output-dir /path/to/output

# 方式4：使用短参数
python analyze_comprehensive_v3.py -i ./data -o ./reports

# 查看帮助信息
python analyze_comprehensive_v3.py --help
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
from utils_data_analysis import (
    read_excel_data,      # 读取Excel文件并解析数据
    find_metric,          # 根据关键词查找指标
    get_metric_value,     # 获取特定队伍和指标的数值
    check_excel_structure, # 检查Excel文件结构
    diagnose_missing_data  # 诊断缺失数据
)
```

**使用示例**：
```python
from utils_data_analysis import read_excel_data, get_metric_value

# 读取数据
metrics_dict, teams = read_excel_data('results-pr03.xls')

# 获取指标值
cash = get_metric_value(metrics_dict, '现金', '做大做强')
```

---

### 3. 结果分析方法.md

**方法论文档** - 完整的方法论3.0版本说明文档

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
python analyze_comprehensive_v3.py --input-dir /path/to/data --output-dir /path/to/output
```

**默认配置**：
如果不指定参数，脚本默认从项目上级目录的 `结果/` 文件夹读取数据，输出到 `分析/` 文件夹。

**通过代码配置**（高级用法）：
如果需要修改默认路径，可以编辑 `analyze_comprehensive_v3.py` 中的默认配置：
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

根据需要修改 `analyze_comprehensive_v3.py` 中的配置：
- `BASE_DIR` - 数据文件路径
- `TEAM_NAME_MAPPING` - 队伍名称映射
- `THRESHOLDS` - 分析阈值配置

### 5. 运行分析

```bash
# 进入项目目录
cd cesim18th

# 方式1：使用默认路径
python analyze_comprehensive_v3.py

# 方式2：指定自定义路径
python analyze_comprehensive_v3.py --input-dir ./data --output-dir ./reports
```

**注意**：如果使用自定义路径，确保数据目录中包含以下文件：
- `results-ir00.xls` - 初始回合
- `results-pr01.xls` - 第一回合
- `results-pr02.xls` - 第二回合
- `results-pr03.xls` - 第三回合

### 6. 查看报告

分析完成后，在配置的输出目录中查看生成的报告文件。

---

## 📝 版本历史

- **v3.0** (2025-11-27)
  - 完全重构方法论文档结构
  - 优化指标提取逻辑
  - 校准阈值配置
  - 完善分析流程
  - 支持多回合数据对比分析
  - 增强数据验证和异常检测

- **v2.0** (已废弃)
  - 旧版本，已移除相关脚本和报告

---

## ⚠️ 注意事项

1. **数据路径**：确保数据文件路径配置正确，文件存在且可读
2. **编码问题**：脚本使用UTF-8编码，确保系统支持中文显示
3. **数据完整性**：如有数据缺失或异常，会标记在报告中，请检查数据异常原因分析报告
4. **指标映射**：如遇到指标提取失败，请参考方法论文档中的"指标名称映射体系"章节
5. **Excel格式**：确保Excel文件格式正确，建议使用标准的Cesim导出格式
6. **队伍名称**：如果队伍名称包含特殊字符或需要映射，请在脚本中配置 `TEAM_NAME_MAPPING`

---

## 🔍 故障排查

### 问题：指标提取为N/A

**原因**：指标名称不匹配

**解决**：
1. 查看 `数据异常原因分析报告.md` 了解具体问题
2. 参考方法论文档中的"指标名称映射体系"章节
3. 更新 `utils_data_analysis.py` 中的匹配逻辑

### 问题：数据完整性验证误差大

**原因**：使用了错误的验证公式

**解决**：
- 已在使用正确的公式：`总资产 = 权益合计 + 负债合计`
- 如果仍有问题，请检查Excel中的实际科目结构

### 问题：报告生成失败

**原因**：可能是路径或编码问题

**解决**：
1. 检查输出目录是否存在
2. 确保有写入权限
3. 检查文件编码设置

---

## 📚 相关文档

- [结果分析方法.md](结果分析方法.md) - 完整的方法论文档（v3.0），包含详细的分析流程、指标说明、公式库等
- 分析执行总结 - 每次运行后生成的执行情况总结
- 数据异常原因分析报告 - 数据问题诊断和解决方案报告

**方法论文档包含**：
- 方法论概述和流程
- 数据基础建设指南
- 自身诊断分析方法
- 竞争分析解码技巧
- 决策支持体系
- 标准输出模板
- 工具与方法库（代码示例、公式库）
- 迭代与优化机制

---

## 🤝 贡献与反馈

如有问题或建议，请：
1. 查看方法论文档中的"迭代与优化"章节
2. 参考现有的分析报告和诊断报告
3. 根据方法论进行改进和优化
4. 提交Issue或Pull Request到本仓库

## 📄 许可证

本项目仅供学习和研究使用。

---

**最后更新**：2025-11-27  
**当前版本**：3.0  
**仓库地址**：https://github.com/rhmoon-github/cesimAnalyze
