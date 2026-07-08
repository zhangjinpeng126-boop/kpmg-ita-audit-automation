# 🔍 优衣库销售系统 IT 审计自动化分析

> **KPMG ITA 未来之翼** — 2026 毕马威信息技术审计竞赛  
> 第四题：自动化审计工具的应用现状与趋势  
> **基于 Python CAATs 的数据完整性审计方案**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 📋 项目概述

本项目以**优衣库（UNIQLO）中国区**销售点系统（POS）为审计对象，使用 **Python 计算机辅助审计技术（CAATs）** 替代传统 ACL/IDEA 审计软件，对销售数据进行 **100% 全量分析**，实现了从数据加载、Benford 定律检测、异常交易识别到审计报告生成的全流程自动化。

### 🎯 核心亮点

| 维度 | 手工审计 | Python 自动化 | 提升 |
|------|----------|---------------|------|
| 数据覆盖度 | 5% | 100% | **20x** |
| 处理时间 | ~2,880 分钟 | ~10 分钟 | **快 288x** |
| 异常检出率 | ~20% | ~100% | **5x** |
| 人工成本 | 基准 100 | 15 | **降 85%** |

---

## 📁 项目结构

```
kpmg-ita-audit-automation/
├── README.md                          # 项目说明文档
├── requirements.txt                   # Python 依赖包
├── .gitignore
├── src/
│   └── uniqlo_it_audit.py            # 🔧 审计自动化主程序
├── data/
│   └── uniqlo_sales.csv              # 📊 优衣库销售数据集
├── outputs/
│   ├── uniqlo_audit_report.docx      # 📝 自动生成的 Word 审计报告
│   ├── benford_uniqlo_analysis.png   # 📈 Benford 定律检测图
│   └── audit_efficiency_comparison.png # 📊 手工 vs 自动化效率对比图
└── presentation/
    └── From Sampling to Intelligence.pptx  # 🎯 竞赛答辩 PPT
```

---

## 🔬 审计程序

### 1️⃣ 数据加载与清洗
- 自动检测文件编码（UTF-8 / GBK / GB2312）
- 缺失值检测与剔除
- 中文日期格式标准化（"2023年1月2日" → datetime）

### 2️⃣ Benford 定律检测（CAATs 数据完整性测试）
- 提取销售金额首位有效数字（1-9）
- 对比实际分布 vs Benford 理论分布 `P(d) = log₁₀(1 + 1/d)`
- **卡方拟合优度检验（χ²）** 量化偏离程度
- 智能 P 值格式化（处理浮点下溢至 ~1e-308 以下）
- 生成专业级可视化图表（300 DPI 打印清晰度）

### 3️⃣ 异常交易检测（三类规则）
| 异常类型 | 检测规则 | IT 审计意义 |
|----------|----------|-------------|
| 负利润交易 | 利润 < 0 | 折扣配置异常 / 退货流程缺陷 / 价格主数据错误 |
| 异常金额（超 3σ） | 按产品类别分组，超出均值±3σ | 数据录入错误 / 系统 Bug / 非授权操作 |
| 整数金额（整百元） | 金额为 100/500/1000 等 | 疑似 POS 测试数据未清理 / 人为取整 |

### 4️⃣ 效率对比分析
- 四维度对比：覆盖度 / 处理时间 / 异常检出率 / 人工成本
- 2×2 分面子图（small multiples），避免量纲差异

### 5️⃣ 自动生成 Word 审计报告
- 完整的七节结构化审计报告
- 表格 + 图表 + IT 控制改进建议
- 符合毕马威 ITA 比赛格式规范

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行审计程序

```bash
cd src
python uniqlo_it_audit.py
```

程序将自动完成以下步骤：
1. 加载并清洗 `data/uniqlo_sales.csv`
2. 执行 Benford 定律分析 → 生成 `outputs/benford_uniqlo_analysis.png`
3. 检测异常交易 → 控制台输出 Top 20 异常清单
4. 生成效率对比图 → `outputs/audit_efficiency_comparison.png`
5. 生成 Word 审计报告 → `outputs/uniqlo_audit_report.docx`

---

## 🧠 关键技术决策

### Benford 定律的小 P 值陷阱
> ⚠️ 本案例样本量近 **2 万条**，卡方检验统计效力极高。即使微小的业务偏差（如零售业 59/79/99 元定价策略）也会导致 P 值趋近于零。  
> **审计判断结论**：偏差符合零售定价习惯 → 判定为业务原因导致的可解释偏差，而非数据完整性缺陷。

### 中文字体自动适配
程序自动检测系统可用的中文字体（SimHei / Microsoft YaHei / Noto Sans SC 等），确保审计图表中文标签正常显示。

### P 值智能格式化
当 χ² 极大时，P 值可能低于 IEEE 754 double 精度下限（~1e-308），scipy 返回 `0.0`。本程序使用对数生存函数手动估算量级，确保审计结论的统计严谨性。

---

## 📊 数据来源

和鲸社区（Heywhale）公开数据集《优衣库销售数据》（已脱敏）

> ⚠️ **免责声明**：本数据集已经脱敏处理，不包含真实客户个人信息，仅供审计方法演示和学术研究使用。

---

## 👤 作者

- **GitHub**: [@zhangjinpeng126-boop](https://github.com/zhangjinpeng126-boop)
- **比赛**: 2026 毕马威 ITA 未来之翼

---

## 📄 License

MIT License — 详见 [LICENSE](LICENSE)
