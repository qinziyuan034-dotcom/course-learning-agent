# 高数 PPT 智能学习助手 · M1–M6 模块状态汇总

> 生成于 2026-10-09，工作区 `E:\agent 开发`
> 说明：本汇总如实记录各模块**当前可离线运行的能力与质量上限**，不夸大成"全功能成品"。

## 总览

| 模块 | 职责 | 状态 | 质量/上限 |
|---|---|---|---|
| M1 PPT 解析 | PPT → Slide JSON（文本/公式/例题/溯源页码） | ✅ 可运行 | 公式文本为双通道片段级 |
| M2 内容问答·RAG | 按标题层级分块，混合检索（BM25+TF-IDF）Top-K，页码溯源 | ✅ 可运行 | 词面+浅语义，片段级准度 |
| M3 按指令出题 | 题型族参数化出题 + SymPy 三重校验 + 题库入库 | ✅ 可运行 | 换元积分题型族 |
| M4 答案批改 | 模式A（简答等价）/模式B（过程逐步定位错步） | ✅ 可运行 | SymPy 证据 |
| M5 错误诊断 | 6 类错误字典 + 纠正 + 同类练习 | ✅ 可运行 | 启发式分类 |
| M6 数学验证 | SymPy 单一事实源（等价/求导/反向验证） | ✅ 可运行 | 确定性符号证据 |

**数据主线**：M1 解析 → M2 检索溯源 → M3 出题入库 → M4 批改 → M5 诊断 → 回写薄弱点档案 → 反哺 M3 倾斜。已端到端打通并验证。

---

## 各模块交付物

### M1 PPT 解析
- `m1_ppt_parser.py`：文本层（python-pptx 走 shape，标题按字号推断）+ OLE 公式计数 + WMF 关联 + CEQU 回填 + OCR 双通道
- `m1_equation_decoder.py`：MathType CEQU 解码器（olefile 读 OLE 流的 Equation Native 流，扫描定位 + 递归解析）
- `wmf_to_png.py`：288 张 WMF → 高清 PNG
- `ocr_all.py`：RapidOCR 批量识别
- **`slides.json`**：52 页，279 个公式对象（CEQU 234 + OCR 234，双通道 197），39/52 页有标题

**质量如实**：公式文本为**片段级**（`sin`、`xlnx`、`fu`），非 LaTeX 整式。`text_confidence` 诚实标注 `fragment`/`dual_channel`/`empty`，不臆造。

### M2 内容问答·RAG
- `m2_rag_search.py`：分块（每页 1 chunk，剥页码，题型关键词补强）+ 双通道（手写 BM25 + sklearn TF-IDF 余弦）加权融合 Top-K + 相关性阈值声明 `not_in_ppt`
- 中文分词：纯 Python char 1/2-gram + 数学符号 token（无 jieba）

**质量如实**：检索**可运行、可溯源**，但词面/浅语义准度为片段级——数学密集内容里个别 Top-1 可能偏页（Top-K=5 内正确页通常仍命中）。"真语义向量"（bge/text2vec）需装 embedding 模型，属后续升级项，未做。

### M3 按指令出题
- `m3_validator.py`：SymPy 三重校验（有解 / 答案正确 / 难度符合），难度按嵌套层数 + 复合结构估算 L1-L3
- `m3_generator.py`：换元积分题型族（凑微分/复合凑微分/三角代换）参数化出题，去重，入库
- `m3_weakness_and_demo.py`：薄弱点倾斜接口 + 端到端 demo（M2 检索定位 → M3 出题 → 校验 → 入库）
- `m3_weakness_loop.py`：薄弱点闭环演示
- **`question_bank.json`**：题库（每题含 `_validation` 证据 + `evidence` + `ppt_pages` 溯源）

**质量如实**：出题是**参数化模板族**（换元积分一族），非"Agnes 自由生成"。难度为启发式估算，个别边界题可能 ±1 级。扩展其它题型（导数/极限/级数）需加 `_TEMPLATES`。

### M4 答案批改
- `m4_grading.py`：模式 A（简答 SymPy 等价，允许 +C）+ 模式 B（过程逐步比对定位首错步），等价判断**委托 M6 单一事实源**

**质量如实**：等价判断有 SymPy 符号证据；过程批改要求每行一个可解析表达式（不支持自然语言，同 AGENT_DESIGN §6 边界）。

### M5 错误诊断
- `m5_diagnosis.py`：6 类错误字典（换元未回代/漏常数C/符号错误/公式记错/计算错误/定义域错误），每种给纠正 + 同类练习

**质量如实**：按差值结构启发式归类，典型模式准确（工作流真实案例 `eˣ(x+1)`→`符号错误` 判对），新奇错误形态落兜底"计算错误"。错误集是 6 类封闭集。

### M6 数学验证
- `m6_math_verify.py`：SymPy 工具集（`antiderivative`/`check_derivative`/`equivalent`/`equivalent_strict`/`validate_indefinite`），所有结论带符号证据

**质量如实**：确定性可复算。

---

## 文件清单

```
AGENT_DESIGN.md            设计文档 (M1-M6 规格)
slides.json                M1 输出 (课件结构化)
question_bank.json         M3 输出 (题库)

m1_ppt_parser.py           M1 主解析器
m1_equation_decoder.py     M1 CEQU 解码器
wmf_to_png.py              M1 WMF→PNG
ocr_all.py                 M1 OCR 批量识别
m2_rag_search.py           M2 RAG 检索
m3_validator.py            M3 三重校验器
m3_generator.py            M3 出题器
m3_weakness_and_demo.py    M3 薄弱点倾斜 + demo
m3_weakness_loop.py        M3 闭环演示
m4_grading.py              M4 批改
m5_diagnosis.py            M5 诊断
m6_math_verify.py          M6 验证引擎 (单一事实源)
main.py                    编排层 (M1→M6 一条数据流 + 档案闭环)

换元积分法.pptx            转换后源文件 (52 页)
wmf_export/ + wmf_export/png/    288 张 WMF + 288 张 PNG
weakness_profile.json      薄弱点档案 (M4/M5 → M3 持久化通道, 运行时生成)
```

## 未做 / 升级项（如实）

1. **M2 真语义向量**：需装 embedding 模型（bge/text2vec，~几百 MB），当前为词面+浅语义。
2. **M1 公式整式还原**：需公式专用 OCR（LaTeX-OCR/Pix2Tex，torch）或视觉大模型，当前为片段级双通道。
3. **M3 自由出题**：参数化题型族可换元积分一族，Agnes 自由生成新题需接编排模型层（本机不代做）。
4. **题型扩展**：M3/M5 目前覆盖换元积分，导数/极限/级数/微分方程需扩展模板 + 错误字典。

## 编码/运行约定（Windows）

- Python 用 `python -X utf8`（中文输出）
- 复杂脚本写文件，不内联进 `python -c`（PS 引号嵌套解析失败）
- PS 管道会吞 stderr，取真实报错用 `2>&1` 直接看
- 中文路径在 PS 输出乱码，但字节级文件操作正常
