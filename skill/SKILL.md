---
name: educoder-report
description: >
  Automate the full workflow of scraping Educoder (头歌) experiment code and generating a
  formatted discrete mathematics programming practice report (.docx). Use this skill whenever
  the user mentions Educoder reports, discrete math reports, 离散数学报告, scraping experiment
  code from educoder.net, or generating programming practice reports from classroom homework.
  Also triggers for requests to batch-download submitted code from Educoder classrooms.
---

# Educoder 实验报告自动生成器

完整的 Educoder 实验代码爬取 + 离散数学编程实践报告生成工作流。

## 前置条件

```bash
pip install playwright python-docx
python -m playwright install chromium
```

## 工作流概览

```
实验报告生成流程:
  Step 1: 爬取代码 → educoder_data.json
  Step 2: 截取截图 → screenshots/
  Step 3: 生成报告 → 离散数学编程实践报告.docx
```

## Step 1: 爬取实验代码

使用 `scripts/educoder_scraper.py` 从 Educoder 课堂爬取所有实验代码。

```bash
python scripts/educoder_scraper.py <手机号> <密码> <课堂ID> [输出文件]
```

**原理：**
1. Playwright 启动 Chromium，自动填写手机号和密码登录
2. 拦截 API 获取实验列表（含自动翻页，处理超过 20 个实验的课堂）
3. 逐个访问实验评论页，点击"复制代码"按钮
4. 通过系统剪贴板读取完整学生代码
5. 增量保存到 JSON 文件

**输出格式：**
```json
[
  {
    "name": "推理与证明",
    "homework_id": 3597377,
    "student_work_id": 274562335,
    "challenge_count": 4,
    "codes": [{"code": "from sympy import symbols\n..."}]
  }
]
```

**关键细节：**
- 课堂 ID 即 URL 中的标识符：`https://www.educoder.net/classrooms/aimfyw6z/shixun_homework` 中的 `aimfyw6z`
- 默认只抓取已截止（"已截止"）的实验，进行中的实验只返回模板代码
- 如需修改，编辑脚本中的 `to_scrape` 逻辑
- 抓取过程自动跳过未提交的实验

## Step 2: 截取通过截图

使用 Playwright 截取每个实验评论页的完整截图，作为报告中"100%通过"的证明。

```python
# 对每个实验访问 comment 页面并截图
page.goto(f'https://www.educoder.net/classrooms/{classroom_id}/shixun_homework/{hw_id}/{sw_id}/comment')
page.screenshot(path=f'{name}.png', full_page=True)
```

截图保存到 `screenshots/` 目录，按实验序号命名。

## Step 3: 生成报告 docx

使用 `scripts/build_report.py` 基于爬取的数据生成格式化报告。

### 报告结构

按照参考模板格式，生成包含以下内容的 .docx 文件：

```
封面（专业、班级、学号、姓名、完成日期）
目录（TOC 域，打开 Word 后右键更新）
1 数理逻辑
  1.1 实验目的
  1.2 实验内容
    1.2.1 推理与证明
    1.2.2 命题逻辑
    1.2.3 逻辑与AI知识表示
  1.3 实验过程
    1.3.1 推理与证明
      - 解题思路（约100字）
      - 代码（Times New Roman 小四 单倍行距）
      - 通过截图
    ...
2 集合及其运算
...
9 总结与课程建议
  9.1 课程总结
  9.2 问题及建议
```

### 格式规范

**必须严格遵循以下格式：**

| 元素 | 字体 | 字号 | 行距 | 其他 |
|------|------|------|------|------|
| 标题1 | 黑体 | 三号(16pt) | - | 加粗，黑色 |
| 标题2 | 黑体 | 四号(14pt) | - | 黑色 |
| 标题3 | 黑体 | 小四(12pt) | - | 黑色 |
| 正文 | 宋体 | 小四(12pt) | 1.5倍 | 首行缩进2字符 |
| 代码 | Times New Roman | 小四(12pt) | 单倍 | 使用"代码"样式 |

### 代码处理

- 使用"复制代码"按钮 + 剪贴板 API 获取完整代码（不是从 textarea 或 pre/code 元素读取）
- 代码以"代码"样式插入文档，字体 Times New Roman 小四，单倍行距
- 每段代码前附解题思路分析（约100字），描述算法原理和关键步骤
- 代码中的空行和注释去除后再插入

### 解题分析编写规范

为每道题目的每一关编写解题分析，要求：
- 每条 80-120 字（中英文混合计数）
- 包含：算法名称、核心数据结构、关键步骤、复杂度分析
- 描述具体的技术细节，不泛泛而谈
- 参考格式："使用xxx算法/数据结构，通过xxx步骤实现xxx。时间复杂度O(xxx)，空间复杂度O(xxx)。"

### 实验分类映射

将实验按 Educoder 课堂左侧分类组织：

```python
CATEGORIES = {
    '数理逻辑': ['推理与证明', '命题逻辑', '逻辑与AI知识表示'],
    '集合及其运算': ['集合的表示和基本运算'],
    '关系表示及其运算': ['特殊关系', '二元关系'],
    '函数': ['函数的运算', '函数的判断'],
    '初等数论': ['除法和模运算', '素数和最大公因数', '同余方程', '密码学'],
    '组合数学': ['计数基础', '鸽巢原理', '排列与组合', '生成函数'],
    '图论': ['图论的相关定义', '图论的应用', '特殊图', '树'],
    '选做': ['数论'],
}
```

### 目录生成

使用 Word TOC 域代码插入目录。在封面后插入：

```python
# 插入 TOC 域: \o "1-3" 表示收集 1-3 级标题
p = doc.add_paragraph()
run = p.add_run()
fldChar = OxmlElement('w:fldChar')
fldChar.set(qn('w:fldCharType'), 'begin')
run._r.append(fldChar)
# ... instrText: ' TOC \\o "1-3" \\h \\z '
# ... separator + end
```

打开生成的 docx 后，右键 TOC 区域 → "更新域" → "更新整个目录" 即可生成完整目录。

## 完整脚本示例

以下是一个完整的端到端示例，展示如何将三步串联：

```python
# 1. 爬取数据
import subprocess
subprocess.run([
    'python', 'scripts/educoder_scraper.py',
    '13800000000', 'password', 'classroom_id', 'data.json'
])

# 2. 截取截图（参考 Step 2 中的 Playwright 代码）

# 3. 生成报告
# 修改 build_report.py 中的配置（分类映射、解题分析等）后运行
subprocess.run(['python', 'scripts/build_report.py'])
```

## 常见问题

**Q: 为什么进行中的实验只有模板代码？**
A: Educoder 在实验截止前隐藏学生的完整提交代码，comment 页面只显示代码模板。等实验状态变为"已截止"后重新抓取即可。

**Q: 为什么代码是空的？**
A: 检查 Educoder 页面是否有"复制代码"按钮。部分旧版实验可能没有此按钮，需要从 textarea 元素回退提取。

**Q: 截图插入后报告文件过大？**
A: 可在 `build_report.py` 中调整 `Inches(5.5)` 参数减小图片宽度，或将 full_page 截图改为仅截取视口区域。

**Q: 如何适配其他课程？**
A: 修改 `CATEGORIES`、`ANALYSIS`、`CHALLENGE_NAMES` 三个字典以匹配目标课程的实验名称和分类。
