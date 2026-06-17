# Educoder 实验代码爬虫

自动登录 [Educoder（头歌）](https://www.educoder.net) 实践教学平台，爬取课堂中所有已提交的实验代码。

## 安装

```bash
pip install playwright
python -m playwright install chromium
```

## 用法

```bash
python educoder_scraper.py <手机号> <密码> <课堂ID> [输出文件]
```

### 示例

```bash
python educoder_scraper.py 13800000000 mypassword aimfyw6z
python educoder_scraper.py 13800000000 mypassword aimfyw6z output.json
```

### 获取课堂 ID

课堂 ID 即 URL 中的标识符：

```
https://www.educoder.net/classrooms/aimfyw6z/shixun_homework
                            ^^^^^^^^
                            课堂 ID
```

## 输出

生成 JSON 文件，包含每个实验的名称、关卡数、完整代码：

```json
[
  {
    "name": "推理与证明",
    "homework_id": 3597377,
    "student_work_id": 274562335,
    "challenge_count": 4,
    "codes": [
      {
        "code": "from sympy import symbols\n..."
      }
    ]
  }
]
```

## 原理

1. 使用 Playwright 启动 Chromium 浏览器
2. 自动填写手机号和密码登录
3. 拦截 API 获取实验列表（含自动翻页）
4. 逐个访问实验评论页，点击"复制代码"按钮
5. 通过系统剪贴板读取完整学生代码
6. 增量保存到 JSON 文件

## 注意事项

- 需要安装 Chromium（`playwright install chromium`）
- 进行中的实验只能获取模板代码，完整代码请等待实验截止
- 默认使用 headless 模式运行，如需观察过程可将 `headless=True` 改为 `headless=False`
