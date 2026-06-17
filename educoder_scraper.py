"""
Educoder 实验代码爬虫
用法: python educoder_scraper.py <手机号> <密码> <课堂ID> [输出文件]

示例: python educoder_scraper.py 13800000000 mypassword aimfyw6z
"""
import json, time, os, re, sys
from playwright.sync_api import sync_playwright


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def login(page, username, password):
    """登录 Educoder，返回是否成功"""
    page.goto("https://www.educoder.net/login", wait_until="networkidle")
    page.wait_for_timeout(3000)
    page.locator('#login').click()
    page.wait_for_timeout(200)
    page.locator('#login').fill('')
    page.locator('#login').type(username, delay=50)
    page.locator('#password').click()
    page.wait_for_timeout(200)
    page.locator('#password').fill('')
    page.locator('#password').type(password, delay=50)
    page.locator('button:has-text("登")').first.click()
    page.wait_for_timeout(5000)
    return "login" not in page.url


def get_homework_list(page, classroom_id):
    """从 API 拦截获取实验列表（含翻页），返回 [{name, homework_id, student_work_id, ...}]"""
    all_homeworks = []

    def on_response(resp):
        if 'homework_commons.json' in resp.url:
            try:
                body = resp.json()
                hw_list = body.get('homeworks', [])
                for hw in hw_list:
                    if hw not in all_homeworks:
                        all_homeworks.append(hw)
            except:
                pass

    page.on('response', on_response)

    url = f"https://www.educoder.net/classrooms/{classroom_id}/shixun_homework"
    page.goto(url, wait_until="networkidle")
    page.wait_for_timeout(10000)

    # 检查是否有第2页
    page2_btn = page.locator('.ant-pagination-item-2').first
    if page2_btn.is_visible():
        log("  发现第2页，正在翻页...")
        page2_btn.click()
        page.wait_for_timeout(5000)

        # 检查是否有第3页
        page3_btn = page.locator('.ant-pagination-item-3').first
        if page3_btn.is_visible():
            page3_btn.click()
            page.wait_for_timeout(5000)

    if not all_homeworks:
        return None

    result = []
    for hw in all_homeworks:
        name = hw.get('name', '')
        status = hw.get('status', [])
        is_submitted = '已截止' in str(status)
        result.append({
            'name': name,
            'homework_id': hw.get('homework_id', ''),
            'student_work_id': hw.get('student_work_id', ''),
            'challenge_count': hw.get('challenge_count', 0),
            'finished_count': hw.get('finished_challenge_count', 0),
            'is_submitted': is_submitted,
            'status': status,
            'upper_category': hw.get('upper_category_name', ''),
        })
        log(f"  {'[已提交]' if is_submitted else '[进行中]'} {name} "
            f"({hw.get('finished_challenge_count',0)}/{hw.get('challenge_count',0)})")

    return result


def scrape_experiment(page, classroom_id, hw_id, sw_id):
    """抓取单个实验的全部代码（通过点击'复制代码'按钮 + 剪贴板）"""
    url = f"https://www.educoder.net/classrooms/{classroom_id}/shixun_homework/{hw_id}/{sw_id}/comment"
    page.goto(url, wait_until="networkidle")
    page.wait_for_timeout(8000)

    copy_btns = page.locator('text=复制代码').all()

    if not copy_btns:
        # 可能未提交，尝试 textarea 提取
        codes = []
        for ta in page.locator('textarea').all():
            try:
                val = ta.input_value()
                if val.strip() and len(val.strip()) > 20:
                    codes.append(val.strip())
            except:
                pass
        return codes

    page.context.grant_permissions(['clipboard-read', 'clipboard-write'])
    all_codes = []

    for i in range(len(copy_btns)):
        try:
            btn = page.locator('text=复制代码').nth(i)
            if not btn.is_visible():
                continue
            btn.click()
            page.wait_for_timeout(1500)
            code = page.evaluate("() => navigator.clipboard.readText()")
            if code and len(code.strip()) > 20:
                all_codes.append(code.strip())
        except Exception as e:
            log(f"    第{i+1}关复制失败: {str(e)[:100]}")

    return all_codes


def scrape_classroom(username, password, classroom_id, output_file):
    """主函数：登录 -> 获取实验列表 -> 逐个抓取 -> 保存"""
    log(f"课堂: {classroom_id}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1400, "height": 900},
            locale="zh-CN"
        )
        context.grant_permissions(['clipboard-read', 'clipboard-write'])
        page = context.new_page()

        # 1. 登录
        log("登录中...")
        if not login(page, username, password):
            log("登录失败，请检查账号密码")
            browser.close()
            return False
        log("登录成功")

        # 2. 获取实验列表
        log("获取实验列表...")
        experiments = get_homework_list(page, classroom_id)
        if experiments is None:
            log("获取实验列表失败")
            browser.close()
            return False

        submitted = [e for e in experiments if e['is_submitted']]
        in_progress = [e for e in experiments if not e['is_submitted']]
        log(f"共 {len(experiments)} 个实验，其中 {len(submitted)} 个已提交, {len(in_progress)} 个进行中")
        log("提示：进行中的实验只能抓取到模板代码，完整代码需实验截止后获取\n")

        # 默认抓取已提交的，若全部提交中则抓全部
        to_scrape = submitted if submitted else experiments

        # 3. 逐个抓取
        all_data = []
        for i, exp in enumerate(to_scrape):
            name = exp['name']
            hw_id = exp['homework_id']
            sw_id = exp['student_work_id']
            challenge_count = exp['challenge_count']

            log(f"\n[{i+1}/{len(to_scrape)}] {name} ({challenge_count}关)")
            codes = scrape_experiment(page, classroom_id, hw_id, sw_id)
            log(f"  获取 {len(codes)} 个代码块")

            all_data.append({
                'name': name,
                'homework_id': hw_id,
                'student_work_id': sw_id,
                'challenge_count': challenge_count,
                'codes': [{'code': c} for c in codes]
            })

            # 增量保存
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)

            page.wait_for_timeout(1000)

        browser.close()

    # 4. 汇总
    total_codes = sum(len(exp.get('codes', [])) for exp in all_data)
    log(f"\n完成: {len(all_data)} 个实验, {total_codes} 个代码块")
    log(f"保存至: {output_file}")
    return True


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    username = sys.argv[1]
    password = sys.argv[2]
    classroom_id = sys.argv[3]
    output_file = sys.argv[4] if len(sys.argv) > 4 else f"educoder_{classroom_id}.json"

    success = scrape_classroom(username, password, classroom_id, output_file)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
