import time
import os
import pyperclip
import re
import pywinauto.mouse
import win32gui
import win32con
from pywinauto import Desktop, Application
from playwright.sync_api import sync_playwright

def sanitize_filename(name):
    return "".join([c for c in name if c.isalpha() or c.isdigit() or c in (' ', '-', '_', '.')]).rstrip()

def force_focus_window(hwnd):
    try:
        if not win32gui.IsWindow(hwnd):
            print(f"  [Focus] 窗口句柄 {hwnd} 已失效", flush=True)
            return

        # 如果最小化了，恢复它
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        # 尝试最简单直接的置顶
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        
        # 使用 BringWindowToTop 替代 SetForegroundWindow 以减少被系统拦截的概率
        win32gui.BringWindowToTop(hwnd)
        
        # 尝试设置焦点
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            # SetForegroundWindow 经常会因为系统权限拦截而报错，忽略它
            pass
            
        time.sleep(0.2)
    except Exception as e:
        print(f"  [Focus] 尝试激活窗口失败: {e}", flush=True)

def main():
    print("脚本开始执行...", flush=True)
    
    # 用户输入控制
    print("\n" + "="*30)
    print("请输入要保存的文章数量：")
    print("  - 数字 (如 3): 保存最新的 3 篇")
    print("  - all: 保存所有可见及滚动可见的文章")
    print("  - 直接回车: 默认保存 10 篇")
    print("="*30)
    user_input = input("请输入: ").strip().lower()
    
    if user_input == "all":
        target_count = float('inf')
        print("设定为保存【所有】文章。", flush=True)
    elif user_input.isdigit():
        target_count = int(user_input)
        print(f"设定为保存最新的【{target_count}】篇文章。", flush=True)
    else:
        target_count = 10
        print("使用默认设置：保存最新的【10】篇文章。", flush=True)

    # 寻找“公众号”窗口
    print("\n正在寻找 '公众号' 窗口 (使用 win32 API)...", flush=True)
    main_win = None
    
    # 使用 win32gui 快速查找句柄
    def find_wechat_window(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "公众号" in title:
                ctx.append(hwnd)

    hwnds = []
    win32gui.EnumWindows(find_wechat_window, hwnds)
    
    if hwnds:
        print(f"找到 {len(hwnds)} 个候选窗口，使用第一个...", flush=True)
        # 找到句柄后，再用 UIA 连接
        try:
            main_win = Desktop(backend="uia").window(handle=hwnds[0])
            print(f"已连接到窗口: {main_win.window_text()}", flush=True)
        except Exception as e:
            print(f"连接窗口失败: {e}", flush=True)
            return
    else:
        print("未找到 '公众号' 窗口。请确保该窗口已打开（如图所示）。", flush=True)
        return

    try:
        force_focus_window(main_win.handle)
        print(f"已锁定并置顶窗口: {main_win.window_text()}", flush=True)
        
    except Exception as e:
        print(f"查找窗口时出错: {e}", flush=True)
        return

    # 锚点定位法：查找日期文本作为文章定位参考
    date_patterns = [
        r"^今天", r"^昨天", r"^星期[一二三四五六日]", 
        r"^\d{1,2}月\d{1,2}日", r"^\d{4}年\d{1,2}月\d{1,2}日",
        r"^\d+小时前", r"^\d+分钟前"
    ]

    def get_anchors():
        anchors = []
        for child in main_win.descendants(control_type="Text"):
            text = child.window_text()
            if not text: continue
            is_date = False
            for pat in date_patterns:
                if re.search(pat, text.strip()):
                    is_date = True
                    break
            if is_date:
                rect = child.rectangle()
                if rect.width() > 0 and rect.height() > 0:
                    anchors.append((child, text, rect))
        anchors.sort(key=lambda x: x[2].top)
        return anchors

    # 确保 results 目录存在
    if not os.path.exists("results"):
        os.makedirs("results")
    
    processed_urls = set()
    processed_count = 0
    last_anchors_signature = None # 用于检测是否到底

    with sync_playwright() as p:
        print("启动 Playwright 浏览器...", flush=True)
        browser = p.chromium.launch(headless=True)
        
        while processed_count < target_count:
            print(f"\n--- 正在扫描文章列表 (已处理: {processed_count}/{target_count}) ---", flush=True)
            force_focus_window(main_win.handle)
            time.sleep(0.5)
            
            potential_clicks = get_anchors()
            
            # 检查是否到底（如果锚点没变，说明滚动无效或到底了）
            current_signature = [(a[1], a[2].top) for a in potential_clicks]
            if current_signature == last_anchors_signature:
                print("文章列表已到底部或无法滚动，停止执行。", flush=True)
                break
            last_anchors_signature = current_signature

            if not potential_clicks:
                print("未找到任何文章锚点，尝试滚动...", flush=True)
                main_win.set_focus()
                # 使用 {PGDN} 替代 {PAGEDOWN}，pywinauto 的正确代码是 {PGDN}
                main_win.type_keys("{PGDN}")
                time.sleep(2)
                continue

            found_new_in_current_view = False
            for i, (item, text, rect) in enumerate(potential_clicks):
                if processed_count >= target_count:
                    break
                
                print(f"\n[检查锚点] {text} at {rect}", flush=True)
                
                # 计算点击位置
                click_point = (rect.left + 50, rect.top - 50)
                
                # 确保点击位置在窗口内
                wr = main_win.rectangle()
                if click_point[1] < wr.top + 50 or click_point[1] > wr.bottom - 20:
                    # 只有在还没处理够的情况下，才考虑跳过边缘锚点
                    print("  锚点位置在窗口边缘，跳过并待滚动后处理。", flush=True)
                    continue

                # 记录点击前的句柄
                before_handles = {w.handle for w in Desktop(backend="win32").windows()}
                
                # 点击文章
                force_focus_window(main_win.handle)
                pywinauto.mouse.click(button='left', coords=click_point)
                
                article_win = None
                for _ in range(15): # 最多等 7.5 秒
                    time.sleep(0.5)
                    current_handles = {w.handle for w in Desktop(backend="win32").windows()}
                    new_handles = current_handles - before_handles
                    if new_handles:
                        for h in new_handles:
                            try:
                                temp_win = Desktop(backend="uia").window(handle=h)
                                if temp_win.is_visible():
                                    r = temp_win.rectangle()
                                    if r.width() > 300 and r.height() > 300:
                                        article_win = temp_win
                                        break
                            except: continue
                    if article_win: break

                if not article_win:
                    print("  ❌ 点击后未检测到文章窗口。", flush=True)
                    continue

                # 获取 URL 并检查重复
                url = None
                try:
                    force_focus_window(article_win.handle)
                    # 寻找“更多”并点击“复制链接”
                    ar = article_win.rectangle()
                    # 记录菜单弹出前的句柄
                    before_menu_handles = {w.handle for w in Desktop(backend="win32").windows()}
                    
                    # 寻找更多按钮
                    more_btn = None
                    for child in article_win.descendants(control_type="Button"):
                        if child.window_text() in ["更多", "More"]:
                            more_btn = child; break
                    
                    if not more_btn:
                        pywinauto.mouse.click(button='left', coords=(ar.right - 120, ar.top + 25))
                    else:
                        more_btn.click_input()
                    
                    # 寻找复制链接
                    copy_btn = None
                    for _ in range(10):
                        time.sleep(0.4)
                        curr_m_h = {w.handle for w in Desktop(backend="win32").windows()}
                        new_m_h = curr_m_h - before_menu_handles
                        if new_m_h:
                            for h in new_m_h:
                                try:
                                    m_win = Desktop(backend="uia").window(handle=h)
                                    items = m_win.descendants(control_type="MenuItem")
                                    for mi in items:
                                        if mi.window_text() in ["复制链接", "Copy Link"]:
                                            copy_btn = mi; break
                                    if copy_btn: break
                                except: continue
                        if copy_btn: break
                    
                    if copy_btn:
                        copy_btn.click_input()
                        time.sleep(0.5)
                        url = pyperclip.paste()
                except Exception as e:
                    print(f"  获取 URL 出错: {e}", flush=True)

                # 关闭文章窗口
                try:
                    if article_win and win32gui.IsWindow(article_win.handle):
                        win32gui.PostMessage(article_win.handle, win32con.WM_CLOSE, 0, 0)
                except: pass

                if not url or not url.startswith("http"):
                    print("  无法获取有效 URL，跳过。", flush=True)
                    continue

                if url in processed_urls:
                    print(f"  [已跳过] 该文章已处理过: {url[:60]}...", flush=True)
                    continue
                
                # 处理新文章
                processed_urls.add(url)
                found_new_in_current_view = True
                print(f"  [新文章] 正在生成 PDF (第 {processed_count+1} 篇)...", flush=True)
                
                # --- PDF 生成逻辑 ---
                page = browser.new_page()
                try:
                    page.goto(url, wait_until="networkidle", timeout=60000)
                    page.evaluate("""async () => {
                        await new Promise((resolve) => {
                            let totalHeight = 0; let distance = 400;
                            let timer = setInterval(() => {
                                let scrollHeight = document.body.scrollHeight;
                                window.scrollBy(0, distance); totalHeight += distance;
                                if (totalHeight >= scrollHeight) { clearInterval(timer); resolve(); }
                            }, 200);
                        });
                    }""")
                    page.evaluate("window.scrollTo(0, 0)")
                    time.sleep(2)
                    title = page.title()
                    safe_name = sanitize_filename(title or f"article_{processed_count}")
                    if not safe_name.strip(): safe_name = f"article_{int(time.time())}"
                    pdf_path = os.path.join("results", f"{safe_name}.pdf")
                    page.pdf(path=pdf_path, print_background=True)
                    print(f"  ✅ 已保存: {pdf_path}", flush=True)
                    processed_count += 1
                except Exception as pe:
                    print(f"  ❌ PDF 生成失败: {pe}", flush=True)
                finally:
                    page.close()
                # ------------------

            if processed_count < target_count:
                print("\n正在向下滚动以加载更多文章...", flush=True)
                main_win.set_focus()
                # 使用 {PGDN} 替代 {PAGEDOWN}，pywinauto 的正确代码是 {PGDN}
                main_win.type_keys("{PGDN}")
                time.sleep(2.5) # 等待加载
            
        browser.close()
        print(f"\n任务完成！共处理 {processed_count} 篇文章。", flush=True)

if __name__ == "__main__":
    main()
