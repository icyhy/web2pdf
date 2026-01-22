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
    
    # 寻找“公众号”窗口
    print("正在寻找 '公众号' 窗口 (使用 win32 API)...", flush=True)
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

    # 寻找文章列表
    # 策略调整：由于无法直接识别 ListItem，改用“锚点定位法”
    # 文章列表通常包含日期（今天、昨天、星期X、X月X日），通过查找这些日期文本，然后点击其上方区域来打开文章
    print("正在扫描日期文本以定位文章...", flush=True)
    
    date_patterns = [
        r"^今天", r"^昨天", r"^星期[一二三四五六日]", 
        r"^\d{1,2}月\d{1,2}日", r"^\d{4}年\d{1,2}月\d{1,2}日",
        r"^\d+小时前", r"^\d+分钟前"
    ]
    
    potential_clicks = []
    
    # 遍历所有 Text 控件
    for child in main_win.descendants(control_type="Text"):
        text = child.window_text()
        if not text: continue
        
        # 检查是否匹配日期格式
        is_date = False
        for pat in date_patterns:
            if re.search(pat, text.strip()):
                is_date = True
                break
        
        if is_date:
            rect = child.rectangle()
            # 排除无效坐标
            if rect.width() > 0 and rect.height() > 0:
                # 排除明显不是文章列表的日期（比如顶部的系统时间，虽然微信窗口里通常没有）
                # 这里假设文章列表的日期都在某个区域内，暂时不做严格过滤
                potential_clicks.append((child, text, rect))

    # 按 Y 坐标排序（从上到下）
    potential_clicks.sort(key=lambda x: x[2].top)
    
    print(f"找到 {len(potential_clicks)} 个潜在的文章日期锚点。", flush=True)
    if not potential_clicks:
        print("未找到任何日期文本，无法定位文章。请确保窗口已滚动到文章列表区域。", flush=True)
        return

    # 确保 results 目录存在
    if not os.path.exists("results"):
        os.makedirs("results")
    
    # 获取当前所有窗口句柄集合（用于后续检测新窗口）
    
    with sync_playwright() as p:
        # 启动浏览器
        print("启动 Playwright 浏览器...", flush=True)
        browser = p.chromium.launch(headless=True)
        
        processed_count = 0
        
        for i, (item, text, rect) in enumerate(potential_clicks):
            if processed_count >= 10: 
                print("已处理 10 篇，停止执行。", flush=True)
                break
            
            print(f"\n[{i+1}] 锚点文本: {text}, 坐标: {rect}", flush=True)
            
            # 关键步骤：每次点击前，强制激活公众号列表窗口，防止被之前的文章窗口遮挡
            print("  正在将公众号窗口置于顶层...", flush=True)
            force_focus_window(main_win.handle)
            time.sleep(1.0) # 给一点时间刷新
            
            # 记录点击前的所有窗口句柄，用于检测弹窗
            # 优化：使用 win32 backend 快速获取句柄，比 uia 快得多
            before_handles = {w.handle for w in Desktop(backend="win32").windows()}
            print(f"  当前可见窗口数: {len(before_handles)}", flush=True)

            try:
                # 计算点击位置
                click_point = (rect.left + 50, rect.top - 50)
                print(f"  计算点击位置: {click_point}", flush=True)
                
                # 确保点击位置在窗口内
                if click_point[1] < main_win.rectangle().top:
                    print("  点击位置超出窗口顶部，跳过。", flush=True)
                    continue

                # 点击文章
                pywinauto.mouse.click(button='left', coords=click_point)
                print("  已点击文章，等待窗口弹出...", flush=True)
                
                article_win = None
                # 轮询检测新窗口
                for _ in range(20): # 最多等 10 秒
                    time.sleep(0.5)
                    # 同样使用 win32 backend 快速检测句柄变化
                    current_handles = {w.handle for w in Desktop(backend="win32").windows()}
                    new_handles = current_handles - before_handles
                    
                    if new_handles:
                        print(f"  监测到新句柄: {new_handles}", flush=True)
                        for h in new_handles:
                            try:
                                # 找到新句柄后，再用 UIA 连接它进行后续操作
                                # 这样只连接一个特定窗口，速度非常快
                                temp_win = Desktop(backend="uia").window(handle=h)
                                if temp_win.is_visible():
                                    r = temp_win.rectangle()
                                    # 文章窗口通常较大
                                    if r.width() > 300 and r.height() > 300:
                                        article_win = temp_win
                                        break
                            except:
                                continue
                    if article_win:
                        break

                if not article_win:
                    print("  ❌ 未检测到文章阅读窗口。根据要求，停止脚本执行。", flush=True)
                    browser.close()
                    return
                
                # 找到窗口后，激活它
                print(f"  已找到文章窗口: {article_win.window_text()}，正在处理...", flush=True)
                force_focus_window(article_win.handle)
                time.sleep(0.5)
                
                # 在 Wrapper 对象上查找控件，需要遍历 descendants
                # 寻找“更多”按钮
                more_btn = None
                print("  正在搜寻‘更多’按钮...", flush=True)
                
                # 获取窗口矩形，用于后续定位
                ar = article_win.rectangle()
                
                # 策略 1: 按名称搜寻
                for child in article_win.descendants(control_type="Button"):
                    name = child.window_text()
                    if name in ["更多", "More"]:
                        more_btn = child
                        break
                
                # 策略 2: 如果策略 1 失败，寻找右上角区域的按钮
                if not more_btn:
                    for child in article_win.descendants(control_type="Button"):
                        r = child.rectangle()
                        # 检查是否在右上角 (顶部 60 像素内，右侧 250 像素内)
                        if r.top >= ar.top and r.top < ar.top + 60 and \
                           r.right > ar.right - 250 and r.right <= ar.right - 100: # 排除关闭/最大化/最小化按钮
                            print(f"  -> 发现右上角疑似按钮: '{child.window_text()}' at {r}", flush=True)
                            more_btn = child
                            break
                
                # 记录点击“更多”前的句柄，用于快速锁定菜单
                before_menu_handles = {w.handle for w in Desktop(backend="win32").windows()}
                
                if not more_btn:
                    print("  无法定位‘更多’按钮控件，尝试坐标点击右上角三个点区域...", flush=True)
                    more_coords = (ar.right - 120, ar.top + 25)
                    pywinauto.mouse.click(button='left', coords=more_coords)
                else:
                    print(f"  找到‘更多’按钮: '{more_btn.window_text()}'，正在点击...", flush=True)
                    more_btn.click_input()
                
                print("  正在寻找‘复制链接’菜单项...", flush=True)
                copy_btn = None
                try:
                    # 策略：通过句柄差异快速锁定新弹出的菜单窗口
                    for attempt in range(10): # 最多等 5 秒
                        time.sleep(0.5)
                        current_menu_handles = {w.handle for w in Desktop(backend="win32").windows()}
                        new_menu_handles = current_menu_handles - before_menu_handles
                        
                        if new_menu_handles:
                            for h in new_menu_handles:
                                try:
                                    # 菜单窗口通常是小窗口且可见
                                    temp_menu = Desktop(backend="uia").window(handle=h)
                                    if temp_menu.is_visible():
                                        # 遍历其子项寻找“复制链接”
                                        items = temp_menu.descendants(control_type="MenuItem")
                                        for mi in items:
                                            if mi.window_text() in ["复制链接", "Copy Link"]:
                                                copy_btn = mi
                                                break
                                    if copy_btn: break
                                except:
                                    continue
                        
                        if copy_btn: break
                        
                        # 备选：如果没检测到新句柄，直接在文章窗口下搜寻（有些版本菜单不是独立窗口）
                        if attempt > 4: # 2.5 秒后尝试备选
                            items = article_win.descendants(control_type="MenuItem")
                            for mi in items:
                                if mi.window_text() in ["复制链接", "Copy Link"]:
                                    copy_btn = mi
                                    break
                        if copy_btn: break
                except Exception as e:
                    print(f"  寻找菜单项时出错: {e}", flush=True)
                    
                if copy_btn:
                    copy_btn.click_input()
                    print("  已点击‘复制链接’", flush=True)
                    time.sleep(0.5)
                    
                    url = pyperclip.paste()
                    print(f"  获取到 URL: {url[:50]}...", flush=True)
                    
                    if url.startswith("http"):
                        print("  正在生成 PDF...", flush=True)
                        page = browser.new_page()
                        try:
                            # 增加超时时间并等待网络空闲
                            page.goto(url, wait_until="networkidle", timeout=60000)
                            
                            # 解决配图加载不全：模拟人工平滑滚动，触发所有懒加载图片
                            print("  正在触发图片懒加载...", flush=True)
                            page.evaluate("""
                                async () => {
                                    await new Promise((resolve) => {
                                        let totalHeight = 0;
                                        let distance = 400; // 每次滚动 400 像素
                                        let timer = setInterval(() => {
                                            let scrollHeight = document.body.scrollHeight;
                                            window.scrollBy(0, distance);
                                            totalHeight += distance;
                                            if (totalHeight >= scrollHeight) {
                                                clearInterval(timer);
                                                resolve();
                                            }
                                        }, 200); // 每 200 毫秒滚动一次
                                    });
                                }
                            """)
                            
                            # 滚动回顶部并额外等待一会，确保所有图片完成渲染
                            page.evaluate("window.scrollTo(0, 0)")
                            time.sleep(2) 
                            
                            # 获取真实标题（来自网页，比窗口标题准确）
                            article_title = page.title()
                            print(f"  获取到文章真实标题: {article_title}", flush=True)
                            
                            safe_name = sanitize_filename(article_title or f"article_{i}")
                            # 再次防止标题为空
                            if not safe_name or safe_name.isspace():
                                safe_name = f"article_{i}_{int(time.time())}"
                                
                            pdf_path = os.path.join("results", f"{safe_name}.pdf")
                            
                            # 打印 PDF
                            page.pdf(path=pdf_path, print_background=True)
                            print(f"  ✅ PDF 已保存: {pdf_path}", flush=True)
                            processed_count += 1
                        except Exception as pe:
                            print(f"  ❌ PDF 生成失败: {pe}", flush=True)
                        finally:
                            page.close()
                    else:
                        print("  剪贴板内容不是有效的 URL。", flush=True)
                else:
                    print("  未找到‘复制链接’菜单项。", flush=True)

                # 关闭文章窗口
                # 用户要求：处理完一篇文章后再点击下一篇文章，需要关闭文章阅读窗口
                print("  正在关闭文章窗口...", flush=True)
                try:
                    # 使用 win32gui 关闭窗口，更加鲁棒
                    if win32gui.IsWindow(article_win.handle):
                        win32gui.PostMessage(article_win.handle, win32con.WM_CLOSE, 0, 0)
                except Exception as e:
                    print(f"  关闭窗口失败: {e}", flush=True)
                
                time.sleep(1.5) # 等待窗口关闭动画
                
            except Exception as e:
                print(f"  处理过程中出错: {e}", flush=True)
                try:
                    # 再次检查，防止误关
                    if win32gui.IsWindow(article_win.handle):
                         win32gui.PostMessage(article_win.handle, win32con.WM_CLOSE, 0, 0)
                except:
                    pass

        browser.close()
        print(f"\n任务完成！共处理 {processed_count} 篇文章。", flush=True)

if __name__ == "__main__":
    main()
