# -*- coding: utf-8 -*-
"""Quick screenshot of updated industry page"""
import sys, os, subprocess, time, socket
sys.stdout.reconfigure(encoding='utf-8')

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
output_dir = os.path.join(desktop, "仪表板截图")
os.makedirs(output_dir, exist_ok=True)

def port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

port = 8522
if port_in_use(port):
    for p in range(8523, 8600):
        if not port_in_use(p): port = p; break

os.system("taskkill /f /im streamlit.exe 2>nul")
time.sleep(1)

workdir = r"D:\狗哥AI学习\案例-数据驾驶舱需求\财务经营分析仪表板"
proc = subprocess.Popen(
    ["python", "-m", "streamlit", "run", "streamlit_app.py",
     "--server.port", str(port), "--server.headless", "true"],
    cwd=workdir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

base_url = f"http://localhost:{port}"
for i in range(30):
    time.sleep(1)
    if port_in_use(port):
        try:
            import urllib.request
            urllib.request.urlopen(base_url, timeout=3)
            break
        except: pass
time.sleep(3)

from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={"width": 1920, "height": 1080})
    page = context.new_page()

    # Industry page
    url = base_url + "/行业对比"
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        time.sleep(3)
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.5)
        fp = os.path.join(output_dir, "同行业对比_更新版.png")
        page.screenshot(path=fp, full_page=True)
        print(f"Saved: {fp}")
    except Exception as e:
        print(f"Error: {e}")

    browser.close()

proc.kill()
print("Done!")
