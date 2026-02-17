from playwright.sync_api import sync_playwright
from reportlab.pdfgen import canvas
from datetime import datetime
import sys
import re
import os
import subprocess
import win32gui
import win32con
import time

"""""""""""""""""""""
SETTING ENVIRONMENTS
"""""""""""""""""""""

TARGET_URL = "https://192.168.200.10"
IMG_SELECTOR = "img[src='/logo2.png']"
SEARCH_SELECTOR = "button:has([aria-label='search'])"
MAX_WAIT = 180  # 180 วินาที (3 นาที)
now = datetime.now()
day = now.day
month = now.month
buddhist_year = now.year + 543
year = str(buddhist_year)[-2:]
folder_name = f"{int(day)}.{int(month)}.{year}"
print("📁 โฟลเดอร์ย่อย:", folder_name)
print(type(folder_name))
download_dir = os.path.join(r"C:\Users\User\Downloads",folder_name)
ADOBE = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

"""""""""""""""""""""
CREATE FUNCTIONS
"""""""""""""""""""""

def minimize_playwright_chrome(wait=0.5):
    """
    Minimize Chrome window ที่ Playwright เปิด
    """
    time.sleep(wait)

    def enum_handler(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if "Chrome" in title:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)

    win32gui.EnumWindows(enum_handler, None)

def lock_user_input(page):
    page.evaluate("""
    () => {
        if (document.getElementById('user-lock')) return;

        const overlay = document.createElement('div');
        overlay.id = 'user-lock';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100%';
        overlay.style.height = '100%';
        overlay.style.zIndex = '999999';
        overlay.style.background = 'rgba(0,0,0,0)'; // โปร่งใส
        overlay.style.cursor = 'not-allowed';
        overlay.style.pointerEvents = 'all';

        document.body.appendChild(overlay);
    }
    """)
    print("🔒 ล็อกการคลิกจากผู้ใช้แล้ว")

def unlock_user_input(page):
    page.evaluate("""
    () => {
        document.getElementById('user-lock')?.remove();
    }
    """)
    print("🔓 ปลดล็อกการคลิกจากผู้ใช้แล้ว")

def print_blank_page(path='blank_page.pdf'):
    c = canvas.Canvas(path, pagesize=(595.2756, 841.8898))
    c.showPage()
    c.save()
    os.startfile(os.path.abspath(path), "print")
    
def print_pdf_adobe(pdf_path, printer=None, wait=10):
    if not os.path.exists(ADOBE):
        raise FileNotFoundError("❌ ไม่พบ Acrobat.exe")

    if not os.path.exists(pdf_path):
        raise FileNotFoundError("❌ ไม่พบไฟล์ PDF")

    cmd = [ADOBE, "/t", pdf_path]

    if printer:
        cmd.append(printer)

    subprocess.Popen(cmd)

    # รอให้ Acrobat ส่งงานพิมพ์เข้า queue
    time.sleep(wait)
    
def get_text_any_frame(page, selector, timeout=20000):
    """
    ดึง text จาก selector ไม่ว่าจะอยู่ main page หรือ iframe
    """
    end_time = time.time() + timeout / 1000

    while time.time() < end_time:
        for frame in page.frames:
            try:
                el = frame.query_selector(selector)
                if el:
                    text = el.text_content()
                    if text:
                        return text.strip()
            except:
                pass
        time.sleep(0.3)

    raise TimeoutError(f"❌ ไม่พบ element: {selector}")

def clean_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '', name)

"""""""""""""""""""""
MAIN CODE
"""""""""""""""""""""

MAX_RETRY = 5          # ลองเปิดสูงสุด 5 ครั้ง
WAIT_BETWEEN = 10      # รอ 10 วินาทีก่อนลองใหม่
GOTO_TIMEOUT = 30_000  # timeout ต่อครั้ง (ms)

with sync_playwright() as p:
    browser = p.chromium.launch(channel="chrome", headless=False)
    context = browser.new_context(ignore_https_errors=True)
    page = context.new_page()
    success = False
    
    for attempt in range(1, MAX_RETRY + 1):
        try:
            print(f"🌐 เปิดเว็บ (ครั้งที่ {attempt}/{MAX_RETRY}) → {TARGET_URL}")
            page.goto(TARGET_URL, timeout=GOTO_TIMEOUT)
            print("⏳ รอโหลดหน้า...")
            page.wait_for_load_state("networkidle", timeout=GOTO_TIMEOUT)
            print("✅ เปิดเว็บสำเร็จ")
            success = True
            break
        except TimeoutError:
            print("❌ โหลดหน้าไม่สำเร็จ (Timeout)")
        except Exception as e:
            print("❌ เปิดเว็บไม่สำเร็จ:", e)
        if attempt < MAX_RETRY:
            print(f"🔁 รอ {WAIT_BETWEEN} วินาที แล้วลองใหม่...\n")
            time.sleep(WAIT_BETWEEN)
        else:
            print("⛔ เปิดเว็บไม่สำเร็จครบจำนวนครั้งที่กำหนด")
    if not success:
        print("🛑 ยกเลิกการทำงานของโปรแกรม")
        browser.close()
        sys.exit(1)
    # 👉 ถ้าเปิดสำเร็จ โปรแกรมจะมาทำงานต่อด้านล่าง
    print("🚀 เริ่มทำงานขั้นถัดไป...")
    try:
        while True:
            # page.click("a:has-text('ข่าวรับ')")
            print("🔍 เช็คปุ่ม...")
            print("⏳ รอโหลดหน้า...")
            # 🔥 บังคับแตะ page
            page.evaluate("() => document.title")
            page.wait_for_selector(IMG_SELECTOR, timeout=10_000)
            print("✅ พบรูป logo2.png → ทำงานต่อ")
            while True:
                try :
                    time.sleep(1)
                    page.evaluate("() => document.title")  # กัน Chrome ถูกปิด
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.type("bart8", delay=100)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.type("Arty82526/", delay=100)
                    page.keyboard.press("Enter")
                    time.sleep(1)
                    while True:
                        try:
                            start_check = datetime.now()
                            print("start_check:", start_check)
                            try:
                                page.evaluate("() => document.title")  # กัน Chrome ถูกปิด
                                page.wait_for_selector("a:has-text('ข่าวรับ')",state="visible",timeout=30_000)
                                page.click("a:has-text('ข่าวรับ')")
                                time.sleep(1)
                                page.wait_for_selector(".ant-select-selection-item", timeout=10000)
                                page.click(".ant-select-selection-item")
                                # เลือก 50 / page
                                page.wait_for_selector("div[title='50 / page']", timeout=10000)
                                page.click("div[title='50 / page']")
                                print("🔍 กำลังหาปุ่ม เปิดอ่าน ...")
                                page.wait_for_selector("button:has-text('เปิดอ่าน')",timeout=10_000)
                                page.click("button:has-text('เปิดอ่าน')")
                                print("✅ พบปุ่ม เปิดอ่าน → คลิก")
                                break
                            except TimeoutError:
                                page.reload()
                                time.sleep(1)
                                page.evaluate("() => document.title")  # กัน Chrome ถูกปิด
                                page.keyboard.press("Tab")
                                time.sleep(1)
                                page.keyboard.type("bart8", delay=100)
                                page.keyboard.press("Tab")
                                time.sleep(1)
                                page.keyboard.type("Arty82526/", delay=100)
                                page.keyboard.press("Enter")
                                time.sleep(1)
                        except:
                            page.evaluate("() => document.title")  # กัน Chrome ถูกปิด
                            print("❌ ยังไม่พบ เปิดอ่าน")
                            print("⏳ รอ 30 นาที แล้ว refresh หน้า")
                            time.sleep(500)  # รอ 30 นาที
                            elapsed = datetime.now()
                            print("elapsed:", elapsed)
                            page.reload()
                            page.wait_for_load_state("networkidle")                
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.type("โอภาส", delay=100)
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.type("Arty82526/", delay=100)
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.press("Tab")
                    time.sleep(1)
                    page.keyboard.press("Enter")
                    time.sleep(1)    
                    page.wait_for_selector("label:has-text('ที่ข่าว :')", timeout=20000)
                    print("✅ พบ label ที่ข่าว : แล้ว")
                    
                    # ================== ใช้งาน ==================

                    # 🔹 ดึง "ที่ข่าว"
                    at_news = get_text_any_frame(page, "#news_atNews")
                    print("📌 ที่ข่าว1:", at_news)

                    # 🔹 ดึง "เรื่อง"
                    title_new = get_text_any_frame(page, "#news_titleNews")
                    print("📌 เรื่อง1:", title_new)
                           
                    if title_new == "":
                        page.wait_for_selector("#news_atNews", timeout=20000)
                        at_news = page.inner_text("#news_atNews").strip()
                        print("📌 ที่ข่าว2:", at_news)
                            
                        page.wait_for_selector("#news_titleNews", timeout=20000)
                        title_new = page.inner_text("#news_titleNews").strip()
                        print("📌 เรื่อง2:", title_new)
                    
                    if title_new == "":
                        page.wait_for_selector("#news_atNews", state="attached", timeout=20000)
                        at_news = page.locator("#news_atNews").text_content()
                        at_news = at_news.strip() if at_news else ""
                        print("📌 ที่ข่าว3:", at_news)

                        page.wait_for_selector("#news_titleNews", state="attached", timeout=20000)
                        title_new = page.locator("#news_titleNews").text_content()
                        title_new = title_new.strip() if title_new else ""
                        print("📌 เรื่อง3:", title_new)
                    
                    if title_new == "":
                        print("❌ ไม่พบ เรื่อง → โหลดหน้าใหม่")
                        time.sleep(500)
                        break
                    with page.expect_download() as download_info:
                        page.locator("img[src*='atfile3.png']").click()

                    download = download_info.value
                    safe_at_news = clean_filename(at_news)
                    safe_title_new = clean_filename(title_new)
                    new_filename = f"{safe_at_news}{safe_title_new}.pdf"
                    # 📂 โฟลเดอร์ปลายทาง (เปลี่ยนได้)
                    os.makedirs(download_dir, exist_ok=True)
                    full_path = os.path.join(download_dir, new_filename)
                    download.save_as(full_path)

                    try:
                        print("📄 suggested filename:", download.suggested_filename)
                        print("📂 saved to:", full_path)
                        print("🖨️ กำลังพิมพ์เอกสาร...")
                        print_pdf_adobe(full_path)    
                        time.sleep(1)
                        print_blank_page()
                        print("✅ สั่งพิมพ์เอกสารเรียบร้อย")
                    except Exception as e:  
                        print("❌ มีข้อผิดพลาดในการบันทึกหรือพิมพ์เอกสาร")
                        print(str(e))
                        break
                 
                    page.click("a:has-text('ข่าวรับ')")
                    start_time = time.time()
                    print(start_time)
                    
                    while True:
                        try:
                            page.wait_for_selector(SEARCH_SELECTOR, timeout=5000)
                            print("✅ พบปุ่ม Search แล้ว")
                            break
                        except TimeoutError:
                            elapsed = time.time() - start_time
                            if elapsed >= MAX_WAIT:
                                raise TimeoutError("❌ รอปุ่ม Search เกิน 3 นาทีแล้ว")
                            print("⏳ ยังไม่พบปุ่ม Search → รอต่อ")
                            time.sleep(1800)  # รอ 30 นาที
                except Exception as e:
                    print("❌ Chrome ถูกปิด → หยุด script")
                    print(str(e))
                    break
    except Exception as e:
        print("❌ Chrome ถูกปิด → หยุด script")
        print(str(e))
    except KeyboardInterrupt:
        print("🛑 หยุดโปรแกรมด้วยมือ (Ctrl+C)")
    finally:
        print("🧹 script จบการทำงาน")
        unlock_user_input(page)
        sys.exit(0)
