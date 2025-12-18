import os
import pickle
import base64
import re
import configparser
import tkinter as tk
from tkinter import messagebox, ttk
from time import sleep
from datetime import datetime as dt, timezone, timedelta

# Selenium & Google API
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --- 設定保存用 ---
CONFIG_FILE = 'settings.ini'

def save_settings(user, pw):
    config = configparser.ConfigParser()
    config['USER'] = {'username': user, 'password': pw}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        config.write(f)

def load_settings():
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE, encoding='utf-8')
        return config['USER'].get('username', ''), config['USER'].get('password', '')
    return '', ''

class ManabaEngine:
    def __init__(self, user, pw, log_func, progress_func):
        self.user = user
        self.pw = pw
        self.log = log_func
        self.progress = progress_func
        self.calendar_id = 'primary'
        self.sig = "[manaba-auto]"

    def run(self):
        try:
            self.log("--- 同期プロセス開始 ---")
            self.progress(5)
            
            # STEP1: manabaスキャン
            self.log("【1/2】manabaから課題を取得しています...")
            tasks, submitted = self.fetch_manaba()
            self.log(f"-> 未提出課題: {len(tasks)}件、提出済み: {len(submitted)}件を検出")
            self.progress(70)
            
            # STEP2: カレンダー同期
            self.log("【2/2】Googleカレンダーと同期しています...")
            self.sync_calendar(tasks, submitted)
            
            self.progress(100)
            self.log("--- すべての工程が完了しました ---")
            messagebox.showinfo("完了", f"同期完了！\n未提出課題 {len(tasks)}件を整理しました。")
        except Exception as e:
            self.log(f"✖ エラーが発生しました: {e}")
            messagebox.showerror("エラー", str(e))
        finally:
            self.progress(0)

    def fetch_manaba(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--lang=ja-JP')
        options.add_argument('--window-size=1000,700')
        # options.add_argument('--headless') # 画面を見たくない場合はコメントアウトを外す

        driver = webdriver.Chrome(options=options)
        auth = base64.b64encode(f"{self.user}:{self.pw}".encode()).decode()
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"Authorization": f"Basic {auth}"}})
        
        driver.get('https://slms.mi.sanno.ac.jp/ct/home')
        try:
            wait = WebDriverWait(driver, 5)
            wait.until(EC.presence_of_element_located((By.ID, "mainuserid"))).send_keys(self.user)
            driver.find_element(By.NAME, "password").send_keys(self.pw + Keys.ENTER)
        except: pass

        links = driver.find_elements(By.CSS_SELECTOR, 'td.course a[href*="course_"]:not(.courseweekly-fav)')
        urls = list(dict.fromkeys([l.get_attribute('href') for l in links]))
        
        results = {}
        submitted_list = []
        targets = [('_report', 'レポート'), ('_query', '小テスト'), ('_survey', 'アンケート')]

        total = len(urls)
        for i, base_url in enumerate(urls):
            self.progress(10 + (i / total * 60))
            driver.get(base_url)
            try:
                name = driver.find_element(By.ID, 'coursename').text
                self.log(f" > 解析中: {name}")
                for suffix, label in targets:
                    driver.get(base_url + suffix)
                    rows = driver.find_elements(By.TAG_NAME, 'tr')
                    for row in rows:
                        t = row.text
                        m = re.findall(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', t)
                        if '未提出' in t and '受付中' in t and m:
                            deadline = sorted(m)[-1]
                            results[(f"【提出：{label}】{name}", deadline)] = 1
                        elif any(x in t for x in ['提出済み', '回答済み', '済']):
                            submitted_list.append(f"{label}】{name}")
            except: continue
        
        driver.quit()
        final_tasks = [(t, dt.strptime(d, '%Y-%m-%d %H:%M').strftime("%Y-%m-%dT%H:%M:00")) for t, d in results.keys()]
        return final_tasks, list(set(submitted_list))

    def _get_calendar_service(self):
        import socket
        socket.setdefaulttimeout(30)
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token: creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try: creds.refresh(Request())
                except: creds = None
            if not creds:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', ['https://www.googleapis.com/auth/calendar'])
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token: pickle.dump(creds, token)
        return build('calendar', 'v3', credentials=creds)

    def sync_calendar(self, tasks, submitted_titles):
        service = self._get_calendar_service()
        now = dt.now(timezone.utc)
        
        self.log(">> 既存の予定を確認中...")
        events = service.events().list(calendarId=self.calendar_id, timeMin=(now - timedelta(days=60)).isoformat(), singleEvents=True).execute().get('items', [])
        
        processed_keys = {}
        for ev in events:
            summary = ev.get('summary', '')
            if "【提出：" not in summary: continue
            
            start_iso = ev['start'].get('dateTime', '')[:19]
            category_with_name = summary.split('：')[-1]
            ev_dt = dt.fromisoformat(ev['start'].get('dateTime').replace('Z', '+00:00'))
            
            # 提出済み、または期限切れの予定を削除
            if category_with_name in submitted_titles or ev_dt < now:
                service.events().delete(calendarId=self.calendar_id, eventId=ev['id']).execute()
                self.log(f" [削除済/期限切れ] {summary}")
                continue
            
            processed_keys[(summary.split('】')[-1], start_iso)] = ev['id']
            self.log(f" [継続] {summary}")

        # 新規課題を追加
        for title, deadline in tasks:
            if (title.split('】')[-1], deadline) not in processed_keys:
                event = {
                    'summary': title, 'description': self.sig,
                    'start': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'end': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'colorId': '3', # 緑色
                    'reminders': {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 60}]}
                }
                service.events().insert(calendarId=self.calendar_id, body=event).execute()
                self.log(f" [新規追加] {title}")

class SimpleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("manaba 同期ツール (改良版)")
        self.root.geometry("450x600")
        
        frame = tk.Frame(root, pady=15)
        frame.pack()
        
        tk.Label(frame, text="manaba ID:", font=("Yu Gothic", 10)).grid(row=0, column=0, sticky="e", padx=5)
        self.ent_user = tk.Entry(frame, width=25, font=("Consolas", 10))
        self.ent_user.grid(row=0, column=1, pady=2)
        
        tk.Label(frame, text="パスワード:", font=("Yu Gothic", 10)).grid(row=1, column=0, sticky="e", padx=5)
        self.ent_pw = tk.Entry(frame, width=25, show="*", font=("Consolas", 10))
        self.ent_pw.grid(row=1, column=1, pady=2)
        
        u, p = load_settings()
        self.ent_user.insert(0, u)
        self.ent_pw.insert(0, p)
        
        self.btn = tk.Button(root, text="同期を開始", command=self.start, width=25, height=2, bg="#4CAF50", fg="white", font=("Yu Gothic", 10, "bold"))
        self.btn.pack(pady=10)
        
        self.pb = ttk.Progressbar(root, length=380, mode='determinate')
        self.pb.pack(pady=5)
        
        self.log_box = tk.Text(root, height=18, width=55, font=("Yu Gothic", 9), bg="#F5F5F5", padx=10, pady=10)
        self.log_box.pack(padx=15, pady=10)

    def set_progress(self, val):
        self.pb['value'] = val
        self.root.update()

    def add_log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.root.update()

    def start(self):
        user, pw = self.ent_user.get(), self.ent_pw.get()
        if not user or not pw: return
        save_settings(user, pw)
        self.btn.config(state=tk.DISABLED, bg="#9E9E9E")
        self.log_box.delete('1.0', tk.END)
        engine = ManabaEngine(user, pw, self.add_log, self.set_progress)
        engine.run()
        self.btn.config(state=tk.NORMAL, bg="#4CAF50")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleApp(root)
    root.mainloop()