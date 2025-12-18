import os
import pickle
import base64
import re
import configparser
import threading  # 追加: 非同期処理用
import tkinter as tk
from tkinter import messagebox, ttk
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
        # 処理全体をtry-catchで囲み、最後に必ずドライバを閉じるようにする
        driver = None
        try:
            self.log("--- 同期プロセス開始 ---")
            self.progress(5)
            
            # STEP1: manabaスキャン
            self.log("【1/2】manabaから課題を取得しています...")
            
            # ブラウザ設定
            options = webdriver.ChromeOptions()
            options.add_argument('--lang=ja-JP')
            # 画面を表示したくない場合は以下のコメントを外す
            # options.add_argument('--headless') 
            
            driver = webdriver.Chrome(options=options)
            tasks, submitted = self.fetch_manaba(driver)
            
            # ドライバーはここで用済みなので閉じる
            driver.quit()
            driver = None

            self.log(f"-> 未提出課題: {len(tasks)}件、提出済み: {len(submitted)}件を検出")
            self.progress(60)
            
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
            if driver:
                driver.quit()
            self.progress(0)

    def fetch_manaba(self, driver):
        # Basic認証用ヘッダー
        auth = base64.b64encode(f"{self.user}:{self.pw}".encode()).decode()
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"Authorization": f"Basic {auth}"}})
        
        driver.get('https://slms.mi.sanno.ac.jp/ct/home')
        
        # ログイン処理（要素待機）
        try:
            wait = WebDriverWait(driver, 10)
            # すでにログイン済みでない場合のみ入力
            if len(driver.find_elements(By.ID, "mainuserid")) > 0:
                driver.find_element(By.ID, "mainuserid").send_keys(self.user)
                driver.find_element(By.NAME, "password").send_keys(self.pw + Keys.ENTER)
        except:
            pass # 既にログイン済み、あるいはBasic認証で通過した場合

        # ログイン成功判定（コース一覧があるか）
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'td.course')))
        except:
            raise Exception("manabaへのログインに失敗しました。ID/PWを確認してください。")

        links = driver.find_elements(By.CSS_SELECTOR, 'td.course a[href*="course_"]:not(.courseweekly-fav)')
        urls = list(dict.fromkeys([l.get_attribute('href') for l in links]))
        
        results = {}
        submitted_list = []
        targets = [('_report', 'レポート'), ('_query', '小テスト'), ('_survey', 'アンケート')]

        total = len(urls)
        for i, base_url in enumerate(urls):
            # プログレスバー更新 (10% -> 60%)
            current_progress = 10 + (i / total * 50)
            self.progress(current_progress)
            
            driver.get(base_url)
            try:
                name_elem = driver.find_elements(By.ID, 'coursename')
                if not name_elem: continue
                name = name_elem[0].text
                
                self.log(f" > 解析中: {name}")
                
                for suffix, label in targets:
                    driver.get(base_url + suffix)
                    rows = driver.find_elements(By.TAG_NAME, 'tr')
                    for row in rows:
                        t = row.text
                        m = re.findall(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', t)
                        
                        # 課題特定ロジック
                        if '未提出' in t and '受付中' in t and m:
                            deadline = sorted(m)[-1]
                            key = (f"【提出：{label}】{name}", deadline)
                            results[key] = 1
                        elif any(x in t for x in ['提出済み', '回答済み', '済']):
                            submitted_list.append(f"{label}】{name}")
            except Exception as e:
                print(f"Error parsing course: {e}")
                continue
        
        final_tasks = [(t, dt.strptime(d, '%Y-%m-%d %H:%M').strftime("%Y-%m-%dT%H:%M:00")) for t, d in results.keys()]
        return final_tasks, list(set(submitted_list))

    def _get_calendar_service(self):
        import socket
        socket.setdefaulttimeout(30)
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                try: creds = pickle.load(token)
                except: pass
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try: creds.refresh(Request())
                except: creds = None
            
            if not creds:
                if not os.path.exists('credentials.json'):
                    raise Exception("credentials.json が見つかりません。")
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', ['https://www.googleapis.com/auth/calendar'])
                creds = flow.run_local_server(port=0)
            
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        return build('calendar', 'v3', credentials=creds)

    def sync_calendar(self, tasks, submitted_titles):
        service = self._get_calendar_service()
        now = dt.now(timezone.utc)
        
        self.log(">> 既存の予定を確認中...")
        events_result = service.events().list(
            calendarId=self.calendar_id, 
            timeMin=(now - timedelta(days=60)).isoformat(), 
            singleEvents=True
        ).execute()
        events = events_result.get('items', [])
        
        processed_keys = {}
        for ev in events:
            summary = ev.get('summary', '')
            # 自ツールが作った予定のみ対象にする（簡易判定）
            if "【提出：" not in summary: continue
            
            start_iso = ev['start'].get('dateTime', '')[:19]
            category_with_name = summary.split('：')[-1] # "レポート】科目名" の部分
            
            ev_dt_str = ev['start'].get('dateTime')
            if ev_dt_str:
                ev_dt = dt.fromisoformat(ev_dt_str.replace('Z', '+00:00'))
                
                # 提出済み、または期限切れの予定を削除
                if category_with_name in submitted_titles or ev_dt < now:
                    try:
                        service.events().delete(calendarId=self.calendar_id, eventId=ev['id']).execute()
                        self.log(f" [削除済/期限切れ] {summary}")
                    except: pass
                    continue
            
            # 重複チェック用キー
            task_key = (summary.split('】')[-1], start_iso)
            processed_keys[task_key] = ev['id']

        # 新規課題を追加
        for title, deadline in tasks:
            # タイトルから科目名抽出
            check_key = (title.split('】')[-1], deadline)
            
            if check_key not in processed_keys:
                event = {
                    'summary': title,
                    'description': self.sig,
                    'start': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'end': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'colorId': '11', # 赤色(目立つように)
                    'reminders': {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 60}]}
                }
                service.events().insert(calendarId=self.calendar_id, body=event).execute()
                self.log(f" [新規追加] {title}")
            else:
                 self.log(f" [継続] {title}")

class SimpleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("manaba 同期ツール (Thread版)")
        self.root.geometry("480x600")
        
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
        
        self.btn = tk.Button(root, text="同期を開始", command=self.start_thread, width=25, height=2, bg="#4CAF50", fg="white", font=("Yu Gothic", 10, "bold"))
        self.btn.pack(pady=10)
        
        self.pb = ttk.Progressbar(root, length=400, mode='determinate')
        self.pb.pack(pady=5)
        
        # スクロールバー付きテキストボックス
        log_frame = tk.Frame(root)
        log_frame.pack(padx=15, pady=10, fill='both', expand=True)
        self.log_box = tk.Text(log_frame, height=15, width=55, font=("Yu Gothic", 9), bg="#F5F5F5", padx=5, pady=5)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.config(yscrollcommand=scrollbar.set)
        self.log_box.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill='y')

    def set_progress(self, val):
        # 別スレッドからGUIを操作するため、root.afterを使うのが安全
        self.root.after(0, lambda: self._update_progress(val))

    def _update_progress(self, val):
        self.pb['value'] = val

    def add_log(self, msg):
        self.root.after(0, lambda: self._update_log(msg))

    def _update_log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    def start_thread(self):
        """ ボタンが押されたらここが呼ばれる """
        user, pw = self.ent_user.get(), self.ent_pw.get()
        if not user or not pw: 
            messagebox.showwarning("入力エラー", "IDとパスワードを入力してください")
            return
            
        save_settings(user, pw)
        
        # UIロック
        self.btn.config(state=tk.DISABLED, bg="#9E9E9E", text="実行中...")
        self.log_box.delete('1.0', tk.END)
        
        # スレッド開始
        thread = threading.Thread(target=self.run_logic, args=(user, pw))
        thread.daemon = True # アプリ終了時に強制終了できるようにする
        thread.start()

    def run_logic(self, user, pw):
        """ 別スレッドで動く実処理 """
        engine = ManabaEngine(user, pw, self.add_log, self.set_progress)
        engine.run()
        
        # 処理が終わったらボタンを戻す
        self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.btn.config(state=tk.NORMAL, bg="#4CAF50", text="同期を開始")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleApp(root)
    root.mainloop()