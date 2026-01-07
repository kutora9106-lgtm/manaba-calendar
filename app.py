import streamlit as st
import os
import base64
import re
import time
from datetime import datetime as dt, timezone, timedelta

# Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# Google API
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

# --- ページ設定 ---
st.set_page_config(page_title="manaba 自動連携ツール", layout="centered")

# --- クラス定義: ロジックの中核 ---
class ManabaEngine:
    def __init__(self, user, pw, log_container, progress_bar, credentials):
        self.user = user
        self.pw = pw
        self.log_container = log_container
        self.progress_bar = progress_bar
        self.credentials = credentials
        self.calendar_id = 'primary'
        self.sig = "[manaba-auto]"
        self.logs = []

    def log(self, message):
        """ログを画面に出力"""
        timestamp = dt.now().strftime("%H:%M:%S")
        self.logs.append(f"[{timestamp}] {message}")
        self.log_container.text("\n".join(self.logs))
        print(f"[{timestamp}] {message}")

    def update_progress(self, value):
        """プログレスバーを更新 (0-100)"""
        self.progress_bar.progress(int(value))

    def run(self):
        driver = None
        try:
            self.log("--- 同期プロセス開始 ---")
            self.update_progress(5)
            
            # STEP1: manabaスキャン
            self.log("【1/2】manabaから課題を取得しています...")
            
            # ブラウザ設定 (Streamlit Cloud向け)
            options = Options()
            options.add_argument("--headless")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1280,720")
            options.add_argument("--lang=ja-JP")
            
            service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
            driver = webdriver.Chrome(service=service, options=options)
            
            tasks, submitted = self.fetch_manaba(driver)
            
            driver.quit()
            driver = None

            self.log(f"-> 未提出課題: {len(tasks)}件、提出済み: {len(submitted)}件を検出")
            self.update_progress(60)
            
            # STEP2: カレンダー同期
            self.log("【2/2】Googleカレンダーと同期しています...")
            self.sync_calendar(tasks, submitted)
            
            self.update_progress(100)
            self.log("--- すべての工程が完了しました ---")
            st.success(f"同期完了！ 未提出課題 {len(tasks)}件を整理しました。")
            
        except Exception as e:
            self.log(f"✖ エラーが発生しました: {e}")
            st.error(f"エラー: {e}")
        finally:
            if driver:
                driver.quit()

    def fetch_manaba(self, driver):
        auth = base64.b64encode(f"{self.user}:{self.pw}".encode()).decode()
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"Authorization": f"Basic {auth}"}})
        
        driver.get('https://slms.mi.sanno.ac.jp/ct/home')
        
        try:
            wait = WebDriverWait(driver, 10)
            if len(driver.find_elements(By.ID, "mainuserid")) > 0:
                driver.find_element(By.ID, "mainuserid").send_keys(self.user)
                driver.find_element(By.NAME, "password").send_keys(self.pw + Keys.ENTER)
        except: pass 

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
            self.update_progress(10 + (i / total * 50))
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
                        if '未提出' in t and '受付中' in t and m:
                            deadline = sorted(m)[-1]
                            results[(f"【提出：{label}】{name}", deadline)] = 1
                        elif any(x in t for x in ['提出済み', '回答済み', '済']):
                            submitted_list.append(f"{label}】{name}")
            except: continue
        
        final_tasks = [(t, dt.strptime(d, '%Y-%m-%d %H:%M').strftime("%Y-%m-%dT%H:%M:00")) for t, d in results.keys()]
        return final_tasks, list(set(submitted_list))

    def _get_calendar_service(self):
        return build('calendar', 'v3', credentials=self.credentials)

    def sync_calendar(self, tasks, submitted_titles):
        service = self._get_calendar_service()
        now = dt.now(timezone.utc)
        
        self.log(">> 既存の予定を確認中...")
        events_result = service.events().list(calendarId=self.calendar_id, timeMin=(now - timedelta(days=60)).isoformat(), singleEvents=True).execute()
        events = events_result.get('items', [])
        
        processed_keys = {}
        for ev in events:
            summary = ev.get('summary', '')
            if "【提出：" not in summary: continue
            start_iso = ev['start'].get('dateTime', '')[:19]
            category_with_name = summary.split('：')[-1]
            ev_dt_str = ev['start'].get('dateTime')
            if ev_dt_str:
                ev_dt = dt.fromisoformat(ev_dt_str.replace('Z', '+00:00'))
                if category_with_name in submitted_titles or ev_dt < now:
                    try:
                        service.events().delete(calendarId=self.calendar_id, eventId=ev['id']).execute()
                        self.log(f" [削除済/期限切れ] {summary}")
                    except: pass
                    continue
            processed_keys[(summary.split('】')[-1], start_iso)] = ev['id']

        for title, deadline in tasks:
            if (title.split('】')[-1], deadline) not in processed_keys:
                event = {
                    'summary': title, 'description': self.sig,
                    'start': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'end': {'dateTime': deadline, 'timeZone': 'Asia/Tokyo'},
                    'colorId': '11',
                    'reminders': {'useDefault': False, 'overrides': [{'method': 'popup', 'minutes': 60}]}
                }
                service.events().insert(calendarId=self.calendar_id, body=event).execute()
                self.log(f" [新規追加] {title}")
            else:
                 self.log(f" [継続] {title}")

# --- メイン画面 ---
st.title("manaba 自動連携ツール (Web版)")
st.markdown("manabaの未提出課題を取得し、Googleカレンダーに同期します。")

# --- OAuth認証フロー ---
SCOPES = ['https://www.googleapis.com/auth/calendar']

if 'credentials' not in st.session_state:
    st.session_state.credentials = None

def get_flow():
    # secrets.toml が正しく読み込めているかチェック
    if "google_oauth" not in st.secrets:
        st.error("エラー: SecretsにGoogleカレンダーの設定が見つかりません。")
        st.info("""
        プロジェクトフォルダ内の `.streamlit/secrets.toml` を確認してください。
        ```toml
        [google_oauth]
        client_id = "..."
        client_secret = "..."
        redirect_uri = "http://localhost:8501"
        ```
        """)
        st.stop()

    conf = st.secrets["google_oauth"]
    return Flow.from_client_config(
        {
            "web": {
                "client_id": conf["client_id"],
                "client_secret": conf["client_secret"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        },
        scopes=SCOPES,
        redirect_uri=conf["redirect_uri"]
    )

# 1. 認証コードがURLにある場合（Googleからのリダイレクト戻り）
if "code" in st.query_params and not st.session_state.credentials:
    try:
        code = st.query_params["code"]
        flow = get_flow()
        flow.fetch_token(code=code)
        st.session_state.credentials = flow.credentials
        st.query_params.clear() # URLをクリーンにする
        st.rerun()
    except Exception as e:
        st.error(f"認証エラー: {e}")

# 2. 未ログイン時：ログインボタンを表示
if not st.session_state.credentials:
    st.warning("まずはGoogleカレンダーへのアクセスを許可してください。")
    flow = get_flow()
    auth_url, _ = flow.authorization_url(prompt='consent')
    st.link_button("Googleでログイン", auth_url)

# 3. ログイン済み：manabaフォームを表示
else:
    st.success("Googleログイン済み")
    if st.button("ログアウト"):
        st.session_state.credentials = None
        st.rerun()

    with st.form("login_form"):
        user_id = st.text_input("manaba ユーザーID")
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("同期を開始")

    if submitted:
        if not user_id or not password:
            st.error("IDとパスワードを入力してください")
        else:
            st.subheader("実行ログ")
            progress_bar = st.progress(0)
            log_area = st.empty()
            
            # 認証情報を渡してエンジンを起動
            engine = ManabaEngine(user_id, password, log_area, progress_bar, st.session_state.credentials)
            engine.run()
