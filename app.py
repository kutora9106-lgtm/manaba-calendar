import streamlit as st
import os
import pickle
# Selenium関連
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- 画面の作成 ---
st.title("manaba 自動連携ツール (Web版)")
st.write("IDとパスワードを入力して実行してください。")

user_id = st.text_input("ユーザーID")
password = st.text_input("パスワード", type="password")

if st.button("実行開始"):
    if not user_id or not password:
        st.error("IDとパスワードを入力してください")
    else:
        st.info("サーバー上でブラウザを起動して処理を開始します...")
        
        # --- WebサーバーでSeleniumを動かすための設定 ---
        chrome_options = Options()
        chrome_options.add_argument("--headless") # 画面を出さない
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        try:
            # サーバー側で自動でブラウザを準備する設定
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            
            # ここに manaba へのログイン処理などを書く
            st.success("ブラウザの起動に成功しました！")
            driver.quit()
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

