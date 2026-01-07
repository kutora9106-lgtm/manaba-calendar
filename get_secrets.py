import pickle
import os
import json

token_data = {}

# 1. token.pickle からトークン情報を読み込む
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
        token_data = json.loads(creds.to_json())
else:
    print("エラー: token.pickle が見つかりません。")
    exit()

# 2. credentials.json から不足しているID/Secretを補完する
if os.path.exists('credentials.json'):
    with open('credentials.json', 'r') as f:
        c_json = json.load(f)
        # 'installed' または 'web' キーの下にある情報を取得
        app_info = c_json.get('installed') or c_json.get('web')
        if app_info:
            if 'client_id' not in token_data or not token_data['client_id']:
                token_data['client_id'] = app_info.get('client_id')
            if 'client_secret' not in token_data or not token_data['client_secret']:
                token_data['client_secret'] = app_info.get('client_secret')
            if 'token_uri' not in token_data:
                token_data['token_uri'] = "https://oauth2.googleapis.com/token"

# 3. 出力
print("\n--- ↓ ここから下をコピーして Streamlit Cloud の Secrets に上書きしてください ↓ ---\n")
print("[google_calendar]")
for key, value in token_data.items():
    if value is not None:
        print(f'{key} = {json.dumps(value)}')
print("\n--- ↑ ここまでコピーしてください ↑ ---\n")