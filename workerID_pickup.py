from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor
import json
import re
import os


#パスを取得
os.chdir(os.path.dirname(os.path.abspath(__file__)))

#コンフィグを取得
config_path = 'config_CW1.json'
with open(config_path, encoding='utf-8') as f:
    config = json.load(f)

#コンフィグを反映
spreadsheet_id = config['spreadsheet_id']
sheet_name = config['sheet_name']
col_URL = int(config['col_URL'])
col_ID = int(config['col_ID'])
start_row = int(config['start_row'])
username = config['username']
password = config['password']

# Google Sheets APIの認証情報を取得するための設定
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# Google Sheets APIの認証情報を指定する
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    'credentials.json', scope)

# Google Sheetsにアクセスするためのオブジェクトを作成する
gc = gspread.authorize(credentials)

# Google Sheetsのシートを開く
worksheet = gc.open_by_key(spreadsheet_id).worksheet(sheet_name)

# ChromeDriverManager().remove()
options = Options()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(options = options, service = service)

#waitを定義
wait = WebDriverWait(driver, 5)

#失敗リストを定義
failure = []

# jupyterlabを起動したら、実行する②

#CWにログインする
def login_crowdworks(driver, username, password):
    try:
        # CrowdWorks へのログインページにアクセス
        driver.get('https://crowdworks.jp/login?ref=toppage_hedder')

        # ユーザ名の入力欄が表示されるまで待機して入力
        input_id = wait.until(EC.presence_of_element_located((By.NAME, "username")))
        input_id.send_keys(username)

        # パスワードの入力欄が表示されるまで待機して入力
        input_pass = wait.until(EC.presence_of_element_located((By.NAME, "password")))
        input_pass.send_keys(password)

        # ログインボタンが表示されるまで待機してクリック
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[type="submit"]')))
        login_button.click()

        try:
            h1 =wait.until(EC.presence_of_element_located((By.XPATH, 
                '//*[@id="ContentHeader"]/div/div/h1')))
            if h1.text == 'このページは表示できません':
                login_crowdworks(driver, username, password)
            else:
                print('ログイン成功！')
        except:
            print('ログイン成功！')
            pass
    except:
        print('ログイン失敗')

# スプシからURLを取得する
def get_URL():
    print("スプシ展開完了")
    # URLとインデックスを取得
    worker_URLs = worksheet.col_values(col_URL)[start_row - 1:]
    worker_IDs = worksheet.col_values(col_ID)[start_row -1:]
    print('URL取得完了')
    #行番号とURLのディクショナリを返す
    rows_and_URLs = {i + start_row : (worker_URL, worker_preID) for i, (worker_URL, worker_preID) in enumerate(zip(worker_URLs, worker_IDs))}
    return rows_and_URLs

# URLを開きIDを取得する
def get_ID(i, worker_URL):
    try:
        #ワーカーページを開く
        driver.get(worker_URL)
        #IDを含むURLを取得
        worker_URL_containsID = wait.until(
            EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[3]/td/a[1]"))).get_attribute('href')
        #IDを含むURLからIDを抽出
        worker_ID = re.search(r'\d+', worker_URL_containsID).group()
        print(f'ID取得完了:{worker_URL}:{worker_ID}')
        #IDを返す
        return worker_ID
    except Exception as e:
        print(f'ID取得失敗:{worker_URL}-{e}')
        failure.append(i)
        return None

'''# IDをスプシに書き込む
def input_ID(i, worker_ID):
    try:
        # IDを対応する行に入力する
        worksheet.update_cell(i, col_ID, str(worker_ID))
        print(f"{i}行目:完了")
    except Exception as e:
        print(f'スプシ更新失敗:{i}行目-{e}')'''

# 📌 スプレッドシートを一括更新
def update_sheet(batch_data):
    if batch_data:
        worksheet.batch_update(batch_data)
        print(f"スプレッドシート更新完了: {len(batch_data)} 件")

# 📌 1ユーザーの処理（並列実行用）
def process_worker(i, worker_URL, worker_preID):
    if worker_preID:  # すでにIDがある場合はスキップ
        return None

    worker_ID = get_ID(i, worker_URL)
    return (i, worker_ID) if worker_ID else None

# 全体の処理
# 📌 メイン処理
def main():
    try:
        # ✅ 1. CrowdWorks にログイン
        if not login_crowdworks(driver, username, password):
            print("ログインに失敗したため処理を中断します")
            return
        
        # ✅ 2. スプレッドシートから URL 取得
        rows_and_URLs = get_URL()

        # ✅ 3. ID 取得を並列処理
        with ThreadPoolExecutor(max_workers=5) as executor:
            results = list(executor.map(lambda item: process_worker(*item), rows_and_URLs.items()))

        # ✅ 4. ID をスプレッドシートに一括更新
        batch_data = [{"range": f"{sheet_name}!{col_ID}{i}", "values": [[worker_ID]]} 
                      for i, worker_ID in results if worker_ID]
        update_sheet(batch_data)

    finally:
        driver.quit()
        
# 📌 実行
if __name__ == '__main__':
    main()

'''#CWログイン
login_crowdworks(driver, username, password)
# URLの取得
rows_and_URLs = get_URL()
# 取得した要素全体にID取得＆入力
for i, worker_URL, worker_preID in rows_and_URLs.items():
    if worker_preID:  # すでにあるIDをスキップ
        continue
    worker_ID = get_ID(i, str(worker_URL))
    if worker_ID:
        input_ID(i, worker_ID)
#終了処理
driver.quit()
exit()'''