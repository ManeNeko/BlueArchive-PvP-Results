import os
import json
import re
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# === 定数定義 ===
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
ATTACK_WORKSHEET_NAME = '入力・攻撃'
DEFENSE_WORKSHEET_NAME = '入力・防衛'


# === 設定ファイル関連 ===
def load_api_config(script_dir):
    """API設定ファイルを読み込む"""
    api_dir = os.path.join(script_dir, "SpreadsheetAPI")
    
    # ファイル存在確認
    api_json_path = os.path.join(api_dir, "api.json")
    ss_txt_path = os.path.join(api_dir, "SS.txt")
    
    if not os.path.exists(api_json_path):
        print("エラー: api.jsonが存在しません。処理を終了します。")
        print(f"期待されるパス: {api_json_path}")
        return None, None
    
    if not os.path.exists(ss_txt_path):
        print("エラー: SS.txtが存在しません。処理を終了します。")
        print(f"期待されるパス: {ss_txt_path}")
        return None, None
    
    try:
        # スプレッドシートIDを読み込み
        with open(ss_txt_path, 'r', encoding='utf-8') as f:
            ss_content = f.read().strip()
        
        # URLからIDを抽出またはIDをそのまま使用
        if 'docs.google.com/spreadsheets' in ss_content:
            match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', ss_content)
            if match:
                spreadsheet_id = match.group(1)
            else:
                print("エラー: スプレッドシートURLからIDを抽出できませんでした。")
                return None, None
        else:
            spreadsheet_id = ss_content
        
        print(f"スプレッドシートID: {spreadsheet_id}")
        return api_json_path, spreadsheet_id
        
    except Exception as e:
        print(f"設定ファイル読み込みエラー: {e}")
        return None, None


# === Google Sheets API関連 ===
def authenticate_google_sheets(service_account_file):
    """Google Sheets APIの認証を行う"""
    try:
        credentials = Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials)
        return service
    except Exception as e:
        print(f"認証エラー: {e}")
        return None


def get_last_row(service, spreadsheet_id, worksheet_name):
    """指定されたワークシートのB列の最終行を取得"""
    try:
        range_name = f"{worksheet_name}!B:B"
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_name).execute()
        
        values = result.get('values', [])
        return len(values)
    except HttpError as e:
        print(f"最終行取得エラー: {e}")
        return 0


def write_data_to_sheets(service, spreadsheet_id, worksheet_name, data_rows, start_row):
    """スプレッドシートにデータを一括書き込み"""
    try:
        if not data_rows:
            print("書き込むデータがありません")
            return False
        
        # データの型変換処理
        converted_data = []
        for row in data_rows:
            converted_row = []
            for i, cell in enumerate(row):
                if i == 0:  # 日付列
                    converted_row.append(cell)
                elif i == 2:  # 真偽値列
                    if cell.upper() == 'TRUE':
                        converted_row.append(True)
                    elif cell.upper() == 'FALSE':
                        converted_row.append(False)
                    else:
                        converted_row.append(cell)
                else:
                    converted_row.append(cell)
            converted_data.append(converted_row)
        
        # 書き込み範囲を計算
        end_row = start_row + len(converted_data) - 1
        num_cols = len(converted_data[0]) if converted_data else 0
        
        # 列番号をアルファベットに変換
        def num_to_col_letters(num):
            result = ""
            while num > 0:
                num -= 1
                result = chr(num % 26 + ord('A')) + result
                num //= 26
            return result
        
        end_col = num_to_col_letters(num_cols)
        range_name = f"{worksheet_name}!A{start_row}:{end_col}{end_row}"
        
        # データを書き込み
        body = {'values': converted_data}
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        updated_cells = result.get('updatedCells', 0)
        print(f"スプレッドシート「{worksheet_name}」に {updated_cells} セルのデータを書き込みました")
        print(f"書き込み範囲: {range_name}")
        return True
        
    except HttpError as e:
        print(f"スプレッドシート書き込みエラー: {e}")
        return False


# === ファイル処理関連 ===
def read_file(file_path):
    """ファイルを読み込んでリストに変換"""
    data_rows = []
    
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if line:  # 空行でない場合
                    columns = line.split('\t')
                    data_rows.append(columns)
        
        print(f"ファイルから {len(data_rows)} 行のデータを読み込みました: {file_path}")
        print("記入中...")
        return data_rows
    
    except FileNotFoundError:
        print(f"ファイルが見つかりません: {file_path}")
        return []
    except Exception as e:
        print(f"ファイル読み込みエラー: {e}")
        return []


def process_file_data(service, spreadsheet_id, file_path, worksheet_name, file_type):
    """ファイルデータを処理してスプレッドシートに書き込み"""
    if not os.path.exists(file_path):
        return False
    
    print(f"\n--- {file_type}データの処理 ---")
    data = read_file(file_path)
    if not data:
        return False
    
    # シートの最終行を取得
    last_row = get_last_row(service, spreadsheet_id, worksheet_name)
    start_row = last_row + 1
    print(f"{file_type}シート - 最終行: {last_row}, 書き込み開始行: {start_row}")
    
    # シートにデータを書き込み
    if write_data_to_sheets(service, spreadsheet_id, worksheet_name, data, start_row):
        print(f"{file_type}データの転記が完了しました")
        #テキストファイル削除
        os.remove(file_path)
        return True
    else:
        print(f"{file_type}データの転記に失敗しました")
        return False


# === メイン処理 ===
def main(script_dir=None):
    """メイン処理"""
    if script_dir is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # srcフォルダ内から実行された場合は、親ディレクトリを取得
        if os.path.basename(script_dir) == 'src':
            script_dir = os.path.dirname(script_dir)
    
    print("=== Google Sheetsデータ転記プログラム ===")
    
    # API設定を読み込み
    print("API設定を読み込み中...")
    service_account_file, spreadsheet_id = load_api_config(script_dir)
    if not service_account_file or not spreadsheet_id:
        return
    
    # ファイルパス定義
    file_attack = os.path.join(script_dir, "リザルト_攻撃.txt")
    file_defense = os.path.join(script_dir, "リザルト_防衛.txt")
    
    # ファイルの存在確認
    attack_exists = os.path.exists(file_attack)
    defense_exists = os.path.exists(file_defense)
    
    if not attack_exists and not defense_exists:
        print("攻撃ファイルと防衛ファイルの両方が存在しません。処理を終了します。")
        return
    
    # Google Sheets APIの認証
    print("Google Sheets APIに接続中...")
    service = authenticate_google_sheets(service_account_file)
    if not service:
        print("Google Sheets APIの認証に失敗しました")
        return
    print("認証成功！")
    
    # データ処理
    success_count = 0
    total_files = 0
    
    if attack_exists:
        total_files += 1
        if process_file_data(service, spreadsheet_id, file_attack, ATTACK_WORKSHEET_NAME, "攻撃"):
            success_count += 1
    
    if defense_exists:
        total_files += 1
        if process_file_data(service, spreadsheet_id, file_defense, DEFENSE_WORKSHEET_NAME, "防衛"):
            success_count += 1
    
    # 結果の表示
    if success_count == total_files:
        print(f"\nデータ転記が完了しました！ ({success_count}/{total_files})")
    else:
        print(f"\nデータ転記に失敗しました ({success_count}/{total_files})")


if __name__ == "__main__":
    main()
