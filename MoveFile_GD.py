import pygsheets
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# 設定 Google Sheets API 和 Google Drive API 認證
SHEET_SERVICE_ACCOUNT_FILE = 'D:/DMM_list/rk-dmm-scrapying@dmm-scrapying.iam.gserviceaccount.com/dmm-scrapying-ca95a23a4216.json'  # 使用者A的Sheet API憑證
DRIVE_SERVICE_ACCOUNT_FILE = 'D:/DMM_list/rk-dmm-scrapying@dmm-scrapying.iam.gserviceaccount.com/dmm-scrapying-ca95a23a4216.json'  # 使用者B的Drive API憑證

# Google Sheets 認證
sheet_credentials = Credentials.from_service_account_file(SHEET_SERVICE_ACCOUNT_FILE)
sheet_client = pygsheets.authorize(service_account_file=SHEET_SERVICE_ACCOUNT_FILE)
spreadsheet_id = '1CSJW28pvLHj9w3L1fiJhZp3opGdHLPdq1EdhqwaqIzs'
sh = sheet_client.open_by_key(spreadsheet_id)

# Google Drive 認證
drive_credentials = Credentials.from_service_account_file(DRIVE_SERVICE_ACCOUNT_FILE)
drive_service = build('drive', 'v3', credentials=drive_credentials)

# 設定 Team Drive ID 和目標資料夾 ID
TEAM_DRIVE_ID = '0AOj7YJoxqthlUk9PVA'
PARENT_FOLDER_ID = '0AOj7YJoxqthlUk9PVA'  # 根據需要設定搜尋和創建資料夾的目標資料夾ID

# 創建資料夾
def create_folder(name, parent_id=None):
    try:
        folder_metadata = {
            'name': name,
            'driveId': TEAM_DRIVE_ID,
            'mimeType': 'application/vnd.google-apps.folder',
        }
        if parent_id:
            folder_metadata['parents'] = [parent_id]
        
        folder = drive_service.files().create(body=folder_metadata, fields='id, name', supportsAllDrives=True).execute()
        return folder.get('id')
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

# 搜尋資料夾
def search_folder(name, parent_id=None):
    try:
        query = f"mimeType = 'application/vnd.google-apps.folder' and name = '{name}'"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        results = drive_service.files().list(
            q=query, 
            spaces='drive', 
            fields='files(id, name, parents)', 
            supportsAllDrives=True, 
            includeItemsFromAllDrives=True
        ).execute()
        items = results.get('files', [])
        return items
    except HttpError as error:
        print(f'An error occurred while searching folder: {error}')
        return []
    
# 搜尋檔案
def search_file(name, parent_id=None):
    try:
        query = f"name = '{name}'"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        results = drive_service.files().list(q=query, corpora='drive', driveId='0AOj7YJoxqthlUk9PVA', fields='files(id, name, parents)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        items = results.get('files', [])
        return items
    except HttpError as error:
        print(f'An error occurred while searching file: {error}')
        return []

# 移動檔案
def move_file(file_id, folder_id):
    try:
        file = drive_service.files().get(fileId=file_id, fields='parents', supportsAllDrives=True).execute()
        if not file:
            print(f"檔案 '{file_id}' 不存在，跳過移動。")
            return
        # 先獲取檔案的父資料夾
        #file = drive_service.files().get(fileId=file_id, fields='parents').execute()
        
        previous_parents = ",".join(file.get('parents'))
        # 移動檔案到新資料夾
        #file = drive_service.files().update(fileId=file_id, addParents=folder_id, removeParents=previous_parents, fields='id, parents', supportsAllDrives=True).execute()
        drive_service.files().update(
            fileId=file_id,
            addParents=folder_id,
            removeParents=previous_parents,
            fields='id, parents',
            supportsAllDrives=True
        ).execute()
    except HttpError as error:
        print(f'An error occurred while moving file: {error}')

# 主程式
def main():
    # 讀取試算表中的所有分頁名稱
    worksheets = sh.worksheets()
    exclude_sheets = ["女優列表", "あいだゆあ", "高井桃", "高樹マリア", "川島和津実", "桜朱音"]  # 在此列表中加入要排除的分頁名稱
    #team_drive_root_id = "0AOj7YJoxqthlUk9PVA"

    for worksheet in worksheets:
        actress_name = worksheet.title
        print(actress_name)
        if actress_name in exclude_sheets:
            print(f"跳過分頁：{actress_name}")
            continue
        
        # 創建 "非個人作品" 資料夾
        non_personal_folder_name = f"{actress_name}_非個人作品"
        existing_non_personal_folders = search_folder(non_personal_folder_name, PARENT_FOLDER_ID)
        if existing_non_personal_folders:
            print(f"資料夾 '{non_personal_folder_name}' 已存在，跳過創建。")
            non_personal_folder_id = existing_non_personal_folders[0]['id']
        else:
            non_personal_folder_id = create_folder(non_personal_folder_name, PARENT_FOLDER_ID)
        
        # 創建 "個人作品" 資料夾
        personal_folder_name = f"{actress_name}_個人作品"
        existing_personal_folders = search_folder(personal_folder_name, PARENT_FOLDER_ID)
        if existing_personal_folders:
            print(f"資料夾 '{personal_folder_name}' 已存在，跳過創建。")
            personal_folder_id = existing_personal_folders[0]['id']
        else:
            personal_folder_id = create_folder(personal_folder_name, PARENT_FOLDER_ID)
        
        if personal_folder_id and non_personal_folder_id:
            # 獲取該分頁中 C 欄和 H 欄的所有值（從 C2 和 H2 開始）
            video_codes = worksheet.get_col(3, include_tailing_empty=False)[1:]
            tags = worksheet.get_col(8, include_tailing_empty=False)[1:]

            for code, tag in zip(video_codes, tags):
                if not code:
                    continue
                file_name = f"{code}.mp4"
                print(file_name)
                # 搜尋檔案
                files = search_file(file_name)
                print(files)
                for file in files:
                    if tag == "ベスト・総集編":
                        target_folder_id = non_personal_folder_id
                    else:
                        target_folder_id = personal_folder_id

                    if target_folder_id in file.get('parents', []):
                        print(f"檔案 '{file_name}' 已存在於資料夾中，跳過移動。")
                    else:
                        move_file(file['id'], target_folder_id)
                        print(f"已將檔案 '{file_name}' 移動到資料夾。")
        else:
            print(f"無法創建或取得資料夾的 ID。")

if __name__ == "__main__":
    main()
