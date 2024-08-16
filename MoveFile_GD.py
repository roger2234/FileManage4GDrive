import pygsheets
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from itertools import zip_longest

# 設定 Google Sheets API 和 Google Drive API 認證
SHEET_SERVICE_ACCOUNT_FILE = '$USER_A API.json'  # 使用者A的Sheet API憑證
DRIVE_SERVICE_ACCOUNT_FILE = '$USER_B API.json'  # 使用者B的Drive API憑證

# Google Sheets 認證
sheet_credentials = Credentials.from_service_account_file(SHEET_SERVICE_ACCOUNT_FILE)
sheet_client = pygsheets.authorize(service_account_file=SHEET_SERVICE_ACCOUNT_FILE)
 # 填入要讀取及編輯的Google Sheet
spreadsheet_id = '$USER_A sheet ID'
sh = sheet_client.open_by_key(spreadsheet_id)

# Google Drive 認證
drive_credentials = Credentials.from_service_account_file(DRIVE_SERVICE_ACCOUNT_FILE)
drive_service = build('drive', 'v3', credentials=drive_credentials)

# 設定 Team Drive ID 和目標資料夾 ID，這邊都為Team Drive根目錄的ID
TEAM_DRIVE_ID = '$root_id'
PARENT_FOLDER_ID = '$parent_folder_id'  # 根據需要設定搜尋和創建資料夾的目標資料夾ID

# 建立資料夾
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
        print(f'搜尋資料夾時發生錯誤: {error}')
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
        print(f'搜尋檔案時發生錯誤: {error}')
        return []
    
# 取得目前檔案所在資料夾
def get_current_folder_name(target_file):
    parents = target_file.get('parents', [])
    parent_id = parents[0]
    # 透過 parent_id 取得資料夾資訊
    folder_info = drive_service.files().get(
        fileId=parent_id, 
        fields='id, name', 
        supportsAllDrives=True
    ).execute()
        
    current_folder_name = folder_info.get('name')
    print(f"檔案： '{target_file['name']}' 目前位於資料夾: {current_folder_name} (ID: {parent_id})")
    return current_folder_name

# 移動檔案
def move_file(file_id, folder_id):
    try:
        file = drive_service.files().get(fileId=file_id, fields='parents', supportsAllDrives=True).execute()
        if not file:
            print(f"檔案 '{file_id}' 不存在，跳過移動。")
            return
        
        previous_parents = ",".join(file.get('parents'))

        drive_service.files().update(
            fileId=file_id,
            addParents=folder_id,
            removeParents=previous_parents,
            fields='id, parents',
            supportsAllDrives=True
        ).execute()
    except HttpError as error:
        print(f'移動檔案時發生錯誤: {error}')

# 利用資料夾ID取得資料夾名稱
def get_folder_name_by_id(folder_id):
    try:
        # 使用 Drive API 取得資料夾的元資料(Metadata)
        folder = drive_service.files().get(fileId=folder_id, fields='id, name', supportsAllDrives=True).execute()
        # 取得資料夾名稱
        folder_name = folder.get('name')
        return folder_name
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None
    
# 主程式
def main():
    # 讀取試算表中的所有分頁名稱
    worksheets = sh.worksheets()
    print(f"目前所有存在工作表： {worksheets}") 

    # 設定排除工作表
    exclude_sheets = ["$Worksheet Name"]  # 在此列表中加入要排除的分頁名稱
    
    #print(f"排除工作表： {exclude_sheets}") 

    for worksheet in worksheets:
        actress_name = worksheet.title
        print(f"目前處理工作表： {actress_name}")
        if actress_name in exclude_sheets:
            print(f"跳過此工作表： {actress_name}")
            continue
        
        # 建立 "非個人作品" 資料夾
        non_personal_folder_name = f"{actress_name}_非個人作品"
        existing_non_personal_folders = search_folder(non_personal_folder_name, PARENT_FOLDER_ID)
        if existing_non_personal_folders:
            print(f"資料夾： '{non_personal_folder_name}'  已存在，跳過建立資料夾。")
            non_personal_folder_id = existing_non_personal_folders[0]['id']
        else:
            non_personal_folder_id = create_folder(non_personal_folder_name, PARENT_FOLDER_ID)
        
        # 建立 "個人作品" 資料夾
        personal_folder_name = f"{actress_name}_個人作品"
        existing_personal_folders = search_folder(personal_folder_name, PARENT_FOLDER_ID)
        if existing_personal_folders:
            print(f"資料夾： '{personal_folder_name}'  已存在，跳過建立資料夾。")
            personal_folder_id = existing_personal_folders[0]['id']
        else:
            personal_folder_id = create_folder(personal_folder_name, PARENT_FOLDER_ID)
        
        if personal_folder_id and non_personal_folder_id:
            # 獲取該分頁中C欄"品番"的值
            video_codes = worksheet.get_col(5, include_tailing_empty=False)[1:]

            # 讀取G欄"単体作品"的值
            single_tags = worksheet.get_col(9, include_tailing_empty=False)[1:]

            # 讀取H欄"ベスト・総集編"的值
            nonsingle_tags = worksheet.get_col(10, include_tailing_empty=False)[1:]
 
            # 取得品番數量，即影片數量。
            num_video_codes = len(video_codes)
            print(f"影片(品番)數量為：{num_video_codes}")

            for row, (code, single_tag, nonsingle_tag) in enumerate(zip_longest(video_codes, single_tags, nonsingle_tags), start = 2):
                if not single_tag:  # 處理空字串或None
                    single_tag = ""
                if not nonsingle_tag:  # 處理空字串或None
                    nonsingle_tag = ""
                print(f"處理第'{row}'列，品番: {code}, I欄: {single_tag}, J欄: {nonsingle_tag}")
                
                if not code:
                    print(f"Skipping row {row} because the code is empty.")
                    continue
                file_name = f"{code}.mp4"
                print(f"Searching for file: {file_name}")

                # 搜尋檔案
                files = search_file(file_name)

                # 寫入到以"actress_name"命名的.txt文件
                with open(f"{actress_name}處理過的品番.txt", "a", encoding="utf-8") as f:
                    f.write(f"{file_name.replace('.mp4', '')}\n")

                # 如果沒有搜尋到該檔案，則在A欄下載者標記。
                cell_value = worksheet.cell(f'A{row}').value

                if not files:
                    print(f" 沒有找到檔案： '{file_name}'。")
                    if cell_value == '':
                        worksheet.update_value(f'A{row}', 'Not in PB')
                        print(f"更新 'A{row}' 資料為： Not in PB 。")
                    continue

                for file in files:                    
                    print(f"找到檔案: {file['name']}, ID: {file['id']}")

                    # 如果I欄為"単体作品"，或者是J欄為空，則移動到個人作品，其他則丟到非個人作品。
                    target_folder_id = personal_folder_id if single_tag == "単体作品" or  nonsingle_tag == "" else non_personal_folder_id
                    print(f"品番：{code}, 単体作品為 '{single_tag}', ベスト・総集編為 '{nonsingle_tag}'")
                    
                    # 根據Folder ID取得資料夾名稱
                    target_folder_name = get_folder_name_by_id(target_folder_id)

                    # 取得目前檔案所在的資料夾
                    current_folder_name = get_current_folder_name(file)
                    
                    # 宣告一個字串
                    str = "個人作品"
                    
                    # 如果字串個人作品不存在目前檔案的資料夾名稱，則往下跑。即：如果檔案還已被歸類到制定的資料夾名稱就不移動檔案，避免同檔案一直被移動。
                    if str not in current_folder_name:
                        if target_folder_id in file.get('parents', []):
                            print(f"檔案 '{file_name}' 已存在於資料夾：'{target_folder_name}' 中，跳過移動。")
                        else:
                            move_file(file['id'], target_folder_id)
                            print(f"已將檔案 '{file_name}' 移動到資料夾：'{target_folder_name}'。")

                        print(f"檔案 '{file_name}' 已存在或被移動到'{target_folder_name}' 。")
            
                        if cell_value == '':
                            worksheet.update_value(f'A{row}', target_folder_name)
                            print(f"更新 'A{row}' 資料為： {target_folder_name} 。")
                    else:
                        print(f"檔案 '{file_name}' 已存在資料夾 '{target_folder_name}' 中，不移動。")
                        if cell_value == '':
                            if current_folder_name == target_folder_name:
                                worksheet.update_value(f'A{row}', current_folder_name) 
                            else:
                                worksheet.update_value(f'A{row}', f"已歸類到資料夾：{current_folder_name}")
        else:
            print(f"無法建立或取得資料夾的 ID。")

if __name__ == "__main__":
    main()
