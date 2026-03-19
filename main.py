import eel
import os
from datetime import datetime
from shutil import copy
import pythoncom
import win32com.client as win32
from win32com.client import constants
import re
###GPLv3###
from fitz import open as fitzopen
###GPLv3###
from zipfile import ZipFile
from shutil import rmtree
import sqlite3
from typing import Dict, List, Tuple
import imagehash
from PIL import Image
import time
import cv2
import numpy as np
from skimage.metrics import structural_similarity as ssim

# 指定前端檔案所在的資料夾
eel.init('web')

global time_stamp_dir
# 定義常見的照片副檔名
PHOTO_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.gif', '.bmp')
COMPARISON_THRESHOLD = 15

# 假設所有照片都在這個目錄下 (請根據您的實際情況修改)
# BASE_DIR = os.path.join(os.getcwd(), "compare")
def log_error(error_message):
    # 設定檔名與時間戳記
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 使用 'a' 模式（append）才不會覆蓋掉舊的紀錄
    with open("error_log.txt", "a", encoding="utf-8") as f:
        f.write(f"--- 錯誤發生時間: {timestamp} ---\n")
        f.write(error_message)
        f.write("\n\n")



def docTransfer(file_path_abs, new_path_abs):
    print("docTransfer start")
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    word = win32.gencache.EnsureDispatch('Word.Application')
    print("below word")
    doc = word.Documents.Open(file_path_abs)
    doc.Activate ()
    # Save and Close
    word.ActiveDocument.SaveAs(
        new_path_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)
    print("docTransfer end")




@eel.expose
def hello_python(msg):
    print(f"Python 收到訊息: {msg}")
    # 從 Python 呼叫前端的 JS 函式
    eel.say_hello_js("Python 說：我收到你的訊息了！")

@eel.expose
def isDirExist():
    
    dir_name = "待比對檔案資料夾"
    exists = os.path.exists(dir_name)
    if exists:
        eel.say_hello_js(f"資料夾 '{dir_name}' 存在: {exists}")
        return True
    else:
        print("failed in is dir exist")
        eel.say_hello_js(f"檢查資料夾 '{dir_name}' 不存在，請建立後再重新執行。")
        log_error("待比對檔案資料夾未建立。else: on def isDirExist")
        return False
    # return exists
@eel.expose
def create_timestamp_dir():
    try:
        time = datetime.now().strftime('%Y%m%d_%H%M%S')
        # print(time)
        os.mkdir(time)
        # return time
        docDir = 'doc'
        docxDir = 'docx'
        pdfDir = 'pdf'
        # 20241121新增
        excelDir = 'excel'
        imgDir = 'extracted_images'
        os.makedirs(time + '\\'+ docDir, exist_ok=True)
        os.makedirs(time + '\\'+ docxDir, exist_ok=True)
        os.makedirs(time + '\\'+ pdfDir, exist_ok=True)
        os.makedirs(time + '\\'+ excelDir, exist_ok=True)
        os.makedirs(time + '\\'+ imgDir, exist_ok=True)
        global time_stamp_dir
        time_stamp_dir = time
        return True
    except Exception as e:
        log_error(e)
        return False
@eel.expose
def file_extention_classify():
    try:
        global time_stamp_dir
        for file in os.listdir('待比對檔案資料夾'):
            print(file)
            fullPath = os.path.abspath('待比對檔案資料夾' +'\\'+file)
            extention = file.split('.')[-1].lower()


            docDir = time_stamp_dir + '\\' + 'doc'
            docxDir = time_stamp_dir + '\\' + 'docx'
            pdfDir = time_stamp_dir + '\\' + 'pdf'
            excelDir = time_stamp_dir + '\\' + 'excel'
            imgDir = time_stamp_dir + '\\' + 'extracted_images'
            # 檔名為.doc結尾的移到 doc 資料夾
            if extention == 'doc':
                newPath = os.path.abspath(docDir+'\\'+file)
            # 檔名為.docx結尾的移到 docx 資料夾
            elif extention == 'docx':
                newPath = os.path.abspath(docxDir+'\\'+file)
            elif extention == 'pdf':
                newPath = os.path.abspath(pdfDir+'\\'+file)
            elif extention == 'xlsx':
                newPath = os.path.abspath(excelDir+'\\'+file)
            elif (extention== 'jpg' or extention=='jpeg' or extention=='png'
                or extention=='gif' or extention=='tif'
                or extention=='tiff'or extention=='bmp' or extention=='jfif'):
                newPath = os.path.abspath(imgDir+'\\'+file)
            else:
                print('檔案格式不符合需求，跳過處理', file)
                continue
            copy(fullPath,newPath)
        print("files_classify end")
        return True
    except Exception as e:
        log_error(e)
        return False
@eel.expose
def doc_to_docx():
    try:
        global time_stamp_dir
        doc_dir = time_stamp_dir + '\\' + 'doc'
        docx_dir = time_stamp_dir + '\\' + 'docx'
        for file in os.listdir(doc_dir):
            # print(file)# 2424.doc

            file_path_abs = os.path.abspath(doc_dir + '\\' + file)
            # print(file_path_abs)

            new = os.path.abspath(docx_dir + '\\' + file)
            new_file_abs = re.sub(r'\.\w+$', '.docx', new)
            # print(new_file_abs)
            print("--------------call docTransfer")
            docTransfer(file_path_abs, new_file_abs)
        return True
    except Exception as e:
        log_error(e)
        return False
@eel.expose
def extract_images():
    global time_stamp_dir
    pdf_dir = time_stamp_dir + '\\' + 'pdf'




def extract_images_from_single_pdf(fileName, pdf_path, output_folder):
    # Open the PDF file
    pdf_document = fitzopen(pdf_path)
    # print("under pdf_document")
    # Iterate over each page in the PDF
    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)
        
        # Extract images from the page
        image_list = page.get_images(full=True)
        for image_index, img in enumerate(image_list):
            xref = img[0]
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]
            image_filename = f"{output_folder}/{fileName}_page_{page_number + 1}_image_{image_index + 1}.jpeg"
            
            # Save the image
            with open(image_filename, "wb") as img_file:
                img_file.write(image_bytes)
            print(f"Saved image {image_filename}")

    # Close the PDF document
    pdf_document.close()

@eel.expose
def extract_image_from_pdfs():
    try:
        print("in eel.expose extract_image_from_pdfs")
        global time_stamp_dir
        for pdf in os.listdir(time_stamp_dir +'\\'+ 'pdf'):
            print(pdf)
            fullPath = os.path.abspath(time_stamp_dir +'\\'+ 'pdf' +'\\'+pdf)
            fileName = os.path.splitext(pdf)[0]
            output_folder = os.path.abspath(time_stamp_dir +'\\'+ 'extracted_images')
            extract_images_from_single_pdf(fileName, fullPath, output_folder)
        return True
    except Exception as e:
        log_error(e)
        return False



def extract_images_from_single_docx(fileName, docx_path, output_dir):
   
    # Ensure the output directory exists
    # os.makedirs(output_dir, exist_ok=True)
    
    # Open the .docx file as a zip file
    with ZipFile(docx_path, 'r') as docx_zip:
        # List all files in the zip
        for file_name in docx_zip.namelist():
            # Check if the file is an image (located in word/media/)
            if file_name.startswith('word/media/') and not file_name.endswith('.emf'):
                # Extract the image
                docx_zip.extract(file_name, output_dir)
                # Get the file path
                file_path = os.path.join(output_dir, file_name)
                # New file name (optional: rename based on your logic)
                new_file_name = os.path.basename(file_path)


                # 在這行更名，連結原本doc檔名與image名
                new_file_name = fileName + new_file_name
                # print('new file name: ', new_file_name)

                new_file_path = os.path.join(output_dir, new_file_name)
                
                # Rename the file (just for demonstration)
                # print("before os.rename")
                # print('file_path ', file_path)

                os.rename(file_path, new_file_path)
                # print(f'Extracted and renamed image: {new_file_path}')



# seems ok
@eel.expose
def extract_image_from_docxs():
    try:
        print("in eel.expose extract_image_from_docxs")
        global time_stamp_dir
        for docx in os.listdir(time_stamp_dir + '\\' + 'docx'):
            print(docx)
            fullPath = os.path.abspath(time_stamp_dir +'\\'+ 'docx' +'\\'+docx)
            fileName = os.path.splitext(docx)[0]
            output_folder = os.path.abspath(time_stamp_dir +'\\'+ 'extracted_images')
            print("-----------before extract_images_from_single_docx----------")
            print("parameters: ", fileName, fullPath, output_folder)
            extract_images_from_single_docx(fileName, fullPath, output_folder)
        
        # 清除掉extracted_images下的word資料夾
        # delete_dir_path = 'extracted_images/word'
        delete_dir_path = time_stamp_dir + '\\' + 'extracted_images/word'
        try:
            rmtree(delete_dir_path)
        except Exception as e:
            print(e)
        return True
    except Exception as e:
        log_error(e)
        return False


def extract_images_from_single_excel(xlsx_path, output_folder):
    print('First line in function extract_images_from_excel')
    print(xlsx_path)
    excel_name = os.path.basename(xlsx_path).split('.')[0]
    print(excel_name)


    # Open the Excel file as a ZIP archive
    # Open the .docx file as a zip file
    with ZipFile(xlsx_path, 'r') as zip_ref:
        # Ensure the output directory exists
        os.makedirs(output_folder, exist_ok=True)

        # List all files in the ZIP archive
        for file_name in zip_ref.namelist():
            # We're interested in files within the 'xl/media/' folder (where images are stored)
            if file_name.startswith('xl/media/'):
                # Extract the image
                img_data = zip_ref.read(file_name)
                img_name = os.path.join(output_folder, excel_name + '_' + os.path.basename(file_name))

                #連結excel檔名與圖片名稱
                
                # img_name = img_name + excel_name
                # print(img_name)
                # Save the image to the output folder
                with open(img_name, 'wb') as img_file:
                    img_file.write(img_data)
                print(f'Extracted: {img_name}')

@eel.expose
def extract_image_from_excels():
    try:
        print("in eel.expose extract_image_from_excels")
        global time_stamp_dir
        print("從excel中提取照片")
        excel_path = time_stamp_dir + '\\' + 'excel'
        output_dir = time_stamp_dir + '\\' + 'extracted_images'
        for file in os.listdir(excel_path):
            print(file)
            filePath = os.path.abspath(excel_path+'\\'+file)
            extract_images_from_single_excel(filePath, output_dir)
        return True
    except Exception as e:
        log_error(e)
        return False
def create_db(dir_name):
    # Connect to (or create) a SQLite database
    
    # conn = sqlite3.connect('image_features.db')
    conn = sqlite3.connect(dir_name + '\\' + 'image_features.db')

    cursor = conn.cursor()

    # Create a table to store image features
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS features (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        image_name TEXT UNIQUE,
        feature BLOB
    )
    ''')
    conn.commit()


@eel.expose
def buildDB():
    print("-------create db start  -----------")
    # info_change("開始創建資料庫")
    global time_stamp_dir
    create_db(time_stamp_dir)
    # except Exception as e:
        # write_error_log(e)
        # messagebox.showinfo("showinfo","發生錯誤，請檢查錯誤訊息.txt")
        # return
    # info_change("資料庫創建完成")
    print("-------create db end  -----------")
# 啟動應用程式，指定首頁檔案與視窗大小

def load_and_hash_photos(directory: str) -> Dict[str, imagehash.ImageHash]:
    """
    掃描指定目錄及其子目錄下的所有照片，計算並返回檔名和 pHash 值的字典。
    """
    photo_hashes: Dict[str, imagehash.ImageHash] = {}
    
    print(f"--- 開始掃描目錄並計算 pHash: {directory} ---")
    
    # 使用 os.walk 遞迴地遍歷目錄和子目錄
    for root, _, files in os.walk(directory):
        for filename in files:
            # 檢查檔案副檔名是否為照片
            if filename.lower().endswith(PHOTO_EXTENSIONS):
                # 構造完整的檔案路徑
                full_path = os.path.join(root, filename)
                
                try:
                    # 1. I/O 讀取操作 (這是耗時部分，但只需要執行一次)
                    img = Image.open(full_path)
                    
                    # 2. 計算 pHash
                    # 將檔名作為 key，pHash 物件作為 value
                    photo_hashes[filename] = imagehash.phash(img)
                    
                    # 釋放圖片物件，防止記憶體過度佔用 (特別是處理大量圖片時)
                    img.close() 
                    
                except Exception as e:
                    print(f"警告：無法處理檔案 {filename} ({e})")
                    continue

    print(f"--- pHash 計算完成，共處理 {len(photo_hashes)} 張照片 ---")
    return photo_hashes
def find_similar_photos(photo_hashes: Dict[str, imagehash.ImageHash], 
                       threshold: int = COMPARISON_THRESHOLD) -> List[Tuple[str, str, int]]:
    """
    對字典中所有照片的 pHash 值進行兩兩比對。

    Args:
        photo_hashes: 包含 {檔名: pHash 值} 的字典。
        threshold: 漢明距離的閾值。距離小於或等於此值的照片將被視為相似。

    Returns:
        一個包含相似照片對結果的列表，格式為 [(檔名A, 檔名B, 漢明距離), ...]
    """
    
    # 將字典的鍵 (檔名) 和值 (Hash) 分別轉換為列表
    # 這樣可以方便地使用索引 i 和 j 進行迭代
    filenames = list(photo_hashes.keys())
    hashes = list(photo_hashes.values())
    
    num_photos = len(filenames)
    similar_pairs = []
    
    print(f"--- 開始兩兩比對 {num_photos} 張照片 ({num_photos * (num_photos - 1) // 2} 次比較) ---")
    start_time = time.time()

    # 使用嵌套迴圈進行兩兩比對
    # 外層迴圈 i 從 0 到 num_photos - 2
    for i in range(num_photos):
        
        # 內層迴圈 j 從 i + 1 開始，確保：
        # 1. 避免重複比較 (A vs B 後不比較 B vs A)
        # 2. 避免與自身比較 (i != j)
        for j in range(i + 1, num_photos):
            
            # 取出檔名和 Hash 值
            filename_i = filenames[i]
            hash_i = hashes[i]
            
            filename_j = filenames[j]
            hash_j = hashes[j]
            
            # 這是純 CPU 運算，速度極快
            distance = hash_i - hash_j
            
            # 檢查漢明距離是否小於閾值
            if distance <= threshold:
                similar_pairs.append((filename_i, filename_j, distance))
                
        # 簡單的進度輸出 (每處理 100 張照片輸出一次)
        if (i + 1) % 100 == 0:
            print(f"  > 已處理 {i + 1} / {num_photos} 張照片...")

    end_time = time.time()
    
    # 輸出計時結果
    print(f"--- 比對完成，總共耗時: {end_time - start_time:.4f} 秒 ---")
    
    return similar_pairs
def cv_imread_chinese_path(file_path):
    
    # 使用 NumPy 和 cv2.imdecode() 來安全讀取包含中文路徑的圖片。
    
    try:
        # 1. 使用 Python 內建的 open() 函式以二進制模式 ('rb') 讀取整個檔案
        with open(file_path, 'rb') as f:
            # 將檔案內容讀取為一個位元組數組 (bytes array)
            binary_data = f.read()
        
        # 2. 將位元組數組轉換為 NumPy 陣列
        #    np.frombuffer() 從緩衝區創建一個陣列
        np_array = np.frombuffer(binary_data, np.uint8)
        
        # 3. 使用 cv2.imdecode() 從記憶體中的緩衝區解碼圖片
        #    cv2.IMREAD_COLOR 確保讀取為彩色圖片 (BGR 格式)
        img = cv2.imdecode(np_array, cv2.IMREAD_COLOR)
        
        if img is None:
            print(f"警告：成功讀取檔案但無法解碼，可能是圖片格式不受支援或檔案損壞。")
            
        return img
        
    except FileNotFoundError:
        print(f"錯誤：找不到檔案路徑：{file_path}")
        return None
    except Exception as e:
        print(f"讀取或解碼圖片時發生錯誤：{e}")
        return None
def compare_images(imagePair):
    BASE_DIR = os.path.join(os.getcwd(), time_stamp_dir, "extracted_images")
    img1_file_name = imagePair[0]
    img2_file_name = imagePair[1]
    img1_path = os.path.join(BASE_DIR, img1_file_name)
    img2_path = os.path.join(BASE_DIR, img2_file_name)
    try:
        # Load images
        # img1 = cv2.imread(img1_path)
        # img2 = cv2.imread(img2_path)
        # load images support chinese character path
        img1 = cv_imread_chinese_path(img1_path)
        img2 = cv_imread_chinese_path(img2_path)
        
        # Convert to grayscale
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        
        # Resize to same dimensions
        h = min(gray1.shape[0], gray2.shape[0])
        w = min(gray1.shape[1], gray2.shape[1])
        gray1 = cv2.resize(gray1, (w, h))
        gray2 = cv2.resize(gray2, (w, h))
        
        # Calculate SSIM (Structural Similarity Index)
        score, _ = ssim(gray1, gray2, full=True)
        percentage = score * 100
        
        # print(f"Similarity: {percentage:.2f}%")
        
        # self.result_label.config(text=f"Similarity: {percentage:.2f}%")
        if percentage>60.0:
            return (img1_file_name, img2_file_name, percentage)
        else:
            return None
    except Exception as e:
    # messagebox.showerror("Error", f"Failed to compare images: {str(e)}")
        print(f"Failed to compare images: {str(e)}")
def sort_and_write_results(results_list: List, output_filename: str = "comparison_results.txt"):
    """
    對比對結果列表進行排序，並將結果寫入 TXT 檔案。

    Args:
        results_list: 包含 (圖1檔名, 圖2檔名, 分數) 的列表。
        output_filename: 輸出檔案的名稱。
    """
    print(f"--- 開始排序 {len(results_list)} 筆資料 ---")
    
    # 1. 排序操作
    # key=lambda x: x[0] 指定以 tuple 的第一個元素 (圖1檔名) 進行排序
    # reverse=True 設置為降冪排序 (從 Z 到 A)
    sorted_results = sorted(
        results_list, 
        # key=lambda x: x[0], 
        key=lambda x: x[2], 
        reverse=True
    )
    
    print("排序完成：依圖1檔名降冪排序 (Z -> A)。")

    # 2. 寫入 TXT 檔案
    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            # 寫入標題行
            f.write("圖1檔名,圖2檔名,比對後相似程度\n")
            
            # 遍歷排序後的結果，並將每一行寫入檔案
            for filename1, filename2, score in sorted_results:
                # 將 tuple 格式化為 CSV 格式的字串
                rounded_score = round(float(score), 2)
                line = f"{filename1},{filename2},{rounded_score}\n"
                f.write(line)
        
        print(f"✅ 成功將結果寫入檔案: {output_filename}")
        
    except IOError as e:
        print(f"❌ 檔案寫入錯誤: {e}")
@eel.expose
def PHashCompare():
    try:
        print("phash compare hello world")
        # 傳入欲比對照片資料夾位置，回傳dict of photos' name and phash pair 
        global time_stamp_dir
        # print(time_stamp_dir)
        BASE_DIR = os.path.join(os.getcwd(), time_stamp_dir, "extracted_images")
        # print(BASE_DIR)
        photos_hash_dict = load_and_hash_photos(BASE_DIR)
        # print(photos_hash_dict)

        # 上述字典檔可用

        found_similarities = find_similar_photos(photos_hash_dict, COMPARISON_THRESHOLD)
        print("\n--- 相似照片對結果 (距離 <= {}) ---".format(COMPARISON_THRESHOLD))
        if found_similarities:
            for name1, name2, distance in found_similarities:
                print(f"相似: {name1} vs {name2} | 距離: {distance}")
            print(str(len(found_similarities)) + "組hash相似照片對")
        else:
            print("沒有找到漢明距離小於閾值的照片對。")


        # 以下對phash低於篩選值的照片對進行ssim分析
        result_list = []
        for pair in found_similarities:
            result = compare_images((pair[0], pair[1]))
            if result is not None:
                result_list.append(result)
        print(result_list)
        print(len(result_list))
        sort_and_write_results(result_list)
        return True
    except Exception as e:
        log_error(e)
        return False



eel.start('index.html', size=(800, 600))