import sys
import os
import time
import logging
from typing import List, Tuple, Optional
import win32print
import win32api
import win32com.client
import fitz

# --- 定数定義 ---
SUPPORTED_EXTENSIONS: Tuple[str, ...] = ('.docx', '.doc', '.xlsx', '.xls', '.pdf')
PRINTER_JOB_WAIT_SECONDS: int = 5
LOG_FILE_NAME: str = 'print_log.log'

# win32print constants
DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2

def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE_NAME, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def is_windows() -> bool:
    return sys.platform == "win32"

def validate_file_path(file_path: str) -> bool:
    if not os.path.exists(file_path):
        logging.error(f"ファイルが見つかりません: {file_path}")
        return False
    if not os.path.isfile(file_path):
        logging.error(f"指定されたパスはファイルではありません: {file_path}")
        return False
    if not file_path.lower().endswith(SUPPORTED_EXTENSIONS):
        logging.warning(f"対応していないファイル形式です: {file_path}")
        return False
    return True

def apply_printer_settings(printer_name: str, orientation: int):
    """プリンタのDevModeを直接操作して方向を強制設定する"""
    # PRINTER_ALL_ACCESSが必要
    hPrinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
    try:
        # GetPrinter level 2 returns a dictionary with 'pDevMode'
        info = win32print.GetPrinter(hPrinter, 2)
        devmode = info['pDevMode']
        
        # 方向を設定
        devmode.Orientation = orientation
        devmode.Fields |= DM_ORIENTATION
        
        # 設定を反映
        # SetPrinterに渡すinfo辞書を更新
        win32print.SetPrinter(hPrinter, 2, info, 0)
        logging.info(f"プリンタ設定を適用しました: {'横' if orientation == DMORIENT_LANDSCAPE else '縦'}")
    except Exception as e:
        logging.error(f"プリンタ設定の適用に失敗しました: {e}")
    finally:
        win32print.ClosePrinter(hPrinter)

def print_excel(file_path: str):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = None
    try:
        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(abs_path)
        ws = wb.ActiveSheet
        
        # Excel側の設定を取得
        # 1: Portrait, 2: Landscape
        orient = ws.PageSetup.Orientation
        printer_name = win32print.GetDefaultPrinter()
        
        # プリンタ側を強制設定
        apply_printer_settings(printer_name, orient)
        
        logging.info(f"Excel印刷実行: {file_path} (方向: {'横' if orient == 2 else '縦'})")
        ws.PrintOut()
        return True
    finally:
        if wb:
            wb.Close(False)
        excel.Quit()

def print_word(file_path: str):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        abs_path = os.path.abspath(file_path)
        doc = word.Documents.Open(abs_path)
        
        # Word側の設定を取得
        # 0: Portrait, 1: Landscape
        word_orient = doc.PageSetup.Orientation
        printer_orient = DMORIENT_LANDSCAPE if word_orient == 1 else DMORIENT_PORTRAIT
        
        printer_name = win32print.GetDefaultPrinter()
        apply_printer_settings(printer_name, printer_orient)
        
        logging.info(f"Word印刷実行: {file_path} (方向: {'横' if printer_orient == DMORIENT_LANDSCAPE else '縦'})")
        doc.PrintOut()
        return True
    finally:
        if doc:
            doc.Close(False)
        word.Quit()

def print_pdf(file_path: str):
    doc = fitz.open(file_path)
    page = doc[0]
    is_landscape = page.rect.width > page.rect.height
    doc.close()
    
    orientation = DMORIENT_LANDSCAPE if is_landscape else DMORIENT_PORTRAIT
    printer_name = win32print.GetDefaultPrinter()
    
    apply_printer_settings(printer_name, orientation)
    
    logging.info(f"PDF印刷実行: {file_path} (方向: {'横' if is_landscape else '縦'})")
    win32api.ShellExecute(0, "print", file_path, None, ".", 0)
    time.sleep(PRINTER_JOB_WAIT_SECONDS)
    return True

def print_file(file_path: str) -> bool:
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ('.xlsx', '.xls'):
            return print_excel(file_path)
        elif ext in ('.docx', '.doc'):
            return print_word(file_path)
        elif ext == '.pdf':
            return print_pdf(file_path)
        return False
    except Exception as e:
        logging.error(f"{file_path} の印刷に失敗しました。", exc_info=True)
        return False

def main() -> None:
    setup_logging()
    logging.info("印刷処理を開始します（プリンタ設定強制適用版）。")

    if not is_windows():
        logging.critical("このスクリプトはWindows環境でのみ実行できます。")
        return

    files_to_print: List[str] = sys.argv[1:]
    if not files_to_print:
        logging.warning("印刷するファイルが指定されていません。")
        return

    for file_path in files_to_print:
        if validate_file_path(file_path):
            print_file(file_path)

    logging.info("すべての処理が完了しました。")

if __name__ == "__main__":
    main()
    time.sleep(5)
