import sys
import os
import time
import logging
import threading
from typing import List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import win32print
import win32api
import win32com.client
import pythoncom
import fitz

# --- 定数定義 ---
SUPPORTED_EXTENSIONS: Tuple[str, ...] = ('.docx', '.doc', '.xlsx', '.xls', '.pdf')
PRINTER_JOB_TIMEOUT_SECONDS: int = 60
LOG_FILE_NAME: str = 'print_log.log'

# win32print constants
DM_ORIENTATION = 0x00000001
DM_PAPERSIZE = 0x00000002
DMORIENT_PORTRAIT = 1
DMORIENT_LANDSCAPE = 2

# スレッドごとのCOM初期化用ロック
com_lock = threading.Lock()

class COMApplicationManager:
    """Excel/Wordアプリケーションのインスタンスを管理し再利用する"""
    _excel_app = None
    _word_app = None
    _lock = threading.Lock()

    @classmethod
    def ensure_apps_running(cls):
        """メインスレッドでアプリケーションを起動しておく"""
        with cls._lock:
            if cls._excel_app is None:
                try:
                    cls._excel_app = win32com.client.Dispatch("Excel.Application")
                    cls._excel_app.Visible = False
                    cls._excel_app.DisplayAlerts = False
                except Exception as e:
                    logging.error(f"Excelの起動に失敗しました: {e}")
            
            if cls._word_app is None:
                try:
                    cls._word_app = win32com.client.Dispatch("Word.Application")
                    cls._word_app.Visible = False
                    cls._word_app.DisplayAlerts = 0
                except Exception as e:
                    logging.error(f"Wordの起動に失敗しました: {e}")

    @classmethod
    def get_excel_app(cls):
        # 各スレッドでDispatchを呼び出すことで、既存のインスタンスへのプロキシを取得する
        # (CoInitialize済みのスレッドである必要がある)
        return win32com.client.Dispatch("Excel.Application")

    @classmethod
    def get_word_app(cls):
        return win32com.client.Dispatch("Word.Application")

    @classmethod
    def quit_all(cls):
        with cls._lock:
            if cls._excel_app:
                try:
                    cls._excel_app.Quit()
                except:
                    pass
                cls._excel_app = None
            if cls._word_app:
                try:
                    cls._word_app.Quit()
                except:
                    pass
                cls._word_app = None
            pythoncom.CoUninitialize()

class PrinterSettingsManager:
    """プリンタ設定のキャッシュと適用を管理する"""
    _cache = {}
    _lock = threading.Lock()

    @classmethod
    def apply_settings(cls, printer_name: str, orientation: int):
        cache_key = (printer_name, orientation)
        with cls._lock:
            if cls._cache.get(printer_name) == orientation:
                return # 既に設定済みならスキップ

            hPrinter = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
            try:
                info = win32print.GetPrinter(hPrinter, 2)
                devmode = info['pDevMode']
                devmode.Orientation = orientation
                devmode.Fields |= DM_ORIENTATION
                win32print.SetPrinter(hPrinter, 2, info, 0)
                cls._cache[printer_name] = orientation
                logging.info(f"プリンタ設定を適用しました: {printer_name} ({'横' if orientation == DMORIENT_LANDSCAPE else '縦'})")
            except Exception as e:
                logging.error(f"プリンタ設定の適用に失敗しました: {e}")
            finally:
                win32print.ClosePrinter(hPrinter)

def wait_for_print_job(file_path: str, timeout: int = PRINTER_JOB_TIMEOUT_SECONDS):
    """印刷ジョブがキューに登録され、完了するのを監視する"""
    printer_name = win32print.GetDefaultPrinter()
    file_name = os.path.basename(file_path)
    start_time = time.time()
    
    logging.info(f"印刷ジョブの監視を開始: {file_name}")
    
    while time.time() - start_time < timeout:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            # ジョブ一覧を取得
            jobs = win32print.EnumJobs(hPrinter, 0, -1, 1)
            # 自分のファイルに関連するジョブを探す
            target_job = None
            for job in jobs:
                if file_name in str(job.get('pDocument', '')):
                    target_job = job
                    break
            
            if not target_job:
                # ジョブが見つからない場合、既に完了したか、まだ登録されていない
                # 少し待ってから再確認（登録待ち）
                if time.time() - start_time > 5: # 5秒経ってもなければ完了とみなす
                    logging.info(f"印刷ジョブ完了（またはキューから消失）: {file_name}")
                    return True
            else:
                status = target_job.get('Status', 0)
                # 0 は正常に処理中または待機中
                if status & win32print.JOB_STATUS_ERROR:
                    logging.error(f"印刷ジョブエラー検出: {file_name}")
                    return False
                if status & win32print.JOB_STATUS_DELETING:
                    logging.info(f"印刷ジョブ削除中: {file_name}")
                    return True
        finally:
            win32print.ClosePrinter(hPrinter)
        
        time.sleep(1)
    
    logging.warning(f"印刷ジョブ監視タイムアウト: {file_name}")
    return False

def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s',
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

def print_excel(file_path: str):
    pythoncom.CoInitialize()
    try:
        excel = COMApplicationManager.get_excel_app()
        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
        try:
            ws = wb.ActiveSheet
            orient = ws.PageSetup.Orientation
            printer_name = win32print.GetDefaultPrinter()
            
            PrinterSettingsManager.apply_settings(printer_name, orient)
            
            logging.info(f"Excel印刷実行: {file_path} (方向: {'横' if orient == 2 else '縦'})")
            ws.PrintOut()
            return wait_for_print_job(file_path)
        finally:
            wb.Close(False)
    except Exception as e:
        logging.error(f"Excel印刷エラー ({file_path}): {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def print_word(file_path: str):
    pythoncom.CoInitialize()
    try:
        word = COMApplicationManager.get_word_app()
        abs_path = os.path.abspath(file_path)
        doc = word.Documents.Open(abs_path, ReadOnly=True, Visible=False)
        try:
            word_orient = doc.PageSetup.Orientation
            printer_orient = DMORIENT_LANDSCAPE if word_orient == 1 else DMORIENT_PORTRAIT
            printer_name = win32print.GetDefaultPrinter()
            
            PrinterSettingsManager.apply_settings(printer_name, printer_orient)
            
            logging.info(f"Word印刷実行: {file_path} (方向: {'横' if printer_orient == DMORIENT_LANDSCAPE else '縦'})")
            doc.PrintOut()
            return wait_for_print_job(file_path)
        finally:
            doc.Close(False)
    except Exception as e:
        logging.error(f"Word印刷エラー ({file_path}): {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

def print_pdf(file_path: str):
    try:
        doc = fitz.open(file_path)
        page = doc[0]
        is_landscape = page.rect.width > page.rect.height
        doc.close()
        
        orientation = DMORIENT_LANDSCAPE if is_landscape else DMORIENT_PORTRAIT
        printer_name = win32print.GetDefaultPrinter()
        
        PrinterSettingsManager.apply_settings(printer_name, orientation)
        
        logging.info(f"PDF印刷実行: {file_path} (方向: {'横' if is_landscape else '縦'})")
        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
        return wait_for_print_job(file_path)
    except Exception as e:
        logging.error(f"PDF印刷エラー ({file_path}): {e}")
        return False

def print_file(file_path: str) -> bool:
    if not validate_file_path(file_path):
        return False
    
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        return print_excel(file_path)
    elif ext in ('.docx', '.doc'):
        return print_word(file_path)
    elif ext == '.pdf':
        return print_pdf(file_path)
    return False

def main() -> None:
    setup_logging()
    logging.info("印刷処理を開始します（パフォーマンス改善版）。")

    if not is_windows():
        logging.critical("このスクリプトはWindows環境でのみ実行できます。")
        return

    files_to_print: List[str] = sys.argv[1:]
    if not files_to_print:
        logging.warning("印刷するファイルが指定されていません。")
        return

    # COMの初期化とアプリの事前起動
    pythoncom.CoInitialize()
    COMApplicationManager.ensure_apps_running()

    # 並列実行の設定
    # Excel/WordはCOMの制約があるため、あまり多くしすぎない
    max_workers = min(os.cpu_count() or 4, 4)
    
    results = []
    with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="PrintWorker") as executor:
        future_to_file = {executor.submit(print_file, f): f for f in files_to_print}
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                success = future.result()
                results.append((file_path, success))
            except Exception as e:
                logging.error(f"予期せぬエラー ({file_path}): {e}")
                results.append((file_path, False))

    # 終了処理
    COMApplicationManager.quit_all()
    pythoncom.CoUninitialize()

    logging.info("--- 印刷結果サマリー ---")
    for file_path, success in results:
        status = "成功" if success else "失敗"
        logging.info(f"{status}: {file_path}")
    logging.info("すべての処理が完了しました。")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.critical(f"致命的なエラーが発生しました: {e}", exc_info=True)
    
    print("\n処理が完了しました。このウィンドウを閉じるには Enter キーを押してください...")
    input()
