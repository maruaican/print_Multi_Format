import sys
import os
import time
import logging
from pathlib import Path

import pythoncom
import win32com.client
import win32api


# ===== 設定 =====

SUPPORTED = [".docx", ".doc", ".xlsx", ".xls", ".pdf"]

LOG_FILE = "print_log.log"


# ===== ログ =====

def setup_logging():

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


# ===== Excel 印刷 =====

def print_excel(excel, file_path: Path):

    wb = excel.Workbooks.Open(str(file_path), ReadOnly=True)

    try:

        ws = wb.ActiveSheet

        logging.info(f"Excel印刷: {file_path.name}")

        ws.PrintOut()

    finally:

        wb.Close(False)


# ===== Word 印刷 =====

def print_word(word, file_path: Path):

    doc = word.Documents.Open(str(file_path), ReadOnly=True)

    try:

        logging.info(f"Word印刷: {file_path.name}")

        doc.PrintOut()

    finally:

        doc.Close(False)


# ===== PDF 印刷 =====

def print_pdf(file_path: Path):

    logging.info(f"PDF印刷: {file_path.name}")

    win32api.ShellExecute(
        0,
        "print",
        str(file_path),
        None,
        ".",
        0
    )


# ===== ファイル処理 =====

def process_file(excel, word, file_path: Path):

    ext = file_path.suffix.lower()

    if ext in [".xlsx", ".xls"]:

        print_excel(excel, file_path)

    elif ext in [".docx", ".doc"]:

        print_word(word, file_path)

    elif ext == ".pdf":

        print_pdf(file_path)

    else:

        logging.warning(f"未対応形式: {file_path}")


# ===== ファイル検証 =====

def validate_file(file_path: Path):

    if not file_path.exists():

        logging.error(f"ファイルが存在しません: {file_path}")
        return False

    if file_path.suffix.lower() not in SUPPORTED:

        logging.warning(f"未対応拡張子: {file_path}")
        return False

    return True


# ===== Office起動 =====

def start_office():

    excel = win32com.client.Dispatch("Excel.Application")
    word = win32com.client.Dispatch("Word.Application")

    excel.Visible = False
    word.Visible = False

    excel.DisplayAlerts = False
    word.DisplayAlerts = False

    return excel, word


# ===== Office終了 =====

def close_office(excel, word):

    try:
        excel.Quit()
    except:
        pass

    try:
        word.Quit()
    except:
        pass


# ===== メイン =====

def main():

    setup_logging()

    logging.info("印刷処理開始")

    if sys.platform != "win32":

        logging.error("Windows専用スクリプトです")
        return

    args = sys.argv[1:]

    if not args:

        logging.warning("印刷対象ファイルが指定されていません")
        return

    files = [Path(a) for a in args]

    pythoncom.CoInitialize()

    excel, word = start_office()

    results = []

    try:

        for file_path in files:

            if not validate_file(file_path):

                results.append((file_path, False))
                continue

            try:

                process_file(excel, word, file_path)

                results.append((file_path, True))

            except Exception as e:

                logging.error(f"印刷失敗: {file_path} : {e}")

                results.append((file_path, False))

    finally:

        close_office(excel, word)

        pythoncom.CoUninitialize()

    logging.info("----- 印刷結果 -----")

    for file, result in results:

        status = "成功" if result else "失敗"

        logging.info(f"{status} : {file}")

    logging.info("すべての処理が完了しました")


# ===== エントリーポイント =====

if __name__ == "__main__":

    try:

        main()

    except Exception as e:

        logging.exception("致命的エラー")

    print("\n処理完了（3秒後に閉じます）")

    time.sleep(3)