import sys
import time
import logging
import traceback
from pathlib import Path

import pythoncom
import win32com.client
import win32api


# ==========================
# ログ設定
# ==========================

def setup_logging():

    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent

    log_file = base / "print_Multi_Format.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )


# ==========================
# 入力展開（フォルダ対応）
# ==========================

def expand_inputs(paths):

    result = []

    for p in paths:

        path = Path(p)

        if path.is_dir():

            for f in path.rglob("*"):

                if f.suffix.lower() in [
                    ".doc", ".docx",
                    ".xls", ".xlsx",
                    ".pdf"
                ]:
                    result.append(str(f))

        else:

            result.append(str(path))

    return result


# ==========================
# Word印刷
# ==========================

def print_word(word, file):

    doc = None

    try:

        logging.info(f"WORD PRINT START {file.name}")

        doc = word.Documents.Open(str(file), ReadOnly=True)

        doc.PrintOut(Background=False)

        logging.info("WORD PRINT SENT")

    finally:

        if doc:
            doc.Close(False)


# ==========================
# Excel印刷
# ==========================

def print_excel(excel, file):

    wb = None

    try:

        logging.info(f"EXCEL PRINT START {file.name}")

        wb = excel.Workbooks.Open(str(file), ReadOnly=True)

        wb.PrintOut()

        time.sleep(2)

        logging.info("EXCEL PRINT SENT")

    finally:

        if wb:
            wb.Close(False)


# ==========================
# PDF印刷
# ==========================

def print_pdf(file):

    logging.info(f"PDF PRINT START {file.name}")

    win32api.ShellExecute(
        0,
        "print",
        str(file),
        None,
        str(file.parent),
        0
    )

    time.sleep(4)

    logging.info("PDF PRINT SENT")


# ==========================
# メイン処理
# ==========================

def process_files(files):

    pythoncom.CoInitialize()

    excel = None
    word = None

    try:

        logging.info("START OFFICE")

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        logging.info("EXCEL OK")

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        logging.info("WORD OK")

        for f in files:

            path = Path(f).resolve()

            logging.info(f"PROCESS {path}")

            if not path.exists():
                raise FileNotFoundError(path)

            ext = path.suffix.lower()

            if ext in [".doc", ".docx"]:
                print_word(word, path)

            elif ext in [".xls", ".xlsx"]:
                print_excel(excel, path)

            elif ext == ".pdf":
                print_pdf(path)

            else:
                logging.warning(f"UNSUPPORTED FILE {path.name}")

    finally:

        time.sleep(2)

        if excel:
            excel.Quit()

        if word:
            word.Quit()

        pythoncom.CoUninitialize()


# ==========================
# エントリーポイント
# ==========================

def main():

    setup_logging()

    logging.info("PROGRAM START")

    try:

        # ダブルクリック起動
        if len(sys.argv) == 1:

            print()
            print("このツールはファイルをドラッグして使用します。")
            print()
            print("使い方")
            print("1. Word / Excel / PDF ファイルを用意")
            print("2. そのファイルをこのEXEにドラッグ")
            print()
            print("Enterキーを押すと終了します")

            input()
            return False

        files = expand_inputs(sys.argv[1:])

        if len(files) == 0:

            print("印刷対象ファイルが見つかりません")
            input()
            return False

        process_files(files)

        logging.info("PROGRAM END")

        print()
        print("処理が完了しました")

        return True

    except Exception:

        error_text = traceback.format_exc()

        logging.error(error_text)

        print()
        print("*** エラーが発生しました ***")
        print()
        print(error_text)

        return False


# ==========================
# 実行
# ==========================

if __name__ == "__main__":

    success = main()

    if success:

        print("3秒後に自動終了します...")
        time.sleep(3)

    else:

        print()
        print("Enterキーを押すと終了します")
        input()
