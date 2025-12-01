import sys
import os
import time
import logging
from typing import List, Tuple

# --- 定数定義 ---
# 対応する拡張子
SUPPORTED_EXTENSIONS: Tuple[str, ...] = ('.docx', '.doc', '.xlsx', '.xls', '.pdf')
# プリンタジョブの待機時間（秒）
PRINTER_JOB_WAIT_SECONDS: int = 5
# ログファイル名
LOG_FILE_NAME: str = 'print_log.log'

def setup_logging() -> None:
    """ロギングを設定する"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE_NAME, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

def is_windows() -> bool:
    """実行環境がWindowsであるかを確認する"""
    return sys.platform == "win32"

def validate_file_path(file_path: str) -> bool:
    """
    ファイルパスが有効かどうかを検証する
    - 存在するか
    - ファイルであるか
    - 対応する拡張子か
    """
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

def print_file(file_path: str) -> bool:
    """ファイルを印刷する"""
    try:
        # win32apiをインポート
        import win32api
        logging.info(f"印刷中: {file_path}")
        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
        time.sleep(PRINTER_JOB_WAIT_SECONDS)
        return True
    except ImportError:
        logging.critical("pywin32がインストールされていません。pip install pywin32 を実行してください。")
        return False
    except Exception as e:
        logging.error(f"{file_path} の印刷に失敗しました。", exc_info=True)
        return False

def main() -> None:
    """メイン処理"""
    setup_logging()
    logging.info("印刷処理を開始します。")

    if not is_windows():
        logging.critical("このスクリプトはWindows環境でのみ実行できます。")
        return

    files_to_print: List[str] = sys.argv[1:]
    if not files_to_print:
        logging.warning("印刷するファイルが指定されていません。")
        logging.info("プログラムを終了します。")
        return

    success_files: List[str] = []
    failed_files: List[str] = []

    for file_path in files_to_print:
        if validate_file_path(file_path):
            if print_file(file_path):
                success_files.append(file_path)
            else:
                failed_files.append(file_path)
        else:
            failed_files.append(file_path)

    # --- 処理結果のサマリー ---
    logging.info("=" * 30)
    logging.info("すべての印刷処理が完了しました。")
    logging.info(f"成功: {len(success_files)}件")
    for path in success_files:
        logging.info(f"  - {path}")
    logging.info(f"失敗: {len(failed_files)}件")
    for path in failed_files:
        logging.info(f"  - {path}")
    logging.info("=" * 30)

if __name__ == "__main__":
    main()
    
    print("5秒後に自動的に画面を閉じます...")
    time.sleep(5)
