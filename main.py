import win32com.client
from pathlib import Path

# 1. 検索対象のフォルダパスを指定（※ご自身の環境に合わせて変更してください）
folder_path_str = r"C:\path\to\your\folder"
folder_path = Path(folder_path_str).resolve()

# 全ブックの合計ページ数を格納する変数
total_all_pages = 0

# Excelアプリケーションの起動
excel = win32com.client.Dispatch("Excel.Application")

# 安全対策1: Excel画面を非表示にし、警告ポップアップ（リンク更新確認など）を出さない
excel.Visible = False
excel.DisplayAlerts = False

print(f"フォルダ「{folder_path}」内の .xlsx ファイルを検索します...\n")
print("-" * 50)

try:
    # フォルダ内の .xlsx ファイルをすべて取得（サブフォルダは含めない場合）
    # ※サブフォルダも含めたい場合は glob("*.xlsx") を rglob("*.xlsx") に変更します
    for file_path in folder_path.glob("*.xlsx"):

        # 安全対策2: Excelが裏で作る一時ファイル（~$から始まるファイル）はスキップする
        if file_path.name.startswith("~$"):
            continue

        book_total_pages = 0
        print(f"■ ファイル: {file_path.name}")

        try:
            # 安全対策3: ReadOnly=True で読み取り専用として開く
            # UpdateLinks=0 で外部参照リンクの更新を無効化し、意図しないデータ変化を防ぐ
            wb = excel.Workbooks.Open(
                Filename=str(file_path), ReadOnly=True, UpdateLinks=0
            )

            # ブック内のすべてのシートをループ
            for ws in wb.Worksheets:
                page_count = ws.PageSetup.Pages.Count
                print(f"  - シート「{ws.Name}」: {page_count} ページ")
                book_total_pages += page_count

            print(f"  >> このファイルの合計: {book_total_pages} ページ\n")
            total_all_pages += book_total_pages

        except Exception as e:
            print(f"  [エラー] {file_path.name} の処理中にエラーが発生しました: {e}\n")

        finally:
            # 安全対策4: 万が一処理が途中で失敗しても、SaveChanges=False で絶対に保存せずに閉じる
            if "wb" in locals():
                wb.Close(SaveChanges=False)

finally:
    # 確実にExcelアプリケーションを終了させる
    excel.Quit()
    print("-" * 50)
    # 最終結果の表示
    print(f"【処理完了】フォルダ内の全ページの総数は {total_all_pages} ページです。")
