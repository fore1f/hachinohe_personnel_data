import pandas as pd
import sys
import re
import datetime
import os
import tkinter as tk
from tkinter import filedialog

def to_zenkaku(s):
    """半角数字を全角数字に変換"""
    zen_table = str.maketrans('0123456789', '０１２３４５６７８９')
    return str(s).translate(zen_table)

def format_retirement_date(date_str, current_month):
    """（3月31日付け）のような文字からルールに従って日付文字列を生成"""
    if not isinstance(date_str, str):
        return ""
    match = re.search(r'(\d+)月(\d+)日', date_str)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        
        if month == current_month:
            return f"{to_zenkaku(day)}日"
        else:
            return f"{to_zenkaku(month)}月{to_zenkaku(day)}日"
    return ""

def format_old_job(job_str):
    """旧職の文字列の括弧を・に変換し、両端に（）をつける"""
    if not isinstance(job_str, str) or pd.isna(job_str) or not str(job_str).strip():
        return ""
    s = str(job_str)
    # 全ての括弧を「・」に置換
    s = s.replace('(', '・').replace(')', '・').replace('（', '・').replace('）', '・')
    # 連続する「・」を1つにまとめる
    s = re.sub(r'・+', '・', s)
    # 先頭と末尾の「・」を削除
    s = s.strip('・')
    
    if not s:
        return ""
    return s

def replace_brackets(val):
    """文字列内の括弧をすべて「・」に変換する"""
    if not isinstance(val, str):
        return val
    s = str(val)
    s = s.replace('(', '・').replace(')', '・').replace('（', '・').replace('）', '・')
    s = re.sub(r'・+', '・', s)
    return s.strip('・')

def clean_str(val):
    """NaNなどを空文字列に変換し、空白や特定の文字を置換し、括弧を変換"""
    if pd.isna(val):
        return ""
    # 空白を削除し、読点（、）を中点（・）に変換する
    # ㈱ は削除する
    s = str(val).strip().replace("、", "・").replace("㈱", "")
    return replace_brackets(s)

def check_name_space(name):
    """氏名の中に全角スペースがあるかチェックし、なければ警告を表示"""
    if not isinstance(name, str) or not name:
        return
    if "　" not in name:
        print(f"【警告】氏名「{name}」に姓名の区切り（全角スペース）がありません。記入漏れの可能性があります。")

def main():
    print("エクセルファイル選択ダイアログを開いています...")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    initial_dir = "C:\\Users\\tked1\\py\\01_hachinohe_personnel\\hachinohe"
    if not os.path.exists(initial_dir):
        initial_dir = "C:\\"
        
    file_path = filedialog.askopenfilename(
        title="抽出する人事データのエクセルファイルを選択してください",
        initialdir=initial_dir,
        filetypes=[("Excelファイル", "*.xlsx *.xls"), ("すべてのファイル", "*.*")]
    )
    
    if not file_path:
        print("ファイルの選択がキャンセルされました。処理を終了します。")
        sys.exit(0)
        
    print(f"選択されたファイル: {file_path}")
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(os.path.dirname(file_path), f"{base_name}_抽出結果.txt")
    
    # 実行時の月を取得
    current_month = datetime.datetime.now().month

    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Excelファイルの読み込みに失敗しました: {e}")
        sys.exit(1)

    output_lines = []
    
    current_kyu = ""
    current_kyoku = ""
    current_bu = ""
    retirement_date_str = "３１日" # デフォルト
    
    # 基本の異動区分
    basic_idou = ["転任", "昇任", "転任・昇任", "配置換", "兼務"]

    for index, row in df.iterrows():
        colA = clean_str(row[0])
        colB = clean_str(row[1])
        colC = clean_str(row[2])
        colD = clean_str(row[3])
        colE = clean_str(row[4])
        
        # 空行はスキップ
        if not any([colA, colB, colC, colD, colE]):
            continue
            
        # 見出し行はスキップ
        if colA == "異動区分":
            continue

        # ヘッダ行の判定（「< 部長級 >」など）
        if colA.startswith("<") and colA.endswith(">"):
            # 級名や局名の変化をチェック
            new_kyu = colA
            new_kyoku = colC if colC.startswith("[") and colC.endswith("]") else current_kyoku
            
            # 退職日の取得 (E列に入っている想定)
            if "月" in colE and "日" in colE:
                parsed_date = format_retirement_date(colE, current_month)
                if parsed_date:
                    retirement_date_str = parsed_date
            
            # 局名が変わったら出力
            if new_kyoku and new_kyoku != current_kyoku:
                output_lines.append("") # 見やすく空行
                output_lines.append(f"【{new_kyoku.strip('[]')}】")
                current_kyoku = new_kyoku
                current_bu = ""  # 局が変わったら部を引き継がないようリセット
            
            # 級名が変わったら出力
            if new_kyu and new_kyu != current_kyu:
                output_lines.append(f"{new_kyu.strip('<> ').strip()}")
                current_kyu = new_kyu
            continue

        # データ行の処理
        if colA:
            # 氏名の全角スペースチェック
            check_name_space(colD)
            
            # 部名の更新
            if colB:
                # []を削除
                current_bu = colB.replace('[', '').replace(']', '')
            
            # 部長級・次長級の例外対応
            display_bu = ""
            if "部長級" not in current_kyu and "次長級" not in current_kyu:
                display_bu = current_bu

            # 採用・再任用の特別処理
            is_recruitment = "採用" in colA or "再任用" in colA
            
            if is_recruitment:
                # 採用・再任用の場合は採用等の形式にするため、旧職欄を上書き
                label = "採用" if "採用" in colA else "再任用"
                formatted_old_job = label
                display_colA = "" # カテゴリとしては表示しない
            else:
                formatted_old_job = format_old_job(colE)
                display_colA = colA
            
            # タブによる区切りを追加
            sep_old = "\t" if formatted_old_job else ""
            sep_name = "\t" if colD else ""

            if colA == "退職":
                # 退職処理
                # 新職部分（colC）に「（公立学校へ）」等がある場合、それがclean_strで抽出されているため結合する
                c_part = f"・{colC}" if colC else ""
                line = f"\t退職・{retirement_date_str}{c_part}{sep_old}{formatted_old_job}{sep_name}{colD}"
                output_lines.append(line)
            elif is_recruitment:
                # 採用・再任用の出力形式
                line = f"\t{display_bu}{colC}{sep_old}{formatted_old_job}{sep_name}{colD}"
                output_lines.append(line)
            else:
                if colA not in basic_idou:
                    # 出向などの処理
                    _bu = "" if "出向" in colA else display_bu
                    
                    # 「降任（役職定年）」の特別処理（クリーン後の「降任・役職定年」を「役職定年」に置換）
                    display_colA = "役職定年" if colA == "降任・役職定年" else colA
                    
                    if display_colA == "出向":
                        # 「出向」の場合は「へ出向」にし、・を入れない
                        line = f"\t{_bu}{colC}へ出向{sep_old}{formatted_old_job}{sep_name}{colD}"
                    else:
                        line = f"\t{_bu}{colC}・{display_colA}{sep_old}{formatted_old_job}{sep_name}{colD}"
                    output_lines.append(line)
                else:
                    # 基本の抽出
                    line = f"\t{display_bu}{colC}{sep_old}{formatted_old_job}{sep_name}{colD}"
                    output_lines.append(line)

    # 結果をファイルに出力
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            for line in output_lines:
                f.write(line + "\n")
        print(f"抽出が完了しました: {output_path}")
    except Exception as e:
        print(f"ファイルの書き込みに失敗しました: {e}")

if __name__ == "__main__":
    main()
