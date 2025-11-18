import csv
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import chardet
import re # 郵便番号のハイフン削除用
import unicodedata # 半角→全角変換用
from PIL import Image, ImageDraw, ImageFont

# --- 定数設定 ---
TEMP_CSV_PATH = "temp.csv" # 一時ファイル名
DEFAULT_TITLE = '様' # 敬称が空の場合に補完する値

# 🚨 画像生成スクリプトが期待する、厳密なヘッダー名と出力順序
OUTPUT_FIELDS_ORDER = [
    '氏名', '氏名２', '郵便番号', '住所１', '住所２',
    '敬称', '敬称２'
]

# 元のCSVが持つ可能性のあるヘッダーの候補マッピング
HEADER_CANDIDATES = {
    '氏名': ['氏名'],
    '氏名２': ['連名', '氏名２'],
    '郵便番号': ['郵便番号', '郵便番号１'],
    '住所１': ['住所1', '住所１'],
    '住所２': ['住所2', '住所２'],
    '敬称': ['敬称'],
    '敬称２': ['敬称2', '敬称２']
}

# Tkinterのルートウィンドウを作成（ユーザーには見えない）
root = tk.Tk()
root.withdraw() # メインウィンドウを非表示にする

# --- プログレスウィンドウ関連の変数と関数 ---
progress_window = None
progress_label = None
progress_bar = None
total_files_to_process = 0
current_file_count = 0

def create_progress_window():
    """処理進行状況を示すプログレスウィンドウを作成する。"""
    global progress_window, progress_label, progress_bar
    progress_window = tk.Toplevel(root)
    progress_window.title("処理中...")
    progress_window.geometry("400x120")
    progress_window.resizable(False, False)
    progress_window.attributes("-topmost", True) # 最前面に表示
    tk.Label(progress_window, text="ハガキ画像を生成中...", font=("Arial", 12)).pack(pady=10)
    progress_label = tk.Label(progress_window, text="準備中...", font=("Arial", 10))
    progress_label.pack(pady=5)
    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=5)
    progress_window.protocol("WM_DELETE_WINDOW", lambda: None) # ユーザーによるクローズを無効化
    root.update_idletasks()
    root.update()

def update_progress(current_name, processed_count, total_count):
    """プログレスバーとラベルを更新する。"""
    if progress_window and progress_label and progress_bar:
        progress_label.config(text=f"処理中: {current_name}\n({processed_count}/{total_count} 件)")
        if total_count > 0:
            progress_bar["value"] = (processed_count / total_count) * 100
        else:
            progress_bar["value"] = 0
        root.update_idletasks()
        root.update()

def destroy_progress_window():
    """プログレスウィンドウを破棄する。"""
    global progress_window
    if progress_window:
        progress_window.destroy()
        progress_window = None
        
def _clean_up_temp_csv():
    """一時ファイル（Temp.csv）を削除する。"""
    if os.path.exists(TEMP_CSV_PATH):
        try:
            os.remove(TEMP_CSV_PATH)
            #print(f"一時ファイル {TEMP_CSV_PATH} を削除しました。")
            return True
        except Exception as e:
            print(f"一時ファイル {TEMP_CSV_PATH} の削除に失敗しました: {e}")
            return False
    return True # ファイルが存在しない場合は成功と見なす

# --- テンプレート生成に関する設定 ---
GENERATE_TEMPLATE = True # Trueにするとテンプレート画像を自動生成する
AUTO_TEMPLATE_FILENAME = 'generated_postcard_template.jpg' # 自動生成されるテンプレートのファイル名
TEMPLATE_DPI = 300       # テンプレートのDPI (Dots Per Inch) - 印刷品質に影響

# ハガキの物理的なサイズ (mm)
POSTCARD_WIDTH_MM = 100 # 短辺が幅
POSTCARD_HEIGHT_MM = 148 # 長辺が高さ

# --- テンプレート画像を自動生成する関数 (郵便番号枠なし) ---
def generate_postcard_template(filename, dpi, width_mm, height_mm):
    """
    指定されたDPIとサイズで、白い背景のハガキテンプレート画像を自動生成する。
    郵便番号枠は描画しない。
    """
    print(f"ハガキテンプレートを自動生成します: {filename}")

    # ピクセルサイズの計算
    width_px = int(width_mm / 25.4 * dpi)
    height_px = int(height_mm / 25.4 * dpi)

    # 白い背景の画像を作成
    img = Image.new('RGB', (width_px, height_px), (255, 255, 255))
    
    # 画像を保存
    img.save(filename, 'JPEG', quality=95)
    print(f"テンプレート「{filename}」を生成しました。")
    return filename

# --- 設定項目 ---
# テンプレートの自動生成設定に基づいてパスを設定
if GENERATE_TEMPLATE:
    TEMPLATE_IMAGE_PATH = AUTO_TEMPLATE_FILENAME # 自動生成されるテンプレートを使用
else:
    TEMPLATE_IMAGE_PATH = 'template_postcard.jpg' # 手動で用意したテンプレートを使用

# 同梱するNoto Sans JPフォントのファイル名
FONT_FILENAME = 'NotoSansJP-Regular.ttf'

# PyInstallerでバンドルされたファイルを安全に参照するためのパス取得
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FONT_PATH = os.path.join(BASE_DIR, FONT_FILENAME)

# フォントサイズの設定 (DPI300のハガキテンプレートを想定した初期値)
COMMON_NAME_FONT_SIZE = 130
NAME_FONT_SIZE = COMMON_NAME_FONT_SIZE
NAME2_FONT_SIZE = COMMON_NAME_FONT_SIZE
TITLE_FONT_SIZE = COMMON_NAME_FONT_SIZE
ADDRESS_FONT_SIZE = 68
ZIP_FONT_SIZE = 86

TEXT_COLOR = (0, 0, 0) # テキストの色 (黒)

# --- 郵便番号枠の物理寸法 (mm) (縦向きハガキの場合) ---
ZIP_BOX_HEIGHT_MM = 8.0 # 各数字枠の高さ
ZIP_BOX_INDIVIDUAL_WIDTH_MM = 5.7 # 各数字枠の幅
ZIP_BOX_INNER_GAP_MM = 1.3 # 各枠の間の隙間 (数字間の隙間)

# ハガキの基準位置からのマージン
ZIP_TOP_MARGIN_MM = 11.0 # ハガキ上端から郵便番号枠上端まで
ZIP_LEFT_MARGIN_MM = 46.0 # ハガキ左端から一番左の枠の左端まで

# --- 描画位置の計算 (ピクセル) ---
PIXELS_PER_MM = TEMPLATE_DPI / 25.4

CALC_POSTCARD_WIDTH_PX = int(POSTCARD_WIDTH_MM * PIXELS_PER_MM)
CALC_POSTCARD_HEIGHT_PX = int(POSTCARD_HEIGHT_MM * PIXELS_PER_MM)

CALC_ZIP_BOX_HEIGHT_PX = int(ZIP_BOX_HEIGHT_MM * PIXELS_PER_MM)
CALC_ZIP_BOX_INDIVIDUAL_WIDTH_PX = int(ZIP_BOX_INDIVIDUAL_WIDTH_MM * PIXELS_PER_MM)
CALC_ZIP_BOX_INNER_GAP_PX = int(ZIP_BOX_INNER_GAP_MM * PIXELS_PER_MM)

CALC_ZIP_TOP_MARGIN_PX = int(ZIP_TOP_MARGIN_MM * PIXELS_PER_MM)
CALC_ZIP_LEFT_MARGIN_PX = int(ZIP_LEFT_MARGIN_MM * PIXELS_PER_MM)

# 郵便番号の描画Y座標（テキストの垂直中央揃えを考慮）
ZIP_TEXT_VERTICAL_OFFSET_IN_BOX_PX = (CALC_ZIP_BOX_HEIGHT_PX - ZIP_FONT_SIZE) / 2
if ZIP_TEXT_VERTICAL_OFFSET_IN_BOX_PX < 0:
    ZIP_TEXT_VERTICAL_OFFSET_IN_BOX_PX = 0

# 郵便番号の共通Y座標 (ハガキの上端から)
ZIP_COMMON_Y = CALC_ZIP_TOP_MARGIN_PX + ZIP_TEXT_VERTICAL_OFFSET_IN_BOX_PX

# 郵便番号の描画開始X座標 (ハガキの左端を基準に計算)
ZIP_OVERALL_LEFT_X_PX = CALC_ZIP_LEFT_MARGIN_PX

# 郵便番号の各数字間のオフセット（描画後に加えるオフセット）
ZIP_CHAR_OFFSETS = [35, 35, 36, 35, 35, 34, 0] # 微調整済みの値

# 描画位置の調整 (DPI300のハガキテンプレート、縦向きを想定した初期値)
ADDRESS_COL1_X = 1000
ADDRESS_COL2_X = 900
ADDRESS_LINE_Y_START = 300
ADDRESS_CHAR_Y_SPACING = 70

# 氏名の開始位置 (中央からやや左、上から下へ)
NAME_COL1_X = 650
NAME_LINE_Y_START = 350
NAME_CHAR_Y_SPACING = 150

# --- レイアウト調整用オフセット ---
OFFSET_NAME1_Y_FROM_SURNAME_END_AFTER_SPACE = 100

# 氏名1の名前が配置される列(NAME_COL1_X)から、氏名2の名前が配置される列までのX方向のオフセット。
# 負の値を指定すると左にずれます。
OFFSET_NAME2_X_FROM_NAME1_COL = -150

# 敬称のY座標オフセット（氏名1の名前の最終Y座標から）
OFFSET_TITLE_Y_FROM_NAME_END = 10

# --- ヘルパー関数 ---
def _convert_halfwidth_to_fullwidth_all(text):
    """
    文字列中の全ての半角文字（英数字、記号、カタカナ、スペース）を全角に変換する。
    """
    if text is None: text = ""
    fullwidth_chars = []
    for char in text:
        # 半角アルファベットを全角に変換
        if 'a' <= char <= 'z':
            fullwidth_chars.append(chr(ord(char) - ord('a') + ord('ａ')))
        elif 'A' <= char <= 'Z':
            fullwidth_chars.append(chr(ord(char) - ord('A') + ord('Ａ')))
        # 半角数字を全角に変換
        elif '0' <= char <= '9':
            fullwidth_chars.append(chr(ord(char) - ord('0') + ord('０')))
        # 半角スペースを全角スペースに変換
        elif char == ' ':
            fullwidth_chars.append('　')
        # その他の文字（半角カタカナ、特定の記号など）はNFKC正規化を適用
        else:
            fullwidth_chars.append(unicodedata.normalize('NFKC', char))
            
    return "".join(fullwidth_chars)


def _convert_address_numbers_and_hyphens(text):
    """
    住所内の数字を漢数字に変換し、ハイフン（半角・全角問わず）を全角縦棒に変換する。
    ただし、数字の直後に特定の文字が続く場合は、その数字は全角数字のままにする。
    """
    kanji_map = {
        '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
        '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
    }

    # フェーズ1: 漢数字変換をスキップすべき部分を特定し、プレースホルダーに置き換える
    skip_pattern = re.compile(
        r'[０-９]+'  # 1桁以上の全角数字
        r'([a-zA-ZＡ-Ｚａ-ｚァ-ヶア-ンーヴｱ-ﾝｦ-ﾟ階号室棟]+)' # その後に続く特定の文字群
    )

    preserved_parts = {} # プレースホルダーと元の文字列のマッピング
    placeholder_idx = 0

    def replace_with_placeholder(match):
        nonlocal placeholder_idx
        full_match = match.group(0)  
        placeholder = f"__PLACEHOLDER_{placeholder_idx}__"
        preserved_parts[placeholder] = full_match
        placeholder_idx += 1
        return placeholder

    # まず、スキップすべきパターンをプレースホルダーに置き換え
    temp_text = skip_pattern.sub(replace_with_placeholder, text)

    # フェーズ2: プレースホルダー以外の部分に対して通常の変換を行う
    converted_temp_text = ""
    i = 0
    while i < len(temp_text):
        char = temp_text[i]
        
        # プレースホルダーかどうかのチェック
        if char == '_' and temp_text[i:i+16].startswith('__PLACEHOLDER_'):
            end_idx = temp_text.find('__', i + 1)
            if end_idx != -1:
                end_idx += 2
                placeholder = temp_text[i:end_idx]
                converted_temp_text += placeholder
                i = end_idx
                continue
            
        # 数字と判断される文字 (全角数字)
        if '０' <= char <= '９':
            converted_temp_text += kanji_map[chr(ord(char) - ord('０') + ord('0'))]
        # ハイフン（全角・半角長音符）を縦棒に変換
        elif char in ['-', '－', 'ー']:
            converted_temp_text += '｜'
        else:
            converted_temp_text += char
        i += 1
    
    # フェーズ3: プレースホルダーを元の文字列に戻す
    final_text = converted_temp_text
    for placeholder, original_text in preserved_parts.items():
        final_text = final_text.replace(placeholder, original_text)

    return final_text


def _normalize_name_spacing(name_fullwidth_input):
    """
    氏名のスペースを正規化する。
    前後の全角スペースをトリムし、内部の連続する全角スペースを1つにまとめる。
    """
    if name_fullwidth_input is None: name_fullwidth_input = ""
    if not isinstance(name_fullwidth_input, str): name_fullwidth_input = "" 
    normalized_name = name_fullwidth_input.strip('　')
    normalized_name = re.sub(r'　+', '　', normalized_name)
    return normalized_name


def draw_vertical_text(img_obj, draw_obj, text, font, start_x, start_y, char_y_spacing, text_color=(0,0,0)):
    """
    縦書きでテキストを描画するヘルパー関数。
    一文字ずつ描画し、縦に積み重ねる。
    Returns the Y-coordinate after the last character is drawn.
    """
    current_y = start_y
    for char in text:
        bbox = draw_obj.textbbox((0, 0), char, font=font)
        char_width = bbox[2] - bbox[0]
        
        char_draw_x = start_x - (char_width / 2)
        
        draw_obj.text((char_draw_x, current_y), char, font=font, fill=text_color)
        current_y += char_y_spacing
    return current_y

def draw_horizontal_zip_code(draw_obj, text, font, start_x, start_y, char_offsets, text_color=(0,0,0)):
    """
    郵便番号を横書きで、個別の桁間隔を考慮して描画する。
    """
    current_x = start_x
    for i, char in enumerate(text):
        draw_obj.text((current_x, start_y), char, font=font, fill=text_color)
        char_length = font.getlength(char)
        current_x += char_length
        
        if i < len(char_offsets):
            current_x += char_offsets[i]
    return current_x

# --- ヘルパー関数 (ヘッダー正規化) ---
def _normalize_header(header):
    kanji_map = {'1': '１', '2': '２', '3': '３', '4': '４', '5': '五', '6': '六', '7': '七', '8': '八', '9': '九', '0': '０'}
    def replace_half_to_full(s):
        if s is None: return ""
        s_str = str(s).strip()
        if s_str == "": return ""
        # 半角数字を全角数字に変換（ヘッダー名内でのみ）
        return "".join(kanji_map.get(c, c) for c in s_str)
    
    normalized = []
    for i, h in enumerate(header):
        processed_h = replace_half_to_full(h)
        if processed_h == "":
            normalized.append(f"COL_{i+1}")
        else:
            normalized.append(processed_h)
    return normalized

# ----------------------------------------------------------------------
#                         メイン処理関数 (makePostalCardPDF)
# ----------------------------------------------------------------------

def makePostalCardPDF():
    """
    Temp.csvからデータを読み込み、ハガキ画像を生成してPDFにまとめるロジック。
    """
    global total_files_to_process

    # --- テンプレートの準備 ---
    if GENERATE_TEMPLATE:
        try:
            global TEMPLATE_IMAGE_PATH
            TEMPLATE_IMAGE_PATH = generate_postcard_template(
                AUTO_TEMPLATE_FILENAME, TEMPLATE_DPI,
                POSTCARD_WIDTH_MM, POSTCARD_HEIGHT_MM
            )
        except Exception as e:
            messagebox.showerror("エラー", f"テンプレートの自動生成に失敗しました: {e}\n「GENERATE_TEMPLATE」をFalseに設定し、手動でテンプレート画像を用意してください。")
            _clean_up_temp_csv() # 異常終了時も削除
            return
    else:
        if not os.path.exists(TEMPLATE_IMAGE_PATH):
            messagebox.showerror("エラー", f"指定されたテンプレート画像が見つかりません: {TEMPLATE_IMAGE_PATH}\n「TEMPLATE_IMAGE_PATH」の設定と、ファイルが存在するか確認してください。")
            _clean_up_temp_csv() # 異常終了時も削除
            return

    # --- フォントのロード ---
    global name_font, name2_font, title_font, address_font, zip_font
    print(f"フォント「{FONT_FILENAME}」をロードしています...")
    try:
        font_index = 0
        name_font = ImageFont.truetype(FONT_PATH, COMMON_NAME_FONT_SIZE, index=font_index)
        name2_font = ImageFont.truetype(FONT_PATH, COMMON_NAME_FONT_SIZE, index=font_index)
        title_font = ImageFont.truetype(FONT_PATH, COMMON_NAME_FONT_SIZE, index=font_index)
        address_font = ImageFont.truetype(FONT_PATH, ADDRESS_FONT_SIZE, index=font_index)
        zip_font = ImageFont.truetype(FONT_PATH, ZIP_FONT_SIZE, index=font_index)

        print(f"フォント「{os.path.basename(FONT_PATH)}」をロードしました。")
    except IOError:
        messagebox.showerror("エラー", f"フォントファイルが見つからないか、読み込めません: {FONT_PATH}\n指定されたフォントファイルがスクリプトと同じディレクトリにあるか、PyInstallerの--add-dataオプションで正しくバンドルされているか確認してください。")
        _clean_up_temp_csv() # 異常終了時も削除
        return
    except Exception as e:
        messagebox.showerror("エラー", f"フォントのロード中に予期せぬエラーが発生しました: {e}")
        _clean_up_temp_csv() # 異常終了時も削除
        return

    # --- PDFファイル名の指定と保存場所の選択 ---
    messagebox.showinfo("PDFファイル保存", "生成されたPDFファイルの保存先とファイル名を指定してください。")
    output_pdf_path = filedialog.asksaveasfilename(
        title="PDFファイルを保存",
        defaultextension=".pdf",
        filetypes=[("PDFファイル", "*.pdf"), ("全てのファイル", "*.*")],
        initialfile="generated_postcards.pdf"
    )

    # 🚨 ここでキャンセルをチェックし、キャンセルされたらTemp.csvを削除して終了
    if not output_pdf_path:
        messagebox.showwarning("処理中断", "PDFファイルの保存先が指定されませんでした。スクリプトを終了します。")
        _clean_up_temp_csv() # ⬅️ **キャンセル時のクリーンアップを追加**
        return

    # 選択されたパスからディレクトリを抽出し、存在しない場合は作成
    OUTPUT_DIR = os.path.dirname(output_pdf_path) or '.'
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"出力フォルダ「{OUTPUT_DIR}」を作成しました。")
    else:
        print(f"出力フォルダ「{OUTPUT_DIR}」を使用します。")

    # --- CSVファイルの読み込みと画像生成 ---
    if not os.path.exists(TEMP_CSV_PATH):
        messagebox.showerror("エラー", f"一時データファイルが見つかりません: {TEMP_CSV_PATH}\n前処理が正しく実行されたか確認してください。")
        return

    # 処理対象のファイル数を事前に取得
    try:
        with open(TEMP_CSV_PATH, 'r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            next(reader, None) # ヘッダー行をスキップ
            total_files_to_process = sum(1 for row in reader)
            
    except Exception as e:
        messagebox.showerror("エラー", f"一時CSVファイルの読み込み中にエラーが発生しました: {e}")
        _clean_up_temp_csv() # 異常終了時も削除
        return

    create_progress_window()
    pdf_pages = []

    try:
        # 一時CSVからデータ読み込み
        with open(TEMP_CSV_PATH, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            rows = list(reader)

        for i, row in enumerate(rows):
            current_name_for_progress = row.get('氏名', '不明')
            update_progress(current_name_for_progress, i + 1, total_files_to_process)

            try:
                img = Image.open(TEMPLATE_IMAGE_PATH).copy()
                img = img.convert("RGB")
            except FileNotFoundError:
                messagebox.showerror("致命的なエラー", f"テンプレート画像「{TEMPLATE_IMAGE_PATH}」が見つかりません。")
                destroy_progress_window()
                raise # エラーを上位に送る

            draw = ImageDraw.Draw(img)

            # CSVから必要なデータを取り出す (Temp.csvのヘッダー名を使用)
            name1_raw = row.get('氏名', '').strip()
            name2_raw = row.get('氏名２', '').strip()
            zip_code_raw = row.get('郵便番号', '').strip()
            address1_raw = row.get('住所１', '').strip()
            address2_raw = row.get('住所２', '').strip()
            title = row.get('敬称', DEFAULT_TITLE).strip()
            title2 = row.get('敬称２', '').strip()

            # データ変換処理 (Temp.csvのデータは未変換なので、ここで最終変換を行う)
            address1_fullwidth = _convert_halfwidth_to_fullwidth_all(address1_raw)
            address2_fullwidth = _convert_halfwidth_to_fullwidth_all(address2_raw)
            address1_final = _convert_address_numbers_and_hyphens(address1_fullwidth)
            address2_final = _convert_address_numbers_and_hyphens(address2_fullwidth)
            name1_fullwidth = _convert_halfwidth_to_fullwidth_all(name1_raw)
            name1_final = _normalize_name_spacing(name1_fullwidth)
            name2_full_converted = _convert_halfwidth_to_fullwidth_all(name2_raw)
            
            # 敬称のデータも念のため全角変換
            title_final = _convert_halfwidth_to_fullwidth_all(title)
            title2_final = _convert_halfwidth_to_fullwidth_all(title2)


            # --- 郵便番号の処理 (横書き) ---
            zip_code = re.sub(r'[^0-9]', '', zip_code_raw)
            if len(zip_code) == 7:
                draw_horizontal_zip_code(draw, zip_code, zip_font,
                                         ZIP_OVERALL_LEFT_X_PX, ZIP_COMMON_Y,
                                         ZIP_CHAR_OFFSETS, TEXT_COLOR)

            # --- 住所の描画 (縦書き) ---
            draw_vertical_text(img, draw, address1_final, address_font, ADDRESS_COL1_X, ADDRESS_LINE_Y_START, ADDRESS_CHAR_Y_SPACING, TEXT_COLOR)
            if address2_final:
                draw_vertical_text(img, draw, address2_final, address_font, ADDRESS_COL2_X, ADDRESS_LINE_Y_START, ADDRESS_CHAR_Y_SPACING, TEXT_COLOR)

            # --- 氏名全体の描画ロジック ---
            # 氏名1の名字と名前を分割
            surname1_part, name1_first_name_part = name1_final.split('　', 1) if '　' in name1_final else (name1_final, "")
            # 氏名2の名字と名前を分割
            surname2_part, name2_first_name_part = name2_full_converted.split('　', 1) if '　' in name2_full_converted else ("", name2_full_converted)

            # 各氏名の名字の描画後のY座標を計算
            temp_y_after_surname1 = NAME_LINE_Y_START + len(surname1_part) * NAME_CHAR_Y_SPACING
            temp_y_after_surname2 = NAME_LINE_Y_START + (len(surname2_part) * NAME_CHAR_Y_SPACING if surname2_part else 0)

            # 名前（名字の後ろの部分）の開始Y座標を、長い方の名字に合わせて統一
            # 🚨 連名の縦の揃えを決定する重要な計算
            unified_name_start_y = max(temp_y_after_surname1, temp_y_after_surname2) + OFFSET_NAME1_Y_FROM_SURNAME_END_AFTER_SPACE

            # 1. 氏名1の名字を描画
            draw_vertical_text(img, draw, surname1_part, name_font, NAME_COL1_X, NAME_LINE_Y_START, NAME_CHAR_Y_SPACING, TEXT_COLOR)
            # 2. 氏名1の名前を描画
            if name1_first_name_part:
                draw_vertical_text(img, draw, name1_first_name_part, name_font, NAME_COL1_X, unified_name_start_y, NAME_CHAR_Y_SPACING, TEXT_COLOR)

            # 3. 氏名2（連名）の描画ロジック
            if name2_full_converted:
                name2_draw_x = NAME_COL1_X + OFFSET_NAME2_X_FROM_NAME1_COL
                if surname2_part:
                    draw_vertical_text(img, draw, surname2_part, name2_font, name2_draw_x, NAME_LINE_Y_START, NAME_CHAR_Y_SPACING, TEXT_COLOR)
                if name2_first_name_part:
                    draw_vertical_text(img, draw, name2_first_name_part, name2_font, name2_draw_x, unified_name_start_y, NAME_CHAR_Y_SPACING, TEXT_COLOR)
            
            # 敬称のY座標決定
            # 氏名1の最終Y座標
            final_y_name1_end = unified_name_start_y + (len(name1_first_name_part) * NAME_CHAR_Y_SPACING if name1_first_name_part else len(surname1_part) * NAME_CHAR_Y_SPACING)
            
            # 氏名2の最終Y座標
            final_y_name2_end = unified_name_start_y + (len(name2_first_name_part) * NAME_CHAR_Y_SPACING if name2_full_converted and name2_first_name_part else (temp_y_after_surname2 if surname2_part else NAME_LINE_Y_START))
            
            # 敬称の開始Y座標は、氏名1と氏名2のより下にある氏名の終端Y座標にオフセット (10) を加える
            unified_title_start_y = max(final_y_name1_end, final_y_name2_end) + OFFSET_TITLE_Y_FROM_NAME_END

            # 🚨 最終修正: 氏名1にスペースがない場合（=企業名など）に敬称の位置を上に調整
            
            # 敬称1の調整
            adjusted_title1_y = unified_title_start_y
            
            # 氏名1の元のデータ（name1_final）に全角スペースが含まれないかどうかをチェック
            if '　' not in name1_final:
                # 企業名など、長い名前の後ろに敬称が来た場合、上に移動させる
                # 調整値を -600 に設定
                adjusted_title1_y -= 600 # <- 最終調整値 600
            
            # 敬称2の調整 (連名の敬称は基本調整しない)
            adjusted_title2_y = unified_title_start_y
            
            # --- 描画 ---
            
            # 4. 敬称1の描画 (調整後のY座標を使用)
            draw_vertical_text(img, draw, title_final, title_font, NAME_COL1_X, adjusted_title1_y, TITLE_FONT_SIZE, TEXT_COLOR)
            
            # 5. 敬称2の描画（存在する場合のみ）
            if title2_final:
                title_draw_x_sub = NAME_COL1_X + OFFSET_NAME2_X_FROM_NAME1_COL
                # 調整後のY座標を使用
                draw_vertical_text(img, draw, title2_final, title_font, title_draw_x_sub, adjusted_title2_y, TITLE_FONT_SIZE, TEXT_COLOR)

            pdf_pages.append(img)
            print(f"「{name1_raw}」様のハガキ画像を一時的に生成しました。")

        # 全ての画像を1つのPDFファイルに保存
        if pdf_pages:
            if len(pdf_pages) > 1:
                pdf_pages[0].save(output_pdf_path, save_all=True, append_images=pdf_pages[1:], resolution=TEMPLATE_DPI)
            else:
                pdf_pages[0].save(output_pdf_path, resolution=TEMPLATE_DPI)
            print(f"\n全てのハガキを1つのPDFファイルにまとめました: {output_pdf_path}")
        else:
            print("\n生成されたハガキがありませんでした。")


    except Exception as e:
        messagebox.showerror("エラー", f"スクリプト実行中に予期せぬエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
    finally:
        destroy_progress_window()
        # 最終的なクリーンアップ
        _clean_up_temp_csv() # ⬅️ **正常/異常終了時のクリーンアップを実行**

    messagebox.showinfo("処理完了", f"全てのハガキ画像の生成が完了し、PDFファイルが作成されました！\n\n「{output_pdf_path}」に保存されています。試し印刷して位置を確認してください。")
    print("\n全てのハガキ画像の生成が完了しました。")


# ----------------------------------------------------------------------
#                             メイン実行ブロック
# ----------------------------------------------------------------------

def process_and_create_temp_csv():
    """
    元のCSVを選択し、ヘッダーを正規化し、データを抽出し、Temp.csvを作成する。
    """
    root = tk.Tk()
    root.withdraw() 
    messagebox.showinfo("postcard_generator", "バージョン 1.1.1 2025/11/18")
    messagebox.showinfo("CSVファイル選択", "住所録CSVファイルを選択してください。")
    CSV_FILE_PATH = filedialog.askopenfilename(
        title="住所録CSVファイルを選択",
        filetypes=[("CSVファイル", "*.csv"), ("全てのファイル", "*.*")]
    )
    if not CSV_FILE_PATH:
        messagebox.showwarning("処理中断", "CSVファイルが選択されませんでした。スクリプトを終了します。")
        return False

    detected_encoding = None
    
    try:
        # エンコーディングの検出とヘッダーの正規化
        with open(CSV_FILE_PATH, 'rb') as f:
            raw_data = f.read(4096)
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        
        with open(CSV_FILE_PATH, 'r', encoding=detected_encoding) as csvfile:
            reader = csv.reader(csvfile)
            raw_header = next(reader, None) 
            if raw_header is None: 
                raise Exception("CSVファイルが空か、ヘッダー行が読み込めませんでした。")
            normalized_header = _normalize_header(raw_header)

        # CSVデータをメモリに読み込み
        rows_raw = []
        with open(CSV_FILE_PATH, 'r', encoding=detected_encoding) as csvfile:
            # 最初の行をヘッダーとしてスキップしないように、fieldnamesには正規化されたヘッダーを使用
            # DictReaderは最初の行をヘッダーとして読むため、fieldnamesに正規化済みヘッダーを渡し、next(reader_dict)でヘッダー行をスキップする
            reader_dict = csv.DictReader(csvfile, fieldnames=normalized_header)
            next(reader_dict) 
            rows_raw = list(reader_dict)

        rows_processed = []
        
        # 元のCSVが持っている有効なフィールド名（COL_Xを除く）をセットで保持
        original_fields_set = set([h.strip() for h in normalized_header if not h.startswith('COL_')])

        # --- データ取得ヘルパー関数 ---
        def get_data(row, keys):
            for key in keys:
                data = row.get(key)
                if data is None: continue
                data = str(data).strip()
                if data.lower() not in ('none', 'null', ''):
                    return data
            return ""
        
        for i, row in enumerate(rows_raw):
            
            if row is None or not isinstance(row, dict):
                continue
            
            # --- 氏名データ取得ロジック (整形・結合ロジック) ---
            # 既存の '氏名' 列を優先
            name1_raw_data = get_data(row, HEADER_CANDIDATES.get('氏名', []))
            
            final_name1 = ""
            if name1_raw_data:
                final_name1 = name1_raw_data
            else:
                # '姓' + '名' を結合するロジック（元のコードにはないが、将来的な対応のため保持）
                surname_data = get_data(row, ['姓'])
                firstname_data = get_data(row, ['名'])
                if surname_data or firstname_data:
                    # ここでは半角スペースで結合し、makePostalCardPDF内で全角に変換する
                    combined_name = f"{surname_data} {firstname_data}".strip(' ')
                    final_name1 = combined_name
            
            name2_raw_data = get_data(row, HEADER_CANDIDATES.get('氏名２', []))

            # --- その他のデータ取得 ---
            zip_code_data = get_data(row, HEADER_CANDIDATES.get('郵便番号', []))
            address1_data = get_data(row, HEADER_CANDIDATES.get('住所１', []))
            address2_data = get_data(row, HEADER_CANDIDATES.get('住所２', []))
            
            honorific_data = get_data(row, HEADER_CANDIDATES.get('敬称', []))
            honorific2_data = get_data(row, HEADER_CANDIDATES.get('敬称２', []))
            
            # --- 敬称データ補完ロジック（敬称１のみを補完） ---
            # '敬称' 列が存在し、かつデータがない場合、強制的に「様」を挿入
            if honorific_data == "" and any(c in original_fields_set for c in HEADER_CANDIDATES.get('敬称', [])):
                honorific_data = DEFAULT_TITLE
            
            # 氏名、連名、郵便番号、住所1のいずれもデータがない場合はスキップ
            if not any([final_name1, name2_raw_data, zip_code_data, address1_data]):
                continue

            # --- 出力用構造に全データを格納する（厳密なヘッダー名を使用） ---
            output_row = {
                '氏名': final_name1, 
                '氏名２': name2_raw_data, 
                '郵便番号': zip_code_data, 
                '住所１': address1_data, 
                '住所２': address2_data, 
                '敬称': honorific_data,
                '敬称２': honorific2_data
            }
            rows_processed.append(output_row)
            
        # 5. 出力用ヘッダーの決定 (元のCSVに存在した場合のみ出力)
        temp_csv_header = []
        
        for final_field in OUTPUT_FIELDS_ORDER:
            # '氏名' は無条件で採用
            if final_field == '氏名':
                temp_csv_header.append(final_field)
                continue
            
            is_present = False
            
            # HEADER_CANDIDATESに定義されているフィールドの場合
            if final_field in HEADER_CANDIDATES:
                # 候補となる元のヘッダー名のいずれかが存在したかチェック
                for candidate in HEADER_CANDIDATES[final_field]:
                    if candidate in original_fields_set:
                        is_present = True
                        break
            
            # 存在した場合のみ、最終ヘッダーに追加
            if is_present:
                temp_csv_header.append(final_field)
                
        with open(TEMP_CSV_PATH, 'w', encoding='utf-8', newline='') as tmpfile:
            writer = csv.DictWriter(tmpfile, fieldnames=temp_csv_header, extrasaction='ignore')
            writer.writeheader() 
            writer.writerows(rows_processed)
            
    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました: {e}\nファイル形式やデータ内容を確認してください。")
        if os.path.exists(TEMP_CSV_PATH):
            print(f"【エラー時】一時ファイルが残っています: {TEMP_CSV_PATH}")
        sys.exit()
        
    return True

if __name__ == '__main__':
    # 1. CSV前処理とTemp.csvの作成
    if process_and_create_temp_csv():
        # 2. Temp.csvを読み込んでPDF生成
        makePostalCardPDF()
