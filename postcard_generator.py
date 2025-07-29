import csv
from PIL import Image, ImageDraw, ImageFont
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re # 郵便番号のハイフン削除用
import unicodedata # 半角→全角変換用
import chardet # エンコーディング自動検出用

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
# ※ここにある「パス」や「座標」「フォントサイズ」は、お使いのテンプレート画像や
# プリンターの出力結果に合わせて適宜調整してください。

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

CALC_ZIP_TOP_MARGIN_PX = int(ZIP_TOP_MARGIN_MM * PIXELS_PER_MM) # ここを修正しました
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

# --- 敬称の設定 ---
DEFAULT_TITLE = '様' # デフォルトの敬称

# --- ヘルパー関数 ---
def _convert_halfwidth_to_fullwidth_all(text):
    """
    文字列中の全ての半角文字（英数字、記号、カタカナ、スペース）を全角に変換する。
    特に半角アルファベットの変換を強化。
    """
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
    ただし、数字の直後にアルファベットやカタカナ、特定の単位を表す漢字が続く場合は、
    その数字は漢数字に変換せず、全角数字のままにする。
    この関数は、_convert_halfwidth_to_fullwidth_all の後に呼び出されることを想定。
    これにより、半角ハイフンは先に全角ハイフンに変換され、その後縦棒になる。
    """
    kanji_map = {
        '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
        '5': '五', '6': '六', '7': '七', '8': '八', '9': '九',
    }

    # フェーズ1: 漢数字変換をスキップすべき部分を特定し、プレースホルダーに置き換える
    # パターン: 1桁以上の全角数字が続き、その直後に特定の非数字文字が1つ以上続く
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
        if char == '_' and temp_text[i:i+16].startswith('__PLACEHOLDER_'): # __PLACEHOLDER_XX__ の長さ
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
    normalized_name = name_fullwidth_input.strip('　')
    normalized_name = re.sub(r'　+', '　', normalized_name)
    return normalized_name


def draw_vertical_text(img_obj, draw_obj, text, font, start_x, start_y, char_y_spacing, text_color=(0,0,0)):
    """
    縦書きでテキストを描画するヘルパー関数。
    一文字ずつ描画し、縦に積み重ねる。
    全角文字（日本語、漢数字、全角縦棒、全角アルファベットなど）は回転しない。
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
    Returns the X-coordinate after the last character is drawn and its offset is applied.
    """
    current_x = start_x
    for i, char in enumerate(text):
        draw_obj.text((current_x, start_y), char, font=font, fill=text_color)
        char_length = font.getlength(char)
        current_x += char_length
        
        if i < len(char_offsets):
            current_x += char_offsets[i]
    return current_x


# --- テンプレートの準備 ---
if GENERATE_TEMPLATE:
    try:
        TEMPLATE_IMAGE_PATH = generate_postcard_template(
            AUTO_TEMPLATE_FILENAME, TEMPLATE_DPI,
            POSTCARD_WIDTH_MM, POSTCARD_HEIGHT_MM
        )
    except Exception as e:
        messagebox.showerror("エラー", f"テンプレートの自動生成に失敗しました: {e}\n「GENERATE_TEMPLATE」をFalseに設定し、手動でテンプレート画像を用意してください。")
        sys.exit()
else:
    if not os.path.exists(TEMPLATE_IMAGE_PATH):
        messagebox.showerror("エラー", f"指定されたテンプレート画像が見つかりません: {TEMPLATE_IMAGE_PATH}\n「TEMPLATE_IMAGE_PATH」の設定と、ファイルが存在するか確認してください。")
        sys.exit()

# --- フォントのロード ---
name_font = None
name2_font = None
title_font = None
address_font = None
zip_font = None

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
    sys.exit()
except Exception as e:
    messagebox.showerror("エラー", f"フォントのロード中に予期せぬエラーが発生しました: {e}")
    sys.exit()

# --- CSVファイルの選択 ---
messagebox.showinfo("CSVファイル選択", "次に、住所録CSVファイルを選択してください。\n\nCSVファイルには以下のヘッダーが必要です:\n氏名,郵便番号,住所１\n\nオプションで連名用: 氏名２\nオプションで敬称個別指定用: 敬称\nオプションで連名用の敬称: 敬称２\nオプションで住所詳細: 住所２\n\n**全ての半角文字（英数字、カタカナ、記号、スペースを含む）は自動的に全角に変換されます。\n氏名１に含まれる全角スペースは自動的に1つに正規化されます。複数のスペースを入れすぎるとレイアウトが崩れる可能性があります。\n氏名２には、名字（スペース区切りで）と名前を入力してください。名字がない場合は名前のみで構いません。\n住所中の半角・全角ハイフンは自動で縦棒に、半角数字は漢数字に変換されます。**")

CSV_FILE_PATH = filedialog.askopenfilename(
    title="住所録CSVファイルを選択",
    filetypes=[("CSVファイル", "*.csv"), ("全てのファイル", "*.*")]
)

if not CSV_FILE_PATH:
    messagebox.showwarning("処理中断", "CSVファイルが選択されませんでした。スクリプトを終了します。")
    sys.exit()

# --- PDFファイル名の指定と保存場所の選択 ---
messagebox.showinfo("PDFファイル保存", "生成されたPDFファイルの保存先とファイル名を指定してください。")

output_pdf_path = filedialog.asksaveasfilename(
    title="PDFファイルを保存",
    defaultextension=".pdf",
    filetypes=[("PDFファイル", "*.pdf"), ("全てのファイル", "*.*")],
    initialfile="generated_postcards.pdf"
)

if not output_pdf_path:
    messagebox.showwarning("処理中断", "PDFファイルの保存先が指定されませんでした。スクリプトを終了します。")
    sys.exit()

# 選択されたパスからディレクトリを抽出し、存在しない場合は作成
OUTPUT_DIR = os.path.dirname(output_pdf_path)
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    print(f"出力フォルダ「{OUTPUT_DIR}」を作成しました。")
else:
    print(f"出力フォルダ「{OUTPUT_DIR}」を使用します。")


# --- CSVファイルの読み込みと画像生成 ---
print(f"CSVファイル「{CSV_FILE_PATH}」を読み込み、ハガキ画像を生成します...")

# 処理対象のファイル数を事前に取得
detected_encoding = None
try:
    # ファイルのバイナリを読み込み、エンコーディングを検出
    with open(CSV_FILE_PATH, 'rb') as f:
        # 先頭の数バイトを読み込むことで、小さいファイルでも検出精度を上げる
        raw_data = f.read(4096) # Read up to 4KB for detection
    
    result = chardet.detect(raw_data)
    
    # 信頼度が高い場合はそのエンコーディングを使用
    if result['confidence'] > 0.9: # 信頼度をやや高めに設定
        detected_encoding = result['encoding']
        print(f"CSVファイルのエンコーディングを {detected_encoding} と高信頼度で検出しました (信頼度: {result['confidence']:.2f})")
    else:
        # 信頼度が低い場合は、一般的なエンコーディングで試行
        print(f"CSVファイルのエンコーディング検出の信頼度が低いです ({result['encoding']} 信頼度: {result['confidence']:.2f})。")
        print("一般的なエンコーディング（UTF-8, Shift-JISなど）で順次試行します...")
        
        encodings_to_try = ['utf-8', 'shift_jis', 'cp932', 'euc_jp'] # 試すエンコーディングのリスト

        for enc in encodings_to_try:
            try:
                with open(CSV_FILE_PATH, 'r', encoding=enc) as f:
                    # ヘッダーを読み込み、エラーがなければ成功とみなす
                    reader = csv.reader(f)
                    header = next(reader)
                    detected_encoding = enc
                    print(f"CSVファイルを '{enc}' エンコーディングで正常に読み込めました。")
                    break # 成功したらループを抜ける
            except UnicodeDecodeError:
                print(f"'{enc}' エンコーディングでの読み込みに失敗しました。")
                continue
            except Exception as e:
                print(f"'{enc}' エンコーディングでの読み込み中に予期せぬエラー: {e}")
                continue
        
        if not detected_encoding:
            raise Exception("適切なエンコーディングを自動検出できませんでした。ファイルが破損しているか、対応していないエンコーディングかもしれません。")

    # 検出されたエンコーディングでCSVを読み込み、行数を取得
    with open(CSV_FILE_PATH, 'r', encoding=detected_encoding) as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader, None) # ヘッダー行をスキップ
        total_files_to_process = sum(1 for row in reader)
        
except Exception as e:
    messagebox.showerror("エラー", f"CSVファイルの読み込み中にエラーが発生しました: {e}\nファイル形式やエンコーディングが正しいか確認してください。")
    sys.exit()

create_progress_window() # プログレスウィンドウを表示

# PDF用の画像リスト
pdf_pages = []

try:
    # 検出されたエンコーディングで再度CSVを開く
    with open(CSV_FILE_PATH, 'r', encoding=detected_encoding) as csvfile:
        reader = csv.DictReader(csvfile)
        rows = list(reader) # 全行を読み込み

    for i, row in enumerate(rows):
        current_name_for_progress = row.get('氏名', '不明') # プログレスバー表示用
        update_progress(current_name_for_progress, i + 1, total_files_to_process)

        # テンプレート画像を開く (自動生成されたものまたは手動で用意したもの)
        try:
            img = Image.open(TEMPLATE_IMAGE_PATH).copy()
            img = img.convert("RGB")
        except FileNotFoundError:
            messagebox.showerror("致命的なエラー", f"テンプレート画像「{TEMPLATE_IMAGE_PATH}」が見つかりません。\n自動生成に失敗したか、手動で指定したファイルが存在しません。")
            destroy_progress_window()
            sys.exit()

        draw = ImageDraw.Draw(img) # ImageDrawオブジェクトはここで作成

        # CSVから必要なデータを取り出す
        name1_raw = row.get('氏名', '').strip()
        name2_raw = row.get('氏名２', '').strip()
        zip_code_raw = row.get('郵便番号', '').strip()
        address1_raw = row.get('住所１', '').strip()
        address2_raw = row.get('住所２', '').strip()
        title = row.get('敬称', DEFAULT_TITLE).strip()
        title2 = row.get('敬称２', '').strip() # 新しい敬称２の取得

        # --- 全ての半角文字を全角に変換する処理を最初に適用 ---
        address1_fullwidth = _convert_halfwidth_to_fullwidth_all(address1_raw)
        address2_fullwidth = _convert_halfwidth_to_fullwidth_all(address2_raw)

        # --- 住所の数字を漢数字に変換し、半角・全角ハイフンを全角縦棒に変換 ---
        address1_final = _convert_address_numbers_and_hyphens(address1_fullwidth)
        address2_final = _convert_address_numbers_and_hyphens(address2_fullwidth)

        # --- 氏名の全角変換とスペース正規化 ---
        name1_fullwidth = _convert_halfwidth_to_fullwidth_all(name1_raw)
        name1_final = _normalize_name_spacing(name1_fullwidth)

        name2_full_converted = _convert_halfwidth_to_fullwidth_all(name2_raw) # 氏名2も全角変換

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
        # 氏名1の処理
        surname1_part = ""
        name1_first_name_part = ""
        if '　' in name1_final:
            name1_parts = name1_final.split('　', 1)
            surname1_part = name1_parts[0]
            name1_first_name_part = name1_parts[1]
        else:
            surname1_part = name1_final

        # 氏名2の処理
        surname2_part = ""
        name2_first_name_part = ""
        if '　' in name2_full_converted:
            name2_parts = name2_full_converted.split('　', 1)
            surname2_part = name2_parts[0]
            name2_first_name_part = name2_parts[1]
        else:
            name2_first_name_part = name2_full_converted # スペースがない場合は全体を名前とみなす

        # 各氏名の描画に必要な縦方向の長さを計算
        # 名字の長さ + スペースの長さ + 名前の長さ
        len_name1_full = len(surname1_part) * NAME_CHAR_Y_SPACING
        if name1_first_name_part:
            len_name1_full += OFFSET_NAME1_Y_FROM_SURNAME_END_AFTER_SPACE + len(name1_first_name_part) * NAME_CHAR_Y_SPACING

        len_name2_full = 0
        if name2_full_converted:
            len_name2_full = len(surname2_part) * NAME_CHAR_Y_SPACING
            if name2_first_name_part:
                len_name2_full += OFFSET_NAME1_Y_FROM_SURNAME_END_AFTER_SPACE + len(name2_first_name_part) * NAME_CHAR_Y_SPACING
        
        # 名前開始のY座標を揃えるための基準Y座標を決定
        # 名字部分が長い場合も考慮し、全体として長くなる方に合わせる
        
        # 仮描画で名字の最終Y座標を取得 (実際に描画はしない)
        # 氏名1と氏名2それぞれの名字の終端Y座標を計算
        temp_y_after_surname1 = NAME_LINE_Y_START + len(surname1_part) * NAME_CHAR_Y_SPACING
        temp_y_after_surname2 = NAME_LINE_Y_START 
        if surname2_part:
            temp_y_after_surname2 = NAME_LINE_Y_START + len(surname2_part) * NAME_CHAR_Y_SPACING

        # 名前部分が始まるY座標は、長い方の名字の終端にオフセットを加えた位置に揃える
        unified_name_start_y = max(temp_y_after_surname1, temp_y_after_surname2) + OFFSET_NAME1_Y_FROM_SURNAME_END_AFTER_SPACE


        # 1. 氏名1の名字を描画
        draw_vertical_text(img, draw, surname1_part, name_font, NAME_COL1_X, NAME_LINE_Y_START, NAME_CHAR_Y_SPACING, TEXT_COLOR)
        
        # 2. 氏名1の名前を描画 (統一された開始Y座標を使用)
        if name1_first_name_part:
            draw_vertical_text(img, draw, name1_first_name_part, name_font, NAME_COL1_X, unified_name_start_y, NAME_CHAR_Y_SPACING, TEXT_COLOR)

        # 3. 氏名2（連名）の描画ロジック
        if name2_full_converted:
            name2_draw_x = NAME_COL1_X + OFFSET_NAME2_X_FROM_NAME1_COL
            
            # 氏名2の名字が存在する場合に描画
            if surname2_part:
                draw_vertical_text(img, draw, surname2_part, name2_font, name2_draw_x, NAME_LINE_Y_START, NAME_CHAR_Y_SPACING, TEXT_COLOR)

            # 氏名2の名前を描画（統一された開始Y座標を使用）
            if name2_first_name_part:
                draw_vertical_text(img, draw, name2_first_name_part, name2_font, name2_draw_x, unified_name_start_y, NAME_CHAR_Y_SPACING, TEXT_COLOR)
        
        # 敬称のY座標を揃えるための基準Y座標を決定
        # 氏名1の最終Y座標を正確に計算
        final_y_name1_end = NAME_LINE_Y_START + len(surname1_part) * NAME_CHAR_Y_SPACING
        if name1_first_name_part:
            final_y_name1_end = unified_name_start_y + len(name1_first_name_part) * NAME_CHAR_Y_SPACING

        # 氏名2の最終Y座標を正確に計算
        final_y_name2_end = NAME_LINE_Y_START # 初期値
        if name2_full_converted:
            if surname2_part:
                final_y_name2_end = NAME_LINE_Y_START + len(surname2_part) * NAME_CHAR_Y_SPACING # 名字の最終Y
            if name2_first_name_part:
                final_y_name2_end = unified_name_start_y + len(name2_first_name_part) * NAME_CHAR_Y_SPACING # 名前の最終Y

        # 敬称のY座標は、氏名1と氏名2のより下にある氏名の終端Y座標に合わせる
        unified_title_start_y = max(final_y_name1_end, final_y_name2_end) + OFFSET_TITLE_Y_FROM_NAME_END
        
        # 4. 敬称1の描画
        draw_vertical_text(img, draw, title, title_font, NAME_COL1_X, unified_title_start_y, TITLE_FONT_SIZE, TEXT_COLOR)
        
        # 5. 敬称2の描画（存在する場合のみ）
        if title2:
            title_draw_x_sub = NAME_COL1_X + OFFSET_NAME2_X_FROM_NAME1_COL # 氏名2の名前と同じX座標
            draw_vertical_text(img, draw, title2, title_font, title_draw_x_sub, unified_title_start_y, TITLE_FONT_SIZE, TEXT_COLOR)

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

messagebox.showinfo("処理完了", f"全てのハガキ画像の生成が完了し、PDFファイルが作成されました！\n\n「{output_pdf_path}」に保存されています。試し印刷して位置を確認してください。")
print("\n全てのハガキ画像の生成が完了しました。")
print(f"「{output_pdf_path}」に生成されたPDFファイルが保存されています。試し印刷して位置を確認してください。")
