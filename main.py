from pdf2jpg import pdf2jpg
from functions import get_targets, get_box3, get_box1, get_box2, delete_line, get_table, get_name, crop2, delete_line2

from pathlib import Path, PurePath
import os
import pathlib
import pyocr
from PIL import Image
import pandas as pd


pdf_files = Path('pdfs').resolve().glob('*.pdf')
img_dir = Path('out_img').resolve()

# poppler環境変数登録
poppler_path = r'C:\Users\g2945\poppler\bin'
if poppler_path not in os.environ["PATH"].split(os.pathsep):
    os.environ['PATH'] += os.pathsep + str(poppler_path)

# tesseract環境変数登録
tesseract_path = r'C:\Users\g2945\tesseract'
if tesseract_path not in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + tesseract_path

# 変換元のpdfﾃﾞｨﾚｸﾄﾘがあるか判定
if not Path('pdfs').resolve().exists():
    print('<pdfsファイル>を作成してください。')

# out_imgﾃﾞｨﾚｸﾄﾘの中身確認
out_img_list = []
for jpg in list(img_dir.glob('*.jpg')):
    out_img_list.append(jpg.stem)
# out_imgにpdf名がない場合に変換実行
for pdf in pdf_files:
    if not pdf.stem in out_img_list:
        pdf2jpg(pdf_files, img_dir)
    else:
        print('<', pdf.stem, '> はすでに変換済みです', sep='')



tools = pyocr.get_available_tools()
tool = tools[0]
text_builder = pyocr.builders.TextBuilder()
table_builder = pyocr.builders.DigitBuilder(tesseract_layout=6)


# ファイル内の変換対象全取得
targets = get_targets(img_dir)

# 本体
base_cols = ['工事名', '事務所', '入札日', '予定価格', '調査基準価格', '社名', '基礎点',
        '施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']
table_cols = ['CPD', '同種類似工事施工経験', '工事成績', '優良技術者表彰', '小計1', '企業：同種類似工事施工経験',
              '企業：工事成績', '企業：工事に係る表彰', '企業：近隣地域での施工実績', '企業：災害支援に係る表彰等',
              '企業：事故及び不誠実な行為等に対する評価', '企業：小計', '企業：AS舗装、海上作業船団施工体制', '企業：小計2',
              '小計2', '評価点合計', 'A:加算点', '品質確保の実効性', '施工体制確保の実効性', 'B:施工体制評価点合計', 'A+B']
cols = base_cols + table_cols

df = pd.DataFrame(index=[], columns=cols)
df = df.set_index('社名')

text_dir = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'out_txts')
if not text_dir.exists():
    text_dir.mkdir()
base_path = pathlib.PurePath.joinpath(text_dir, 'base_df.csv')


result_data = []
result_eva = []
result_table = []

for i, imgs in enumerate(targets):
    print(i+1, '/', len(targets), 'ファイル変換中', sep='')
    for img in imgs:
        im = Image.open(img)
        # 1ページ目（調書）読み取り
        if '1' in img.stem:
            print('1ページ目（調書）読み取り中')
            # BOX3を読み取り
            box3 = get_box3(tool, text_builder, im)
            # BOX1を読み取り
            box1 = get_box1(tool, text_builder, im)
            result_data.append(box3 + box1)
            # BOX2を読み取り
            box2 = get_box2(tool, text_builder, im)
            result_eva.append(box2)
        elif '_cut' in img.stem:
            continue
        # 2ページ目(表)読み取り
        else:
            print('2ページ目(表)読み取り中')
            c_list = crop2(img)
            for i, crop_img in enumerate(c_list):
                no_line_path = delete_line2(crop_img, img, i)
                no_line_img = Image.open(no_line_path)
                # os.remove(no_line_path)
                table_text = tool.image_to_string(no_line_img, lang="jpn", builder=table_builder).split()
                if len(table_text) < 18:
                    continue
                result_table.append(table_text)

    # まとめてdf化
for rd, r_e, rt in zip(result_data, result_eva, result_table):
    for rr_e, rrt in zip(r_e, rt):
        r = rd + rr_e + rrt
        if len(r) <= 33:
            c = 33 - len(r)
            for i in range(c):
                r2 = r.append('読み取り失敗')
                sd = pd.Series(r2, index=cols)
        elif len(r) >= 33:
            c = len(r) - 33
            for i in range(c):
                r2 = r.pop(i+12)
                sd = pd.Series(r2, index=cols)

        sd = pd.Series(r, index=cols)
        df = df.append(sd, ignore_index=True)
        df.to_csv(base_path)

# 大本のファイル保存
# df = df.set_index('社名')


# print(result_eva)
# print(result_data)