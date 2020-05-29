from pdf2jpg import pdf2jpg
from jpg2text import jpg2text
from functions import get_targets, get_box3, get_box1, get_box2

from pathlib import Path, PurePath
import re
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
builder = pyocr.builders.TextBuilder()

# ファイル内の変換対象全取得
targets = get_targets(img_dir)

# 本体
cols = ['工事名', '事務所', '入札日', '予定価格', '調査基準価格', '社名', '基礎点',
        '施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']

df = pd.DataFrame(index=[], columns=cols)
df = df.set_index('社名')

df_q = pd.DataFrame(index=[], columns=cols)

text_dir = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'out_txts')
text_path = pathlib.PurePath.joinpath(text_dir, 'base_df.csv')


result_data = []
result_eva = []
for i, imgs in enumerate(targets):
    print(i+1, '/', len(targets), 'ファイル変換中', sep='')
    for img in imgs:
        im = Image.open(img)
        if '1' in img.stem:
            # BOX3を読み取り
            box3 = get_box3(tool, builder, im)
            # BOX1を読み取り
            box1 = get_box1(tool, builder, im)
            result_data.append(box3 + box1)
            # BOX2を読み取り
            box2 = get_box2(tool, builder, im)
            result_eva.append(box2)

    # まとめてdf化
for r in result_eva:
    for r_d in result_data:
        r = r_d + r
        if len(r) == len(cols):
            sd = pd.Series(r, index=cols)
            df = df.append(sd, ignore_index=True)

        else:
            dummy = ['認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗']
            r = r_d + dummy
            sd = pd.Series(r, index=cols)
            df = df.append(sd, ignore_index=True)
print(df)

# print(result_eva)
# print(result_data)