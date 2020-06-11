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
              '小計2', '評価点合計', 'A:加算点', '品質確保の実効性', '施工体制確保の実効性', 'B:施工体制評価点合計', 'A+B', 'flag']
cols = base_cols + table_cols


text_dir = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'out_txts')
if not text_dir.exists():
    text_dir.mkdir()
base_path = pathlib.PurePath.joinpath(text_dir, 'base_df.csv')
excel_path = pathlib.PurePath.joinpath(text_dir, 'base_df.xlsx')

result_data = []
result_eva = []
result_table = []

if base_path.exists(): # 既存のファイルに結果を追記
    df = pd.read_csv(base_path)
    base = pd.read_csv(base_path)
    flag_list = list(set(base['flag'].values.tolist()))
    for i, imgs in enumerate(targets):
        print(i + 1, '/', len(targets), 'ファイル確認中', sep='')
        flag = imgs[0].parent.stem
        if flag in flag_list:
            continue
        else:
            print(i + 1, '/', len(targets), 'ファイル変換中', sep='')
            for img in imgs:
                flag = img.parent.stem
                im = Image.open(img)
                # 1ページ目（調書）読み取り
                if 'pg-1' in img.stem:
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
                elif 'pg-2' in img.stem:
                    print('2ページ目(表)読み取り中')
                    result_tables = []
                    c_list = crop2(img)
                    for i, crop_img in enumerate(c_list):
                        no_line_path = delete_line2(crop_img, img, i)
                        no_line_img = Image.open(no_line_path)
                        table_text = tool.image_to_string(no_line_img, lang="jpn", builder=table_builder).split()
                        if len(table_text) < 18:
                            continue
                        table_text.append(flag)
                        result_tables.append(table_text)
                    result_table.append(result_tables)

    count = -1
    for eva, table in zip(result_eva, result_table):
        count += 1
        for r_e, rt in zip(eva, table):
            rd = result_data[count]
            r = rd + r_e + rt
            if len(r) < 34:
                c = 34 - len(r)
                for i in range(c):
                    r.insert(-1, '読み取り失敗')
            elif len(r) > 34:
                c = len(r) - 34
                for i in range(c):
                    r.pop(i + 12)
            sd = pd.Series(r, index=cols)
            df = df.append(sd, ignore_index=True)
    df = df.set_index('社名')
    df.to_csv(base_path, mode='a')
    df.to_excel(excel_path, mode='a')

else: # 新規で保存先を作成
    df = pd.DataFrame(index=[], columns=cols)
    for i, imgs in enumerate(targets):
        print(i + 1, '/', len(targets), 'ファイル変換中', sep='')
        for img in imgs:
            flag = img.parent.stem
            im = Image.open(img)
            # 1ページ目（調書）読み取り
            if 'pg-1' in img.stem:
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
            elif 'pg-2' in img.stem:
                print('2ページ目(表)読み取り中')
                result_tables = []
                c_list = crop2(img)
                for i, crop_img in enumerate(c_list):
                    no_line_path = delete_line2(crop_img, img, i)
                    no_line_img = Image.open(no_line_path)
                    table_text = tool.image_to_string(no_line_img, lang="jpn", builder=table_builder).split()
                    if len(table_text) < 18:
                        continue
                    table_text.append(flag)
                    result_tables.append(table_text)
                result_table.append(result_tables)
    count = -1
    for eva, table in zip(result_eva, result_table):
        count += 1
        for r_e, rt in zip(eva, table):
            rd = result_data[count]
            r = rd + r_e + rt
            if len(r) < 34:
                c = 34 - len(r)
                for i in range(c):
                    r.insert(-1, '読み取り失敗')
            elif len(r) > 34:
                c = len(r) - 34
                for i in range(c):
                    r.pop(i + 12)
            sd = pd.Series(r, index=cols)
            df = df.append(sd, ignore_index=True)
    df = df.set_index('社名')
    df.to_csv(base_path)
    df.to_excel(excel_path)

#社名ごとに保存
base = pd.read_csv(base_path)
name_list = list(base['社名'])
# set() 社名リストから重複分を削除
c_names = list(set(name_list))
for c_name in c_names:
    name_path = pathlib.PurePath.joinpath(text_dir, '{}_df.csv'.format(c_name))
    name_excel = pathlib.PurePath.joinpath(text_dir, '{}_df.xlsx'.format(c_name))
    if name_path.exists():
        name_data = base.loc[base['社名'] == c_name]
        name_data.to_csv(name_path, mode='a')
        name_data.to_excel(name_excel, mode='a')
    else:
        name_data = base.loc[base['社名'] == c_name]
        name_data.to_csv(name_path)
        name_data.to_excel(name_excel)