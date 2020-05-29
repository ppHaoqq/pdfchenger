import re
import os
import pathlib
import pyocr
from PIL import Image
import pandas as pd

from functions import get_targets, get_box3, get_box1, get_box2


def jpg2text():
    # tesseract環境変数登録
    tesseract_path = (r'C:\Users\g2945\tesseract')

    if tesseract_path not in os.environ["PATH"].split(os.pathsep):
        os.environ["PATH"] += os.pathsep + tesseract_path

    tools = pyocr.get_available_tools()
    tool = tools[0]
    builder = pyocr.builders.TextBuilder()

    # ファイル内の変換対象全取得
    targets = get_targets()

    # 本体
    cols = ['工事名', '事務所', '入札日', '予定価格', '調査基準価格', '社名', '基礎点',
            '施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']

    df = pd.DataFrame(index=[], columns=cols)
    df = df.set_index('社名')

    df_q = pd.DataFrame(index=[], columns=cols)

    text_dir = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'out_txts')
    text_path = pathlib.PurePath.joinpath(text_dir, 'base_df.csv')

    for imgs in targets:
        result_data = []
        result_eva = []
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
    #     for r in result_eva:
    #         r = result_data + r
    #         if len(r) == len(cols):
    #             sd = pd.Series(r, index=cols)
    #             df = df.append(sd, ignore_index=True)
    #
    #         else:
    #             dummy = ['認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗', '認識失敗']
    #             r = result_data + dummy
    #             sd = pd.Series(r, index=cols)
    #             df = df.append(sd, ignore_index=True)

    #     # 社名ごとに保存
    #     for name in list(set(list(df['社名']))):
    #         df_q = pd.DataFrame(index=[], columns=cols)
    #         df2 = df.loc[df['社名'] == name]
    #         df_q = df_q.append(df2, ignore_index=True)
    #         try:
    #             df_q.to_csv(pathlib.PurePath.joinpath(text_dir, '{}_点数表.csv'.format(name)))
    #         except:
    #             df_q.to_csv(pathlib.PurePath.joinpath(text_dir, '名称不明_点数表.csv'))
    #
    # # 大本のファイル保存
    # df = df.set_index('社名')
    # df.to_csv(text_path)
    print(result_data)
    print(result_eva)