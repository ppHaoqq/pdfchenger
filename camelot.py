# pdf →　excel ver2.0
# index外したほうが連結しやすいか
import camelot
from pathlib import Path, PurePath
import pandas as pd

table_cols = ['社名', 'CPD', '同種類似工事施工経験', '工事成績', '優良技術者表彰', '小計1', '企業：同種類似工事施工経験',
              '企業：工事成績', '企業：工事に係る表彰', '企業：近隣地域での施工実績', '企業：災害支援に係る表彰等',
              '企業：事故及び不誠実な行為等に対する評価', '企業：小計', '企業：AS舗装、海上作業船団施工体制', '企業：小計2', 'A:加算点',
              '品質確保の実効性', '施工体制確保の実効性', 'B:施工体制評価点合計', 'A+B']
base_cols = ['社名', '基礎点', '施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']

# pdfsの中身全取得
pdf_file = PurePath.joinpath(Path.cwd(), 'pdfs')
targets = list(pdf_file.glob('*.pdf'))
# 　吐き出し先無ければ作る
excel_file = PurePath.joinpath(Path.cwd(), 'excel_data')
if not excel_file.exists():
    excel_file.mkdir()
# 判定用にexcelファイルの中身を取得、STR型にしてlist化
excel_list = [xf.stem for xf in list(excel_file.glob('*.xlsx'))]

# 読み取り、保存
for target in targets:
    name = target.stem
    # 対象のpdfファイル名がexcelファイルに存在する場合処理をスキップ
    if name in excel_list:
        print(name, 'の処理をスキップします。')
        continue
    save_path = PurePath.joinpath(excel_file, '{}.xlsx'.format(name))

    tables1 = camelot.read_pdf(str(target), pages='1', shift_text=[''])
    tables2 = camelot.read_pdf(str(target), pages='2', table_regions=['50, 876, 2202, 1509'])

    df1 = tables1[0].df
    df1 = df1.T
    df1.columns = df1.loc[0]

    df2 = tables1[1].df
    df2 = df2.drop(df2.index[[0, 1, 2]])
    df2 = df2.drop(df2.columns[[7, 8, 9, 10, 11]], axis=1)
    df2.columns = base_cols
    df2.set_index('社名', inplace=True)

    df3 = tables2[0].df
    df_list = []
    for i in range(12):
        dfx = df3.loc[i].str.split('\n')
        _list = []
        for x in dfx:
            if not x == ['']:
                _list.extend(x)
        df_list.append(_list)

    df = pd.DataFrame(index=[], columns=table_cols)
    for x in df_list:
        if len(x) == 2:
            for i in range(18):
                x.append('無効')
        if len(x) == 19:
            x.insert(13, '-')
        if not len(x) == len(table_cols):
            continue

        sd = pd.Series(x, index=df.columns)
        df = df.append(sd, ignore_index=True)
    df.set_index('社名', inplace=True)

    with pd.ExcelWriter(save_path) as writer:
        df1.to_excel(writer, sheet_name='df1')
        df2.to_excel(writer, sheet_name='df2')
        df.to_excel(writer, sheet_name='df3')

    print(name, 'を処理しました。')

# df結合研究
table_cols = ['社名', 'CPD', '同種類似工事施工経験', '工事成績', '優良技術者表彰', '小計1', '企業：同種類似工事施工経験',
              '企業：工事成績', '企業：工事に係る表彰', '企業：近隣地域での施工実績', '企業：災害支援に係る表彰等',
              '企業：事故及び不誠実な行為等に対する評価', '企業：小計', '企業：AS舗装、海上作業船団施工体制', '企業：小計2','A:加算点',
              '品質確保の実効性', '施工体制確保の実効性', 'B:施工体制評価点合計', 'A+B']
base_cols = ['社名', '基礎点','施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']
cols = table_cols + base_cols

df5 = df2.join(df, how='inner')

df2_list = [list(df2.iloc[i]) for i in range(len(df2))]
df2
# dl = list(df5.iloc[0])
# dl2 = list(df1.iloc[1])
# dl3 = dl + dl2
# pd.DataFrame(dl3, columns=cols)