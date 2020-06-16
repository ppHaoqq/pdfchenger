# pdf →　excel ver2.0
import camelot
import re
from pathlib import Path, PurePath
import pandas as pd

base_cols = ['社名', '工事名', '基礎点', '施工体制評価点', '合算', '入札価格', '評価値', '評価値=>基準評価値']
table_cols = ['社名', 'CPD', '同種類似工事施工経験', '工事成績', '優良技術者表彰', '小計1', '企業：同種類似工事施工経験',
              '企業：工事成績', '企業：工事に係る表彰', '企業：近隣地域での施工実績', '企業：災害支援に係る表彰等',
              '企業：事故及び不誠実な行為等に対する評価', '企業：小計', '企業：AS舗装、海上作業船団施工体制', '企業：小計2', 'A:加算点',
              '品質確保の実効性', '施工体制確保の実効性', 'B:施工体制評価点合計', 'A+B']
cols = base_cols + table_cols

# pdfsの中身全取得
pdf_file = PurePath.joinpath(Path.cwd(), 'pdfs')
targets = list(pdf_file.glob('*.pdf'))
# 　吐き出し先無ければ作る
excel_file = PurePath.joinpath(Path.cwd(), 'excel_datas')
if not excel_file.exists():
    excel_file.mkdir()
c_data = PurePath.joinpath(excel_file, 'c_data')
if not c_data.exists():
    c_data.mkdir()

base_save_path = PurePath.joinpath(excel_file, 'base.csv')
if base_save_path.exists():
    old_base = pd.read_csv(base_save_path)
    check_list = list(set(old_base['工事名']))



# 読み取り、保存
for target in targets:
    # pdf名から工事名抽出
    i_name = target.stem
    i_name = re.findall(r'(?<=\（).*?(?=\）)', i_name)[0]

    # <工事名>がbase.csvにある場合は全工程スキップ
    if base_save_path.exists():
        if i_name in check_list:
            print(i_name, 'の処理をスキップします。')
            continue

    save_path = PurePath.joinpath(c_data, '{}.xlsx'.format(i_name))

    # 表読み取り
    tables1 = camelot.read_pdf(str(target), pages='1', shift_text=[''])
    tables2 = camelot.read_pdf(str(target), pages='2', table_regions=['50, 876, 2202, 1509'])
    # 向きを整形
    df1 = tables1[0].df
    df1 = df1.drop(df1.index[[5]])
    df1 = df1.T
    df1.columns = df1.loc[0]
    # いらない行削除
    df2 = tables1[1].df
    df2 = df2.drop(df2.index[[0, 1, 2]])
    df2 = df2.drop(df2.columns[[7, 8, 9, 10, 11]], axis=1)
    df2.insert(1, '工事名', i_name)
    df2.columns = base_cols
    df2.set_index('社名', inplace=True)
    # 一気に読みとった数字を分離
    df3 = tables2[0].df
    df = pd.DataFrame(index=[], columns=table_cols)

    for i in range(12):
        dfx = df3.loc[i].str.split('\n')
        _list = []
        for x in dfx:
            if not x == ['']:
                _list.extend(x)

        if len(_list) == 2:
            for i in range(18):
                _list.append('無効')
        if len(_list) == 19:
            _list.insert(13, '-')
        if not len(_list) == len(table_cols):
            continue

        sd = pd.Series(_list, index=df.columns)
        df = df.append(sd, ignore_index=True)
    df.set_index('社名', inplace=True)

    with pd.ExcelWriter(save_path) as writer:
        df1.to_excel(writer, sheet_name='df1')
        df2.to_excel(writer, sheet_name='df2')
        df.to_excel(writer, sheet_name='df3')

    print(i_name, 'を処理しました。')

# 全体保存

excel_file = PurePath.joinpath(Path.cwd(), 'excel_datas')
if not excel_file.exists():
    excel_file.mkdir()

# 工事ごと＝c_dataの中身取得
excel_target = [xf.stem for xf in list(c_data.glob('*.xlsx'))]
extend_df = pd.DataFrame(index=[])  # , columns=new_cols

for excel in excel_target:
    excel_path = PurePath.joinpath(c_data, '{}.xlsx'.format(excel))
    r_df1 = pd.read_excel(excel_path, sheet_name='df1', index_col=0)
    r_df2 = pd.read_excel(excel_path, sheet_name='df2')
    r_df3 = pd.read_excel(excel_path, sheet_name='df3')

    # 結合用にいったんリスト化
    df1_list = [list(r_df1.iloc[i]) for i in range(len(r_df1))]
    df2_list = [list(r_df2.iloc[i]) for i in range(len(r_df2))]
    df3_list = [list(r_df3.iloc[i]) for i in range(len(r_df3))]

    new_cols = (df1_list[0] + cols)

    # リスト結合、全体DF作成
    for l_df2, l_df3 in zip(df2_list, df3_list):
        ex_list = df1_list[1] + l_df2 + l_df3
        ex_df = pd.DataFrame(ex_list)
        ex_df.index = new_cols
        ex_df = ex_df.T
        extend_df = extend_df.append(ex_df)
# 全体DF保存、追記のためCSVで保存 すでにファイルがある場合はヘッダーなしで追記
if base_save_path.exists():
    extend_df.to_csv(base_save_path, mode='a', header=False, encoding='utf_8_sig')
    print('base.csvに追記しました。')
else:
    extend_df.to_csv(base_save_path, encoding='utf_8_sig')
    print('base.csvを新規作成しました。')


# 各社ごとのデータ保存用ファイル作成
name_data_ = PurePath.joinpath(excel_file, 'name_data')
if not name_data_.exists():
    name_data_.mkdir()
# base.csvにある社名を抜き出して社名ごとに集約、保存
r_base = pd.read_csv(base_save_path)
name_list = list(set(r_base['社名']))
for c_name in name_list:
    name_path = PurePath.joinpath(name_data_, '{}.xlsx'.format(c_name))
    name_data = r_base.loc[r_base['社名'] == c_name]
    name_data.to_excel(name_path)
    print(c_name, '.xlsxを保存しました。', sep='')