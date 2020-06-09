import re
import os
import pyocr
from PIL import Image
import pandas as pd
import cv2
import pathlib
import numpy as np

def get_targets(img_dir):
    # ファイル内の変換対象全取得
    targets = []
    # img_dir = pathlib.Path('out_img').resolve()
    for path in list(img_dir.glob('*')):
        target_path = list(path.glob('*jpg'))
        targets.append(target_path)
    return targets


def get_box3(tool, builder, im):
    # BOX3を読み取り、関数化予定
    result_data = []
    box3 = (416, 198, 937, 239)
    box4 = (416, 283, 623, 311)
    box5 = (416, 325, 665, 350)
    crop_3 = im.crop(box3)
    crop_4 = im.crop(box4)
    crop_5 = im.crop(box5)

    title = tool.image_to_string(crop_3, lang="jpn", builder=builder)
    result_data.append(title)
    office = tool.image_to_string(crop_4, lang="jpn", builder=builder)
    result_data.append(office)
    date = tool.image_to_string(crop_5, lang="jpn", builder=builder)
    result_data.append(date)

    return result_data


def get_box1(tool, builder, im):
    # BOX1を読み取り、関数化予定
    result_data = []
    box1 = (1310, 207, 2071, 247)
    box2 = (1310, 243, 2071, 283)
    crop_1 = im.crop(box1).resize((710, 40), Image.LANCZOS)
    crop_2 = im.crop(box2).resize((710, 40), Image.LANCZOS)

    schedul_price = tool.image_to_string(crop_1, lang="jpn", builder=builder)
    schedul_price = schedul_price.replace(' ', '').replace('\n', '')
    schedul_price = re.sub(r'\(.+\)', '\n', schedul_price).split('\n')
    result_data.append(schedul_price[1])

    base_price = tool.image_to_string(crop_2, lang="jpn", builder=builder)
    base_price = base_price.replace(' ', '').replace('\n', '')
    base_price = re.sub(r'\(.+\)', '\n', base_price).split('\n')
    result_data.append(base_price[1])

    return result_data


# BOX2を読み取り、関数化予定
def get_box2(tool, builder, im):
    result_eva = []
    for count in range(12):
        a = 121
        b = 620
        c = 1440
        d = 693
        box_cal = (a, b + (77 * count), c, d + (71 * count))
        crop_im = im.crop(box_cal)
        crop_text = tool.image_to_string(crop_im, lang="jpn", builder=builder).split()
        if not len(crop_text) == 0:
            name = crop_text[0] + crop_text[1]
            del crop_text[0]
            del crop_text[0]
            crop_text.insert(0, name)
            if len(crop_text) <= 3:
                crop_text.extend(['無効'] * (7 - len(crop_text)))
            result_eva.append(crop_text)
    return result_eva


def delete_line(path):
    # 書き込み用path
    p = (path.stem + '{}').format('_cut.jpg')
    write_path = pathlib.PurePath.joinpath(path.parent, p)
    if pathlib.Path(write_path).exists():
        return write_path
    else:
        # cv2で画像読み込み
        img_array = np.fromfile(path, dtype=np.uint8)
        img = cv2.imdecode(img_array, 0)

        # 処理範囲切り取り
        x, y = 355, 870
        h = int(img.shape[0] - img.shape[0] / 3)  # 2388
        w = int(img.shape[1] - img.shape[1] / 4)  # 1652
        roi = img[y:y + h, x:x + w]
        # 書き込み用画作成
        origin = roi.copy()

        # 白黒反転処理
        roi2 = cv2.bitwise_not(roi)
        # エッジ検出
        edge = cv2.Canny(roi2, 240, 250)
        # 膨張処理
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        edge = cv2.dilate(edge, kernel)
        # 直線検出1回目
        lines1 = cv2.HoughLinesP(edge, rho=1, theta=np.pi / 360, threshold=32, minLineLength=50, maxLineGap=9)
        # 線を消す(白で線を引く)
        for line in lines1:
            x1, y1, x2, y2 = line[0]
            no_lines_img = cv2.line(origin, (x1, y1), (x2, y2), (255, 255, 255), 3)
            imwrite(write_path, no_lines_img)

        # 2回目処理用画像読み込み
        array = np.fromfile(write_path, dtype=np.uint8)
        no_line = cv2.imdecode(array, cv2.IMREAD_COLOR)
        # 書き込み用画作成
        write = no_line.copy()
        # 白黒反転処理
        no_line = cv2.bitwise_not(no_line)
        # エッジ検出
        no_line = cv2.Canny(no_line, 240, 250)
        # 膨張処理
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        no_line = cv2.dilate(no_line, kernel)
        # 直線検出2回目
        lines2 = cv2.HoughLinesP(no_line, rho=1, theta=np.pi / 360, threshold=5, minLineLength=40, maxLineGap=15)
        # 線を消す(白で線を引く)
        for line in lines2:
            x1, y1, x2, y2 = line[0]
            no_lines_img2 = cv2.line(write, (x1, y1), (x2, y2), (255, 255, 255), 3)
            imwrite(write_path, no_lines_img2)

        print(write_path.stem, 'を作成しました。')
        return write_path


# パスに日本語が含まれている場合はcv2.imwrite使えないからこっち（ただのimwrite()）
def imwrite(filename, img, params=None):
    try:
        ext = os.path.splitext(filename)[1]
        result, n = cv2.imencode(ext, img, params)

        if result:
            with open(filename, mode='w+b') as f:
                n.tofile(f)
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return False


def get_table(tool, builder, img):
    table_texts = []
    for i in range(12):
        a = 0
        b = 12
        c = 1754
        d = 48
        crop = (a, b + (50 * i), c, d + (50 * i))
        crop_img = img.crop(crop)
        table_text = tool.image_to_string(crop_img, lang="jpn", builder=builder).split()
        if len(table_text) == 0:
            continue
        table_texts.append(table_text)
    return table_texts


def get_name(tool, builder, im):
    c_names = []
    for i in range(12):
        a = -(101 - 101)
        b = -(878 - 929)
        c = -(355 - 355)
        d = -(926 - 977)

        crop = (101 + (a * i), 878 + (b * i), 355 + (c * i), 926 + (d * i))

        crop_img = im.crop(crop)
        name = tool.image_to_string(crop_img, lang="jpn", builder=builder).split()
        if not name:
            continue
        name = name[0] + name[1]
        c_names.append(name)
    return c_names


# crop
def crop2(path):
    crop_list = []
    img = Image.open(path)
    for i in range(11):
        a = 363
        b = 875
        c = 2095
        d = 923
        crop = (a, (b+(3*i)+(48*i)), c, (d+(3*i)+(48*i)))
        crop_img = img.crop(crop)
        crop_list.append(crop_img)
    return crop_list


# delete_line
def delete_line2(p_img, path, count):
    # pil 2 cv2
    c_img = np.array(p_img, dtype=np.uint8)

    # 書き込み用path
    p = (path.stem + '_cut_{}.jpg').format(count + 1)
    write_path = pathlib.PurePath.joinpath(path.parent, p)
    if pathlib.Path(write_path).exists():
        return write_path
    else:
        # 書き込み用画作成
        origin = c_img.copy()
        # 白黒反転処理
        reverse = cv2.bitwise_not(c_img)
        # エッジ検出
        edge = cv2.Canny(reverse, 240, 250)
        # 膨張処理
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        edge = cv2.dilate(edge, kernel)
        # 直線検出1回目
        lines1 = cv2.HoughLinesP(edge, rho=1, theta=np.pi / 360, threshold=0, minLineLength=22, maxLineGap=5)
        # 線を消す(白で線を引く)
        for line in lines1:
            x1, y1, x2, y2 = line[0]
            no_lines_img = cv2.line(origin, (x1, y1), (x2, y2), (255, 255, 255), 3)
            imwrite(write_path, no_lines_img)

        #         # 2回目処理用画像読み込み
        #         array = np.fromfile(write_path, dtype=np.uint8)
        #         no_line = cv2.imdecode(array, cv2.IMREAD_COLOR)

        #         # 書き込み用画作成
        #         write = no_line.copy()
        #         # 白黒反転処理
        #         no_line = cv2.bitwise_not(no_line)
        #         # エッジ検出
        #         no_line = cv2.Canny(no_line, 240, 250)
        #         # 膨張処理
        #         kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        #         no_line = cv2.dilate(no_line, kernel)
        #         # 直線検出2回目
        #         lines2 = cv2.HoughLinesP(no_line, rho=1, theta=np.pi / 360, threshold=5, minLineLength=40, maxLineGap=15)
        #         # 線を消す(白で線を引く)
        #         for line in lines2:
        #             x1, y1, x2, y2 = line[0]
        #             no_lines_img2 = cv2.line(write, (x1, y1), (x2, y2), (255, 255, 255), 3)
        #             imwrite(write_path, no_lines_img2)

        print(write_path.stem, 'を処理しました。')
        return write_path

