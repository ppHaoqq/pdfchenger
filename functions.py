import re
import os
import pathlib
import pyocr
from PIL import Image
import pandas as pd


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

    # tools = pyocr.get_available_tools()
    # tool = tools[0]
    # builder = pyocr.builders.TextBuilder()

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

    # tools = pyocr.get_available_tools()
    # tool = tools[0]
    # builder = pyocr.builders.TextBuilder()

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