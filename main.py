import pathlib
import pdf2image
import os
import io
import PIL

# JPEG → excell
import os
import pathlib
import pyocr
from PIL import Image



pdfs = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'pdfs')
if not pdfs.exists():
    print('pdfsファイルを作成してください')
pdf_files = pdfs.glob('*.pdf')

img_dir = pathlib.PurePath.joinpath(pathlib.Path.cwd(), 'out_img')
if not img_dir.exists():
    img_dir.mkdir()


poppler = pathlib.Path(r'C:\Users\g2945\poppler\bin')
if not poppler.exists():
    poppler = pathlib.Path(r'C:\poppler\bin')

if not poppler in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + str(poppler)

# pdf→jpg
for pdf_file in pdf_files:
    imgs2 = pdf2image.convert_from_path(pdf_file)
    pgs = str(len(imgs2))
    title = pdf_file.stem + '_' + pgs + 'p'
    title_dir = pathlib.PurePath.joinpath(img_dir, title)
    if not title_dir.exists():
        title_dir.mkdir()
        for pg, img2 in enumerate(imgs2):
            save_path = pathlib.PurePath.joinpath(title_dir, 'pg-{}.jpg'.format(pg+1))
            img2.save(save_path, 'JPEG')
    else:
        print('<', pdf_file.stem, '>はすでに変換済みです', sep='')


# JPEG → excell

# 環境変数にテッセラクトのパス通す
tesseract_path = pathlib.Path(r'C:\Users\g2945\tesseract')
if not tesseract_path.exists():
    tesseract_path = pathlib.Path(r'C:\tesseract')

    if tesseract_path not in os.environ["PATH"].split(os.pathsep):
        os.environ["PATH"] += os.pathsep + str(tesseract_path)

tools = pyocr.get_available_tools()
tool = tools[0]
builder = pyocr.builders.TextBuilder()

# 画像データ取得
imgs = []
img_d_dir = img_dir.glob('*')
for a in img_d_dir:
    img_list = a.glob('*.jpg')
    img = Image.open(list(img_list)[0])
    imgs.append(img)

# 画像データに対してocr実行
results = []
for index, img in enumerate(imgs):
        result = tool.image_to_string(img, lang="jpn", builder=builder)
        results.append(result)


