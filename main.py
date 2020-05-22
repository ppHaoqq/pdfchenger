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

# pdfsファイル作ってね
pdf_files = pathlib.Path('pdfs').glob('*.pdf')

img_dir = pathlib.Path('out_img')
pdf_files = pathlib.Path('pdfs').glob('*.pdf')

for pdf_file in pdf_files:
    imgs2 = pdf2image.convert_from_path(pdf_file)
    pgs = str(len(imgs2))
    title = pdf_file.stem + '_ '+ pgs + 'p'
    title_dir = pathlib.PurePath.joinpath(img_dir, title)
    if not title_dir.exists():
        title_dir.mkdir()
        for pg, img2 in enumerate(imgs2):
            save_path = pathlib.PurePath.joinpath(title_dir, 'pg-{}.jpg'.format(pg+1))
            img2.save(save_path, 'JPEG')
    else:
        print('<', pdf_file.stem, '>はすでに変換済みです', sep='')

# JPEG → excell
tesseract_path = (r'C:\Users\g2945\tesseract')

if tesseract_path not in os.environ["PATH"].split(os.pathsep):
    os.environ["PATH"] += os.pathsep + tesseract_path

tools = pyocr.get_available_tools()
tool = tools[0]

img = Image.open(r'C:\Users\g2945\Desktop\jupyter\OCR\out_img\入札調書（R1-2窪川）_ 2p\pg-1.jpg')

builder = pyocr.builders.TextBuilder()
result = tool.image_to_string(img, lang="jpn", builder=builder)
