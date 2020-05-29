import pathlib
import pdf2image
import os

# pdf→jpeg

# img_dir = pathlib.Path('out_img')
# pdf_files = pathlib.Path('pdfs').glob('*.pdf')

def pdf2jpg(pdf_files, img_dir):

    if not img_dir.exists():
        img_dir.mkdir()

    for pdf_file in pdf_files:
        imgs = pdf2image.convert_from_path(pdf_file)
        pgs = str(len(imgs))
        title = pdf_file.stem + '_ ' + pgs + 'p'
        title_dir = pathlib.PurePath.joinpath(img_dir, title)
        title_dir = title_dir.resolve()
        if not title_dir.exists():
            title_dir.mkdir()
            for pg, img in enumerate(imgs):
                save_path = pathlib.PurePath.joinpath(title_dir, 'pg-{}.jpg'.format(pg+1))
                img.save(save_path, 'JPEG')
        else:
            print('<', pdf_file.stem, '>はすでに変換済みです', sep='')