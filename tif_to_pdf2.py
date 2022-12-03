from reportlab.pdfgen import canvas
import reportlab
from reportlab.lib.pagesizes import letter
import PIL
from PIL import Image
import sys


def convert(tiff, rotate_img = 0): # по умолчанию rotate_img=0 (не переворачивать документ)
    try:  
        if rotate_img != 0 and rotate_img != 1:
            rotate_img = 0
        img = PIL.Image.open(tiff)
        width, height = img.size
        scale = width / height
        scale = scale * 0.9

        out = tiff.replace('.tif','.pdf')
        
        if rotate_img == 1:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(height, width), bottomup=1)
        if rotate_img == 0:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(width, height), bottomup=1)
        

        for page in range(img.n_frames):
            img.seek(page)
            if rotate_img == 1:
                rotate_image = img.rotate(90, expand=True)
                imgPage = reportlab.lib.utils.ImageReader(rotate_image)
                outPDF.drawImage(imgPage, 0, 0,  height, width)
            if rotate_img == 0:
                imgPage = reportlab.lib.utils.ImageReader(img)
                outPDF.drawImage(imgPage, 0, 0,  width, height)    
           
            if page < img.n_frames:
                outPDF.showPage()

        outPDF.save()
        img.close()
        return 1

    except:
        return 0

def convert_with_auto_rotate(tiff, vert = 1): #по умолчанию vert = 1 (вертикальная ориентация картинки)
    try: 
        print(tiff)
        if vert != 0 and vert != 1:
            vert = 0
        
        img = PIL.Image.open(tiff)
        width, height = img.size
        
        print('размеры файла для конвертации', width, height)
        scale = width / height
        scale = scale * 0.9

        if (vert == 1 and scale > 1) or (vert == 0 and scale < 1):
            rotate_img = 1
        else:
            rotate_img = 0             

        out = tiff.replace('.tif','.pdf')
        
        if rotate_img == 1:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(height, width), bottomup=1)
        if rotate_img == 0:
            outPDF = canvas.Canvas(out, pageCompression=1, pagesize=(width, height), bottomup=1)
        

        for page in range(img.n_frames):
            img.seek(page)
            if rotate_img == 1:
                rotate_image = img.rotate(90, expand=True)
                imgPage = reportlab.lib.utils.ImageReader(rotate_image)
                outPDF.drawImage(imgPage, 0, 0,  height, width)
            if rotate_img == 0:
                imgPage = reportlab.lib.utils.ImageReader(img)
                outPDF.drawImage(imgPage, 0, 0,  width, height)    
           
            if page < img.n_frames:
                outPDF.showPage()

        outPDF.save()
        img.close()
        return 1

    except Exception as ex:
        print(ex)
        return 0

if __name__ == "__main__":
    #res = convert(sys.argv[1], 0)
    res = convert(sys.argv[1], 1)
    if res == 0:
        print('ошибка конвертирования')

