import os
from PIL import Image, ImageFilter
from scipy import ndimage, misc
import numpy as np
import pytesseract

def solve_captcha_jurisconsult(image):
    '''Resolve captcha do site jurisconsult'''
    image = image.convert('L')    #CONVERTE EM ESCALA DE CINZA
    image = image.filter(ImageFilter.SHARPEN)
    #image = image.resize((640, 200), Image.NEAREST)
    image = image.point(lambda x: 0 if x<160 else 255, '1')   #ESCURECE AS CORES CINZAS ATÉ CHEGAR NO PRETO
    solution = pytesseract.image_to_string(image, 
                                config='-psm 10000 -c tessedit_char_whitelist=0123456789').replace(" ", "")
    return image, solution

def solve_captcha_pje(image, chop=1):
    '''Resolve captcha do site pje do tjma'''
    image = image.convert('L')    #CONVERTE EM ESCALA DE CINZA
    image = image.filter(ImageFilter.SHARPEN)
    image = image.point(lambda x: 0 if x<180 else 255, '1')     #ESCURECE AS CORES CINZAS ATÉ CHEGAR NO PRETO
    #image.save("captcha_pje/darker_" + filename)
    #image = image.resize((1200, 400), Image.NEAREST)
    width, height = image.size
    data = image.load()

    # Iterate through the rows.
    for y in range(height):
        for x in range(width):

            # Make sure we're on a dark (128) pixel.
            if data[x, y] > 128:
                continue

            # Keep a total of non-white contiguous pixels.
            total = 0

            # Check a sequence ranging from x to image.width.
            for c in range(x, width):

                # If the pixel is dark(128), add it to the total.
                if data[c, y] < 128:
                    total += 1

                # If the pixel is light, stop the sequence.
                else:
                    break

            # If the total is less than the chop, replace everything with white.
            if total <= chop:
                for c in range(total):
                    data[x + c, y] = 255

            # Skip this sequence we just altered.
            x += total


    # Iterate through the columns.
    for x in range(width):
        for y in range(height):

            # Make sure we're on a dark (128) pixel.
            if data[x, y] > 128:
                continue

            # Keep a total of non-white contiguous pixels.
            total = 0

            # Check a sequence ranging from y to image.height.
            for c in range(y, height):

                # If the pixel is dark(128), add it to the total.
                if data[x, c] < 128:
                    total += 1

                # If the pixel is light, stop the sequence.
                else:
                    break

            # If the total is less than the chop, replace everything with white.
            if total <= chop:
                for c in range(total):
                    data[x, y + c] = 255

            # Skip this sequence we just altered.
            y += total

    # image.save("captcha_pje/resultado_" + filename)
    # image2 = ndimage.imread("captcha_pje/resultado_" + filename)
    # image2 = ndimage.binary_erosion(image2, iterations=2).astype(np.int)
    # image2 = ndimage.binary_dilation(image2, iterations=2).astype(np.int)
    # misc.imsave("captcha_pje/resultado2_" + filename, image2)
    # image = Image.open("captcha_pje/resultado2_" + filename)
    # #image = image.resize((1200, 400), Image.NEAREST)
    solution = pytesseract.image_to_string(image, config='-psm 10000 -c tessedit_char_whitelist=123456789').replace(" ", "")
    return image, solution

if __name__ == '__main__':
    for filename in os.listdir(os.getcwd() + "/captcha_pje"):
        if filename.endswith(".jpg"):
            image = Image.open("captcha_pje/" + filename)
            image, solution = solve_captcha_pje(image)
            image.save("captcha_pje/resultado_" + filename)
            print(filename.replace(".jpg", ""), solution, filename.replace(".jpg", "")==solution)