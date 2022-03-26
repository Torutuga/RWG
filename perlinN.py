from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as _gl
import matplotlib.pyplot as plt
from perlin_noise import PerlinNoise
import random as ran

def noiseGiver(sht, y, x):
    seed1 = ran.randint(250, 500)
    noise = PerlinNoise(octaves=6, seed=seed1)
    xpix, ypix = (x+1)*10, (y+1)*10
    pic = [[noise([i/xpix, j/ypix]) for j in range(xpix)] for i in range(ypix)]
    
    for cy in range(1, ypix, 10):
        for cx in range(1, xpix, 10):
            try:
                value = pic[cy+3][cx+3]
            except IndexError:
                pass
            if cx == 1:
                rx = 1
            if cy == 1:
                ry = 1
            if cx > 1:
                rx = int((cx-1)/10)
            if cy > 1:
                ry = int((cy-1)/10)
            sht[_gl(rx)+str(ry)] = int((value+1)*5)

    plt.imshow(pic, cmap='copper')
    plt.show()
