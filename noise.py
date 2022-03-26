from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter as _gl
import os, math, random as ran
from colorama import init

class Zone:
    def __init__(self,cor, elv, grpd):
        self.cor = cor
        self.elv = elv
        self.grpd = grpd

def Roll(x):
    run = True
    while run:
        if x > 3 and x < 8:
            result = ran.randint(x-1, x+1)
        if x > 0 and x < 4:
            result = ran.randint(x-1, x+3)
        if x > 7 and x < 11:
            result = ran.randint(x-2, x+1)
        if result < 11 and result > 0:
            return result
            
def Color(v):
    if not v:
       return ' ' 
    else:
        par, inp = v.split('_')
        return '\033[' + par + 'm' + inp

def Draw(sht, y_cor, x_cor):
    data = ('44;30;2_1','44;30;2_2','44;30;2_3','44;30;2_4','33_5','32_6','32_7','30;1_8','30;1_9','40;1_0')

    for y in range(1, y_cor):
        print('\n', end='')
        for x in range(1, x_cor):
            val = sht[_gl(x)+str(y)].value
            text = Color(data[val-1])
            print(text, end='')

def Context(sht, x, y, mode=0):
    parameter = [(-1,-1),(-1,0),(-1,1),(0,-1),(1,-1),(1,0),(1,1),(0,1)]
    data = []

    for p in parameter:
        try:
            val = sht[_gl(x+p[0]) + str(y+p[1])].value
            if val:
                data.append(val)
        except:
            continue

    result = int(sum(data)/len(data))
    l = 0
    for n in range(len(data)):
        if data[n] > 4:
            l += 1

    if mode == 0:
        return Roll(result)

    if mode == 1:
    ### Choose Water, if too much Land around, turn to Land
        if l >= 5:
            return ran.randint(5, 7) 
        else:
            pass

    if mode == 3:
    ### Connect Lands touching only in diagonal direction
        for n in range(0,8,2):
            try:
                prm = parameter[n]
                point = sht[_gl(x+prm[0]) + str(y+prm[1])].value
                if point > 4:
                    sideA = sht[_gl(x) + str(y+prm[1])].value
                    sideB = sht[_gl(x+prm[0]) + str(y)].value
                    if sideA < 5 and sideB < 5:
                        choice = ran.randint(0,1)
                        if choice == 0:
                            sht[_gl(x) + str(y+prm[1])] = ran.randint(5, 7)
                        if choice == 1:
                            sht[_gl(x+prm[0]) + str(y)] = ran.randint(5, 7)
            except:
                continue
