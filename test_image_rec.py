from PIL import ImageGrab, Image
import cv2
import numpy as np
import autopy3 as ap
import struct
import pygame
import time
import sys
import win32com.client as comclt
from ctypes import windll
dc = windll.user32.GetDC(0)

# constants
SLEEPING_TIME = 0.02
MAX_MOVES = 5
DRAW_CANVAS = False
MAX_TIME = 75  # for some gems, use up to 160s
simulazione = True


def getpixel(x, y):
    pixel2 = windll.gdi32.GetPixel(dc, x, y)
    output = [
        pixel2 & 0xff,
        (pixel2 >> 8) & 0xff,
        (pixel2 >> 16) & 0xff,
        (pixel2 >> 24) & 0xff
    ]
    return output

simulazione = False
"""
lista_pixels = [
    (457, 267),  # prima riga
    (457 + 43, 267),
    (457 + 43 * 2, 267),
    (457 + 43 * 3, 267),
    (457 + 43 * 4, 267),
    (457 + 43 * 5, 267),
    (457 + 43 * 6, 267),
    (457 + 43 * 7, 267),
    (457, 267 + 43),  # seconda riga
    (457 + 43, 267 + 43),
    (457 + 43 * 2, 267 + 43),
    (457 + 43 * 3, 267 + 43),
    (457 + 43 * 4, 267 + 43),
    (457 + 43 * 5, 267 + 43),
    (457 + 43 * 6, 267 + 43),
    (457 + 43 * 7, 267 + 43),
    (457, 267 + 43 * 2),  # terza riga
    (457 + 43, 267 + 43 * 2),
    (457 + 43 * 2, 267 + 43 * 2),
    (457 + 43 * 3, 267 + 43 * 2),
    (457 + 43 * 4, 267 + 43 * 2),
    (457 + 43 * 5, 267 + 43 * 2),
    (457 + 43 * 6, 267 + 43 * 2),
    (457 + 43 * 7, 267 + 43 * 2),
    (457, 267 + 43 * 3),  # quarta riga
    (457 + 43, 267 + 43 * 3),
    (457 + 43 * 2, 267 + 43 * 3),
    (457 + 43 * 3, 267 + 43 * 3),
    (457 + 43 * 4, 267 + 43 * 3),
    (457 + 43 * 5, 267 + 43 * 3),
    (457 + 43 * 6, 267 + 43 * 3),
    (457 + 43 * 7, 267 + 43 * 3),
    (457, 267 + 43 * 4),  # quinta riga
    (457 + 43, 267 + 43 * 4),
    (457 + 43 * 2, 267 + 43 * 4),
    (457 + 43 * 3, 267 + 43 * 4),
    (457 + 43 * 4, 267 + 43 * 4),
    (457 + 43 * 5, 267 + 43 * 4),
    (457 + 43 * 6, 267 + 43 * 4),
    (457 + 43 * 7, 267 + 43 * 4),
    (457, 267 + 43 * 5),  # sesta riga
    (457 + 43, 267 + 43 * 5),
    (457 + 43 * 2, 267 + 43 * 5),
    (457 + 43 * 3, 267 + 43 * 5),
    (457 + 43 * 4, 267 + 43 * 5),
    (457 + 43 * 5, 267 + 43 * 5),
    (457 + 43 * 6, 267 + 43 * 5),
    (457 + 43 * 7, 267 + 43 * 5),
    (457, 267 + 43 * 6),  # settima riga
    (457 + 43, 267 + 43 * 6),
    (457 + 43 * 2, 267 + 43 * 6),
    (457 + 43 * 3, 267 + 43 * 6),
    (457 + 43 * 4, 267 + 43 * 6),
    (457 + 43 * 5, 267 + 43 * 6),
    (457 + 43 * 6, 267 + 43 * 6),
    (457 + 43 * 7, 267 + 43 * 6),
    (457, 267 + 43 * 7),  # ottava riga
    (457 + 43, 267 + 43 * 7),
    (457 + 43 * 2, 267 + 43 * 7),
    (457 + 43 * 3, 267 + 43 * 7),
    (457 + 43 * 4, 267 + 43 * 7),
    (457 + 43 * 5, 267 + 43 * 7),
    (457 + 43 * 6, 267 + 43 * 7),
    (457 + 43 * 7, 267 + 43 * 7),
]

snapshot = ImageGrab.grab()
# img_rgb = Image.open('C:\\Users\\fesposti\\Dropbox\\Miner_Speed\\miner_speed_02.jpg')
# img_rgb = Image.open(snapshot)
pix = snapshot.load()
giallo = []
blu = []
verde = []
rosso = []
viola = []
for index, item in enumerate(lista_pixels):
    # pixel = pix[item[0], item[1]]
    pixel = getpixel(item[0], item[1])
    if 150 <= pixel[0] <= 193 and 40 <= pixel[1] <= 80 and 0 <= pixel[2] <= 30:  # giallo
        giallo.append((index + 1, pixel[0], pixel[1], pixel[2]))
        colore = 'giallo'
    elif 210 <= pixel[0] <= 255 and 220 <= pixel[1] <= 255 and 220 <= pixel[2] <= 255:  # blu
        blu.append((index + 1, pixel[0], pixel[1], pixel[2]))
        colore = 'blu'
    elif 155 <= pixel[0] <= 180 and 220 <= pixel[1] <= 250 and 140 <= pixel[2] <= 180:  # verde
        verde.append((index + 1, pixel[0], pixel[1], pixel[2]))
        colore = 'verde'
    elif 200 <= pixel[0] <= 235 and 100 <= pixel[1] <= 130 and 110 <= pixel[2] <= 135:  # rosso
        rosso.append((index + 1, pixel[0], pixel[1], pixel[2]))
        colore = 'rosso'
    elif 230 <= pixel[0] <= 255 and 140 <= pixel[1] <= 160 and 180 <= pixel[2] <= 220:  # viola
        viola.append((index + 1, pixel[0], pixel[1], pixel[2]))
        colore = 'viola'
    else:
        colore = ''
        print(str(index + 1) + ": " + str(pixel) + " " + colore)
"""


def get_board_init():
    """
    if simulazione:
        img_rgb = cv2.imread('C:\\Users\\fesposti\\Dropbox\\Miner_Speed\\miner_speed_02.jpg')
    else:
        snapshot = ImageGrab.grab()
        save_path = "C:\\Users\\fesposti\\Downloads\\miner_speed_remoto_xx.jpg"
        snapshot.save(save_path)
        img_rgb = cv2.imread(save_path)
    template = cv2.imread('C:\\Users\\fesposti\\Dropbox\\Miner_Speed\\score.jpg')
    res = cv2.matchTemplate(img_rgb, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    threshold = .9
    loc = np.where(res >= threshold)
    # if not np.size(loc[0], 0):
    #     time.sleep(1)
    #     return False
    lista_pixels = [
        (int(loc[1][0] + 338), int(loc[0][0] + 101)),  # prima riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43)),  # seconda riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 2)),  # terza riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 2)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 3)),  # quarta riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 3)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 4)),  # quinta riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 4)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 5)),  # sesta riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 5)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 6)),  # settima riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 6)),
        (int(loc[1][0] + 338), int(loc[0][0] + 101 + 43 * 7)),  # ottava riga
        (int(loc[1][0] + 338 + 43), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 2), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 3), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 4), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 5), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 6), int(loc[0][0] + 101 + 43 * 7)),
        (int(loc[1][0] + 338 + 43 * 7), int(loc[0][0] + 101 + 43 * 7)),
    ]
    """
    lista_pixels = [(457, 277), (500, 277)]
    return lista_pixels


def get_board(lista_pixel):
    board = []
    for y in range(0, 8):
        board.append([])
        for x in range(0, 8):
            print(lista_pixel[y * 8 + x][0], lista_pixel[y * 8 + x][1])
            color = getpixel(lista_pixel[y * 8 + x][0], lista_pixel[y * 8 + x][1])
            board[y].append(color)
    return board


def main():
    # preparing canvas
    if DRAW_CANVAS:
        pygame.init()
        disp = pygame.display.set_mode([320, 320])
        pygame.display.set_caption("Bejeweled Blitz Demo")

    # bring remote to front
    # keys = comclt.Dispatch('WScript.Shell')
    # keys.SendKeys('^+1')
    time.sleep(1)

    # find anchor image
    lista_coordinate = get_board_init()
    t = time.time()

    while (time.time() - t) < MAX_TIME:
        if not getpixel(127, 174) != (2, 101, 2):
            print('Board not valid. Halt')
            exit()
        board = get_board(lista_coordinate)
        if DRAW_CANVAS: draw_board(board)


if __name__ == "__main__":
    main()
