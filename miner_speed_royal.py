"""
Plays a game of Bejeweled Blitz on Facebook.
Created: 24.12.2012
Updated: Aug 2016
@author: Adrianus Kleemans
"""

import autopy3 as ap
# import Quartz.CoreGraphics as CG
import struct
import pygame
import time
import sys
import win32com.client as comclt
from ctypes import windll
import cv2
from PIL import ImageGrab
import numpy as np
import random
import os

# constants
dc = windll.user32.GetDC(0)
SLEEPING_TIME = 0.002
MAX_MOVES = 7
DRAW_CANVAS = True
MAX_TIME = 55  # for some gems, use up to 160s
simulazione = True
test = False
os.environ['SDL_VIDEO_WINDOW_POS'] = str(1000) + "," + str(300)


# input helpers

def left_down():
    ap.mouse.click(ap.mouse.LEFT_BUTTON)
    # ap.mouse.toggle(True, ap.mouse.LEFT_BUTTON)
    time.sleep(SLEEPING_TIME)


def left_up():
    ap.mouse.toggle(False, ap.mouse.LEFT_BUTTON)
    time.sleep(SLEEPING_TIME)


def move_mouse(x, y):
    x, y = board_to_pixel(x, y)
    ap.mouse.move(int(x), int(y + 10))
    time.sleep(SLEEPING_TIME)


def click(x, y):
    move_mouse(x, y)
    left_down()
    left_up()


def move_fields(x0, y0, x1, y1):
    global moves
    move_mouse(x0, y0)
    left_down()
    move_mouse(x1, y1)
    left_down()
    moves += 1


# screen helpers

def getpixel(x, y):
    dc = windll.user32.GetDC(0)
    pixel = windll.gdi32.GetPixel(dc, x, y)
    # print(pixel)
    output = [
        pixel & 0xff,
        (pixel >> 8) & 0xff,
        (pixel >> 16) & 0xff,
        (pixel >> 24) & 0xff
    ]
    return output


def get_pixel(x, y):
    x, y = int(x) * 2, int(y) * 2
    data_format = "BBBB"  # BBBB
    offset = 4 * ((screenshot_width * int(round(y))) + int(round(x)))
    b, g, r, a = struct.unpack_from(data_format, screenshot_data, offset=offset)
    return r, g, b


def get_field(x, y):
    x, y = board_to_pixel(x, y)
    return get_pixel(x, y)


def get_board_init():
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
    if not np.size(loc[0], 0):
        time.sleep(1)
        return False
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
    anchor = [int(loc[1]), int(loc[0])]
    return lista_pixels, anchor


def get_board(lista_pixel):
    board = []
    for y in range(0, 8):
        board.append([])
        for x in range(0, 8):
            # print(lista_pixel[y * 8 + x][0], lista_pixel[y * 8 + x][1])
            pix = getpixel(lista_pixel[y * 8 + x][0], lista_pixel[y * 8 + x][1])
            color = match_color(pix)
            colori = ['giallo', 'verde', 'rosso', 'blu', 'viola']
            if test:
                board[y].append(colori[random.randint(0, 4)])
            else:
                board[y].append(color)
    return board


def match_color(pix):
    names = ['giallo', 'blu', 'verde', 'rosso', 'viola', '.', '.']
    colors = [(170, 60, 25), (230, 235, 235), (165, 235, 160),
              (215, 115, 120), (230, 150, 200), (20, 25, 90), (30, 35, 40)
              ]
    c = 0
    for color in colors:
        diff = 0
        for i in range(3):
            diff += abs(color[i] - pix[i])
        if diff < 90:
            return names[c]
        c += 1
    return 'X'


def board_to_pixel(x, y):
    x = x * 43 + 338 + anchor[0]
    y = y * 43 + 101 + anchor[1]
    return x, y


def same_gem(x0, y0, x1, y1):
    m = 7  # max row
    if x0 < 0 or y0 < 0 or x1 < 0 or y1 < 0 or x0 > m or y0 > m or x1 > m or y1 > m:
        return False
    elif board[y0][x0] == '.' or board[y1][x1] == '.':
        return False
    elif board[y0][x0] == 'X' or board[y1][x1] == 'X':
        return True
    else:
        # if board[y0][x0] == board[y1][x1]:
        #     print(board[y0][x0], board[y1][x1])
        return board[y0][x0] == board[y1][x1]


def board_valid(board):
    whites = 0
    for line in board:
        for field in line:
            if field == 'w':
                whites += 1
    if whites >= 20:
        return False
    else:
        return True


def find_x(board):
    global moves
    moves = 0
    x_positions = []
    for index, item in enumerate(board):
        count_x = [i for i, j in enumerate(item) if j == 'X']
        for item2 in count_x:
            x_positions.append((item2, index))
    if len(x_positions) == 0:
        return False
    elif len(x_positions) == 1:
        # I have only one X, I combine it randomly
        if x_positions[0][0] < 7:
            x = x_positions[0][0] + 1
        else:
            x = x_positions[0][0] - 1
        move_fields(x_positions[0][0], x_positions[0][1], x, x_positions[0][1])
        return True
    elif len(x_positions) > 1:
        # I search for two X next to each others
        for index, item in enumerate(x_positions):
            if index == 0:
                continue
            if (abs(item[0] - x_positions[index - 1][0]) < 1 and abs(item[1] - x_positions[index - 1][1]) <= 1) or (
                    abs(item[0] - x_positions[index - 1][0]) <= 1 and abs(item[1] - x_positions[index - 1][1]) < 1):
                move_fields(item[0], item[1], x_positions[index - 1][0], x_positions[index - 1][1])
                return True



# canvas helpers

def draw_board(board):
    board_size = [320, 320]
    colors = {'giallo': (254, 254, 38), 'w': (254, 254, 254), 'rosso': (254, 29, 59),
              'o': (254, 128, 0), 'viola': (250, 10, 250), 'verde': (50, 254, 50),
              'blu': (20, 112, 232), 'X': (114, 186, 112), '.': (0, 0, 0)}
    signs = []
    for line in board:
        for sign in line:
            signs.append(sign)

    for y in range(8):
        for x in range(8):
            col = colors[board[y][x]]
            # pygame.draw.rect(disp, col, (x * 40, y * 40, 40, 40))
    print('DISEGNO')
    print(board)
    # pygame.display.flip()


# main
def main():
    global anchor
    global disp
    global rects
    global moves
    global board

    print('Starting, please bring window into position')

    # preparing canvas
    if DRAW_CANVAS:
        pygame.init()
        disp = pygame.display.set_mode([320, 320])
        pygame.display.set_caption("Bejeweled Blitz Demo")

    # play
    print("Starting game!")
    # bring remote screen to the front
    keys = comclt.Dispatch('WScript.Shell')
    keys.SendKeys('^+1')
    time.sleep(1)
    lista_coordinate, anchor = get_board_init()
    t = time.time()

    while (time.time() - t) < MAX_TIME:
        if not getpixel(127, 174) != (2, 101, 2):
            print('Board not valid. Halt')
            exit()
        board = get_board(lista_coordinate)
        if DRAW_CANVAS:
            draw_board(board)

        # calculate possible moves
        while find_x(board):
            board = get_board(lista_coordinate)

        moves = 0
        for y in range(0, 8, 1):
            if moves >= MAX_MOVES:
                break
            for x in range(0, 8, 1):
                if moves >= MAX_MOVES:
                    break
                if same_gem(x, y, x - 1, y):  # two gems next to each other, horizontal
                    if same_gem(x, y, x + 1, y - 1): move_fields(x + 1, y, x + 1, y - 1)   # right
                    if same_gem(x, y, x + 2, y): move_fields(x + 1, y, x + 2, y)
                    if same_gem(x, y, x + 1, y + 1): move_fields(x + 1, y, x + 1, y + 1)
                    if same_gem(x, y, x - 2, y - 1): move_fields(x - 2, y, x - 2, y - 1)  # left
                    if same_gem(x, y, x - 2, y + 1): move_fields(x - 2, y, x - 2, y + 1)
                    if same_gem(x, y, x - 3, y): move_fields(x - 2, y, x - 3, y)
                if same_gem(x, y, x, y - 1):  # two gems next to each other, vertical
                    if same_gem(x, y, x + 1, y + 1): move_fields(x, y + 1, x + 1, y + 1)  # below
                    if same_gem(x, y, x, y + 2): move_fields(x, y + 1, x, y + 2)
                    if same_gem(x, y, x - 1, y + 1): move_fields(x, y + 1, x - 1, y + 1)
                    if same_gem(x, y, x - 1, y - 2): move_fields(x, y - 2, x - 1, y - 2)  # above
                    if same_gem(x, y, x + 1, y - 2): move_fields(x, y - 2, x + 1, y - 2)
                    if same_gem(x, y, x, y - 3): move_fields(x, y - 2, x, y - 3)
                if same_gem(x, y, x - 2, y):  # gem in the middle is missing, horizontal
                    if same_gem(x, y, x - 1, y - 1): move_fields(x - 1, y, x - 1, y - 1)
                    if same_gem(x, y, x - 1, y + 1): move_fields(x - 1, y, x - 1, y + 1)
                if same_gem(x, y, x, y - 2):  # gem in the middle is missing, vertical
                    if same_gem(x, y, x - 1, y - 1): move_fields(x, y - 1, x - 1, y - 1)
                    if same_gem(x, y, x + 1, y - 1): move_fields(x, y - 1, x + 1, y - 1)
    print("Game ended.")


if __name__ == "__main__":
    main()
