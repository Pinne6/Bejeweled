from ctypes import windll
dc = windll.user32.GetDC(0)

# ==========================
# GLOBAL VARIABLES
# ==========================
# image anchor to search
image_anchor = 'C:\\Users\\fesposti\\Dropbox\\Word_game\\img\\autohotkey\\new\\submit.jpg'
# image to read if simulation is True
image_simulation = 'C:\\Users\\fesposti\\Downloads\\word_game_remoto_xx.JPG'
# path where to save the screenshot image
image_save_path = 'C:\\Users\\fesposti\\Downloads\\farm_game_remoto_xx.jpg'
# simulation mode True or False
simulation = False
# number of moves
max_moves = 9
# sleeping time
SLEEPING_TIME = 0.1
# ===========================

"""
def getpixel(x, y):
    pixel2 = windll.gdi32.GetPixel(dc, x, y)
    output = [
        pixel2 & 0xff,
        (pixel2 >> 8) & 0xff,
        (pixel2 >> 16) & 0xff,
        (pixel2 >> 24) & 0xff
    ]
    return output

lista_pixels = [
    (340, 276),  # prima riga
    (340 + 55, 276),
    (340 + 55 * 2, 276),
    (340 + 55 * 3, 276),
    (340 + 55 * 4, 276),
    (340 + 55 * 5, 276),
    (340 + 55 * 6, 276),
    (340 + 55 * 7, 276),
    (340, 276 + 54),  # seconda riga
    (340 + 55, 276 + 54),
    (340 + 55 * 2, 276 + 54),
    (340 + 55 * 3, 276 + 54),
    (340 + 55 * 4, 276 + 54),
    (340 + 55 * 5, 276 + 54),
    (340 + 55 * 6, 276 + 54),
    (340 + 55 * 7, 276 + 54),
    (340, 276 + 54 + 55),  # terza riga
    (340 + 55, 276 + 54 + 55),
    (340 + 55 * 2, 276 + 54 + 55),
    (340 + 55 * 3, 276 + 54 + 55),
    (340 + 55 * 4, 276 + 54 + 55),
    (340 + 55 * 5, 276 + 54 + 55),
    (340 + 55 * 6, 276 + 54 + 55),
    (340 + 55 * 7, 276 + 54 + 55),
    (340, 276 + 54 * 2 + 55),  # quarta riga
    (340 + 55, 276 + 54 * 2 + 55),
    (340 + 55 * 2, 276 + 54 * 2 + 55),
    (340 + 55 * 3, 276 + 54 * 2 + 55),
    (340 + 55 * 4, 276 + 54 * 2 + 55),
    (340 + 55 * 5, 276 + 54 * 2 + 55),
    (340 + 55 * 6, 276 + 54 * 2 + 55),
    (340 + 55 * 7, 276 + 54 * 2 + 55),
    (340, 276 + 54 * 2 + 55 * 2),  # quinta riga
    (340 + 55, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 2, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 3, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 4, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 5, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 6, 276 + 54 * 2 + 55 * 2),
    (340 + 55 * 7, 276 + 54 * 2 + 55 * 2),
    (340, 276 + 54 * 3 + 55 * 2),  # sesta riga
    (340 + 55, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 2, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 3, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 4, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 5, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 6, 276 + 54 * 3 + 55 * 2),
    (340 + 55 * 7, 276 + 54 * 3 + 55 * 2),
    (340, 276 + 54 * 3 + 55 * 3),  # settima riga
    (340 + 55, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 2, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 3, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 4, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 5, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 6, 276 + 54 * 3 + 55 * 3),
    (340 + 55 * 7, 276 + 54 * 3 + 55 * 3),
    (340, 276 + 54 * 4 + 55 * 3),  # ottava riga
    (340 + 55, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 2, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 3, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 4, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 5, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 6, 276 + 54 * 4 + 55 * 3),
    (340 + 55 * 7, 276 + 54 * 4 + 55 * 3),
]

for index, item in enumerate(lista_pixels):
    # pixel = pix[item[0], item[1]]
    pixel = getpixel(item[0], item[1])
    colore = ''
    print(str(index + 1) + ": " + str(pixel) + " " + colore)
"""


def left_click():
    ap.mouse.click(ap.mouse.LEFT_BUTTON)
    # ap.mouse.toggle(True, ap.mouse.LEFT_BUTTON)
    time.sleep(SLEEPING_TIME)


def move_mouse(x, y):
    x, y = board_to_pixel(x, y)
    ap.mouse.move(int(x), int(y + 10))
    time.sleep(SLEEPING_TIME)


def click(x, y):
    move_mouse(x, y)
    left_click()


class Grid:
    anchor = []
    grid = []
    items_to_find = []

    def grid_to_pixel(self, x, y):
        x = x * 55 + self.anchor[0]
        if y % 2 == 0:  # even row
            y = (y / 2) * 54 + (y / 2) * 55 + self.anchor[1]
        else:  # odd row
            y = (y // 2 + 1) * 54 + (y // 2) * 55 + self.anchor[1]
        return x, y

    def get_pixel_color(self):
        for y in range(0, 8):
            self.grid.append([])
            for x in range(0, 8):
                x0, y0 = self.grid_to_pixel(x, y)
                item = GridItem()
                item.get_item(x0, y0)
                self.grid[y].append(item)

    def find_anchor(self, anchor, save_path, simu):
        if simulation:
            img_rgb = cv2.imread(simu)
        else:
            snapshot = ImageGrab.grab()
            snapshot.save(save_path)
            img_rgb = cv2.imread(save_path)
        template = cv2.imread(anchor)
        res = cv2.matchTemplate(img_rgb, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        threshold = .8
        loc = np.where(res >= threshold)
        if not np.size(loc[0], 0):
            time.sleep(1)
            return False
        else:
            self.anchor = [(int[loc[1]]), (int[loc[0]])]

    def get_items_to_find(self):
        """
        Check the pixels of the 3 items at the bottom to find in the grid
        :return:
        """
        self.items_to_find = ['sole', 'farina', 'innaffiatoio']


class GridItem:
    item_type = 'X'

    @staticmethod
    def get_color(x, y):
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

    def get_item(self, x, y):
        pix = self.get_color(x, y)
        names = ['acqua', 'sole', 'farina', 'vanga', 'secchiello', 'innaffiatoio']
        colors = [(20, 250, 250), (250, 250, 20), (250, 240, 10), (150, 215, 210), (230, 5, 200), (44, 220, 5)]
        c = 0
        for color in colors:
            diff = 0
            for i in range(3):
                diff += abs(color[i] - pix[i])
            if diff < 90:
                self.item_type = names[c]
            c += 1
        self.item_type = 'X'


class Matcher:
    find_item_count = [0, 0, 0]

    def __init__(self, find_item_c):
        self.find_item_count = find_item_c

    def match_item(self, grid):
        for y in range(0, 8):
            for x in range(0, 8):
                # x
                # -
                # x
                # x
                if x + 0 < 8 and y + 3 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 0][y + 2] and grid[x + 0][y + 2] == grid[x + 0][y + 3] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x, y)
                        click(x, y + 1)
                # - x
                # x
                # x
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 1][y + 0] == grid.grid[x + 0][y + 1] and grid[x + 0][y + 1] == grid[x + 0][y + 2] \
                            and grid.grid[x + 1][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 1, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 1, y + 0)
                        click(x + 0, y + 0)
                # x
                # - x
                # x
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 0][y + 2] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 1, y + 1)
                        click(x + 0, y + 1)
                # x
                # x
                # - x
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 0][y + 1] and grid[x + 0][y + 1] == grid[x + 1][y + 2] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 1, y + 2)
                        click(x + 0, y + 2)
                # x
                # x
                # -
                # x
                if x + 0 < 8 and y + 3 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 0][y + 1] and grid[x + 0][y + 1] == grid[x + 0][y + 3] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 3)
                        click(x + 0, y + 2)
                # x -
                #   x
                #   x
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 1][y + 2] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 0)
                        click(x + 1, y + 0)
                #   x
                # x -
                #   x
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 1][y + 0] == grid.grid[x + 0][y + 1] and grid[x + 0][y + 1] == grid[x + 1][y + 2] \
                            and grid.grid[x + 1][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 1, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 1)
                        click(x + 1, y + 1)
                #   x
                #   x
                # x -
                if x + 1 < 8 and y + 2 < 8:
                    if grid.grid[x + 1][y + 0] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 0][y + 2] \
                            and grid.grid[x + 1][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 1, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 2)
                        click(x + 1, y + 2)
                # xx-x
                if x + 3 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 0] and grid[x + 1][y + 0] == grid[x + 3][y + 0] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 3, y + 0)
                        click(x + 2, y + 0)
                # x--
                # -xx
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 2][y + 1] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 0)
                        click(x + 0, y + 1)
                # -x-
                # x-x
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 1] == grid.grid[x + 1][y + 0] and grid[x + 1][y + 0] == grid[x + 2][y + 1] \
                            and grid.grid[x + 0][y + 1] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 1] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 1, y + 0)
                        click(x + 1, y + 1)
                # --x
                # xx-
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 1] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 2][y + 0] \
                            and grid.grid[x + 0][y + 1] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 1] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 2, y + 0)
                        click(x + 2, y + 1)
                # x-xx
                if x + 3 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 2][y + 0] and grid[x + 2][y + 0] == grid[x + 3][y + 0] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 0)
                        click(x + 1, y + 0)
                # -xx
                # x--
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 1] == grid.grid[x + 1][y + 0] and grid[x + 1][y + 0] == grid[x + 2][y + 0] \
                            and grid.grid[x + 0][y + 1] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 1] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 0, y + 1)
                        click(x + 0, y + 0)
                # x-x
                # -x-
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 1] and grid[x + 1][y + 1] == grid[x + 2][y + 0] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 1, y + 1)
                        click(x + 1, y + 0)
                # xx-
                # --x
                if x + 2 < 8 and y + 1 < 8:
                    if grid.grid[x + 0][y + 0] == grid.grid[x + 1][y + 0] and grid[x + 1][y + 0] == grid[x + 2][y + 1] \
                            and grid.grid[x + 0][y + 0] in grid.items_to_find:
                        for index, item in enumerate(grid.items_to_find):
                            if item == grid.grid[x + 0, y + 0] and (self.find_item_count > 0
                                                                    or sum(self.find_item_count) == 0):
                                self.find_item_count[index] -= 3
                        click(x + 2, y + 1)
                        click(x + 2, y + 0)


def main():
    # make Remote window active using Macro Scheduler
    keys = comclt.Dispatch('WScript.Shell')
    keys.SendKeys('^+1')
    time.sleep(1)

    # create the grid
    grid = Grid()
    # find the anchor image
    grid.find_anchor(image_anchor, image_save_path, image_simulation)
    # populate the grid with GridItem items
    grid.get_pixel_color()

    # find the items and the count
    find_item_c = [10, 12, 15]

    # find matches
    match = Matcher(find_item_c)
    match.match_item(grid)
