import cv2
import win32gui
import win32ui
import win32api
import win32con
from ctypes import windll
from PIL import Image
import numpy as np
from timeit import default_timer as timer
import time
import os
import subprocess

def set_resolution(title):
    hwnd = win32gui.FindWindow(None, title)
    x0, y0, x1, y1 = win32gui.GetWindowRect(hwnd)
    win32gui.MoveWindow(hwnd, x0, y0, 1280, 750, True)

def get_screenshot(title):
    hwnd = win32gui.FindWindow(None, title)
    x0, y0, x1, y1 = win32gui.GetWindowRect(hwnd)

    left, top, right, bot = win32gui.GetClientRect(hwnd)
    w = right - left
    h = bot - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = None
    try:
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
    except ValueError:
        return None

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    return im

def find_template_center(file): 
    ss = get_screenshot('Bloons TD Battles 2')
    if ss is None:
        return None
    cvImg = np.array(ss.convert('L'))
    template = cv2.imread(file, cv2.IMREAD_GRAYSCALE)
    w, h = template.shape[::-1]

    res = cv2.matchTemplate(cvImg, template, cv2.TM_CCOEFF_NORMED)
    threshold = 0.8
    loc = np.where(res >= threshold)
    if len(loc[0]) == 0 or len(loc[1]) == 0:
        return None

    total_x, total_y = 0, 0
    count = 0
    for pt in zip(*loc[::-1]):
        total_x += pt[0]
        total_y += pt[1]
        count = count + 1        

    avg_x = int(total_x / count)
    avg_y = int(total_y / count)

    return int(avg_x + w / 2), int(avg_y + h / 2)

def click(x, y):
    hWnd = win32gui.FindWindow(None, 'Bloons TD Battles 2')
    lParam = win32api.MAKELONG(x, y)

    win32gui.SendMessage(hWnd, win32con.WM_MOUSEMOVE, 0, lParam)
    win32gui.SendMessage(hWnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
    win32gui.SendMessage(hWnd, win32con.WM_LBUTTONUP, 0, lParam)

def drag_and_drop(x, y, x2, y2):
    hWnd = win32gui.FindWindow(None, 'Bloons TD Battles 2')
    lParam = win32api.MAKELONG(x, y)
    lParam2 = win32api.MAKELONG(x2, y2)

    win32gui.SendMessage(hWnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
    win32gui.SendMessage(hWnd, win32con.WM_MOUSEMOVE, 0, lParam2)
    # do a few mouse moves offset to ensure game recognizes unit is in a valid spot
    lParam2 = win32api.MAKELONG(x2 + 1, y2 + 1)
    time.sleep(0.1)
    win32gui.SendMessage(hWnd, win32con.WM_MOUSEMOVE, 0, lParam2)
    lParam2 = win32api.MAKELONG(x2 + 2, y2 + 2)
    time.sleep(0.1)
    win32gui.SendMessage(hWnd, win32con.WM_MOUSEMOVE, 0, lParam2)
    time.sleep(0.1)
    win32gui.SendMessage(hWnd, win32con.WM_LBUTTONUP, 0, lParam2)

def find_and_click(file, attempts=1):
    attempt = 0
    while attempt < attempts:
        pt = find_template_center(file)
        if pt is not None:
            click(pt[0], pt[1])
            return True
        attempt += 1
        time.sleep(1)

def wait_for(file, timeout=0):
    start = timer()

    while True:
        pos = find_template_center(file)
        if pos is not None:
            return pos
        time.sleep(0.5)
        if timeout > 0 and timer() - start > timeout:
            return None

def send_key(title, key):
    hwnd = win32gui.FindWindow(None, title)
    win32gui.SendMessage(hwnd, win32con.WM_CHAR, key, 0)

def restart_game():
    os.system('taskkill /f /im btdb2_game.exe')
    time.sleep(2)
    btdb2_path = 'C:\\Program Files (x86)\\Steam\\steamapps\\common\\Bloons TD Battles 2\\btdb2_game.exe'
    subprocess.call([btdb2_path])
    time.sleep(10)

color_to_placement = {
    (70, 53, 44): (219, 355),
    (4, 206, 197): (219, 355),
    (245, 232, 198): (222, 533),
    (89, 166, 99): (300, 533),
    (203, 174, 111): (300, 533),
    (109, 120, 97): (300, 533),
    (73, 184, 146): (300, 485),
    (80, 175, 83): (300, 493),
    (230, 170, 123): (300, 672),
    (98, 105, 100): (300, 672),
    (87, 71, 64): (300, 252)
}

set_resolution('Bloons TD Battles 2')

games = 0
start_time = timer()
state = 'idle'
while True:
    queue_start = timer()
    while state != 'queued':
        battle_pos = find_template_center('templates/battle.png')
        if battle_pos is not None:
            click(battle_pos[0], battle_pos[1])
            state = 'queued'

        time.sleep(2)

        # discard screen will show after you click battle, discard chest and restart loop    
        discard_pos = find_template_center('templates/discard.png')
        if discard_pos is not None:
            click(discard_pos[0], discard_pos[1])
            state = 'discarding'
            time.sleep(.5)

        if state == 'discarding':
            click(104, 300)
            state = 'idle'

        time.sleep(2)
    
    need_restart = False
    found_battle = False
    while not found_battle:
        if state != 'unit_select':
            ready_pos = find_template_center('templates/ready.png')
            if ready_pos is not None:
                click(ready_pos[0], ready_pos[1])
                queue_start = timer()
                state = 'map_select'
                continue

        if state == 'map_select':
            battle2_pos = find_template_center('templates/battle2.png')
            if battle2_pos is not None:
                click(battle2_pos[0], battle2_pos[1])
                state = 'unit_select'

        # catch disconnects and restart loop
        dc_pos = find_template_center('templates/opp_disconnect.png')
        if dc_pos is not None:
            print('Opponent disconnected, attempting to reset')
            ok_pos = find_template_center('templates/ok2.png')
            if ok_pos is not None or True:
                print('Found ok button')
                click(328, 369)
                need_restart = True
                state = 'idle'
                time.sleep(3)
                break

        dc_pos = find_template_center('templates/disconnect.png')
        if dc_pos is not None:
            print('We disconnected, attempting to reset')
            ok_pos = find_template_center('templates/ok2.png')
            if ok_pos is not None or True:
                print('Found ok button')
                click(328, 369)
                need_restart = True
                state = 'idle'
                time.sleep(3)
                break

        unable_connect_pos = find_template_center('templates/unable_connect.png')
        if unable_connect_pos is not None:
            print('Unable to connect, attempting to reset')
            quit_pos = find_template_center('templates/quit.png')
            if quit_pos is not None:
                click(quit_pos[0], quit_pos[1])
                need_restart = True
                state = 'idle'
                time.sleep(3)
                break

        if timer() - queue_start > 120:
            print("In some unknown state, restarting bloons")
            restart_game()
            time.sleep(10)
            wait_for('templates/battle.png')
            need_restart = True
            state = 'idle'
            break
            
        # once we are in game, we no longer need to worry about disconnection errors
        lock_pos = find_template_center('templates/lock.png')
        if lock_pos is not None:
            found_battle = True
            break

        time.sleep(1)
    
    if need_restart:
        continue

    lock_pos = find_template_center('templates/lock.png')
    if lock_pos is None:
        # dced?
        ok_pos = wait_for('templates/ok.png')
        click(ok_pos[0], ok_pos[1])
        time.sleep(2)
        continue
    side = 'none'
    if lock_pos[0] < 600:
        side = 'left'
    else:
        side = 'right'

    hero_location = (0, 0)
    if side == 'left':
        hero_location = (67, 147)
    elif side == 'right':
        hero_location = (1200, 147)

    time.sleep(1)
    ss = get_screenshot('Bloons TD Battles 2')
    if ss is None:
        continue
    time.sleep(1.8)

    import math
    # compare color at 355, 228 to placement colors and find the one with least difference
    min_diff = 100
    min_color = (0, 0, 0)
    for color in color_to_placement:
        diff = math.sqrt(math.pow(color[0] - ss.getpixel((355, 228))[0], 2) + math.pow(color[1] - ss.getpixel((355, 228))[1], 2) + math.pow(color[2] - ss.getpixel((355, 228))[2], 2))
        if diff < min_diff:
            min_diff = diff
            min_color = color

    placement_loc = None
    try:
        placement_loc = color_to_placement[min_color]
        # mirror placement location if on right side
        if side == 'right':
            placement_loc = (1280 - placement_loc[0] - 9, placement_loc[1])

        drag_and_drop(hero_location[0], hero_location[1], placement_loc[0], placement_loc[1])
    except:
        pass

    surrender_pos = find_template_center('templates/surrender.png')
    if surrender_pos is not None:
        click(surrender_pos[0], surrender_pos[1])
        time.sleep(3)
        checkmark_pos = wait_for('templates/checkmark.png', 2)
        if checkmark_pos is not None:
            click(checkmark_pos[0], checkmark_pos[1])

    ok_pos = wait_for('templates/ok.png')
    click(ok_pos[0], ok_pos[1])
    games += 1
    if games % 5 == 0:
        print('games: ' + str(games) + ' seconds per game: ' + str((timer() - start_time) / games))

    # load times seem to increase as game is left open
    # close and restart btdb2_game.exe every 50 games
    if games % 50 == 0:
        print('Restarting BTDB2')
        restart_game()
        wait_for('templates/battle.png')
        continue

    time.sleep(4)
