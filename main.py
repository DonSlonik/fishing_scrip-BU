import time
import keyboard
import sys
import mss
import mss.tools
import os.path
import winsound
# установочный файл для этой библиотеки Pillow
from PIL import Image
import numpy as np
import mouse
import random
from datetime import date
# установочный файл для этой библиотеки pywin32
from win32gui import GetWindowText, GetForegroundWindow
import win32gui
import win32ui
from ctypes import windll
import ctypes
import win32api
import win32con


# флаг выхода из программы
global flag_stop
flag_stop: bool = False

# флаг создания скринов для отладки
global flag_screen
flag_screen: bool = False


# распечатать информацию про массив
def print_array_info(sta):
    print("количество осей в массиве - " + str(sta.ndim))
    print("размерность каждой оси в массиве - " + str(sta.shape))
    print("тип данных в массиве - " + str(sta.dtype))
    print("=======================")
# =====================================================================================================================


# создать одномерный масив int32 со всеми вариантами цветов из определенного диапазона RGB
def create_lime_array(red, green, blue, red_, green_, blue_):
    color_red = np.arange(red - red_, red + 1)
    color_green = np.arange(green - green_, green + 1)
    color_blue = np.arange(blue - blue_, blue + 1)
    # создать масив нулей для всех комбинаций цвета лайм
    lime_all = np.zeros((len(color_red) * len(color_green) * len(color_blue)), dtype=np.int32)
    count_pix = 0
    for r in color_red:
        for g in color_green:
            for b in color_blue:
                #lime_all[count_pix] = (r * 1) + (g * 256) + (b * 65536)
                lime_all[count_pix] = (r * 65536) + (g * 256) + (b * 1)
                count_pix = count_pix + 1

    # найти ширину и высоту для картинки чтобы поместились все варианты цвета лайм
    w, h = find_mlt_xy((red_ + 1) * (green_ + 1) * (blue_ + 1))

    # создаём матрицу будущей картинки
    I = np.zeros((h, w, 3), dtype=np.uint8)

    # заполняем матрицу масивом всех салатовых цветов
    count_pix = 0
    for x in range(h):
        for y in range(w):
            int_32 = hex(lime_all[count_pix])
            I[x, y, 0] = int("0x" + str(int_32[2:4]), 0)
            I[x, y, 1] = int("0x" + str(int_32[4:6]), 0)
            I[x, y, 2] = int("0x" + str(int_32[6:8]), 0)
            count_pix = count_pix + 1
    # недокументированная функция загнать масив назад в картинку PIL
    new_im = Image.fromarray(I)
    # сохраняем картинку всех салатовых цветов в файл
    new_im.save(full_path + "all_lime_pixels.png")

    return lime_all
# =====================================================================================================================

# скрин фонового окна. перекрывать можно но НЕ СВОРАЧИВАТЬ
def screen_fonovogo_okna():
    print(time.time())
    hwnd = win32gui.FindWindow(None, 'Bless Unleashed')

    # Change the line below depending on whether you want the whole window
    # or just the client area.
    # left, top, right, bot = win32gui.GetClientRect(hwnd)
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bot - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

    saveDC.SelectObject(saveBitMap)

    # Change the line below depending on whether you want the whole window
    # or just the client area.
    # result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
    print(result)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)

    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        # PrintWindow Succeeded
        im.save("test.png")
    print(time.time())

    print(dir(ctypes))
    print(dir(win32gui))
    print(dir(win32ui))
# =====================================================================================================================

# искать зеленые всполохи в центре экрана
def count_lime(monitor):
    # сделать скрин и передать его в двумерный массив, сразу откинуть слой прозрачности из BGRA
    screen_to_array = np.array(mss.mss().grab(monitor), dtype=np.uint8)[:, :, 0:3]
    # в этом месте можно конвертировать RGB в BGR и назад
    # путем перемножения на разные байты [65536, 256, 1] и [1, 256, 65536] соответственно
    # преобразуем трехмерный масив с цветами пикселей BGR в двумерный с цветами в десятиричной системе
    pix_to_dex_24_bit = np.dot(screen_to_array.astype(np.uint8), [1, 256, 65536])
    # преобразовать двумерный массив в одномерный
    array_from_file = pix_to_dex_24_bit.flatten()
    # и делаем ПЕРЕСЕЧЕНИЕ двух ОДНОМЕРНЫХ массивов командой "in1d"
    cross_array = np.in1d(array_from_file, lime_array)
    # подсчитыеваем количество пикселей из первого массива совпавших со значениями в втором
    count_find_pix = np.count_nonzero(cross_array)
    # можно получить одномерный масив искомых пикселей
    # find_pix = array_from_file[cross_array]
    return count_find_pix
# =====================================================================================================================


# найти ДВА наиболее близких друг к другу делителя чьё произведение равно
# произведению ТРЕХ заданных положительных чисел
def find_mlt_xy(mlt_xy):
    # в любом варианте число делится на себя и на единицу
    min_x = mlt_xy
    min_y = 1
    # ищем другие варианты перебором
    for i in range(mlt_xy):
        # если остаток от деления (операция "%") равен нолю, то у нас деление без остатка
        if (mlt_xy % (i + 1)) == 0:
            # делим по модулю, точно зная что остатка не будет
            y = (mlt_xy // (i + 1))
            # если сумма двух множителей меньше предыдудущей суммы минимумов
            if (y + (i + 1)) < min_x:
                # то присвоить новый минимум сумм
                min_x = (y + (i + 1))
                min_y = i + 1
    # делим по модулю, точно зная что остатка не будет
    min_x = mlt_xy // min_y
    # если надо, то разворачиваем размеры так чтобы результат был широким, а не стоял столбиком как в тик-токе
    if min_x < min_y:
        tmp_min_x = min_y
        min_y = min_x
        min_x = tmp_min_x
    return min_x, min_y
# =====================================================================================================================


# флаг остановки скрипта
def key_space():
    global flag_stop
    flag_stop = True
# =====================================================================================================================


# флаг сохранения кусков экрана
def key_f2():
    # сохраняем много скринов всего экрана
    # повторное нажатие останавливает сохранение скринов
    global flag_screen
    if flag_screen:
        flag_screen = False
        winsound.Beep(500, 300)
    else:
        winsound.Beep(2500, 300)
        flag_screen = True
# =====================================================================================================================

# функция сохранения кусков экрана
def save_chunk_screen():
    with mss.mss() as sct:
        # для центра
        sct_img = sct.grab(xy_centre_monitor)
        fname = (full_path
                 + 'Screen/screen_'
                 + str(time.time())
                 + ("_{top}x{left}_{width}x{height}".format(**xy_centre_monitor))
                 + '.png')
        mss.tools.to_png(sct_img.rgb, sct_img.size, output=fname)
        # и повторяем для низа
        sct_img = sct.grab(xy_down_right_monitor)
        fname = (full_path
                 + 'Screen/screen_'
                 + str(time.time())
                 + ("_{top}x{left}_{width}x{height}".format(**xy_down_right_monitor))
                 + '.png')
        mss.tools.to_png(sct_img.rgb, sct_img.size, output=fname)
# =====================================================================================================================


# ======================================================
# ЗАПУСК ОСНОВНОГО ТЕЛА СКРИПТА
# ======================================================

# создаем директорию для сохранения скринов, даже если она существует
full_path = 'Resourse/' + str(date.today()) + '/'
os.makedirs(full_path + "Screen/", exist_ok=True)
print("RUN ...")
winsound.Beep(800, 100)

# регистрируем события
# keyboard.on_press_key('F1', lambda event: key_f1(), suppress=False)
keyboard.on_press_key('F2', lambda event: key_f2(), suppress=False)
# keyboard.on_press_key('F3', lambda event: key_f3(), suppress=False)
keyboard.on_press_key(' ', lambda event: key_space(), suppress=False)

# запретить сканировать зеленые всполохи
green_ok = False
# флаг проводки рыбы
provodka = False
# заглушка я не понял как по другому
emp = False
# флаг действия между проводками, там нужно дергать правую мышку в одну из четырех сторон
flag_action = False

# словари для определения контролируемых точек и размеров экрана
# для широкого монитора
all_dict_2560x1080 = {'Botton_Cast_X': 50, 'Botton_Cast_Y': 34,
                      'Botton_Stop_X': 410, 'Botton_Stop_Y': 34,
                      'Botton_Hook_X': 240, 'Botton_Hook_Y': 34,
                      'Botton_Reel_X': 366, 'Botton_Reel_Y': 34,
                      'orange': np.array([65, 150, 181], dtype=np.uint8),
                      'lime_r': 255, 'lime_g': 255, 'lime_b': 170,
                      'lime_r_': 0, 'lime_g_': 0, 'lime_b_': 60,
                      'monitor': {"left": 0, "top": 0, "width": 2560, "height": 1080},
                      'xy_down_right_monitor': {"left": 1995, "top": 950, "width": 550, "height": 80},
                      'xy_centre_monitor': {"left": 1295, "top": 415, "width": 23, "height": 26},
                      'xy_big_monitor': {"left": 0, "top": 320, "width": 2560, "height": 500}
                      }
# для обычного монитора
all_dict_1920x1080 = {'Botton_Cast_X': 43, 'Botton_Cast_Y': 27,
                      'Botton_Stop_X': 406, 'Botton_Stop_Y': 27,
                      'Botton_Hook_X': 235, 'Botton_Hook_Y': 27,
                      'Botton_Reel_X': 361, 'Botton_Reel_Y': 27,
                      'orange': np.array([65, 150, 181], dtype=np.uint8),
                      'lime_r': 255, 'lime_g': 255, 'lime_b': 170,
                      'lime_r_': 0, 'lime_g_': 0, 'lime_b_': 60,
                      'monitor': {"left": 0, "top": 0, "width": 1920, "height": 1080},
                      'xy_down_right_monitor': {"left": 1360, "top": 955, "width": 550, "height": 75},
                      'xy_centre_monitor': {"left": 975, "top": 415, "width": 24, "height": 27},
                      'xy_big_monitor': {"left": 0, "top": 300, "width": 1920, "height": 450}
                      }

# пока определяем какой словарь использовать без проверок условий
all_dict = all_dict_1920x1080
#all_dict = all_dict_2560x1080


# определяем из уже объявленого словаря

# цвет кнопки мышки в подсказке снизу справа в BGR без альфа канала
orange = all_dict['orange']

# координаты считываемого куска экрана - низ право экрана
xy_down_right_monitor = all_dict['xy_down_right_monitor']
# координаты считываемого куска экрана - центр экрана
xy_centre_monitor = all_dict['xy_centre_monitor']
# координаты бля большого центрального скрина
xy_big_monitor = all_dict['xy_big_monitor']
# создаём массив цвета лайм с дельтой
lime_array = create_lime_array(all_dict['lime_r'], all_dict['lime_g'], all_dict['lime_b'],
                               all_dict['lime_r_'], all_dict['lime_g_'], all_dict['lime_b_'])

def sleep(duration, get_now=time.perf_counter):
    now = get_now()
    end = now + duration
    while now < end:
        now = get_now()

while True:

    # если активное окно не игра, то ничего не делать
    if GetWindowText(GetForegroundWindow()) != "Bless Unleashed":
        # подождать и пойти на следующий цикл "while True:"
        time.sleep(2)
        continue

    # если установлен флаг сохранения кусков экрана
    if flag_screen:
        # то сохранить куски экрана
        save_chunk_screen()

    # хорошая задержка
    time.sleep(0.05)

    if flag_stop:
        # выход
        winsound.Beep(800, 100)
        winsound.Beep(800, 100)
        winsound.Beep(800, 100)
        print("DONE !!!")
        sys.exit()

    # ПЕРВОЕ: сделать скрин нижнего правого угла, передать скрин в тип массив, откинуть альфа канал
    down_monitor = np.array(mss.mss().grab(xy_down_right_monitor), dtype=np.uint8)[:, :, 0:3]

    # забросить удочку если есть соответствующее приглашение
    if (down_monitor[all_dict['Botton_Cast_Y'], all_dict['Botton_Cast_X']] == orange).all() and \
            (down_monitor[all_dict['Botton_Stop_Y'], all_dict['Botton_Stop_X']] == orange).all():
        # разрешить сканировать зеленые всполохи
        green_ok = True
        # зажать левую мышку
        mouse.press(button='left')
        # дать на заброс две секунды плюс рандом
        time.sleep(2 + (random.randint(1, 2) / 10))
        # отпустить мышку
        mouse.release(button='left')
        # пойти на следующий цикл "while True:"
        continue

    # ВТОРОЕ ждать зеленых всплохов
    if green_ok and count_lime(xy_centre_monitor) > 0:
        # заглушка для теста
        #save_chunk_screen()
        #time.sleep(0.5)
        #continue
        mouse.press(button='left')
        green_ok = False
        winsound.Beep(800, 100)
        time.sleep(0.4)
        mouse.release(button='left')

    # ТРЕТЬЕ жать мышку пока есть надпись Reel IN но нет надписи Hook
    if (down_monitor[all_dict['Botton_Hook_Y'], all_dict['Botton_Hook_X']] == orange).all():
        # какая-то пустота
        emp = True
    else:
        if (down_monitor[all_dict['Botton_Reel_Y'], all_dict['Botton_Reel_X']] == orange).all():
            # если проводка идет то пойти на следующий цикл "while True:"
            if provodka:
                continue
            else:
                # включить флаг проводки рыбы на крючке
                provodka = True
                # зажать левую кнопку для начала проводки рыбы
                mouse.press(button='left')
                # пойти на следующий цикл "while True:"
                continue

    # ЧЕТВЕРТОЕ если была начата проводка, но пиксели не распознаны
    # то отпустить мышку и обнулить проводку
    if provodka:
        provodka = False
        # print("проводка окончена")
        mouse.release(button='left')
        # по окончанию проводки поднять флаг проверки необходимого действия
        flag_action = True
        # пойти на следующий цикл "while True:"
        continue

    if flag_action:
        #print('Действие выполнено - ' + str(time.time()))
        # сделать большой скрин центральной части экрана, откинуть альфа канал
        pic_big_monitor = mss.mss().grab(xy_big_monitor)
        big_monitor = np.array(pic_big_monitor, dtype=np.uint8)[:, :, 0:3]
        # перикинуть масив из BGR (3 х uint8) в dex (1 х int32)
        big_monitor = np.dot(big_monitor.astype(np.uint8), [1, 256, 65536])
        # цвет искомого пикселя восклицательного знака
        color_exclamation_point = np.dot(np.array([156, 5, 5], dtype=np.uint8).astype(np.uint8), [65536, 256, 1])
        # цвет искомого пикселя стрелок
        color_arrow = np.dot(np.array([255, 179, 43], dtype=np.uint8).astype(np.uint8), [65536, 256, 1])
        # шаг для поиска пикселя, если маленький - то медленный поиск, если больше 5 - то может не найти
        step = 5
        end_y = big_monitor.shape[0] - step
        end_x = big_monitor.shape[1] - step
        # координаты найденного пикселя
        # обнуляем стартовые значения для входа в цикл поиска
        find_x = 0
        find_y = 0
        flag_find_arrow = False
        flag_find_exclamation_point = False
        x = 0
        y = 0
        while y < end_y:
            while x < end_x:
                # если найден нужный пиксель, запомнить координаты, поднять флаг найденого пикселя, прервать оба цикла перебора
                if big_monitor[y, x] == color_exclamation_point:
                    flag_find_exclamation_point = True
                    find_y = y
                    find_x = x
                    y = end_y
                    continue
                # если найден нужный пиксель, запомнить координаты, поднять флаг найденого пикселя
                if big_monitor[y, x] == color_arrow:
                    flag_find_arrow = True
                    find_y = y
                    find_x = x
                    y = end_y
                    continue
                x = x + step
            x = 0
            y = y + step

        if flag_find_exclamation_point:
            print('знак - ' + str(time.time()) +' - '+ str(find_x) + ':' + str(find_y))
        else:
            if flag_find_arrow:
                # от координат найденого пикселя отмеряем квадрат будущего точечного поиска
                tmp_start_y = find_y - 14
                tmp_start_x = find_x - 14
                tmp_end_y = find_y + 14
                tmp_end_x = find_x + 14
                # проверяем что границы полученного квадрата поиска не выходят за границы сделанного скрина
                if tmp_start_y < 0:
                    tmp_start_y = 0
                if tmp_start_x < 0:
                    tmp_start_x = 0
                if tmp_end_y > all_dict['xy_big_monitor']['height']:
                    tmp_end_y = all_dict['xy_big_monitor']['height']
                if tmp_end_x > all_dict['xy_big_monitor']['width']:
                    tmp_end_x = all_dict['xy_big_monitor']['width']

                arrow_find_y = 0
                arrow_find_x = 0
                # запускаем точный поиск верхнего левого пикселя
                while tmp_start_y < tmp_end_y:
                    while tmp_start_x < tmp_end_x:
                        # действие сравнения
                        if big_monitor[tmp_start_y, tmp_start_x] == color_arrow:
                            # если искомый пиксель найден, запомнить координаты, выскочить из циклов
                            arrow_find_y = tmp_start_y
                            arrow_find_x = tmp_start_x
                            tmp_start_y = tmp_end_y
                            continue
                        tmp_start_x = tmp_start_x + 1
                    tmp_start_x = 0
                    tmp_start_y = tmp_start_y + 1

                text_arrow = 'FIND'
                if big_monitor[arrow_find_y, (arrow_find_x+1)] != color_arrow:
                    # дернуть мышку вверх
                    mouse.press(button='right')
                    for move in range(750):
                        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 0, -1, 0, 0)
                        sleep(0.001)
                    mouse.release(button='right')
                    text_arrow = 'UP'
                elif big_monitor[(arrow_find_y+2), arrow_find_x] != color_arrow:
                    # дернуть мышку вниз
                    mouse.press(button='right')
                    for move in range(750):
                        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 0, 1, 0, 0)
                        sleep(0.001)
                    mouse.release(button='right')
                    text_arrow = 'DOWN'
                elif big_monitor[(arrow_find_y + 1), (arrow_find_x - 1)] != color_arrow:
                    # дернуть мышку вправо
                    mouse.press(button='right')
                    for move in range(750):
                        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 1, 0, 0, 0)
                        sleep(0.001)
                    mouse.release(button='right')
                    text_arrow = 'RIGHT'
                elif big_monitor[(arrow_find_y + 3), (arrow_find_x + 2)] != color_arrow:
                    # дернуть мышку влево
                    mouse.press(button='right')
                    for move in range(750):
                        win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, -1, 0, 0, 0)
                        sleep(0.001)
                    mouse.release(button='right')
                    text_arrow = 'LEFT'
                else:
                    text_arrow = 'NONE'

                # для теста сохраняем большой скрин со стрелкой
                fname = (full_path
                         + 'Screen/screen_'
                         + str(time.time())
                         + ("_{top}x{left}_{width}x{height}".format(**xy_centre_monitor))
                         + '.'
                         + str(arrow_find_x)
                         + '-'
                         + str(arrow_find_y)
                         + '.'
                         + text_arrow
                         + '.png')
                mss.tools.to_png(pic_big_monitor.rgb, pic_big_monitor.size, output=fname)

                print('стрелка - ' + text_arrow + ' - ' + str(time.time()) + ' - ' + str(find_x) + ':' + str(find_y))
            else:
                print('NONE - ' + str(time.time()))
                time.sleep(1.5)
                keyboard.press_and_release('F9')
                time.sleep(0.5)
        # опустить флаг необходимого действия
        flag_action = False

# мусорная корзина
"""



import win32gui
import win32ui

def background_screenshot(hwnd, width, height):
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj=win32ui.CreateDCFromHandle(wDC)
    cDC=dcObj.CreateCompatibleDC()
    dataBitMap = win32ui.CreateBitmap()
    dataBitMap.CreateCompatibleBitmap(dcObj, width, height)
    cDC.SelectObject(dataBitMap)
    cDC.BitBlt((0,0),(width, height) , dcObj, (0,0), win32con.SRCCOPY)
    dataBitMap.SaveBitmapFile(cDC, 'screenshot.bmp')
    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(dataBitMap.GetHandle())

hwnd = win32gui.FindWindow(None, windowname)
background_screenshot(hwnd, 1280, 780)


"""
