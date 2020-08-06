import os
import xlrd
import datetime
import re
import random
from mutagen.mp3 import MP3

# DATES
# -------------------------------------------------------------------------------------------------
CURRENT_DAY = datetime.datetime.now().date()
NEXT_DAY = CURRENT_DAY + datetime.timedelta(days=1)

PLAYLIST_DATE_FOR_TOMORROW = (CURRENT_DAY + datetime.timedelta(days=1)).strftime('%d%m%Y')
NEXT_PLAYLIST_DATE = (CURRENT_DAY + datetime.timedelta(days=2)).strftime('%d_%m_%Y')

# DIRS
# -------------------------------------------------------------------------------------------------
BASE_DIR = os.getcwd()
MEDIA_DIR = os.path.join(BASE_DIR, 'Archive_2018')
PLAYLIST_DIR = os.path.join('D:', 'Playlist Radioboss')

KIEV_STUDIO_DIR_TODAY = os.path.join(
    BASE_DIR,
    'Kievskaya Studia',
    '!{}'.format(CURRENT_DAY.strftime('%m %Y'))
)

KIEV_STUDIO_DIR_TOMORROW = os.path.join(
    BASE_DIR,
    'Kievskaya Studia',
    '!{}'.format(NEXT_DAY.strftime('%m %Y'))
)

KHARKOV_STUDIO_DIR_TODAY = os.path.join(
    BASE_DIR,
    '1 SREDNIE VOLNI ONLINE',
    '{}'.format(CURRENT_DAY.strftime('%m-%Y'))
)

KHARKOV_STUDIO_DIR_TOMORROW = os.path.join(
    BASE_DIR,
    '1 SREDNIE VOLNI ONLINE',
    '{}'.format(NEXT_DAY.strftime('%m-%Y'))
)


# EXCEL settings
# -------------------------------------------------------------------------------------------------
FILE_NAME = os.path.join(BASE_DIR, '08-2020 Расписание онлайн вещания (август).xlsx')
EXCEL_PAGE_NAME = ((datetime.datetime.now() + datetime.timedelta(days=1)).date()).strftime('%-d.%m')


# NAMES FOR FILES
# -------------------------------------------------------------------------------------------------
MUZBLOCKS = [
    'muzblok_01_time_12.15.mp3',
    'muzblok_02_time_14.43.mp3',
    'muzblok_03_time_13.42.mp3',
    'muzblok_04_time_14.57.mp3',
    'muzblok_05_time_13.20.mp3',
    'muzblok_06_time_14.19.mp3',
    'muzblok_07_time_14.21.mp3',
    'muzblok_08_time_13.55.mp3',
    'muzblok_09_time_14.20.mp3',
    'muzblok_10_time_14.33.mp3',
    'muzblok_11_time_14.02.mp3',
    'muzblok_12_time_12.40.mp3',
    'muzblok_13_time_14.27.mp3',
    'muzblok_14_time_14.41.mp3',
    'muzblok_15_time_13.58.mp3',
    'muzblok_16_time_14.49.mp3',
    'muzblok_17_time_14.46.mp3',
    'muzblok_18_time_12.40.mp3',
    'muzblok_19_time_14.55.mp3',
    'muzblok_20_time_14.59.mp3',
    'muzblok_21_time_15.43.mp3',
    'muzblok_22_time_15.33.mp3',
    'muzblok_23_time_14.55.mp3',
    'muzblok_24_time_14.16.mp3',
    'muzblok_25_time_16.18.mp3',
    'muzblok_26_time_16.16.mp3',
]

MAIN_AUDIO_FILES = {
    '900 секунд доброты': '900 sekund dobroti_{}.mp3',
    'БА': 'RUS_BST_0420_20200804_1800_BR_.mp3',
    'Библейские искатели': 'a.mp3',
    'Вивчаємо Біблію разом': 'Bible study_{}.mp3',
    'ВЦП': 'a.mp3',
    'Герои': 'a.mp3',
    'ГОДИНА БОЖОГО СЛОВА': 'Online radio blok {}.mp3',
    'Голос друга': 'a.mp3',
    'Джерельце': 'a.mp3',
    'ЖКОЕ': 'a.mp3',
    'ЖН': 'a.mp3',
    'Калейдоскоп': 'a.mp3',
    'МН': 'RUS_BOH_{}_{}_1800_BR_.mp3',
    'Ответственность': 'a.mp3',
    'Погляд ': 'a.mp3',
    'Свет жизни': 'a.mp3',
    'Серебро': 'RUS_SIL_0382_20200805_1930_BR_.mp3',
    'Слово на сегодня': 'a.mp3',
    'Стежинка': 'a.mp3',
    'Суламита': 'RUS_SUL_1107_20200805_1845_BR_.mp3',
    'Табор': 'Tabor uhodit v nebo_412_040820.mp3',
    'Тихие воды': 'a.mp3',
    'Хлеб жизни': 'RUS_BLR_0378_20200804_1815_BR_.mp3',
    'Шалом': 'a.mp3',
    'Шанс // ГВЛ': 'a.mp3',
}

def get_excel_info(file_name, excel_page_name):
    workbook = xlrd.open_workbook(file_name)
    sheet = workbook.sheet_by_name(excel_page_name)
    return sheet

def get_mp3_file_length(media_dir, mp3_file_name):
    mp3_data = MP3(os.path.join(media_dir, mp3_file_name))
    return int(mp3_data.info.length)


def write_playlist_to_file(date, file_data):
    with open(os.path.join(PLAYLIST_DIR, f'playlist for {date}.m3u8'), 'w') as write_file:
        write_file.writelines(file_data)
    print('Плейлист на {} готов!'.format(date))


def main():
    file_data = ['#EXTM3U\n']
    sheet = get_excel_info(FILE_NAME, EXCEL_PAGE_NAME)

    for i in range(sheet.nrows):
        if i < 3:
            continue
        else:
            if 33 > i >= 28 or 66 > i >= 61:
                continue
                # mp3_file_name = MAIN_AUDIO_FILES[sheet.cell_value(i, 5)]
                # load_file_name = sheet.cell_value(i, 5)
                # load_file_number = ''

            if sheet.cell_value(i, 5) == 'муз.блок':
                muzblock_index = random.randrange(0, len(MUZBLOCKS))
                mp3_file_name = MUZBLOCKS[muzblock_index]
                load_file_name = 'Muzblock'
                load_file_number = muzblock_index + 1

            else:
                load_file_name = str(sheet.cell_value(i, 3))
                try:
                    load_file_number = str(round(sheet.cell_value(i, 4)))
                except:
                    load_file_number = str(sheet.cell_value(i, 4))
                    if 'Лекция' in load_file_number:
                        load_file_number = re.sub(r'Лекция', 'L', load_file_number)
                    if 'М.В.' in load_file_number:
                        load_file_number = re.sub(r'M\.B\.', 'M', load_file_number)

                mp3_file_name = f'{load_file_name} {load_file_number}.mp3'
                mp3_file_name = re.sub(r'\s\s', ' ', mp3_file_name)
                mp3_length = get_mp3_file_length(MEDIA_DIR, mp3_file_name)

            full_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

            file_info = f'#EXTINF:{mp3_length},{load_file_name} - {load_file_number}\n'
            file_data.append(file_info)
            file_data.append(f'{full_file_path}\n')

    file_data.append(f'playlist {NEXT_PLAYLIST_DATE}.command')
    write_playlist_to_file(PLAYLIST_DATE_FOR_TOMORROW, file_data)


if __name__ == '__main__':
    main()
