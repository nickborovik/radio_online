import os
import xlrd
import datetime
import re
import random
from mutagen.mp3 import MP3

# DATES
# -------------------------------------------------------------------------------------------------
CURRENT_DAY = datetime.datetime.today().date()
NEXT_DAY = CURRENT_DAY + datetime.timedelta(days=1)

PLAYLIST_DATE_FOR_TOMORROW = (CURRENT_DAY + datetime.timedelta(days=1)).strftime('%d%m%Y')
NEXT_PLAYLIST_DATE = (CURRENT_DAY + datetime.timedelta(days=2)).strftime('%d_%m_%Y')

MONTH = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
         'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

# DIRS
# -------------------------------------------------------------------------------------------------
BASE_DIR = os.getcwd()
MEDIA_DIR = os.path.join(BASE_DIR, 'Archive_2018')
PLAYLIST_DIR = os.path.join('D:\\', 'Playlist Radioboss')

# TEST CASE
# PLAYLIST_DIR = BASE_DIR

KIEV_STUDIO_DIR_TODAY = os.path.join(
    BASE_DIR,
    'Kievskaya Studia',
    f'!{CURRENT_DAY.strftime("%m %Y")}'
)

KIEV_STUDIO_DIR_TOMORROW = os.path.join(
    BASE_DIR,
    'Kievskaya Studia',
    f'!{NEXT_DAY.strftime("%m %Y")}'
)

KHARKOV_STUDIO_DIR_TODAY = os.path.join(
    BASE_DIR,
    '1 SREDNIE VOLNI ONLINE',
    'KharkovTWR'
    f'{CURRENT_DAY.strftime("%m-%Y")}'
)

KHARKOV_STUDIO_DIR_TOMORROW = os.path.join(
    BASE_DIR,
    '1 SREDNIE VOLNI ONLINE',
    f'{NEXT_DAY.strftime("%m-%Y")}'
)


# EXCEL settings
# -------------------------------------------------------------------------------------------------

EXCEL_FILE_NAME = f'{CURRENT_DAY.strftime("%m-%Y")} Расписание онлайн вещания ({MONTH[CURRENT_DAY.month - 1]}).xlsx'
FULL_EXCEL_FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)
EXCEL_PAGE_NAME = f'{NEXT_DAY.day}.{NEXT_DAY.strftime("%m")}'


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
    '900 секунд доброты': ['900_sekund_dobroti_{}.mp3', 'Kharkov'],
    'БА': ['RUS_BST_0420_20200804_1800_BR_.mp3', 'Kiev'],
    'Библейские искатели': ['RUS_TSK_{}.mp3', 'Kiev'],
    'Вивчаємо Біблію разом': ['Bible_study_{}.mp3', 'Kharkov'],
    'ВЦП': ['UKR_PRC_{}.mp3',  'Kiev'],
    'Герои': ['Gde_vi_geroi_{}.mp3', 'Kharkov'],
    'Голос друга': ['BEL_VFR_{}.mp3', 'Kiev'],
    'Джерельце': ['UKR_TLS_{}.mp3', 'Kiev'],
    'ЖКОЕ': ['Zhizn_kak_ona_est_{}.mp3', 'Kharkov'],
    'ЖН': ['UKR_HOPE_{}.mp3', 'Kiev'],
    'Калейдоскоп': ['UKR_KAL_{}.mp3', 'Kiev'],
    'МН': ['RUS_BOH_{}.mp3', 'Kiev'],
    'Ответственность': ['RUS_SPOT_{}.mp3', 'Kiev'],
    'Погляд ': ['UKR_TOV_{}.mp3', 'Kiev'],
    'Свет жизни': ['RUS_IFL_{}.mp3', 'Kiev'],
    'Серебро': ['RUS_SIL_{}.mp3', 'Kiev'],
    'Слово на сегодня': ['RUS_TWT_{}.mp3', 'Kiev'],
    'Стежинка': ['UKR_TLP_{}.mp3', 'Kiev'],
    'Суламита': ['RUS_SUL_{}.mp3', 'Kiev'],
    'Табор': ['Tabor uhodit v nebo_412_040820.mp3', 'Kharkov'],
    'Тихие воды': ['RUS_SWA_{}.mp3', 'Kiev'],
    'Хлеб жизни': ['RUS_BLR_{}.mp3', 'Kiev'],
    'Шалом': ['Shalom_{}.mp3', 'Kharkov'],
    'Шанс // ГВЛ': ['UKR_MAE_{}.mp3', 'Kiev'],
}

def get_excel_info(file_name, excel_page_name):
    workbook = xlrd.open_workbook(file_name)
    sheet = workbook.sheet_by_name(excel_page_name)
    return sheet

def get_mp3_file_length(full_path_to_file):
    mp3_data = MP3(full_path_to_file)
    return int(mp3_data.info.length)


def write_playlist_to_file(date, file_data):
    with open(os.path.join(PLAYLIST_DIR, f'playlist for {date}.m3u8'), 'w') as write_file:
        write_file.writelines(file_data)
    print(f'Плейлист на {date} готов и находится в папке \n{PLAYLIST_DIR}')


def main():
    file_data = ['#EXTM3U\n']
    sheet = get_excel_info(FULL_EXCEL_FILE_PATH, EXCEL_PAGE_NAME)

    for i in range(sheet.nrows):
        if i < 3:
            continue
        else:
            if sheet.cell_value(i, 5) == 'ГОДИНА БОЖОГО СЛОВА':
                load_file_number = str(round(sheet.cell_value(i, 4)))
                mp3_file_name = f'Online radio blok {load_file_number}.mp3'
                load_file_name = 'Online radio blok'
                full_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

            elif 32 > i >= 28:
                date = CURRENT_DAY.strftime('%Y%m%d')
                mp3_file_name = MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][0].format(date)
                load_file_name = MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][0].format(date)
                load_file_number = date
                if MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][1] == 'Kiev':
                    file_dir = KIEV_STUDIO_DIR_TODAY
                else:
                    file_dir = KHARKOV_STUDIO_DIR_TODAY
                full_file_path = os.path.join(file_dir, mp3_file_name)

            elif 65 > i >= 61:
                date = NEXT_DAY.strftime('%Y%m%d')
                mp3_file_name = MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][0].format(date)
                load_file_name = MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][0].format(date)
                load_file_number = date
                if MAIN_AUDIO_FILES[sheet.cell_value(i, 5)][1] == 'Kiev':
                    file_dir = KIEV_STUDIO_DIR_TOMORROW
                else:
                    file_dir = KHARKOV_STUDIO_DIR_TOMORROW
                full_file_path = os.path.join(file_dir, mp3_file_name)

            elif sheet.cell_value(i, 5) == 'муз.блок':
                muzblock_index = random.randrange(0, len(MUZBLOCKS))
                mp3_file_name = MUZBLOCKS[muzblock_index]
                load_file_name = 'Muzblock'
                load_file_number = muzblock_index + 1
                full_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

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

                full_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

            mp3_length = get_mp3_file_length(full_file_path)
            # TEST CASE
            # mp3_length = 960
            file_info = f'#EXTINF:{mp3_length},{load_file_name} - {load_file_number}\n'
            file_data.append(file_info)
            file_data.append(f'{full_file_path}\n')

    file_data.append(f'playlist {NEXT_PLAYLIST_DATE}.command')
    write_playlist_to_file(PLAYLIST_DATE_FOR_TOMORROW, file_data)


if __name__ == '__main__':
    main()
