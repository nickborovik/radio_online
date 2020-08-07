import os
from openpyxl import load_workbook
import datetime
import re
import random
from mutagen.mp3 import MP3

# DATES
# -------------------------------------------------------------------------------------------------
CURRENT_DAY = datetime.datetime.today().date()
NEXT_DAY = CURRENT_DAY + datetime.timedelta(days=1)

PLAYLIST_DATE_FOR_TOMORROW = NEXT_DAY.strftime('%d%m%Y')
NEXT_PLAYLIST_DATE = (NEXT_DAY + datetime.timedelta(days=1)).strftime('%d_%m_%Y')

MONTH = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
         'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

# DIRS
# -------------------------------------------------------------------------------------------------
BASE_DIR = os.getcwd()
MEDIA_DIR = os.path.join(BASE_DIR, 'Archive_2018')
PLAYLIST_DIR = os.path.join('D:\\', 'Playlist Radioboss')

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
    'KharkovTWR',
    '1 SREDNIE VOLNI ONLINE',
    f'{CURRENT_DAY.strftime("%m-%Y")}'
)

KHARKOV_STUDIO_DIR_TOMORROW = os.path.join(
    BASE_DIR,
    'KharkovTWR',
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
    '01 Kharkov time 11 min-a.mp3',
    '02 Kharkov time 11 min-a.mp3',
    '03 Kharkov time 11 min-a.mp3',
    '04 Kharkov time 11 min-a.mp3',
    '05 Kharkov time 11 min-a.mp3',
    '06 Kharkov time 11 min-a.mp3',
    '07 Kharkov time 11 min-a.mp3',
    '08 Kharkov time 11 min-a.mp3',
    '09 Kharkov time 11 min-a.mp3',
    '10 Kharkov time 11 min-a.mp3',
    '11 Kharkov time 9 min-a.mp3',
    '12 Kharkov time 9 min-a.mp3',
    '13 Kharkov time 10 min-a.mp3',
    '14 Kharkov time 10 min-a.mp3',
    '15 Kharkov time 10 min-a.mp3',
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

# MAIN CODE

def get_excel_info(file_name, excel_page_name):
    workbook = load_workbook(file_name)
    sheet = workbook[excel_page_name]
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

    for row in sheet.iter_rows(min_row=4, max_row=69, max_col=6, values_only=True):

        if row[5] == 'муз.блок':
            '''Выбрать случайный музлок и вставить в вместо пустого поля'''
            mp3_file_name = MUZBLOCKS[random.randrange(0, len(MUZBLOCKS))]
            file_name = 'Muzblock'
            full_mp3_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

        elif row[5] == 'ГОДИНА БОЖОГО СЛОВА':
            '''Конкретный случай для передачи Година Божого Слова'''
            load_file_number = row[4]
            file_name = 'Online radio blok'
            mp3_file_name = f'{file_name} {load_file_number}.mp3'
            full_mp3_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

        elif 30 > row[0] >= 26:
            '''Повтор за вчера'''
            date = CURRENT_DAY.strftime('%Y%m%d')
            file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)
            mp3_file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)

            if MAIN_AUDIO_FILES[row[5]][1] == 'Kiev':
                file_dir = KIEV_STUDIO_DIR_TODAY
            else:
                file_dir = KHARKOV_STUDIO_DIR_TODAY
            full_mp3_file_path = os.path.join(file_dir, mp3_file_name)

        elif 63 > row[0] >= 59:
            '''Прямой эфир'''
            date = NEXT_DAY.strftime('%Y%m%d')
            file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)
            mp3_file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)

            if MAIN_AUDIO_FILES[row[5]][1] == 'Kiev':
                file_dir = KIEV_STUDIO_DIR_TOMORROW
            else:
                file_dir = KHARKOV_STUDIO_DIR_TOMORROW
            full_mp3_file_path = os.path.join(file_dir, mp3_file_name)

        else:
            '''Все остальные случаи, где файл из папки Archive_2018'''
            if 'Лекция' in str(row[4]):
                file_number = re.sub(r'Лекция', 'L', row[4])
            elif 'М.В.' in str(row[4]):
                file_number = re.sub(r'M\.B\.', 'M', row[4])
            else:
                file_number = row[4]

            file_name = row[3]
            mp3_file_name = re.sub(r'\s\s', ' ', f'{file_name} {file_number}.mp3')
            full_mp3_file_path = os.path.join(MEDIA_DIR, mp3_file_name)

        mp3_file_length = get_mp3_file_length(full_mp3_file_path)

        file_data.append(f'#EXTINF:{mp3_file_length},{file_name}\n')
        file_data.append(f'{full_mp3_file_path}\n')

    file_data.append(f'playlist {NEXT_PLAYLIST_DATE}.command')

    write_playlist_to_file(PLAYLIST_DATE_FOR_TOMORROW, file_data)


if __name__ == '__main__':
    main()