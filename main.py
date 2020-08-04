import os
import xlrd
import datetime
import re
import random
from mutagen.mp3 import MP3

CURRENT_DAY = datetime.datetime.now().date()
NEXT_DAY = CURRENT_DAY + datetime.timedelta(days=1)

PLAYLIST_DATE_FOR_TOMORROW = (CURRENT_DAY + datetime.timedelta(days=1)).strftime('%d_%m_%Y')
NEXT_PLAYLIST_DATE = (CURRENT_DAY + datetime.timedelta(days=2)).strftime('%d_%m_%Y')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIA_DIR = os.path.join(BASE_DIR, 'Archive_2018')

KIEV_STUDIO_DIR = os.path.join(
    BASE_DIR,
    'INTERNET RADIO',
    'Kievskaya Studia',
    '!{}'.format(NEXT_DAY.strftime('%m-%Y'))
)

KHARKOV_STUDIO_DIR = os.path.join(
    BASE_DIR,
    'INTERNET RADIO',
    'KharkovTWR',
    '1 SREDNIE VOLNI ONLINE',
    '{}'.format(NEXT_DAY.strftime('%m-%Y'))
)

FILE_NAME = '08-2020 Расписание онлайн вещания (август).xlsx'
EXCEL_PAGE_NAME = (datetime.datetime.now().date() + datetime.timedelta(days=1)).strftime('%-d.%m')

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


main_audio_files = {
    '900 секунд доброты': '',
    'БА': 'RUS_BST_0420_20200804_1800_BR_.mp3\n',
    'Библейские искатели': '',
    'Вивчаємо Біблію разом': 'Bible study_{}_{}.mp3\n'.format('000', CURRENT_DAY.strftime('%d%m%y')),
    'ВЦП': '',
    'Герои': '',
    'ГОДИНА БОЖОГО СЛОВА': 'Online radio blok',
    'Голос друга': '',
    'Джерельце': '',
    'ЖКОЕ': '',
    'ЖН': '',
    'Калейдоскоп': '',
    'МН': 'RUS_BOH_{}_{}_1800_BR_.mp3\n',
    'Ответственность': '',
    'Погляд ': '',
    'Свет жизни': '',
    'Серебро': 'RUS_SIL_0382_20200805_1930_BR_.mp3\n',
    'Слово на сегодня': '',
    'Стежинка': '',
    'Суламита': 'RUS_SUL_1107_20200805_1845_BR_.mp3\n',
    'Табор': 'Tabor uhodit v nebo_412_040820.mp3\n',
    'Тихие воды': '',
    'Хлеб жизни': 'RUS_BLR_0378_20200804_1815_BR_.mp3\n',
    'Шалом': '',
    'Шанс // ГВЛ': '',

}
print(EXCEL_PAGE_NAME)

def main():
    file_data = ['#EXTM3U\n']
    workbook = xlrd.open_workbook(FILE_NAME)
    sheet = workbook.sheet_by_name(EXCEL_PAGE_NAME)
    for i in range(sheet.nrows):
        if i < 3:
            continue
        else:
            if 32 > i >= 28 or 66 > i >= 61:

                if sheet.cell_value(i, 5) == 'ГОДИНА БОЖОГО СЛОВА':
                    mp3_file_name = '{} {}.mp3\n'.format(main_audio_files[sheet.cell_value(i, 5)], str(round(sheet.cell_value(i, 4))))
                else:
                    mp3_file_name = main_audio_files[sheet.cell_value(i, 5)]
                load_file_name = sheet.cell_value(i, 5)
                load_file_number = ''

            elif sheet.cell_value(i, 5) == 'муз.блок':
                mp3_file_name = '{}\n'.format(MUZBLOCKS[random.randrange(0, len(MUZBLOCKS))])
                load_file_name = 'Muzblock'
                load_file_number = ''

            else:
                load_file_name = str(sheet.cell_value(i, 3)) + str()
                try:
                    load_file_number = str(round(sheet.cell_value(i, 4)))
                except:
                    load_file_number = str(sheet.cell_value(i, 4))
                mp3_file_name = '{} {}.mp3\n'.format(load_file_name, load_file_number)
                mp3_file_name = re.sub(r'\s\s', ' ', mp3_file_name)
                # audio_data = MP3(os.path.join(MEDIA_DIR, mp3_file_name))
                audio_data = MP3('TBS 0011.mp3')

            file_length = '#EXTINF:{},{} {}\n'.format(int(audio_data.info.length), load_file_name, load_file_number)
            file_data.append(file_length)
            file_data.append(os.path.join(MEDIA_DIR, mp3_file_name))

    file_data.append('playlist {}.command'.format(NEXT_PLAYLIST_DATE))

    with open('playlist for {}.m3u8'.format(PLAYLIST_DATE_FOR_TOMORROW), 'w') as write_file:
        write_file.writelines(file_data)

    print('Плейлист на {} готов!'.format(PLAYLIST_DATE_FOR_TOMORROW))


if __name__ == '__main__':
    main()

