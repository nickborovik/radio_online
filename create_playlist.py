import datetime
import smtplib
from pathlib import Path
from configparser import ConfigParser
from openpyxl import load_workbook
from mutagen.mp3 import MP3
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# STATIC SETTINGS

# DATES

CUR_DAY = datetime.datetime.today().date()
NEXT_DAY = CUR_DAY + datetime.timedelta(days=1)

MONTH = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
         'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

# MAIN DIRS

ROOT = Path('D:/')
BASE_DIR = ROOT / 'INTERNET RADIO'
MEDIA_DIR = BASE_DIR / 'Archive_2018'
CONFIG_DIR = BASE_DIR / 'Playlist_auto_generator'
PLAYLIST_DIR = ROOT / 'Playlist Radioboss'
DO_15_MIN_DIR = MEDIA_DIR / 'domashniy ochag 15 min'
print(DO_15_MIN_DIR)

KIEV_ST_DIR_TODAY = BASE_DIR / 'Kievskaya Studia' / f'!{CUR_DAY.strftime("%m %Y")}'
KIEV_ST_DIR_TOMORROW = BASE_DIR / 'Kievskaya Studia' / f'!{NEXT_DAY.strftime("%m %Y")}'

KHAR_ST_DIR_TODAY = BASE_DIR / 'KharkovTWR' / '1 SREDNIE VOLNI ONLINE' / f'{CUR_DAY.strftime("%m-%Y")}'
KHAR_ST_DIR_TOMORROW = BASE_DIR / 'KharkovTWR' / '1 SREDNIE VOLNI ONLINE' / f'{NEXT_DAY.strftime("%m-%Y")}'

# EXCEL settings

EXCEL_FILE_NAME = f'{NEXT_DAY.strftime("%m-%Y")} Расписание онлайн вещания ({MONTH[NEXT_DAY.month - 1]}).xlsx'
EXCEL_FILE_PATH = BASE_DIR / EXCEL_FILE_NAME
EXCEL_PAGE_NAME = f'{NEXT_DAY.day}.{NEXT_DAY.strftime("%m")}'

# PLAYLIST settings

PLAYLIST_NAME = f'playlist_for_{NEXT_DAY.strftime("%d%m%Y")}.m3u8'
NEXT_PLAYLIST_NAME = f'playlist_for_{(NEXT_DAY + datetime.timedelta(days=1)).strftime("%d%m%Y")}.m3u8'
FULL_PLAYLIST_PATH = PLAYLIST_DIR / PLAYLIST_NAME

# MISTAKE REPORT settings

MISTAKE_SUBJECT = f"Playlist {PLAYLIST_NAME} for online radio TWR not created"

# MP3 FILES

MUZBLOCKS = {
    540: '11 Kharkov time 9 min-a.mp3',
    545: '12 Kharkov time 9 min-a.mp3',
    580: '14 Kharkov time 10 min-a.mp3',
    590: '13 Kharkov time 10 min-a.mp3',
    600: '15 Kharkov time 10 min-a.mp3',
    661: '09 Kharkov time 11 min-a.mp3',
    662: '04 Kharkov time 11 min-a.mp3',
    663: '10 Kharkov time 11 min-a.mp3',
    664: '03 Kharkov time 11 min-a.mp3',
    665: '01 Kharkov time 11 min-a.mp3',
    666: '06 Kharkov time 11 min-a.mp3',
    667: '05 Kharkov time 11 min-a.mp3',
    668: '02 Kharkov time 11 min-a.mp3',
    669: '08 Kharkov time 11 min-a.mp3',
    670: '07 Kharkov time 11 min-a.mp3',
    734: 'muzblok_01_time_12.15.mp3',
    759: 'muzblok_18_time_12.40.mp3',
    760: 'muzblok_12_time_12.40.mp3',
    799: 'muzblok_05_time_13.20.mp3',
    821: 'muzblok_03_time_13.42.mp3',
    834: 'muzblok_08_time_13.55.mp3',
    837: 'muzblok_15_time_13.58.mp3',
    841: 'muzblok_11_time_14.02.mp3',
    855: 'muzblok_24_time_14.16.mp3',
    858: 'muzblok_06_time_14.19.mp3',
    859: 'muzblok_09_time_14.20.mp3',
    860: 'muzblok_07_time_14.21.mp3',
    866: 'muzblok_13_time_14.27.mp3',
    872: 'muzblok_10_time_14.33.mp3',
    880: 'muzblok_14_time_14.41.mp3',
    882: 'muzblok_02_time_14.43.mp3',
    885: 'muzblok_17_time_14.46.mp3',
    888: 'muzblok_16_time_14.49.mp3',
    893: 'muzblok_23_time_14.55.mp3',
    894: 'muzblok_19_time_14.55.mp3',
    896: 'muzblok_04_time_14.57.mp3',
    898: 'muzblok_20_time_14.59.mp3',
    932: 'muzblok_22_time_15.33.mp3',
    942: 'muzblok_21_time_15.43.mp3',
    975: 'muzblok_26_time_16.16.mp3',
    977: 'muzblok_25_time_16.18.mp3',
}

MAIN_AUDIO_FILES = {
    '900 секунд доброты': ['RUS_KIND_{}.mp3', 'Kharkov'],
    'БА': ['RUS_BST_{}.mp3', 'Kiev'],
    'Библейские искатели': ['RUS_TSK_{}.mp3', 'Kiev'],
    'Вивчаємо Біблію разом': ['UKR_SBT_{}.mp3', 'Kharkov'],
    'ВЦП': ['UKR_PRC_{}.mp3', 'Kiev'],
    'Герои': ['RUS_CA_{}.mp3', 'Kharkov'],
    'Голос друга': ['BEL_VFR_{}.mp3', 'Kiev'],
    'Джерельце': ['UKR_TLS_{}.mp3', 'Kiev'],
    'ЖКОЕ': ['RUS_LAI_{}.mp3', 'Kharkov'],
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
    'Табор': ['RUS_RCMO_{}.mp3', 'Kharkov'],
    'Тихие воды': ['RUS_SWA_{}.mp3', 'Kiev'],
    'Хлеб жизни': ['RUS_BLR_{}.mp3', 'Kiev'],
    'Шалом': ['RUS_SHA_{}.mp3', 'Kharkov'],
    'Шанс // ГВЛ': ['UKR_MAE_{}.mp3', 'Kiev'],
}


# MAIN CODE

def send_email_report(subject, body_text):
    """Отправка письма в случае ошибки создания плейлиста"""
    config_path = CONFIG_DIR / "email.ini"

    if config_path.exists():
        cfg = ConfigParser()
        cfg.read(config_path)
    else:
        print("Файл конфигурации email.ini не найден!")
        raise SystemExit

    host = cfg.get("smtp", "host")
    from_addr = cfg.get("smtp", "from_addr")
    password = cfg.get("smtp", "password")
    to_emails = cfg.get("smtp", "to_emails")

    msg = MIMEMultipart()
    msg['From'] = from_addr
    msg['To'] = to_emails
    msg['Subject'] = Header(subject)

    msg.attach(MIMEText(body_text, 'plain', 'cp1251'))

    server = smtplib.SMTP_SSL(host, 465)
    server.ehlo()
    server.login(user=from_addr, password=password)
    server.sendmail(from_addr, [to_emails], msg.as_string())
    server.quit()


def get_excel_info(file_name, excel_page_name):
    """Возвращает лист с книги Excel"""
    if file_name.exists():
        workbook = load_workbook(file_name, data_only=True)
        sheet = workbook[excel_page_name]
        return sheet

    mistake_text = f"Excel файл \n---\n{file_name}\n---\nНе найден\nПожалуйста, проверьте файл в папке 'INTERNET RADIO'"
    send_email_report(MISTAKE_SUBJECT, mistake_text)
    print(f'Excel файл \n{file_name}\nне найден')
    raise SystemExit


def get_muzblock_with_needed_length(row, total_playing_tracks_time):
    """Выбрать самый подходящий музблок"""
    hours, minutes = (23, 59) if row[2] == datetime.time(0, 0) else (row[2].hour, row[2].minute)
    track_end_time = datetime.timedelta(hours=hours, minutes=minutes)
    current_playing_track_time = datetime.timedelta(seconds=total_playing_tracks_time)
    muzblock_needed_length = int((track_end_time - current_playing_track_time).total_seconds())
    track = min(MUZBLOCKS, key=lambda x: abs(x - muzblock_needed_length))
    return MUZBLOCKS[track]


def extract_excel_data(row, total_playing_tracks_time):
    if row[5] == 'муз.блок':
        """Получаем музблок"""
        file_name = 'Muzblock'
        mp3_file_name = get_muzblock_with_needed_length(row, total_playing_tracks_time)
        full_mp3_file_path = MEDIA_DIR / mp3_file_name

    elif row[5] == 'ГОДИНА БОЖОГО СЛОВА':
        """Конкретный случай для передачи Година Божого Слова"""
        file_number = row[4]
        file_name = 'Online radio blok'
        mp3_file_name = f'{file_name} {file_number}.mp3'
        full_mp3_file_path = MEDIA_DIR / mp3_file_name

    elif row[5] == 'ДО (15)':
        """Конкретный случай для передачи Домашний очаг 15 минут"""
        file_number = row[4]
        file_name = row[3]
        mp3_file_name = f'{file_name} {file_number}.mp3'
        # full_mp3_file_path = os.path.join(DO_15_MIN_DIR, mp3_file_name)
        full_mp3_file_path = DO_15_MIN_DIR / mp3_file_name

    elif 30 > row[0] >= 26:
        """Повтор за вчера"""
        date = CUR_DAY.strftime('%Y%m%d')
        file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)
        mp3_file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)

        if MAIN_AUDIO_FILES[row[5]][1] == 'Kiev':
            file_dir = KIEV_ST_DIR_TODAY
        else:
            file_dir = KHAR_ST_DIR_TODAY
        full_mp3_file_path = file_dir / mp3_file_name

    elif 63 > row[0] >= 59:
        """Прямой эфир"""
        date = NEXT_DAY.strftime('%Y%m%d')
        file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)
        mp3_file_name = MAIN_AUDIO_FILES[row[5]][0].format(date)

        if MAIN_AUDIO_FILES[row[5]][1] == 'Kiev':
            file_dir = KIEV_ST_DIR_TOMORROW
        else:
            file_dir = KHAR_ST_DIR_TOMORROW
        full_mp3_file_path = file_dir / mp3_file_name

    else:
        """Все остальные случаи, где файл из папки Archive_2018"""
        if 'Лекция' in str(row[4]):
            file_number = str(row[4]).replace('Лекция', 'L')
        elif 'М.В.' in str(row[4]):
            file_number = str(row[4]).replace('М.В.', 'M')
        else:
            file_number = row[4]

        file_name = row[3]
        mp3_file_name = f'{file_name} {file_number}.mp3'.replace('  ', ' ')
        full_mp3_file_path = MEDIA_DIR / mp3_file_name

    return file_name, full_mp3_file_path


def get_mp3_file_length(full_path_to_file):
    """Возвращает длинну MP3 трека в секундах"""
    if full_path_to_file.exists():
        mp3_data = MP3(full_path_to_file)
        return int(mp3_data.info.length)

    mistake_text = f"MP3 файл \n---\n{full_path_to_file}\n---\nНе найден\nПроверьте наличие файла в папке"
    send_email_report(MISTAKE_SUBJECT, mistake_text)
    print(f'MP3 файл \n{full_path_to_file}\nне найден')
    raise SystemExit


def write_playlist_to_file(playlist_path, file_data):
    """Записать плейлист в файл"""
    with open(playlist_path, 'w') as write_file:
        write_file.writelines(file_data)
    print(f'Плейлист на {NEXT_DAY.strftime("%d.%m.%Y")} готов и находится в папке \n{PLAYLIST_DIR}')


def main():
    playlist_data = ['#EXTM3U\n']
    sheet = get_excel_info(EXCEL_FILE_PATH, EXCEL_PAGE_NAME)
    total_playing_tracks_time = 0

    for row in sheet.iter_rows(min_row=4, max_row=69, max_col=6, values_only=True):

        file_name, full_mp3_file_path = extract_excel_data(row, total_playing_tracks_time)
        mp3_file_length = get_mp3_file_length(full_mp3_file_path)

        total_playing_tracks_time += mp3_file_length
        playlist_data.append(f'#EXTINF:{mp3_file_length},{file_name}\n{full_mp3_file_path}\n')

    playlist_data.append(f'load {PLAYLIST_DIR / NEXT_PLAYLIST_NAME}.command')

    write_playlist_to_file(FULL_PLAYLIST_PATH, playlist_data)


if __name__ == '__main__':
    main()
