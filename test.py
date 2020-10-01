import re
import datetime as dt
from pathlib import Path
from mutagen.mp3 import MP3

MUZBLOCKS = [
    '11 Kharkov time 9 min-a.mp3',
    '12 Kharkov time 9 min-a.mp3',
    '14 Kharkov time 10 min-a.mp3',
    '13 Kharkov time 10 min-a.mp3',
    '15 Kharkov time 10 min-a.mp3',
    '09 Kharkov time 11 min-a.mp3',
    '04 Kharkov time 11 min-a.mp3',
    '10 Kharkov time 11 min-a.mp3',
    '03 Kharkov time 11 min-a.mp3',
    '01 Kharkov time 11 min-a.mp3',
    '06 Kharkov time 11 min-a.mp3',
    '05 Kharkov time 11 min-a.mp3',
    '02 Kharkov time 11 min-a.mp3',
    '08 Kharkov time 11 min-a.mp3',
    '07 Kharkov time 11 min-a.mp3',
    'muzblok_01_time_12.15.mp3',
    'muzblok_18_time_12.40.mp3',
    'muzblok_12_time_12.40.mp3',
    'muzblok_05_time_13.20.mp3',
    'muzblok_03_time_13.42.mp3',
    'muzblok_08_time_13.55.mp3',
    'muzblok_15_time_13.58.mp3',
    'muzblok_11_time_14.02.mp3',
    'muzblok_24_time_14.16.mp3',
    'muzblok_06_time_14.19.mp3',
    'muzblok_09_time_14.20.mp3',
    'muzblok_07_time_14.21.mp3',
    'muzblok_13_time_14.27.mp3',
    'muzblok_10_time_14.33.mp3',
    'muzblok_14_time_14.41.mp3',
    'muzblok_02_time_14.43.mp3',
    'muzblok_17_time_14.46.mp3',
    'muzblok_16_time_14.49.mp3',
    'muzblok_23_time_14.55.mp3',
    'muzblok_19_time_14.55.mp3',
    'muzblok_04_time_14.57.mp3',
    'muzblok_20_time_14.59.mp3',
    'muzblok_22_time_15.33.mp3',
    'muzblok_21_time_15.43.mp3',
    'muzblok_26_time_16.16.mp3',
    'muzblok_25_time_16.18.mp3',
]

BASE_DIR = Path('D:/') / 'Playlist Radioboss'
ARCHIVE_DIR = Path('D:/') / 'INTERNET RADIO' / 'Archive_2018'
# BASE_DIR = Path('.')

def get_playlist_length():
    files = list(BASE_DIR.glob('*.m3u8'))
    for file_name in files:
        file_length = 0
        with open(file_name, 'r') as read_file:
            for line in read_file.readlines():
                if '#EXTINF' in line:
                    file_length += int(re.search(r'\d+', line).group())

        print(file_name, '=', dt.timedelta(seconds=file_length))

def get_muzblock_length():
    for muzblock in MUZBLOCKS:
        mp3 = MP3(ARCHIVE_DIR / muzblock)
        print(int(mp3.info.length))


if __name__ == '__main__':
    get_playlist_length()
    print('-----')
    get_muzblock_length()
