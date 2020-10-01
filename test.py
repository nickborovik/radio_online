import re
import datetime as dt
from pathlib import Path

BASE_DIR = Path('D:/') / 'Playlist Radioboss'
# BASE_DIR = Path('.')
files = list(BASE_DIR.glob('*.m3u8'))
for file_name in files:
    file_length = 0
    with open(file_name, 'r') as read_file:
        for line in read_file.readlines():
            if '#EXTINF' in line:
                file_length += int(re.search(r'\d+', line).group())

    print(file_name, '=', dt.timedelta(seconds=file_length))
