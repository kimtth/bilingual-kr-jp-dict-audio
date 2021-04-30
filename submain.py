import os
import main
from os import listdir
from os.path import isfile, join


if __name__ == '__main__':
    base_dir = os.path.abspath(".")
    save_dir = os.path.join(base_dir, 'ret')
    save_audio_dir = os.path.join(save_dir, 'mp3')
    save_audio_slow_dir = os.path.join(save_dir, 'mp3_s')

    only_files = [f for f in listdir(save_audio_dir) if isfile(join(save_audio_dir, f))]
    for file in only_files:
        dir, filename = os.path.split(file)
        file_name_only = filename.split('.')[0]
        print(filename)
        main.convert_slow_down_audio(save_audio_dir, save_audio_slow_dir, file_name_only, 0.7)

