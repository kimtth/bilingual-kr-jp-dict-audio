import os
import main
import time
from os import listdir
from os.path import isfile, join


def dir_convert_slow_down_audio():
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


def df_to_write_result_to_html():
    base_dir = os.path.abspath(".")
    dir_path = os.path.join(base_dir, 'ret')
    file_path = os.path.join(dir_path, 'conv_20210430-212311-edit.xlsx')
    rtn_df = main.read_excel(file_path)
    time_str = time.strftime("%Y%m%d-%H%M%S")
    main.write_result_to_html(dir_path, 'conv_' + time_str, rtn_df)


if __name__ == '__main__':
    df_to_write_result_to_html()

