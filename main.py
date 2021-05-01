import sys
import traceback

import pandas as pd
import os
import time
from tkinter import messagebox, filedialog

from google_trans_new import google_translator
from gtts import gTTS
from pykakasi import kakasi
import ffmpy


def read_excel(excel_file_path):
    df = pd.read_excel(excel_file_path, engine='openpyxl')
    return df


def convert_japanese_to_hurigana(text):
    kakasi_ = kakasi()
    kakasi_.setMode('J', 'HF')
    conv = kakasi_.getConverter()
    return conv.do(text)


def convert_translation(text, source_lang='ko', target_lang='ja'):
    '''
    https://github.com/lushan88a/google_trans_new
    '''

    translator_ = google_translator()
    translate_text = translator_.translate(text, lang_src=source_lang, lang_tgt=target_lang)
    print(text, ' -> ', translate_text)
    return translate_text


def convert_g_tts_to_mp3(dir_path, file_name, text, lang='ko'):
    tts = gTTS(text, lang=lang)  # , slow=True)
    store_path = '{}\\{}{}'.format(dir_path, file_name, '.mp3')
    tts.save(store_path)


def convert_slow_down_audio(input_dir, output_dir, file_name, speed=1.0):
    # https://stackoverflow.com/questions/45441557/how-to-change-speed-of-a-wav-file-while-retaining-the-sampling-frequency-in-pyth
    sound_path = '{}\\{}{}'.format(input_dir, file_name, '.mp3')
    sound_out_path = '{}\\{}{}'.format(output_dir, file_name, '_s.mp3')
    ff = ffmpy.FFmpeg(global_options=['-y', '-loglevel panic'], inputs={sound_path: None},
                      outputs={sound_out_path: ["-filter:a", "atempo=" + str(speed)]})
    # print(ff.cmd)
    ff.run()


def write_result(dir_path, file_name, dataframe):
    excel_output_path = '{}\\{}{}'.format(dir_path, file_name, '.xlsx')
    print(excel_output_path)

    dataframe.to_excel(excel_output_path, index=False)


def write_result_to_html(dir_path, file_name, dataframe):
    copy_dataframe = dataframe.copy()
    # (0, '=HYPERLINK(".\\mp3\\1-1.mp3"," [ ▶ ] ")') (1, '=HYPERLINK(".\\mp3\\2-2.mp3"," [ ▶ ] ")')
    # <a href="url">link text</a>
    copy_dataframe['Play'] = copy_dataframe['Play'].str.replace('=HYPERLINK(".\\mp3\\',
                                                                '<a href=".\\mp3\\',
                                                                regex=False)
    copy_dataframe['Play'] = copy_dataframe['Play'].str.replace(', " [ ▶ ] ")', '> [ ▶ ] </a>', regex=False)
    copy_dataframe['Slow'] = copy_dataframe['Slow'].str.replace('=HYPERLINK(".\\mp3_s\\',
                                                                '<a href=".\\mp3_s\\',
                                                                regex=False)
    copy_dataframe['Slow'] = copy_dataframe['Slow'].str.replace(', " [ ▶ ] ")', '> [ ▶ ] </a>', regex=False)

    pd.set_option('colheader_justify', 'center')  # FOR TABLE <th>

    html_string = '''
    <html>
      <head><title>HTML Pandas Dataframe with CSS</title></head>
      <link rel="stylesheet" type="text/css" href="df_style.css"/>
      <body>
        {table}
      </body>
    </html>.
    '''

    # OUTPUT AN HTML FILE
    html_output_path = '{}\\{}{}'.format(dir_path, file_name, '.html')
    with open(html_output_path, 'w', encoding='utf-8') as f:
        f.write(html_string.format(table=copy_dataframe.to_html(classes='mystyle')))

    # replace &lt; and &gt;
    with open(html_output_path, 'r', encoding='utf-8') as file:
        filedata = file.read()

    # Replace the target string
    filedata = filedata.replace('&lt;', '<')
    filedata = filedata.replace('&gt;', '>')

    # Write the file out again
    with open(html_output_path, 'w', encoding='utf-8') as file:
        file.write(filedata)


def convert_processes(seq, text):
    converted_values = []
    translated_text = convert_translation(text)
    hurigana_text = convert_japanese_to_hurigana(translated_text)

    converted_values.append(translated_text)
    converted_values.append(hurigana_text)
    audio_path = '{}{}{}'.format(".\\mp3\\", seq, '.mp3')
    audio_slow_path = '{}{}{}'.format(".\\mp3_s\\", seq, '_s.mp3')

    # =HYPERLINK(CONCATENATE(".\mp3\", A2, ".mp3"), "[ ▶] ")
    converted_values.append('=HYPERLINK("' + audio_path + '", " [ ▶ ] ")')
    converted_values.append('=HYPERLINK("' + audio_slow_path + '", " [ ▶ ] ")')

    return converted_values


if __name__ == '__main__':
    file_path = filedialog.askopenfilename()
    print(file_path)

    if file_path:
        try:
            base_dir = os.path.abspath(".")
            save_dir = os.path.join(base_dir, 'ret')
            save_audio_dir = os.path.join(save_dir, 'mp3')
            save_audio_slow_dir = os.path.join(save_dir, 'mp3_s')
            if not os.path.exists(save_dir):
                os.mkdir(save_dir)
            if not os.path.exists(save_audio_dir):
                os.mkdir(save_audio_dir)
            if not os.path.exists(save_audio_slow_dir):
                os.mkdir(save_audio_slow_dir)

            convert_src_df = []
            df = read_excel(file_path)
            for i, row in df.iterrows():
                seq = row['No']
                text = row['Text']
                print(seq, '::')
                converted = convert_processes(seq, text)
                convert_src_df.append(converted)
                convert_g_tts_to_mp3(save_audio_dir, seq, text)
                convert_slow_down_audio(save_audio_dir, save_audio_slow_dir, seq, 0.7)

            converted_df = pd.DataFrame(convert_src_df, columns=['Translation', 'Hurigana', 'Play', 'Slow'])
            rtn_df = pd.concat([df, converted_df], axis=1)

            time_str = time.strftime("%Y%m%d-%H%M%S")
            write_result(save_dir, 'conv_' + time_str, rtn_df)
            write_result_to_html(save_dir, 'conv_' + time_str, rtn_df)
        except:
            print(traceback.format_exc())
            messagebox.showwarning("Warning",
                                   "Wrong.\n" + traceback.format_exc())
            sys.exit()
