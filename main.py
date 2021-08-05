# -*- coding:UTF-8 -*-
# auth：NeWolf
# date 20210727

import os

import xlwt as xlwt

FILE_PATH = 'files'
TEXT_FILE_END = '.info'
EXCEL_FILE_END = '.xls'


def log(param):
    print(param)


def check_text_file():
    if not os.path.exists(FILE_PATH):
        os.mkdir(FILE_PATH)

    listdir = os.listdir(FILE_PATH)
    if len(listdir) == 0:
        log('dir has no file')
        return None

    for index in range(len(listdir)):
        item_file = listdir[index]
        if item_file.endswith(TEXT_FILE_END):
            return item_file


def write_excel(input_text):
    if input_text is None:
        log('input_text is None')
        return
    if not input_text.readable():
        log('input_text can\'t read')

    log(input_text.name.strip())
    text_file_name = input_text.name.strip()
    split_file_name = text_file_name.split('.')
    result_file_name = split_file_name[0] + EXCEL_FILE_END
    result = xlwt.Workbook()
    result_sheet = result.add_sheet('msg')
    result_cols = 0

    lines = input_text.readlines()
    for i in range(len(lines)):
        line = lines[i]
        # line = line.replace('\n', '')
        line_split = line.split('-|')
        if len(line_split) < 5:
            continue

        # if len(line_split) > 6:
        #     log(line_split)
        #     # continue

        # log(line_split)

        for row in range(len(line_split)):
            result_sheet.write(result_cols, row, line_split[row])
        result_cols += 1

    result.save(result_file_name)
    log('suss')


def text2excel():
    text_file = check_text_file()

    if text_file is None:
        log('get text file fail')
        return

    input_text = open(os.path.join(FILE_PATH, text_file), 'r')

    write_excel(input_text)


if __name__ == '__main__':
    text2excel()
