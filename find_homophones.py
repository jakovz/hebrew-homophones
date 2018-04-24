# -*- coding: UTF-8 -*-
import openpyxl as xl
from os import path
from config import *
from itertools import combinations_with_replacement
import progressbar


# we will first ignore words that are written in the same way (analyzing their meaning is way harder).

def add_all_permutations(row, new_sheet, counter, all_char_combinations):
    current_word = row[0].value.encode("UTF-8")
    for combination in progressbar.progressbar(all_char_combinations):
        for char in combination:
            if char in current_word:
                for equivalent_char in HOMOPHONE_HEBREW_CHARS[char]:
                    counter += 1
                    new_sheet["A%d" % counter] = current_word.replace(char, equivalent_char, 1)

                    # note that *all* the possible combinations for every possible replacement are being processed.


def create_all_permutations_excel():
    """
    for efficiency reasons we will first create a worksheet which will contain all the possible permutations for each
    word. then, we will be able to iterate this worksheet and decide for each word in it if it is a legit word in the
    language (by using the original sheet).

    :return: saves a new worksheet.
    """
    excel_file = xl.load_workbook(path.join(EXCEL_FILE_PATH, EXCEL_FILE_NAME), read_only=True)
    excel_sheet = excel_file["sheet1"]
    new_excel_file = xl.Workbook()
    new_sheet = new_excel_file.create_sheet()
    counter = 0
    all_chars_combinations = [combination for combination in
                              combinations_with_replacement(HOMOPHONE_HEBREW_CHARS, MAX_WORD_SIZE)]
    print len(all_chars_combinations)
    rows = excel_sheet['{0}:{1}'.format(excel_sheet.min_row, SEARCH_MAX_ROW)]
    for row in rows:
        counter += 1
        add_all_permutations(row, new_sheet, counter, all_chars_combinations)
    new_excel_file.save(path.join(EXCEL_FILE_PATH, NEW_EXCEL_FILE_NAME))


if __name__ == '__main__':
    create_all_permutations_excel()
