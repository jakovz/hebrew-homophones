# -*- coding: UTF-8 -*-
EXCEL_FILE_PATH = r"C:\Users\Jakov\Downloads"
EXCEL_FILE_NAME = r"hebrew_words.xlsx"
NEW_EXCEL_FILE_NAME = r"new_hebrew_words.xlsx"
MIN_OCCURRENCES = 3
SEARCH_MAX_ROW = 75356  # this max row indicates a word that appeared at least 19 times
MAX_WORD_SIZE = 5  # for efficiency purposes this number is suggested to be the maximum probable size of a word
# this are all the letters which sound the same in hebrew.
HOMOPHONE_HEBREW_CHARS = {
    "ב": ["ו"],
    "ח": ["כ"],
    "ט": ["ת"],
    "ק": ["כ"],
    "ס": ["ש"],
    "א": ["ע"],
    "ו": ["ב"],
    "כ": ["ק", "ח"],
    "ת": ["ט"],
    "ש": ["ס"],
    "ע": ["א"],

}
