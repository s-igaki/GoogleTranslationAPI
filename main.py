# Imports the Google Cloud client library
from oauth2client.client import GoogleCredentials
from google.cloud import translate
import os
import pandas as pd
import openpyxl

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "./json/TranslateApp-dfb1e94a1d3d.json"
ENG_EXCEL_FILE_PATH = './practice_eng.xlsx'
ENG_AND_JA_EXCEL_FILE_PATH = './practice_eng_and_ja.xlsx'
SHEET_NAME = 'sheet_1'

def read_excel(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def write_excel(df, file_path, sheet_name):
    df.to_excel(file_path, index=False,  sheet_name=sheet_name)

def translate_from_en_to_ja(text):
    credentials = GoogleCredentials.get_application_default()
    translate_client = translate.Client()
    json = translate_client.translate(text, source_language='en', target_language='ja', model='nmt')
    return json["translatedText"]

def translate_from_ja_to_en(text):
    credentials = GoogleCredentials.get_application_default()
    translate_client = translate.Client()
    json = translate_client.translate(text, source_language='ja', target_language='en', model='nmt')
    return json["translatedText"]

def main():
    df = read_excel(ENG_EXCEL_FILE_PATH, SHEET_NAME)
    df['Japanese'] = df['English'].map(translate_from_en_to_ja)
    write_excel(df, ENG_AND_JA_EXCEL_FILE_PATH, SHEET_NAME)

if __name__ == "__main__":
    main()
