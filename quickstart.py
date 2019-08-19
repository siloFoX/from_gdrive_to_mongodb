from __future__ import print_function
import pickle
import os.path
import io
import pandas as pd
from pymongo import MongoClient
import csv
import openpyxl
import json

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

#from store_image import store_image

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# 저장된 권한 정보 가져오는 것
def auth():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

# 엑셀 파일 다운로드
def downloadFile(file_id, filepath):
    service = build('drive', 'v3', credentials=auth())
    request = service.files().get_media(fileId = file_id) # I fixed this point
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    with io.open(filepath,'wb') as f:
        fh.seek(0)
        f.write(fh.read())

# 몽고DB 입력
def insertData(reader, collection):
    for row in reader:
        list = {}
        for fieldname in reader.fieldnames:
            list[fieldname]=row[fieldname]
        collection.update_one({'실험날짜':list.get('실험날짜'), '실험자명':list.get('실험자명'), '해당실험기판번호':list.get('해당실험기판번호')},
                              {'$set' : list},
                              upsert=True) 


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """

    service = build('drive', 'v3', credentials=auth())

    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))
            downloadFile(item['id'], item['name'])
            cwd = os.path.join(os.getcwd(), item['name'])

            xlsx = pd.ExcelFile(cwd)
            for sheet in xlsx.sheet_names :
                PC = xlsx.parse(sheet)
                PC.to_csv(sheet + '.csv', encoding='utf-8', index=False)

            os.remove(cwd)


    # << 몽고디비 연결 >>
    client = MongoClient('localhost', 27017)
    # localhost: ip주소
    # 27017 : port 번호

    # db 객체 할당
    db = client['SmartProcess']

    with open('store_image/collection_allocate.json', encoding = 'UTF-8') as f :
        collections = json.load(f)

    collection_list = []
    idx_dict = dict()

    #collection 객체 할당받기
    for idx, val in enumerate(collections.values()) :
        collection_list.append(db.get_collection(val))
        idx_dict[val] = idx

    folder = os.getcwd()

    for filename in os.listdir(folder) :

        fullname = os.path.join(folder, filename)

        if fullname.find('.csv') + 1 :

            csvfile = open(fullname, 'rt', encoding='utf-8')
            reader = csv.DictReader(csvfile)

            insertData(reader, collection_list[idx_dict[collections[filename.split('.')[0]]]])

            csvfile.close()
            os.remove(fullname)


if __name__ == '__main__':
    main()