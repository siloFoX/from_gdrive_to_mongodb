from __future__ import print_function
import pickle
import os.path
import io
import pandas as pd
from pymongo import MongoClient
import csv
import openpyxl

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

from store_image import store_image

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

# 저장된 권한 정보 가져오는 것
def auth():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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
    #request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    #request = service.files().export_media(fileId=file_id,
    #                                       mimeType='text/csv')

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

    #여기부터

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

            record_path, num_data = store_image(cwd, item['name'])
            path_idx = 0

            xlsx = pd.ExcelFile(cwd)
            for sheet in xlsx.sheet_names :
                PC = xlsx.parse(sheet)
                is_exist = False

                for column in PC.columns :
                    if column == '사진' :
                        is_exist = True

                if is_exist :

                    for idx in range(num_data - 1) :

                        if type(PC['사진'][idx]) is float :

                            PC['사진'][idx] = os.path.join(os.getcwd(), record_path[path_idx])
                            path_idx += 1

                PC.to_csv(item['name'] + '_' + sheet + '.csv', encoding='utf-8', index=False)

            os.remove(cwd)


    # << 몽고디비 연결 >>
    client = MongoClient('localhost', 27017)
    # localhost: ip주소
    # 27017 : port 번호

    # db 객체 할당
    db = client['SmartProcess']

    #collection 객체 할당받기
    Sample_preparation = db.Sample_preparation
    Washing1=db.Washing1
    Washing2=db.Washing2
    Washing3=db.Washing3
    Cleaning_blowing=db.Cleaning_blowing
    Pre_sputtering=db.Pre_sputtering
    Sputtering=db.Sputtering
    Heat_treatment=db.Heat_treatment
    PR_hand_blowing=db.PR_hand_blowing
    HMDS_coating=db.HMDS_coating
    PR_coating=db.PR_coating
    PR_baking=db.PR_baking
    PR_cooling=db.PR_cooling
    Chlorobenzene_treatment=db.Chlorobenzene_treatment
    DI_cleaning=db.DI_cleaning
    Blowing_after_chlorobenzene_cleaning=db.Blowing_after_chlorobenzene_cleaning
    Stepper_exposure=db.Stepper_exposure
    Develop=db.Develop
    Cleaning_after_develop=db.Cleaning_after_develop
    Blowing_after_develop_cleaning=db.Blowing_after_develop_cleaning
    Sample_preparation_for_evaporation=db.Sample_preparation_for_evaporation
    evaporation=db.evaporation
    Strip_soaking=db.Strip_soaking
    Strip_spreading=db.Strip_spreading
    Oxide_removal=db.Oxide_removal
    Applying_silver_paste=db.Applying_silver_paste
    Dry_silver_paste=db.Dry_silver_paste
    Measure=db.Measure

    folder = os.getcwd()
    for filename in os.listdir(folder):
        fullname = os.path.join(folder, filename)
        if fullname.find('.csv') is not -1:
            csvfile = open(fullname, 'rt', encoding='utf-8')
            reader = csv.DictReader(csvfile)
            if filename.find('샘플준비.csv') is not -1:
                insertData(reader, Sample_preparation)
            elif filename.find('세정1.csv') is not -1:
                insertData(reader, Washing1)
            elif filename.find('세정2.csv') is not -1:
                insertData(reader, Washing2)
            elif filename.find('세정3.csv') is not -1:
                insertData(reader, Washing3)
            elif filename.find('세정블로윙.csv') is not -1:
                insertData(reader, Cleaning_blowing)
            elif filename.find('프리스퍼터링.csv') is not -1:
                insertData(reader, Pre_sputtering)
            elif filename.find('스퍼터링.csv') is not -1:
                insertData(reader, Sputtering)
            elif filename.find('열처리.csv') is not -1:
                insertData(reader, Heat_treatment)
            elif filename.find('PR핸드블로윙.csv') is not -1:
                insertData(reader, PR_hand_blowing)
            elif filename.find('HMDS코팅.csv') is not -1:
                insertData(reader, HMDS_coating)
            elif filename.find('PR코팅.csv') is not -1:
                insertData(reader, PR_coating)
            elif filename.find('PR베이킹.csv') is not -1:
                insertData(reader, PR_baking)
            elif filename.find('PR쿨링.csv') is not -1:
                insertData(reader, PR_cooling)
            elif filename.find('클로로벤젠처리.csv') is not -1:
                insertData(reader, Chlorobenzene_treatment)
            elif filename.find('클로로벤젠세정후블로윙.csv') is not -1:
                insertData(reader, Blowing_after_chlorobenzene_cleaning)
            elif filename.find('노광.csv') is not -1:
                insertData(reader, Stepper_exposure)
            elif filename.find('현상.csv') is not -1:
                insertData(reader, Develop)
            elif filename.find('현상후세정.csv') is not -1:
                insertData(reader, Cleaning_after_develop)
            elif filename.find('현상액세정후블로윙.csv') is not -1:
                insertData(reader, Blowing_after_develop_cleaning)
            elif filename.find('이베포레이션샘플거치.csv') is not -1:
                insertData(reader, Sample_preparation_for_evaporation)
            elif filename.find('이베포레이션.csv') is not -1:
                insertData(reader, evaporation)
            elif filename.find('스트립-담그기.csv') is not -1:
                insertData(reader, Strip_soaking)
            elif filename.find('스트립-뿌리기.csv') is not -1:
                insertData(reader, Strip_spreading)
            elif filename.find('산화막제거.csv') is not -1:
                insertData(reader, Oxide_removal)
            elif filename.find('실버페이스트도포.csv') is not -1:
                insertData(reader, Applying_silver_paste)
            elif filename.find('실버페이스트건조.csv') is not -1:
                insertData(reader, Dry_silver_paste)
            elif filename.find('측정.csv') is not -1:
                insertData(reader, Measure)
            csvfile.close()
            os.remove(fullname)





if __name__ == '__main__':
    main()


