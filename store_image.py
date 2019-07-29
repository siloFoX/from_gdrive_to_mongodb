import win32com.client as win32
from PIL import ImageGrab
import os

def store_image (path, name) :
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(path)

    # image store at each directory
    img_path = 'image'
    img_path = os.path.join(img_path, name)

    if not os.path.exists(img_path) :
        os.makedirs(img_path)

    record_path = []

    # sheet 는 하나의 페이지(ex : 세정1)
    for sheet in workbook.Worksheets :
        num_data = 1 # data 개수 + 1

        while sheet.Range(chr(65) + str(num_data)).Value != None :
            num_data += 1

        # 알파벳 'A' ~ 'Z' 까지만 검사
        for col in range(65, 91) :
            node = chr(col)

            # 사진 줄만 검사
            if sheet.Range(node + '1').Value == '사진' :
                
                # 2 ~ data 개수 + 1 까지 검사
                for row in range(2, num_data) :

                    # 아무 value도 없는 곳에서만 사진파일을 가져올 것임    
                    if sheet.Range(node + str(row)).Value == None :
                        sheet.Range(node + str(row)).Copy()
                        image = ImageGrab.grabclipboard()

                        name_tmp = name + '_' + sheet.Range('C1').Value + '_' + str(row) + '.jpg'
                        path_tmp = os.path.join(img_path, name_tmp)
                        image.save(path_tmp, 'jpeg')

                        record_path.append(path_tmp)

    excel.Quit()
    num_data -= 1

    return record_path, num_data           

def main () :
    path =  os.getcwd().replace('\'','\\') + '\\'
    record_path, num_data = store_image(path + '190701.xlsx', '190701')

    print(num_data)
    print(record_path)

if __name__ == '__main__' :
    main()