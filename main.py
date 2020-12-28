from PIL import ImageGrab
from datetime import datetime
import pyautogui
import pandas as pd
import face_recognition
import cv2
import numpy as np
import os
import keyboard
import argparse

def create_encodings(images):
    encode_list = []
    for image in images:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(image)[0]
        encode_list.append(encode)
    return encode_list

def screen_capture():
    capture = np.array(ImageGrab.grab())
    capture= cv2.cvtColor(capture, cv2.COLOR_RGB2BGR)
    return capture

def check_attendance(name, frame_idx):
    now = datetime.now()
    timestamp = now.strftime("%Y/%m/%d %H:%M:%S")
    if name in attendance_time:
        if len(attendance_time[name]) < 100:
            attendance_time[name].append((frame_idx, timestamp))
    else:
        attendance_time[name] = [(frame_idx, timestamp)]


if __name__ == "__main__":
    '''
    Sort attendance list by name by default
    Optional:
    - sort by student ID (keyword=sort_by_id)
    - sort by attendance time (keyword=sort_by_time)
    
    Example of student image: 응웬민뚜_16101384.jpg
    '''

    parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('--path', help='학생들의 사진 폴더', default='student_images')
    parser.add_argument('--course_id', help='과목코드', required=True)
    parser.add_argument('--lecture_id', help='강좌번호', required=True)
    parser.add_argument('--sort_by_id', help='학생 번호별로 출석 리스트를 정렬', action='store_true', default=False)
    parser.add_argument('--sort_by_time', help='학생의 출석 시간별로 출석 리스트를 정렬', action='store_true', default=False)
    parser.add_argument('--reverse', help='출석 리스트의 역순으로 정렬', action='store_true', default=False)
    parser.add_argument('--ofile', help='출력 파일의 이름', default=None)
    args = parser.parse_args()

    path = args.path
    course_id = args.course_id
    lecture_id = args.lecture_id
    sort_by_name = False
    if not args.sort_by_id and not args.sort_by_time:
        sort_by_name = True
   
    try:
        df = pd.read_excel("courses_database.xlsx", converters={"과목코드": str, "강좌번호": str})
    except OSError as err:
        print("OS error: {0}".format(err))
        exit(0)

    course_name, prof_name = "NaN", "NaN"
    for index, row in df.iterrows():
        if row["과목코트"] == course_id and row["강좌번호"] == lecture_id:
            course_name = row["과목명"]
            prof_name = row["교수명"]

    if course_name == "NaN" or prof_name == "NaN":
        print("Error: Cannot find course ID " + course_id + " with lecture ID " + lecture_id)
        exit(0)

    #print(course_name, prof_name)
    print("RECORDING ATTENDANCE...")
    print("PLEASE MAKE SURE THAT THE SCREEN DISPLAYS EACH OF THE STUDENTS")

    images = []
    stud_names, stud_id = [], []
    attendance_time = {}
    img_folder = os.listdir(path)
    for file in img_folder:
        # read Hangul
        image = cv2.imdecode(np.fromfile(f'{path}/{file}', dtype=np.uint8), cv2.IMREAD_UNCHANGED)
        images.append(image)
        full_name = os.path.splitext(file)[0].split('_')
        name, id = full_name[0], full_name[1]
        stud_names.append(name)
        stud_id.append(id)
    known_students = create_encodings(images)


    frame_idx = 0
    while not keyboard.is_pressed('esc'):
        screen = screen_capture()
        img = cv2.cvtColor(screen, cv2.COLOR_BGR2RGB)

        faces_locations = face_recognition.face_locations(img)
        encodings = face_recognition.face_encodings(img, faces_locations)

        for encode_face, face_loc in zip(encodings, faces_locations):
            matches = face_recognition.compare_faces(known_students, encode_face)
            face_dist = face_recognition.face_distance(known_students, encode_face)
            match_idx = np.argmin(face_dist)

            # find person with most similar face (lowest face_dist)
            if face_dist[match_idx] < 0.45:
                name = stud_names[match_idx]
                check_attendance(name, frame_idx)
            else:
                name = 'Unknown'
            print(frame_idx, name, face_dist[match_idx])

            #y1, x2, y2, x1 = face_loc
            #y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
            #cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
            #cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
            #cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

        #for Webcam
        #cv2.imshow('screen', img)
        #if cv2.waitKey(1) & 0xFF == ord('q'):
        #    break

        frame_idx += 1

    # export attendance list to csv
    col_time = []
    attendance_check = []
    for idx in range(len(stud_names)):
        name, id = stud_names[idx], stud_id[idx]
        if name in attendance_time:
            attendance_check.append("Y")
            col_time.append(attendance_time[name][0][1])
        else:
            attendance_check.append("")
            col_time.append("NaN")


    if args.sort_by_id:
        for i in range(len(stud_names)):
            for j in range(i + 1, len(stud_names)):
                if stud_id[i] > stud_id[j]:
                    stud_id[i], stud_id[j] = stud_id[j], stud_id[i]
                    stud_names[i], stud_names[j] = stud_names[j], stud_names[i]
                    col_time[i], col_time[j] = col_time[j], col_time[i]
    elif args.sort_by_time:
        for i in range(len(stud_id)):
            for j in range(i + 1, len(stud_id)):
                if col_time[i] > col_time[j]:
                    stud_id[i], stud_id[j] = stud_id[j], stud_id[i]
                    stud_names[i], stud_names[j] = stud_names[j], stud_names[i]
                    col_time[i], col_time[j] = col_time[j], col_time[i]
    elif sort_by_name:
        for i in range(len(stud_id)):
            for j in range(i + 1, len(stud_id)):
                if stud_names[i] > stud_names[j]:
                    stud_id[i], stud_id[j] = stud_id[j], stud_id[i]
                    stud_names[i], stud_names[j] = stud_names[j], stud_names[i]
                    col_time[i], col_time[j] = col_time[j], col_time[i]



    if args.reverse:
        stud_names = stud_names.reverse()
        stud_id = stud_id.reverse()
        col_time = col_time.reverse()

    attendance_data = pd.DataFrame(list(zip(stud_id, stud_names, attendance_check, col_time)), columns=['학번', '이름', '출석', '출석시간'])
    file_name = args.ofile
    if file_name == None:
        file_name = course_id + "_" + lecture_id + ".xlsx"
    attendance_data.to_excel(file_name, index=False)

    print("Saved attendance list as " + file_name)
