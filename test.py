from sklearn.neighbors import KNeighborsClassifier  # type: ignore
import cv2  # type: ignore
import pickle
import numpy as np  # type: ignore
import os
import csv
import time
from datetime import datetime

from win32com.client import Dispatch # type: ignore

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

# Use OpenCV's built-in Haar cascade file
cascade_path = cv2.data.haarcascades + 'haarcascade_frontalface_default.xml'
facesdetect = cv2.CascadeClassifier(cascade_path)

# Check if the cascade file was loaded correctly
if facesdetect.empty():
    raise FileNotFoundError(f"Could not load cascade classifier xml file from path: {cascade_path}")

# Load the labels and faces data
with open(r'data\names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open(r'data\faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

# Initialize the KNeighborsClassifier
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Correct the file path using raw string notation
background_path = r'C:\Face Recognition System\Data\background.png'
imgBackground = cv2.imread(background_path)

# Check if imgBackground was loaded correctly
if imgBackground is None:
    raise FileNotFoundError(f"Background image not found at the specified path: {background_path}")

# Open the video capture
video = cv2.VideoCapture(0)

COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    if not ret:
        print("Failed to capture frame from video source.")
        break

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facesdetect.detectMultiScale(gray, 1.3, 5)
    for (x, y, w, h) in faces:
        crop_img = frame[y:y + h, x:x + w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).reshape(1, -1)
        output = knn.predict(resized_img)
        ts=time.time()
        date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp=datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist=os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 1)

    # Check the size of imgBackground before attempting to assign the frame
    if imgBackground.shape[0] >= 162 + 480 and imgBackground.shape[1] >= 55 + 640:
        imgBackground[162:162 + 480, 55:55 + 640] = frame
    else:
        print("Background image is too small to hold the frame.")

    attendance=[str(output[0]), str(timestamp)]

    cv2.imshow("Frame", imgBackground)
    k=cv2.waitKey(1)
    if k==ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(attendance)
            csvfile.close()
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a") as csvfile:
                writer=csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
            csvfile.close()
    if k == ord('q'):
        break

# Release the video capture and close all OpenCV windows
video.release()
cv2.destroyAllWindows()
