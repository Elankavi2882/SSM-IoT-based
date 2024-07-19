import os 
import cv2 
import face_recognition 
from datetime import datetime, timedelta 
import xlsxwriter 
import console 
known_persons = {} 
known_persons_folder = { 
"Elankavi": "elank", 
"Ramanathan": "raam", 
"Kishor": "kishor", 
"Manikandan": "manii", 
"Vasanth": "vasanth", 
"Magesh": "magez", 
"Haresh": "Harez", 
"Esakki Pandian": "Ezakki", 
"Nambeeshwaran": "Nambee", 
"Gowtham Durai": "Gowtham", 
"Surya Prakash": "Soori", 
"Prakash": "Prakaas", 
"Manikandan": "Mani", 
"Mohammad Asif": "MDAsif", 
"Jashwanth": "Jazz", 
"Hrushikesh": "rishi", 
"Aravindhan": "aravindh", 
"Balaji": "Bala", 
"Saaruhasan": "saaru", 
"Sanidhya Vardhan Singh": "singh", 
"Pranav R A": "gamer", 
"Pranav Nandhakumar": "Pranavv", 
"Ashik Hussain": "Ashiq", 
"Shri Hari": " srihari", 
"Nandhakumar": "nandha", 
"Akhil": "akhil", 
"Vignesh": "vicky", 
"Vishal": "viz", 
"Adhithyan": "aadhi", 
"Adhithya Jaiswal": "adhithya", 
"Karan": "mondrii", 
"Divya Srinath": "DS", 
"Naveen": "naveen", 
} 
for person, folder_path in known_persons_folder.items(): 
 image_paths = [ 
  os.path.join(folder_path, file) 
  for file in os.listdir(folder_path) 
  if file.lower().endswith(".jpg") 
] 
known_persons[person] = [] 
for image_path in image_paths: 
 face_encodings = face_recognition.face_encodings( 
face_recognition.load_image_file(image_path) 
) 
if face_encodings: 
 known_persons[person].append(face_encodings[0] 
video_capture = cv2.VideoCapture(0) 
wait_time_seconds = 10  # Set the time to wait in seconds 
wait_start_time = None
workbook = xlsxwriter.Workbook("output.xlsx") 
worksheet = workbook.add_worksheet() 
hist = [] 
while True:   
ret, frame = video_capture.read() 
face_locations = face_recognition.face_locations(frame) 
face_encodings = face_recognition.face_encodings(frame, face_locations) 
for (top, right, bottom, left), face_encoding in zip( 
face_locations, face_encodings 
): 
matched_person = ( 
None  # Initialize matched_person here to handle cases when no faces match 
) 
for name, known_encodings in known_persons.items(): 
for known_encoding in known_encodings: 
match = face_recognition.compare_faces([known_encoding], face_encoding) 
if ( 
match and match[0] 
):  # Check if match is not empty and the first element is True 
matched_person = name 
break 
if matched_person: 
break 
if matched_person: 
if wait_start_time is None or matched_person != current_person: 
wait_start_time = datetime.now() 
current_person = matched_person 
detection_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S") 
status = "Present" 
hist.append([matched_person, detection_time, status]) 
print(f"{matched_person} present at {detection_time}") 
else: 
else: 
elapsed_time = datetime.now() - wait_start_time 
if elapsed_time.total_seconds() >= wait_time_seconds: 
print( 
f"{matched_person} detected. Waiting time reached. Closing camera." 
) 
video_capture.release() 
cv2.destroyAllWindows() 
exit() 
wait_start_time = None 
cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2) 
font = cv2.FONT_HERSHEY_DUPLEX 
cv2.putText( 
frame, 
matched_person or "Unknown", 
(left + 6, bottom - 6), 
font, 
0.5, 
(255, 255, 255), 
1, 
) 
cv2.imshow("Video", frame) 
# Break the loop if 'q' key is pressed 
if cv2.waitKey(1) & 0xFF == ord("q"): 
break 
video_capture.release() 
cv2.destroyAllWindows() 
worksheet.write(0, 0, "Matched Person") 
worksheet.write(0, 1, "Detected Time") 
worksheet.write(0, 2, "Status") 
row = 1 
col = 0 
hist = tuple(hist) 
for mp, t, s in hist: 
worksheet.write(row, col, mp) 
worksheet.write(row, col + 1, t) 
worksheet.write(row, col + 2, s) 
row += 1 
workbook.close()                