import face_recognition
import cv2
import os
import pandas as pd
import time

# Path to folder containing voter images
voter_images_path = "voters"

# Full path to Excel sheet
excel_path = r"C:\Users\milua\Desktop\ElectionDemo\voters.xlsx"

# Load Excel sheet
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Excel file not found at {excel_path}")

df = pd.read_excel(excel_path)  # Columns: Voter Name | Image | Has Voted

# Dictionary to store voter info
voter_data = {}  # {name: {"encoding": ..., "voted": False}}

# Load voter images and create encodings
for index, row in df.iterrows():
    voter_name = row['Voter Name']
    img_filename = row['Image']
    voted = True if row['Has Voted'] == "Yes" else False
    path = os.path.join(voter_images_path, img_filename)

    if os.path.exists(path):
        image = face_recognition.load_image_file(path)
        encoding = face_recognition.face_encodings(image)
        if encoding:
            voter_data[voter_name] = {"encoding": encoding[0], "voted": voted}
            print(f"✅ Loaded {voter_name}")
        else:
            print(f"⚠ Could not encode {voter_name}")
    else:
        print(f"⚠ Image not found for {voter_name}")

# Dictionary to track start time of voting for each voter
voting_start_time = {}

# Start video capture
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        break

    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

    face_locations = face_recognition.face_locations(rgb_small_frame)
    face_encodings = []
    for loc in face_locations:
        enc = face_recognition.face_encodings(rgb_small_frame, [loc])
        if enc:
            face_encodings.append(enc[0])

    for face_encoding, face_location in zip(face_encodings, face_locations):
        voter_name = "Unknown"
        for name, data in voter_data.items():
            match = face_recognition.compare_faces([data["encoding"]], face_encoding, tolerance=0.5)
            if match[0]:
                voter_name = name
                break

        top, right, bottom, left = face_location
        top *= 4
        right *= 4
        bottom *= 4
        left *= 4

        if voter_name != "Unknown":
            if voter_data[voter_name]["voted"]:
                label = f"{voter_name} - Already Voted"
                color = (0, 0, 255)
            else:
                # First-time voter
                color = (0, 255, 0)
                if voter_name not in voting_start_time:
                    voting_start_time[voter_name] = time.time()  # start timer
                elapsed = time.time() - voting_start_time[voter_name]
                if elapsed < 8:  # keep green live for 8 seconds
                    label = f"{voter_name} - Registered, You Can Vote"
                else:
                    # After 8 seconds, mark as voted
                    voter_data[voter_name]["voted"] = True
                    df.loc[df['Voter Name'] == voter_name, 'Has Voted'] = "Yes"
                    df.to_excel(excel_path, index=False)
                    label = f"{voter_name} - Already Voted"
                    color = (0, 0, 255)

        else:
            label = voter_name
            color = (0, 0, 255)

        cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), color, cv2.FILLED)
        cv2.putText(frame, label, (left + 6, bottom - 6),
                    cv2.FONT_HERSHEY_DUPLEX, 0.8, (255, 255, 255), 1)

    cv2.imshow('Voting Verification', frame)
    
    # Press 'q' to quit manually anytime
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
