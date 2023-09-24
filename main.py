from win32com.client import Dispatch
from ultralytics import YOLO

# Initialize the speech engine
speak = Dispatch("SAPI.SpVoice")

# Load the model
model = YOLO("yolov8n.pt")

# Perform object detection
results = model.predict(show=True, source=1)

# Initialize an empty list to store detected object names
item_detect = []

# Iterate over each result
for res in results:
    for obj in res.boxes:
        # Append the class name instead of class index to the list
        item_detect.append(model.names[obj.cls])

# Speak out each detected object
for item in item_detect:
    speak.Speak("Warning")
    speak.Speak(item)
