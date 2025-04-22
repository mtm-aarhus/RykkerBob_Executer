import json
import os

DATA_FILE = "queue_data.json"

# Default structure if the file doesn't exist
default_data = {
    "out_ListOfProcessedItems": [],
    "out_ListOfErrorMessages": [],
    "out_ListOfFailedCases": []
}

def load_data():
    if not os.path.exists(DATA_FILE):
        return default_data.copy()
    with open(DATA_FILE, "r") as f:
        return json.load(f)

def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)