from datetime import datetime
import spacy
import win32com.client
import re
from spacy.training import Example
from spacy.util import minibatch
import random

# Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY" (used during inference as fallback)
date_patterns = [
    r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
    r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
]

# Try to load the pre-trained model if it exists, otherwise start with a blank model
try:
    nlp = spacy.load("./trained_ner_model")  # Load the previously trained model
    print("Loaded existing model.")
except:
    nlp = spacy.blank("en")  # Start with a blank model if no model exists
    print("No existing model found, starting fresh.")

# Function to use the model to predict the entities in the email body
def get_model_predictions(email_body):
    doc = nlp(email_body)
    project_name = None
    contractor = None
    bid_due_date = None
    
    # Extract the model's predictions for each entity
    for ent in doc.ents:
        if ent.label_ == "PROJECT_NAME":
            project_name = ent.text
        elif ent.label_ == "CONTRACTOR":
            contractor = ent.text
        elif ent.label_ == "BID_DUE_DATE":
            bid_due_date = ent.text
    
    return project_name, contractor, bid_due_date

# Function to convert any date (model prediction or user input) to "MM/DD/YYYY"
def format_date_to_mmddyyyy(date_str):
    if date_str is None:
        return None  # If no date was found or input, return None
    
    try:
        # Try to parse dates in "Month DD, YYYY" format
        date_obj = datetime.strptime(date_str, "%B %d, %Y")
    except ValueError:
        try:
            # Try to parse dates in "MM/DD/YYYY" format
            date_obj = datetime.strptime(date_str, "%m/%d/%Y")
        except ValueError:
            return date_str  # If it can't be parsed, return the original string
    
    # Return the date in "MM/DD/YYYY" format
    return date_obj.strftime("%m/%d/%Y")

# Add the NER pipeline if it doesn't exist
if "ner" not in nlp.pipe_names:
    ner = nlp.add_pipe("ner")
else:
    ner = nlp.get_pipe("ner")  # Get the existing NER pipeline

# Add custom labels to the NER model (only if they're not already present)
ner.add_label("PROJECT_NAME")
ner.add_label("CONTRACTOR")
ner.add_label("BID_DUE_DATE")

# Connect to Outlook and get the namespace
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Function to recursively find folders containing the keyword "Bid"
def find_bid_folders(folder, folder_path=""):
    bid_folders = []
    folder_path = f"{folder_path}\\{folder.Name}" if folder_path else folder.Name
    
    # Check if the current folder's name contains "Bid"
    if "Bid" in folder.Name:
        bid_folders.append((folder_path, folder))
    
    # Recursively search subfolders
    for subfolder in folder.Folders:
        bid_folders.extend(find_bid_folders(subfolder, folder_path))
    
    return bid_folders

# Function to manually prompt for entity information from the user, showing model's predictions
def prompt_for_labels(email_body):
    print("\nEmail Body:\n", email_body)

    # Ask if the email is relevant
    relevant = input("Is this email relevant? (y/n): ").strip().lower()
    if relevant == 'n':
        return None  # Disregard this email

    # Get the model's predictions for this email
    model_project_name, model_contractor, model_bid_due_date = get_model_predictions(email_body)
    
    # Display model's predictions and allow user to confirm or correct
    print(f"\nModel's Prediction - Project Name: {model_project_name}")
    project_name = input(f"Enter the Project Name (or press Enter to confirm '{model_project_name}'): ")
    if not project_name:
        project_name = model_project_name  # Use model's prediction if no input is given
    
    print(f"Model's Prediction - Contractor: {model_contractor}")
    contractor = input(f"Enter the Contractor Name (or press Enter to confirm '{model_contractor}'): ")
    if not contractor:
        contractor = model_contractor  # Use model's prediction if no input is given
    
    print(f"Model's Prediction - Bid Due Date: {model_bid_due_date}")
    bid_due_date = input(f"Enter the Bid Due Date (format: MM/DD/YYYY) (or press Enter to confirm '{model_bid_due_date}'): ")
    if not bid_due_date:
        bid_due_date = model_bid_due_date  # Use model's prediction if no input is given

    # Convert the confirmed or corrected bid due date to "MM/DD/YYYY" (if not None)
    if bid_due_date is not None:
        bid_due_date = format_date_to_mmddyyyy(bid_due_date)

    # Identify entity positions in the email body for annotation
    entities = []
    if project_name:
        start_idx = email_body.lower().find(project_name.lower())
        if start_idx != -1:
            entities.append((start_idx, start_idx + len(project_name), "PROJECT_NAME"))

    if contractor:
        start_idx = email_body.lower().find(contractor.lower())
        if start_idx != -1:
            entities.append((start_idx, start_idx + len(contractor), "CONTRACTOR"))

    if bid_due_date:
        start_idx = email_body.lower().find(bid_due_date.lower())
        if start_idx != -1:
            entities.append((start_idx, start_idx + len(bid_due_date), "BID_DUE_DATE"))

    return entities

# Get emails from a selected folder
bid_folders = []
for folder in outlook.Folders:
    bid_folders.extend(find_bid_folders(folder))

if bid_folders:
    print("Found the following folders with 'Bid' in the name:")
    for i, (path, _) in enumerate(bid_folders):
        print(f"{i+1}. {path}")
    
    selection = int(input("Please enter the number of the folder you want to use: ")) - 1
    selected_folder = bid_folders[selection][1]

    # Retrieve emails from the selected folder
    messages = selected_folder.Items
    print(f"Total emails in the selected folder: {len(messages)}")
else:
    print("No folders with 'Bid' in the name were found.")
    messages = []

# Collect training data by using model predictions and prompting for corrections when needed
TRAIN_DATA = []
message_count = len(messages)
if message_count > 0:
    # Select 5 random emails that have all three entities predicted
    samples = 0
    while samples < 5:
        i = random.randint(0, message_count - 1)
        message = messages[i]
        email_body = message.Body

        # Get the model's predictions
        project_name, contractor, bid_due_date = get_model_predictions(email_body)

        if project_name and contractor and bid_due_date:  # Only take samples where all three entities are predicted
            entities = prompt_for_labels(email_body)
            if entities:
                TRAIN_DATA.append((email_body, {"entities": entities}))
                samples += 1

# Begin training if training data exists
if TRAIN_DATA:
    optimizer = nlp.resume_training()  # Resume training the existing model

    # Train the model for multiple iterations
    for iteration in range(30):
        random.shuffle(TRAIN_DATA)
        losses = {}
        batches = minibatch(TRAIN_DATA, size=8)
        for batch in batches:
            for text, annotations in batch:
                doc = nlp.make_doc(text)
                example = Example.from_dict(doc, annotations)
                nlp.update([example], losses=losses)
        print(f"Iteration {iteration}, Losses: {losses}")

    # Save the updated model to disk
    nlp.to_disk("./trained_ner_model")
    print("Model training completed and saved to disk!")
else:
    print("No training data was generated.")
