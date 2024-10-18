import spacy
import win32com.client
import pandas as pd
from datetime import datetime
from dateutil import parser

# Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY"
date_patterns = [
    r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
    r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
]

# Try to load the Transformer-based NER model or create a new one
try:
    nlp = spacy.load("./trained_ner_model")  # Load Transformer model
    print("Loaded saved NER Model.")
except:
    nlp = spacy.blank("en")  # Start with a blank model if no model exists
    print("No existing model found, starting fresh.")

# Add the NER pipeline if it doesn't exist
if "ner" not in nlp.pipe_names:
    ner = nlp.add_pipe("ner")
else:
    ner = nlp.get_pipe("ner")

# Add custom labels if not already present
ner.add_label("PROJECT_NAME")
ner.add_label("CONTRACTOR")
ner.add_label("BID_DUE_DATE")

# **Initialize the NER model**
nlp.initialize()  # Make sure to call initialize() here

# Connect to Outlook
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

# Validate dates using dateutil.parser
def validate_date(entity):
    try:
        parsed_date = parser.parse(entity, fuzzy=False)
        return True  # If date is valid
    except ValueError:
        return False  # If not a valid date

# Function to get model predictions for project name, contractor, and bid due date
def get_model_predictions(text, confidence_threshold=0.75):
    doc = nlp(text)
    project_name = None
    contractor = None
    bid_due_date = None

    # Extract the model's predictions for each entity with confidence scoring
    for ent in doc.ents:
        confidence = get_confidence(ent)
        if ent.label_ == "PROJECT_NAME" and confidence >= confidence_threshold:
            project_name = ent.text
        elif ent.label_ == "CONTRACTOR" and confidence >= confidence_threshold:
            contractor = ent.text
        elif ent.label_ == "BID_DUE_DATE" and confidence >= confidence_threshold:
            if validate_date(ent.text):
                bid_due_date = ent.text

    return project_name, contractor, bid_due_date

# Get confidence score for entities
def get_confidence(entity):
    return entity.kb_id_ if entity.kb_id_ else 0.0

# Active learning workflow for low-confidence predictions (Verbose Training Mode)
def active_learning_correction(entity_type, prediction):
    print(f"Low-confidence prediction for {entity_type}: '{prediction}'")
    corrected_value = input(f"Please confirm or correct the {entity_type} (press Enter to confirm '{prediction}'): ")
    return corrected_value if corrected_value else prediction

# Assign Orange Category to email in Outlook
def assign_orange_category(message):
    try:
        message.Categories = "Orange Category"
        message.Save()  # Save the email after changing the category
        print("Category set to Orange for email:", message.Subject)
    except Exception as e:
        print(f"Failed to set category for email: {message.Subject}. Error: {e}")

# Prompt user to enable Verbose Training Mode
verbose_training_mode = input("Would you like to enable Verbose Training Mode? (y/n): ").strip().lower() == 'y'

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

    # Ask for the number of most recent emails to process
    num_emails = int(input("How many recent emails would you like to process?: "))
    filtered_messages = [messages[i] for i in range(min(num_emails, len(messages)))]
    messages = filtered_messages

    print(f"Processing {len(messages)} emails after filtering.")

else:
    print("No folders with 'Bid' in the name were found.")
    messages = []

# Collect email data for output
email_data = []
known_project_names = []  # List to track known project names

for message in messages:
    try:
        body = message.Body
        subject = message.Subject

        # Combine subject and body for model predictions
        combined_text = f"{subject}\n{body}"

        # AI/ML-based extraction of relevant information
        project_name, contractor, bid_due_date = get_model_predictions(combined_text)

        # Active learning: Check predictions in verbose training mode
        if verbose_training_mode:
            if not project_name:
                project_name = active_learning_correction("Project Name", project_name)
            if not contractor:
                contractor = active_learning_correction("Contractor", contractor)
            if not bid_due_date:
                bid_due_date = active_learning_correction("Bid Due Date", bid_due_date)

        # Add the project name to the known list if extracted successfully
        if project_name:
            known_project_names.append(project_name)

        email_info = {
            "Subject": subject,
            "Bid Due Date": bid_due_date,
            "Project Name": project_name,
            "General Contractors": contractor,
            "Body": body  # Include body for logging or debugging
        }

        email_data.append(email_info)

        # Assign an orange category to the email after processing
        assign_orange_category(message)
        
    except Exception as e:
        print(f"Failed to process email: {e}")

# Convert email data into a DataFrame
df = pd.DataFrame(email_data)

# Show the extracted data
print(df.head())

# Save the updated data to an Excel file
output_file = 'bid_requests_calendar_new.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Bid Requests', index=False)

print(f"Data saved to {output_file}")

# Save the model after corrections
if verbose_training_mode:
    nlp.to_disk("./trained_ner_model")
    print("Model training completed and saved to disk after corrections.")
