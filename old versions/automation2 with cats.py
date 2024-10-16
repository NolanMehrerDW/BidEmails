import spacy
import win32com.client
import re
import pandas as pd
from datetime import datetime, timedelta
import random
from fuzzywuzzy import fuzz
from dateutil import parser

# Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY"
date_patterns = [
    r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
    r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
]

# Load the trained NER model
nlp = spacy.load("./trained_ner_model")
print("Loaded NER Model...")

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

# Function to extract dates using regex patterns
def extract_dates_with_regex(email_body):
    for pattern in date_patterns:
        match = re.search(pattern, email_body)
        if match:
            return match.group()
    return None  # No date found

def validate_date(entity):
    try:
        parsed_date = parser.parse(entity, fuzzy=False)
        return True  # If date is valid
    except ValueError:
        return False  # If not a valid date

# Function to get model predictions for project name, contractor, and bid due date
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
            # Validate if it's a real date
            if validate_date(ent.text):
                bid_due_date = ent.text

    # Use regex as a fallback if the model doesn't extract the bid due date
    if not bid_due_date:
        bid_due_date = extract_dates_with_regex(email_body)
        if bid_due_date:
            bid_due_date += "?"  # Affix '?' to indicate it was found via regex

    return project_name, contractor, bid_due_date

# Function to assign an orange category to an email
def assign_orange_category(message):
    try:
        # Set the category to "Orange Category"
        message.Categories = "Orange Category"
        message.Save()  # Save the email after changing the category
        print("Category set to Orange for email:", message.Subject)
    except Exception as e:
        print(f"Failed to set category for email: {message.Subject}. Error: {e}")

# Function to find project name using a backup method (matching the subject)
def backup_project_name_from_subject(subject, known_project_names):
    for project_name in known_project_names:
        # Use fuzzy matching to compare the subject with known project names
        if fuzz.partial_ratio(subject.lower(), project_name.lower()) > 45:  # Adjust threshold if needed, 80 was default
            return f"**{project_name}"
    return None

# Function to pass subject line to the model as a last resort
def extract_project_name_from_subject(subject):
    print(f"Attempting to extract project name from subject: {subject}")
    subject_doc = nlp(subject)
    for ent in subject_doc.ents:
        if ent.label_ == "PROJECT_NAME":
            return ent.text
    return None

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

    # Ask the user if they want to filter by number of emails or days
    filter_choice = input("Would you like to filter by number of emails or days? (Enter 'emails' or 'days'): ").strip().lower()

    # Filter by the number of recent emails
    if filter_choice == 'emails':
        num_emails = int(input("How many recent emails would you like to process?: "))
        # Instead of sorting all messages, we will iterate over the first num_emails
        message_count = len(messages)
        filtered_messages = []
        for i in range(min(num_emails, message_count)):
            filtered_messages.append(messages[i])
        messages = filtered_messages
    
    # Filter by the number of days
    elif filter_choice == 'days':
        days = int(input("Enter the number of days for recent emails: "))
        cutoff_date = datetime.now() - timedelta(days=days)
        messages = [msg for msg in messages if msg.ReceivedTime >= cutoff_date]

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
        
        # AI/ML-based extraction of relevant information
        project_name, contractor, bid_due_date = get_model_predictions(body)

        # Backup: If project name is not found, check the subject line against known project names
        if not project_name:
            project_name = backup_project_name_from_subject(subject, known_project_names)

        # Final Backup: If still no project name, use the model on the subject line
        if not project_name:
            project_name = f"&{extract_project_name_from_subject(subject)}"

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

# Convert 'General Contractors' from list to string for proper handling
df = pd.DataFrame(email_data)

# Show gathered data
print(df.head())

# Save the updated data to an Excel file
output_file = 'bid_requests_calendar_new.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Bid Requests', index=False)

print(f"Data saved to {output_file}")
