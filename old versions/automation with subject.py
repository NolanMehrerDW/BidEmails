import spacy
import win32com.client
import re
import pandas as pd
from datetime import datetime, timedelta

# Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY"
date_patterns = [
    r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
    r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
]

# Load the trained NER model
nlp = spacy.load("trained_ner_model")

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
def extract_dates_with_regex(text):
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group()
    return None  # No date found

# Function to get model predictions from the body only
def get_model_predictions_from_body(body):
    doc = nlp(body)
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
for message in messages:
    try:
        body = message.Body
        subject = message.Subject
        
        # Use regex to find dates in both the subject and body
        bid_due_date = extract_dates_with_regex(subject + " " + body)
        
        # AI/ML-based extraction of project name and contractor (focus on body)
        project_name, contractor, _ = get_model_predictions_from_body(body)

        email_info = {
            "Subject": subject,
            "Bid Due Date": bid_due_date,
            "Project Name": project_name,
            "General Contractors": contractor,
            "Body": body  # Include body for logging or debugging
        }

        email_data.append(email_info)
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
