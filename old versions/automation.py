import win32com.client
import re
import pandas as pd
import os
from datetime import datetime, timedelta
import spacy
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Load the trained NER model
nlp = spacy.load("trained_ner_model")
print("NER model: Loaded.")

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

# Search for all folders containing the keyword "Bid"
bid_folders = []
for folder in outlook.Folders:
    bid_folders.extend(find_bid_folders(folder))

# Check if any "Bid" folders were found
if bid_folders:
    print("Found the following folders with 'Bid' in the name:")
    for i, (path, _) in enumerate(bid_folders):
        print(f"{i+1}. {path}")
    
    # Prompt the user to select the correct folder
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
        # Loop through emails, keeping track of the most recent ones
        recent_emails = []
        message_count = len(messages)
        for i in range(min(num_emails, message_count)):
            recent_emails.append(messages[i])
        messages = recent_emails
    
    # Filter by the number of days
    elif filter_choice == 'days':
        days = int(input("Enter the number of days for recent emails: "))
        cutoff_date = datetime.now() - timedelta(days=days)
        messages = [msg for msg in messages if msg.ReceivedTime >= cutoff_date]

    print(f"Processing {len(messages)} emails after filtering.")

else:
    print("No folders with 'Bid' in the name were found.")
    messages = []

# Function to use the trained model for parsing project name, bid due date, and general contractors
def ai_extraction_with_trained_model(email_body):
    doc = nlp(email_body)
    project_name = None
    bid_due_date = None
    general_contractors = []

    for ent in doc.ents:
        if ent.label_ == "PROJECT_NAME":
            project_name = ent.text
        elif ent.label_ == "CONTRACTOR":
            general_contractors.append(ent.text)
        elif ent.label_ == "BID_DUE_DATE":
            bid_due_date = ent.text

    return project_name, bid_due_date, general_contractors

# Function to handle duplicate detection using cosine similarity
def is_duplicate_using_ai(existing_data, new_email_body):
    if not existing_data:
        return False

    # Convert the email bodies to vectors using TF-IDF
    vectorizer = TfidfVectorizer().fit_transform([new_email_body] + [email['Body'] for email in existing_data])
    vectors = vectorizer.toarray()

    # Check cosine similarity of the new email with each existing one
    cosine_similarities = cosine_similarity(vectors[0:1], vectors[1:]).flatten()

    # If there's a high similarity (e.g., > 0.85), consider it a duplicate
    if any(sim > 0.85 for sim in cosine_similarities):
        return True
    
    return False

# Only proceed if messages were successfully retrieved
if messages:
    email_data = []
    for message in messages:
        try:
            body = message.Body
            subject = message.Subject
            
            # AI/ML-based extraction of relevant information
            project_name, bid_due_date, general_contractors = ai_extraction_with_trained_model(body)

            email_info = {
                "Subject": subject,
                "Bid Due Date": bid_due_date,
                "Project Name": project_name,
                "General Contractors": general_contractors,
                "Body": body  # Include body for duplicate detection
            }
            
            # Only add non-duplicate emails using AI-based duplicate handling
            if not is_duplicate_using_ai(email_data, body):
                email_data.append(email_info)
        except Exception as e:
            print(f"Failed to process email: {e}")

    # Convert 'General Contractors' from list to string for proper handling
    df = pd.DataFrame(email_data)
    df['General Contractors'] = df['General Contractors'].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)

    # Show gathered data
    print(df.head())

    # Updated file name to avoid permission issues
    output_file = 'bid_requests_calendar_new.xlsx'

    # If file exists, load existing data and append new data
    if os.path.exists(output_file):
        existing_df = pd.read_excel(output_file)
        # Convert 'General Contractors' in existing dataframe as well
        existing_df['General Contractors'] = existing_df['General Contractors'].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)
        
        df = pd.concat([existing_df, df]).drop_duplicates()

    # Save the updated data to the new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Bid Requests', index=False)

    print(f"Data saved to {output_file}")

else:
    print("No emails retrieved or folder not found.")
