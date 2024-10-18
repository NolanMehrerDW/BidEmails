import openai
import win32com.client
import pandas as pd
from datetime import datetime
from dateutil import parser

# Set up OpenAI API key
openai.api_key = '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'

# Function to extract information using OpenAI's API
def extract_info_with_gpt(email_text):
    prompt = f"""Extract the following information from the email:
    - Project Name
    - Contractor
    - Bid Due Date

    Email: {email_text}
    """
    
    response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an assistant that extracts information from emails."},
            {"role": "user", "content": prompt}
        ],
        max_tokens=200,
        temperature=0.3
    )
    
    extracted_text = response['choices'][0]['message']['content'].strip()
    return parse_extracted_text(extracted_text)

# Function to parse extracted text into structured data
def parse_extracted_text(extracted_text):
    lines = extracted_text.split('\n')
    data = {"Project Name": None, "Contractor": None, "Bid Due Date": None}
    for line in lines:
        if "Project Name:" in line:
            data["Project Name"] = line.split(":", 1)[1].strip()
        elif "Contractor:" in line:
            data["Contractor"] = line.split(":", 1)[1].strip()
        elif "Bid Due Date:" in line:
            data["Bid Due Date"] = line.split(":", 1)[1].strip()
    return data

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
        info = extract_info_with_gpt(combined_text)
        project_name = info["Project Name"]
        contractor = info["Contractor"]
        bid_due_date = info["Bid Due Date"]

        # Active learning: Check predictions in verbose training mode
        if verbose_training_mode:
            if not project_name:
                project_name = input(f"Low-confidence prediction for Project Name: '{project_name}'. Please correct or press Enter to confirm: ") or project_name
            if not contractor:
                contractor = input(f"Low-confidence prediction for Contractor: '{contractor}'. Please correct or press Enter to confirm: ") or contractor
            if not bid_due_date:
                bid_due_date = input(f"Low-confidence prediction for Bid Due Date: '{bid_due_date}'. Please correct or press Enter to confirm: ") or bid_due_date

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
