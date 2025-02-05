import openai
import pandas as pd
from datetime import datetime
import os
import win32com.client
from dateutil import parser
import logging
import json

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load the OpenAI API key from a config file
with open("api_key.txt", "r") as file:
    openai.api_key = file.read().strip()

# Load corrections from a JSON file
def load_corrections():
    try:
        with open("corrections.json", "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}

# Save corrections to a JSON file
def save_corrections(corrections):
    with open("corrections.json", "w") as file:
        json.dump(corrections, file, indent=4)

# Load existing corrections
corrections = load_corrections()

# Function to connect to the OpenAI API and extract relevant information from email content
def extract_email_info(subject, body):
    try:
        logging.info("Extracting information from email with subject: %s", subject)
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an assistant extracting project information from emails."},
                {"role": "user", "content": f"Extract the project name, contractor, bid due date, job walk information (if available), and a brief project description (highlight union or prevailing wage if mentioned) from the following email:\n\nEmail Example:\nSubject: {subject}\nBody: {body}\n\nThe format of the response should be:\nProject Name: ...\nContractor: ...\nBid Due Date: ...\nJob Walk: ...\nDescription: ..."}
            ],
            max_tokens=500,
            temperature=0.2
        )

        content = response.choices[0].message['content']
        lines = content.split('\n')
        project_name, contractor, bid_due_date, job_walk, description = None, None, None, None, None

        # Extract the relevant fields from the response
        for line in lines:
            if line.startswith("Project Name:"):
                project_name = line.replace("Project Name:", "").strip()
            elif line.startswith("Contractor:"):
                contractor = line.replace("Contractor:", "").strip()
            elif line.startswith("Bid Due Date:"):
                bid_due_date = line.replace("Bid Due Date:", "").strip()
            elif line.startswith("Job Walk:"):
                job_walk = line.replace("Job Walk:", "").strip()
            elif line.startswith("Description:"):
                description = line.replace("Description:", "").strip()

        logging.info("Extracted - Project Name: %s, Contractor: %s, Bid Due Date: %s, Job Walk: %s, Description: %s", project_name, contractor, bid_due_date, job_walk, description)
        return project_name, contractor, bid_due_date, job_walk, description

    except Exception as e:
        logging.error(f"Failed to process email: {e}")
        return None, None, None, None, None

# Function to normalize dates to mm/dd/yy format
def normalize_date(date_str):
    try:
        parsed_date = parser.parse(date_str, fuzzy=True)
        return parsed_date.strftime("%m/%d/%y")
    except Exception as e:
        logging.warning("Failed to parse date '%s': %s", date_str, e)
        return "01/01/11**"

# Connect to Outlook and get the namespace
logging.info("Connecting to Outlook...")
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
    logging.info("Total emails in the selected folder: %d", len(messages))

    # Ask for the number of most recent emails to process
    user_input = input("How many recent emails would you like to process? (e.g., '15' or '15,200'): ")
    if "," in user_input:
        num_emails, start_index = map(int, user_input.split(","))
    else:
        num_emails, start_index = int(user_input), 0
    
    filtered_messages = [messages[i] for i in range(start_index, min(start_index + num_emails, len(messages)))]

    logging.info("Processing %d emails after filtering.", len(filtered_messages))
    messages = filtered_messages

else:
    logging.warning("No folders with 'Bid' in the name were found.")
    messages = []

# Collect extracted data
email_data = []
for message in messages:
    try:
        logging.info("Processing email with subject: %s", message.Subject)
        subject = message.Subject
        body = message.Body
        project_name, contractor, bid_due_date, job_walk, description = extract_email_info(subject, body)

        # Normalize bid due date
        bid_due_date = normalize_date(bid_due_date) if bid_due_date else "01/01/11**"

        # Apply corrections if available
        if project_name in corrections:
            corrected_info = corrections[project_name]
            contractor = corrected_info.get("Contractor", contractor)
            bid_due_date = corrected_info.get("Bid Due Date", bid_due_date)
            
        # Mark email with Orange category
        message.Categories = "Orange Category"
        message.Save()
        logging.info("Marked email with subject '%s' as Orange Category.", message.Subject)    

        # Add extracted information to email data list
        email_data.append({
            "Project Name": project_name,
            "Contractor": contractor,
            "Bid Due Date": bid_due_date if bid_due_date != "Not specified" else "",
            "Job Walk": job_walk if job_walk else "",
            "Description": description if description else ""
        })
    except Exception as e:
        logging.error("Failed to process email with subject '%s': %s", message.Subject, e)

# Convert to DataFrame
email_df = pd.DataFrame(email_data)

# Consolidate identical projects and list unique contractors
consolidated_data = (
    email_df.groupby("Project Name", as_index=False)
    .agg({
        "Bid Due Date": "first",
        "Contractor": lambda x: ", ".join(sorted(set(filter(None, [c.strip() for c in x])))),
        "Job Walk": "first",
        "Description": "first"
    })
)

# Save to Excel file
output_file = "bid_requests_calendar.xlsx"
logging.info("Saving data to Excel file: %s", output_file)
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    consolidated_data.to_excel(writer, sheet_name='Bid Requests', index=False)
logging.info("Data successfully saved to %s", output_file)
