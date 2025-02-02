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

def extract_email_info(subject, body):
    """
    Extracts key fields from the email content using the OpenAI API:
      - Project Name
      - Contractor
      - Bid Due Date
      - Job Walk (date/time and GC hosting if known)
      - Project Description (noting union or prevailing wage if applicable)
    """
    try:
        logging.info("Extracting information from email with subject: %s", subject)

        # Extended prompt to include new fields (Job Walk & Project Description)
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an assistant extracting project information from emails."
                    )
                },
                {
                    "role": "user",
                    "content": (
                        f"Extract the following details from the email:\n\n"
                        f"1. Project Name\n"
                        f"2. Contractor\n"
                        f"3. Bid Due Date\n"
                        f"4. Job Walk (date/time, and which GC is hosting if known)\n"
                        f"5. Project Description (especially note union or prevailing wage)\n\n"
                        f"Email Example:\n"
                        f"Subject: {subject}\n"
                        f"Body: {body}\n\n"
                        "The format of the response should be:\n"
                        "Project Name: ...\n"
                        "Contractor: ...\n"
                        "Bid Due Date: ...\n"
                        "Job Walk: ...\n"
                        "Project Description: ..."
                    )
                }
            ],
            max_tokens=450,
            temperature=0.2
        )

        content = response.choices[0].message['content']
        lines = content.split('\n')

        # Default values
        project_name = None
        contractor = None
        bid_due_date = None
        job_walk = None
        project_description = None

        # Extract relevant fields from the API response
        for line in lines:
            lower_line = line.lower()
            if line.startswith("Project Name:"):
                project_name = line.replace("Project Name:", "").strip()
            elif line.startswith("Contractor:"):
                contractor = line.replace("Contractor:", "").strip()
            elif line.startswith("Bid Due Date:"):
                bid_due_date = line.replace("Bid Due Date:", "").strip()
            elif line.startswith("Job Walk:"):
                job_walk = line.replace("Job Walk:", "").strip()
            elif line.startswith("Project Description:"):
                project_description = line.replace("Project Description:", "").strip()

        logging.info(
            "Extracted - Project Name: %s, Contractor: %s, "
            "Bid Due Date: %s, Job Walk: %s, Project Description: %s",
            project_name, contractor, bid_due_date, job_walk, project_description
        )

        return project_name, contractor, bid_due_date, job_walk, project_description

    except Exception as e:
        logging.error(f"Failed to process email: {e}")
        return None, None, None, None, None

def normalize_date(date_str):
    """
    Normalizes a date string to mm/dd/yy format.
    Returns '01/01/11**' if the date cannot be parsed.
    """
    try:
        parsed_date = parser.parse(date_str, fuzzy=True)
        return parsed_date.strftime("%m/%d/%y")
    except Exception as e:
        logging.warning("Failed to parse date '%s': %s", date_str, e)
        return "01/01/11**"

logging.info("Connecting to Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def find_bid_folders(folder, folder_path=""):
    """
    Recursively locates folders containing 'Bid' in the name.
    """
    bid_folders = []
    folder_path = f"{folder_path}\\{folder.Name}" if folder_path else folder.Name
    
    if "Bid" in folder.Name:
        bid_folders.append((folder_path, folder))
    
    for subfolder in folder.Folders:
        bid_folders.extend(find_bid_folders(subfolder, folder_path))
    
    return bid_folders

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

    # Extended logic: if user enters "15,200", we interpret that as "process 15 from the 200th most recent"
    user_input = input("How many recent emails would you like to process? (e.g., '15' or '15,200'): ")
    if "," in user_input:
        num_emails_str, start_index_str = user_input.split(",")
        num_emails = int(num_emails_str)
        start_index = int(start_index_str)
    else:
        num_emails = int(user_input)
        start_index = 0

    # We build from the most recent. The default Items collection is usually oldest to newest,
    # so let's reverse it to ensure we pick the most recent.
    # If your environment is sorted differently, you might need to adjust the logic.
    messages = sorted(messages, key=lambda m: m.ReceivedTime, reverse=True)

    # Slice from start_index to start_index + num_emails
    filtered_messages = messages[start_index:start_index + num_emails]

    logging.info("Processing %d emails after filtering.", len(filtered_messages))
else:
    logging.warning("No folders with 'Bid' in the name were found.")
    filtered_messages = []

email_data = []
for message in filtered_messages:
    try:
        logging.info("Processing email with subject: %s", message.Subject)
        subject = message.Subject
        body = message.Body
        (project_name,
         contractor,
         bid_due_date,
         job_walk,
         project_description) = extract_email_info(subject, body)

        # Normalize bid due date
        bid_due_date = normalize_date(bid_due_date) if bid_due_date else "01/01/11**"

        # Apply corrections if available
        if project_name in corrections:
            corrected_info = corrections[project_name]
            contractor = corrected_info.get("Contractor", contractor)
            bid_due_date = corrected_info.get("Bid Due Date", bid_due_date)
            # If you want to store job_walk and project_description in corrections, you can do so.
            # For example: job_walk = corrected_info.get("Job Walk", job_walk)

        # Mark email with Orange category
        message.Categories = "Orange Category"
        message.Save()
        logging.info("Marked email with subject '%s' as Orange Category.", message.Subject)

        # Add extracted info to the data
        email_data.append({
            "Project Name": project_name,
            "Contractor": contractor if contractor != "Not specified" else "",
            "Bid Due Date": bid_due_date if bid_due_date != "Not specified" else "",
            "Job Walk": job_walk if job_walk else "",
            "Project Description": project_description if project_description else ""
        })
    except Exception as e:
        logging.error("Failed to process email with subject '%s': %s", message.Subject, e)

# Convert to DataFrame
email_df = pd.DataFrame(email_data)

# Clean up missing fields
email_df["Contractor"] = email_df["Contractor"].apply(lambda x: x if x else "Not specified")
email_df["Job Walk"] = email_df["Job Walk"].apply(lambda x: x if x else "No info")
email_df["Project Description"] = email_df["Project Description"].apply(lambda x: x if x else "No info")

# Consolidate identical projects and list unique contractors
consolidated_data = (
    email_df.groupby("Project Name", as_index=False).agg({
        "Bid Due Date": "first",
        "Contractor": lambda x: ", ".join(sorted(set(filter(None, [c.strip() for c in x])))),
        "Job Walk": "first",
        "Project Description": "first"
    })
)

def manual_feedback(consolidated_df):
    """
    Allows the user to manually review potential duplicates and choose the name to keep.
    Any corrections are saved to corrections.json for next time.
    """
    reviewed_projects = []
    for idx, row in consolidated_df.iterrows():
        # This search might find the same row as well, but the user can still confirm or not.
        potential_duplicates = consolidated_df[consolidated_df["Project Name"].str.contains(
            row["Project Name"] if row["Project Name"] else "",
            case=False, na=False
        )]
        if len(potential_duplicates) > 1:
            print(f"Potential duplicates found for project: {row['Project Name']}")
            for _, duplicate in potential_duplicates.iterrows():
                print(f"- {duplicate['Project Name']}")
            keep = input("Which project name would you like to keep? (Leave blank to keep the original): ")
            if keep:
                row["Project Name"] = keep
                corrections[row["Project Name"]] = {
                    "Contractor": row["Contractor"],
                    "Bid Due Date": row["Bid Due Date"]
                    # If you want to store the new fields in corrections as well, add them here.
                }
        reviewed_projects.append(row)
    save_corrections(corrections)
    return pd.DataFrame(reviewed_projects)

consolidated_data = manual_feedback(consolidated_data)

# Save to Excel file
output_file = "bid_requests_calendar.xlsx"
logging.info("Saving data to Excel file: %s", output_file)
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    consolidated_data.to_excel(writer, sheet_name='Bid Requests', index=False)

logging.info("Data successfully saved to %s", output_file)
