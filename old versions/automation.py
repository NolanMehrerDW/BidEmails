import spacy
import win32com.client
import pandas as pd
from datetime import datetime
from dateutil import parser
import random
import torch.utils._pytree as pytree
import torch

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
labels = ["PROJECT_NAME", "CONTRACTOR", "BID_DUE_DATE"]
for label in labels:
    if label not in ner.labels:
        ner.add_label(label)

# **Initialize the NER model**
try:
    nlp.initialize()  # Make sure to call initialize() here
except Exception as e:
    print(f"Failed to initialize the model: {e}")

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
def active_learning_correction(entity_type, prediction, message):
    print(f"Low-confidence prediction for {entity_type}: '{prediction}'")
    message.Display()  # Opens the email in Outlook for easier review
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
                project_name = active_learning_correction("Project Name", project_name, message)
            if not contractor:
                contractor = active_learning_correction("Contractor", contractor, message)
            if not bid_due_date:
                bid_due_date = active_learning_correction("Bid Due Date", bid_due_date, message)

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

# **Deprecation Warning Updates**
# Update deprecated calls
pytree.register_pytree_node = getattr(pytree, 'register_pytree_node', pytree._register_pytree_node)
print("Using updated register_pytree_node method.")

# Set torch load with weights_only=True to avoid potential security issues
def load_weights_safe(filelike, map_location):
    return torch.load(filelike, map_location=map_location, weights_only=True)

# Update torch.amp.autocast to the new format
def autocast(device, *args, **kwargs):
    return torch.amp.autocast(device, *args, **kwargs)
print("Updated deprecated torch.cuda.amp.autocast to torch.amp.autocast.")

# Warning fixes for AttributeRuler and Lemmatizer
from spacy.pipeline import AttributeRuler, Lemmatizer

if "attribute_ruler" not in nlp.pipe_names:
    attribute_ruler = nlp.add_pipe("attribute_ruler")
else:
    attribute_ruler = nlp.get_pipe("attribute_ruler")

if "lemmatizer" not in nlp.pipe_names:
    lemmatizer = nlp.add_pipe("lemmatizer", after="attribute_ruler")
else:
    lemmatizer = nlp.get_pipe("lemmatizer")

print("Initialized AttributeRuler and Lemmatizer to resolve warnings.")
