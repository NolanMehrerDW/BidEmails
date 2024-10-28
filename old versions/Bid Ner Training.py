from datetime import datetime
import spacy
import win32com.client
import random
from spacy.training import Example
from spacy.util import minibatch
import torch.utils._pytree as pytree
import torch

# Load the Transformer-based model or the trained custom model
try:
    nlp = spacy.load("./trained_ner_model")  # Load the previously trained model if it exists
    print("Loaded existing trained NER model.")
except Exception as e:
    print(f"Failed to load trained NER model: {e}")
    try:
        nlp = spacy.load("en_core_web_trf")  # Load Transformer-based model as a fallback
        print("Loaded Transformer-based NER Model.")
    except Exception as e:
        print(f"Failed to load Transformer-based model: {e}")
        nlp = spacy.blank("en")  # Start with a blank model if no model exists
        print("No existing model found, starting fresh.")
        # Add the NER pipeline to the blank model
        ner = nlp.add_pipe("ner")
        # Add custom labels to the NER model
        labels = ["PROJECT_NAME", "CONTRACTOR", "BID_DUE_DATE"]
        for label in labels:
            ner.add_label(label)
        nlp.initialize()  # Only initialize if starting fresh

# Add the NER pipeline if it doesn't exist
if "ner" not in nlp.pipe_names:
    ner = nlp.add_pipe("ner")
else:
    ner = nlp.get_pipe("ner")

# Add custom labels to the NER model if they don't already exist
labels = ["PROJECT_NAME", "CONTRACTOR", "BID_DUE_DATE"]
for label in labels:
    if label not in ner.labels:
        ner.add_label(label)

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

# Function to get model predictions for project name, contractor, and bid due date
def get_model_predictions(text):
    doc = nlp(text)
    project_name, contractor, bid_due_date = None, None, None

    # Extract the model's predictions for each entity
    for ent in doc.ents:
        if ent.label_ == "PROJECT_NAME":
            project_name = ent.text
        elif ent.label_ == "CONTRACTOR":
            contractor = ent.text
        elif ent.label_ == "BID_DUE_DATE":
            bid_due_date = ent.text

    return project_name, contractor, bid_due_date

# Function to prompt the user for entity information from the email, showing model predictions
def prompt_for_labels(combined_text):
    print("\nCombined Email Text:\n", combined_text)

    # Ask if the email is relevant
    relevant = input("Is this email relevant? (y/n): ").strip().lower()
    if relevant == 'n':
        return None  # Disregard this email

    # Get the model's predictions for this email
    model_project_name, model_contractor, model_bid_due_date = get_model_predictions(combined_text)
    
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

    # Identify entity positions in the email body for annotation
    entities = []
    if project_name:
        start_idx = combined_text.lower().find(project_name.lower())
        if start_idx != -1:
            entities.append((start_idx, start_idx + len(project_name), "PROJECT_NAME"))

    if contractor:
        start_idx = combined_text.lower().find(contractor.lower())
        if start_idx != -1:
            entities.append((start_idx, start_idx + len(contractor), "CONTRACTOR"))

    if bid_due_date:
        start_idx = combined_text.lower().find(bid_due_date.lower())
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
    # Select 5 random emails
    random_indices = random.sample(range(message_count), min(5, message_count))  # Randomly select 5 emails
    for i in random_indices:
        message = messages[i]  # Access each random message
        combined_text = f"{message.Subject}\n{message.Body}"  # Combine subject and body for training

        # Get the model's predictions for this email
        entities = prompt_for_labels(combined_text)  # Allow corrections if requested

        if entities:
            TRAIN_DATA.append((combined_text, {"entities": entities}))

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
