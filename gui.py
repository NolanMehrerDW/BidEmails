import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QTextEdit,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
import spacy

class BidEmailsProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Bid Emails Processor')
        self.init_ui()
        self.model_loaded = False
        self.load_model()

    def init_ui(self):
        # Folder Selection
        folder_label = QLabel('Select Folder:')
        self.folder_path = QLineEdit()
        self.folder_path.setReadOnly(True)
        browse_button = QPushButton('Browse')
        browse_button.clicked.connect(self.select_folder)

        folder_layout = QHBoxLayout()
        folder_layout.addWidget(folder_label)
        folder_layout.addWidget(self.folder_path)
        folder_layout.addWidget(browse_button)

        # Buttons
        process_button = QPushButton('Process Emails')
        process_button.clicked.connect(self.process_emails)
        train_button = QPushButton('Train Model')
        train_button.clicked.connect(self.train_model)

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(process_button)
        buttons_layout.addWidget(train_button)

        # Log Output
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        # Main Layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(folder_layout)
        main_layout.addLayout(buttons_layout)
        main_layout.addWidget(self.log_output)

        self.setLayout(main_layout)

    def select_folder(self):
        folder_selected = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder_selected:
            self.folder_path.setText(folder_selected)

    def process_emails(self):
        folder_selected = self.folder_path.text()
        if not folder_selected:
            QMessageBox.warning(self, 'Warning', 'Please select a folder first.')
            return

        self.log_output.append('Processing emails...')
        QApplication.processEvents()  # Update GUI

        try:
            # Insert your email processing code here
            import spacy
            import win32com.client
            import re
            import pandas as pd
            import os
            from datetime import datetime, timedelta
            import random

            # Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY"
            date_patterns = [
                r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
                r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
            ]

            # Load the trained NER model
            nlp = spacy.load("BidEmails/trained_ner_model")
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
                        bid_due_date = ent.text
                
                # Use regex as a fallback if the model doesn't extract the bid due date
                if not bid_due_date:
                    bid_due_date = extract_dates_with_regex(email_body)
                
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
            for message in messages:
                try:
                    body = message.Body
                    subject = message.Subject
                    
                    # AI/ML-based extraction of relevant information
                    project_name, contractor, bid_due_date = get_model_predictions(body)

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


            # Simulate processing
            self.log_output.append(f'Processed emails from folder: {folder_selected}')
            # Display results or update the log
            # self.log_output.append(str(results))
        except Exception as e:
            self.log_output.append(f'Error: {e}')
            QMessageBox.critical(self, 'Error', f'An error occurred: {e}')
        else:
            QMessageBox.information(self, 'Success', 'Email processing completed!')

    def train_model(self):
        self.log_output.append('Starting model training...')
        QApplication.processEvents()  # Update GUI

        try:
            # Insert your model training code here
            from datetime import datetime
            import spacy
            import win32com.client
            import re
            from spacy.training import Example
            from spacy.util import minibatch
            import random

            # Date regex patterns for "Month DD, YYYY" and "MM/DD/YYYY"
            date_patterns = [
                r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b",  # "Month DD, YYYY"
                r"\b\d{1,2}/\d{1,2}/\d{4}\b"  # "MM/DD/YYYY"
            ]

            # Try to load the pre-trained model if it exists, otherwise start with a blank model
            try:
                nlp = spacy.load("BidEmails/trained_ner_model")  # Load the previously trained model
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

            # Function to extract dates using regex patterns
            def extract_dates_with_regex(text):
                for pattern in date_patterns:
                    match = re.search(pattern, text)
                    if match:
                        return match.group()
                return None  # No date found

            # Pre-annotation function using regex to find project names, contractors, and dates
            def pre_annotate(text):
                entities = []

                # Pre-label dates using regex
                date_matches = re.finditer(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b", text)
                for match in date_matches:
                    entities.append((match.start(), match.end(), 'BID_DUE_DATE'))
                
                # Pre-label project names (based on pattern recognition)
                project_matches = re.finditer(r'Project: (.+)', text)
                for match in project_matches:
                    entities.append((match.start(1), match.end(1), 'PROJECT_NAME'))

                # Pre-label contractors (based on pattern recognition)
                contractor_matches = re.finditer(r'Contractor: (.+)', text)
                for match in contractor_matches:
                    entities.append((match.start(1), match.end(1), 'CONTRACTOR'))

                return entities

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

            # Collect training data by pre-annotating with regex and prompting for corrections when needed
            TRAIN_DATA = []
            message_count = len(messages)
            if message_count > 0:
                # Select 5 random emails
                random_indices = random.sample(range(message_count), min(5, message_count))  # Randomly select 5 emails
                for i in random_indices:
                    message = messages[i]  # Access each random message
                    email_body = message.Body

                    # Pre-annotate the entities
                    entities = pre_annotate(email_body)

                    if entities:
                        print(f"\nPre-annotated Entities: {entities}")
                        user_review = input("Would you like to review and correct? (y/n): ").strip().lower()
                        if user_review == 'y':
                            entities = prompt_for_labels(email_body)  # Allow corrections if requested

                        TRAIN_DATA.append((email_body, {"entities": entities}))

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
                nlp.to_disk("BidEmails/trained_ner_model")
                print("Model training completed and saved to disk!")
            else:
                print("No training data was generated.")


            # Simulate training
            self.log_output.append('Model training completed successfully.')
        except Exception as e:
            self.log_output.append(f'Error: {e}')
            QMessageBox.critical(self, 'Error', f'An error occurred: {e}')
        else:
            QMessageBox.information(self, 'Success', 'Model training completed!')

    def load_model(self):
        try:
            self.nlp = spacy.load("BidEmails/trained_ner_model")
            self.model_loaded = True
            self.log_output.append('Loaded existing model.')
        except Exception as e:
            self.log_output.append('No existing model found. Starting with a blank model.')
            self.nlp = spacy.blank("en")
            self.model_loaded = False

def main():
    app = QApplication(sys.argv)
    window = BidEmailsProcessor()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
