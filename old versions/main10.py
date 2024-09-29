import win32com.client
from datetime import datetime, timedelta
import re
import joblib
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.cluster import AgglomerativeClustering
from sklearn.metrics import pairwise_distances
import nltk
from nltk.corpus import stopwords

# Download stopwords if not already downloaded
nltk.download('stopwords')

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Recursive function to list all folders and their subfolders
def list_folders(folder, folder_list, parent_name=""):
    folder_name = f"{parent_name}/{folder.Name}" if parent_name else folder.Name
    folder_list.append((folder_name, folder))
    
    # Recursively add subfolders
    for i in range(folder.Folders.Count):
        sub_folder = folder.Folders.Item(i + 1)
        list_folders(sub_folder, folder_list, folder_name)

# Function to allow the user to select from all available folders
def select_folder():
    folder_list = []
    
    # List all top-level folders
    for i in range(outlook.Folders.Count):
        top_folder = outlook.Folders.Item(i + 1)
        list_folders(top_folder, folder_list)  # Recursively list all subfolders

    # Display all folders for selection
    print("\nAvailable Folders:")
    for index, (folder_name, _) in enumerate(folder_list):
        print(f"{index + 1}: {folder_name}")

    # Input folder selection from the user
    try:
        folder_index = int(input("\nEnter the number of the folder you want to select: ")) - 1
        selected_folder = folder_list[folder_index][1]
        print(f"You selected: {folder_list[folder_index][0]}")
    except (ValueError, IndexError):
        print("Invalid selection. Defaulting to Inbox.")
        selected_folder = outlook.GetDefaultFolder(6)  # Inbox as fallback
    
    return selected_folder

# Function to clean subject lines by removing prefixes
def clean_subject(subject):
    cleaned_subject = re.sub(r'^\s*(RE:|FW:|bid invite:)\s*', '', subject, flags=re.IGNORECASE)
    return cleaned_subject.strip()

# Function to preprocess the text
def preprocess_text(subjects):
    stop_words = set(stopwords.words('english'))
    processed_subjects = []
    
    for subject in subjects:
        # Remove stopwords and convert to lowercase
        words = subject.lower().split()
        filtered_words = [word for word in words if word not in stop_words]
        processed_subjects.append(' '.join(filtered_words))
    
    return processed_subjects

# Function to group similar subjects using Agglomerative Clustering
def load_vectorizer(filename):
    return joblib.load(filename)

def group_similar_subjects(subjects, distance_threshold=0.5, email_indices=None):
    try:
        vectorizer = load_vectorizer('vectorizer.pkl')
        print("Loading...")
    except Exception as e:
        print("Could not load vectorizer, creating a new one.")
        vectorizer = TfidfVectorizer()
    
    X = vectorizer.fit_transform(subjects)
    distance_matrix = pairwise_distances(X.toarray(), metric='euclidean')

    clustering = AgglomerativeClustering(distance_threshold=distance_threshold, n_clusters=None)
    clustering.fit(distance_matrix)

    clustered_subjects = {}
    for idx, label in enumerate(clustering.labels_):
        if label not in clustered_subjects:
            clustered_subjects[label] = []
        clustered_subjects[label].append((subjects[idx], email_indices[idx]))
    
    return clustered_subjects

# Function to get emails from the folder, including flagged emails
def get_emails_from_folder(folder, max_emails):
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by latest first

    compiled_emails = []
    count = 0
    for message in messages:
        if count >= max_emails:
            break  # Stop if we've reached the maximum number of emails requested
        
        try:
            received_time = message.ReceivedTime.replace(tzinfo=None)  # Make the received time offset-naive
            
            # Collect all emails regardless of flagged status
            compiled_emails.append({
                "Sender": message.SenderName,
                "Subject": clean_subject(message.Subject),
                "ReceivedTime": received_time,
                "Body": message.Body[:200],  # Get first 200 characters of the body
                "EntryID": message.EntryID,
                "StoreID": folder.StoreID
            })
            count += 1  # Count every email

        except Exception as e:
            print(f"Error processing email: {e}")
    
    return compiled_emails

# Function to print sorted emails with a body snippet and a link to open the email
def print_sorted_emails(emails):
    for idx, email in enumerate(emails):
        print(f"Email {idx + 1}")
        print(f"Sender: {email['Sender']}")
        print(f"Subject: {email['Subject']}")
        print(f"ReceivedTime: {email['ReceivedTime'].strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Body (snippet): {email['Body']}")
        print(f"Open this email using: open_email({idx + 1})")
        print("=" * 50)

# Function to open an email in Outlook by EntryID and StoreID
def open_email(email_index):
    try:
        selected_email = retrieved_emails[email_index - 1]
        mail_item = outlook.GetItemFromID(selected_email["EntryID"], selected_email["StoreID"])
        mail_item.Display()  # Open the email in the Outlook client
    except Exception as e:
        print(f"Error opening email: {e}")

# Function to update the category of emails
def update_email_category(emails, category="Blue Category"):
    for email in emails:
        try:
            mail_item = outlook.GetItemFromID(email["EntryID"], email["StoreID"])
            mail_item.Categories = category  # Set the category
            mail_item.Save()  # Save the changes
            print(f"Updated category for email: {email['Subject']} to '{category}'")
        except Exception as e:
            print(f"Error updating category for {email['Subject']}: {e}")

# Function to train a model based on past emails
def train_model(training_emails, labels):
    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform(training_emails)
    
    classifier = LogisticRegression()
    classifier.fit(X, labels)

    # Save the model and vectorizer for future use
    joblib.dump(classifier, 'email_classifier.pkl')
    joblib.dump(vectorizer, 'vectorizer.pkl')
    print("Model and vectorizer saved.")

# Function to predict if new emails are relevant
def predict_emails(new_emails):
    classifier = joblib.load('email_classifier.pkl')
    vectorizer = joblib.load('vectorizer.pkl')

    new_X = vectorizer.transform(new_emails)
    predictions = classifier.predict(new_X)
    return predictions

# Main script logic
if __name__ == "__main__":
    # Step 1: Select folder
    selected_folder = select_folder()

    # Step 2: Ask for maximum number of emails to retrieve
    try:
        max_emails = int(input("What is the maximum number of recent emails you want to retrieve? "))
    except ValueError:
        print("Invalid number entered. Defaulting to 10 emails.")
        max_emails = 10  # Default to 10 emails if invalid input

    # Step 3: Get emails from the selected folder
    retrieved_emails = get_emails_from_folder(selected_folder, max_emails)

    # Step 4: Print retrieved emails
    if retrieved_emails:
        print("\nRetrieved Emails:")
        print_sorted_emails(retrieved_emails)
        
        # Step 5: Collect subjects for clustering and prediction
        subjects = [email["Subject"] for email in retrieved_emails]
        processed_subjects = preprocess_text(subjects)
        
        # Step 6: Group similar subjects
        clustered_subjects = group_similar_subjects(processed_subjects, email_indices=[email["EntryID"] for email in retrieved_emails])

        # Display the clustered results
        print("\nGrouped Emails:")
        for label, items in clustered_subjects.items():
            print(f"\nCluster {label}:")
            unique_items = {item[0]: item for item in items}  # Remove duplicates
            for subject, _ in unique_items.values():
                print(f"  Subject: {subject}")

        # Step 7: Train the model with labeled data (manual step)
        if input("\nDo you want to train the model with new labeled data? (y/n): ").strip().lower() == 'y':
            training_emails = [email["Subject"] for email in retrieved_emails]
            labels = [1 if "bid" in email["Subject"].lower() else 0 for email in retrieved_emails]  # Example labels
            train_model(training_emails, labels)

        # Step 8: Predict relevance of the most recent emails
        new_emails = [email["Subject"] for email in retrieved_emails]
        predictions = predict_emails(new_emails)
        print("Predictions for recent emails:", predictions)

        # Step 9: Update email categories based on user input (optional)
        if input("\nDo you want to update the category for these emails? (y/n): ").strip().lower() == 'y':
            update_email_category(retrieved_emails)

    else:
        print("No emails found within the specified criteria.")
