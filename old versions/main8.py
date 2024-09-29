import win32com.client
from datetime import datetime, timedelta
import re
from sklearn.feature_extraction.text import TfidfVectorizer
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
    
    # List all top-level folders (usually "Mailbox - [Your Name]" or other account folders)
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

# Function to clean subject lines by removing prefixes like "RE:" and "FW:"
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
def group_similar_subjects(subjects, distance_threshold=0.5, email_indices=None):
    # Vectorize the subjects
    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform(subjects)

    # Calculate pairwise distances
    distance_matrix = pairwise_distances(X.toarray(), metric='euclidean')

    # Apply Agglomerative clustering
    clustering = AgglomerativeClustering(distance_threshold=distance_threshold, n_clusters=None)
    clustering.fit(distance_matrix)

    # Group subjects by their cluster labels
    clustered_subjects = {}
    for idx, label in enumerate(clustering.labels_):
        if label not in clustered_subjects:
            clustered_subjects[label] = []
        # Append tuple of (subject, email index)
        clustered_subjects[label].append((subjects[idx], email_indices[idx]))
    
    return clustered_subjects

# Function to get a user-specified maximum number of emails from the folder within a time window, ignoring flagged emails
def get_emails_from_folder(folder, max_emails, max_days_back):
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by latest first

    # Calculate the time window (e.g., last X days)
    time_window = datetime.now().replace(tzinfo=None) - timedelta(days=max_days_back)

    compiled_emails = []
    count = 0
    for message in messages:
        if count >= max_emails:
            break  # Stop if we've reached the maximum number of emails requested
        
        try:
            received_time = message.ReceivedTime
            received_time = received_time.replace(tzinfo=None)  # Make the received time offset-naive
            
            # Check if the email was received within the time window
            if received_time >= time_window:
                # Check if the email is not flagged
                is_flagged = False
                try:
                    is_flagged = message.FlagStatus == 1 or message.IsMarkedAsTask
                except Exception as e:
                    print(f"Error checking flagged status: {e}")

                if not is_flagged:
                    # Extract email details
                    sender = message.SenderName
                    subject = clean_subject(message.Subject)  # Clean subject line
                    body = message.Body[:200]  # Get first 200 characters of the body as a snippet
                    entry_id = message.EntryID  # Get the unique EntryID for the message
                    store_id = folder.StoreID  # StoreID is needed for later to open the email

                    # Add details to the compiled list
                    compiled_emails.append({
                        "Sender": sender,
                        "Subject": subject,
                        "ReceivedTime": received_time,
                        "Body": body,  # Snippet of the body
                        "EntryID": entry_id,
                        "StoreID": store_id
                    })
                    
                    count += 1  # Only count unflagged emails

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
        selected_email = retrieved_emails[email_index - 1]  # Email indexing starts at 1 in the print output
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

# Main script logic
if __name__ == "__main__":
    # Step 1: Select folder
    selected_folder = select_folder()

    # Step 2: Ask for maximum number of emails
    try:
        max_emails = int(input("What is the maximum number of unflagged emails you want to retrieve? "))
    except ValueError:
        print("Invalid number entered. Defaulting to 10 emails.")
        max_emails = 10  # Default to 10 emails if invalid input

    # Step 3: Ask for the maximum number of days back (time window)
    try:
        max_days_back = int(input("Retrieve emails from how many days back (maximum)? "))
    except ValueError:
        print("Invalid number entered. Defaulting to 7 days.")
        max_days_back = 7  # Default to 7 days if invalid input

    # Step 4: Get emails from the selected folder
    retrieved_emails = get_emails_from_folder(selected_folder, max_emails, max_days_back)

    # Step 5: Update categories for retrieved emails
    if retrieved_emails:
        update_email_category(retrieved_emails)  # Update the category to "Blue Category"

        # Preprocess and group similar subjects
        subjects = [email["Subject"] for email in retrieved_emails]
        email_indices = [idx + 1 for idx in range(len(retrieved_emails))]  # Create email indices
        processed_subjects = preprocess_text(subjects)

        # Step 6: Print sorted emails
        print_sorted_emails(retrieved_emails)

        # Step 7: Group similar subjects
        clustered_subjects = group_similar_subjects(processed_subjects, email_indices=email_indices)

        # Print clustered subjects at the end
        for cluster, subjects in clustered_subjects.items():
            print(f"\nCluster {cluster + 1}:")
            for subject, email_index in subjects:
                print(f"  {subject} (Email Number: {email_index})")

        # Step 8: View emails if needed
        while True:
            view_email = input("Enter the email number to view it in full (or 'n' to exit): ")
            if view_email.lower() == 'n':
                break
            try:
                email_index = int(view_email)
                open_email(email_index)
            except ValueError:
                print("Please enter a valid email number or 'n' to exit.")
    else:
        print("No unflagged emails found.")
