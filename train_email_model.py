import win32com.client
from sklearn.feature_extraction.text import TfidfVectorizer
import pickle
import re

# Recursive function to list all folders and their subfolders
def list_folders(folder, folder_list, parent_name=""):
    folder_name = f"{parent_name}/{folder.Name}" if parent_name else folder.Name
    folder_list.append(folder_name)
    
    # Recursively add subfolders
    for i in range(folder.Folders.Count):
        sub_folder = folder.Folders.Item(i + 1)
        list_folders(sub_folder, folder_list, folder_name)

# Function to allow the user to select from all available folders
def select_folder():
    folder_list = []
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # List all top-level folders
    for i in range(outlook.Folders.Count):
        top_folder = outlook.Folders.Item(i + 1)
        list_folders(top_folder, folder_list)  # Recursively list all subfolders

    # Display all folders for selection
    print("\nAvailable Folders:")
    for index, folder_name in enumerate(folder_list):
        print(f"{index + 1}: {folder_name}")

    # Input folder selection from the user
    try:
        folder_index = int(input("\nEnter the number of the folder you want to select: ")) - 1
        selected_folder = folder_list[folder_index]
        print(f"You selected: {selected_folder}")
    except (ValueError, IndexError):
        print("Invalid selection. Exiting.")
        return None
    
    # Get the folder object from the selected folder name
    selected_folder_obj = outlook.Folders[selected_folder.split('/')[0]]
    for subfolder in selected_folder.split('/')[1:]:
        selected_folder_obj = selected_folder_obj.Folders[subfolder]

    return selected_folder_obj

# Function to get emails from a specified folder
def get_emails_from_folder(folder, max_emails):
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by latest first

    compiled_emails = []
    count = 0
    for message in messages:
        if count >= max_emails:
            break  # Stop if we've reached the maximum number of emails requested
        
        try:
            # Extract email details
            subject = clean_subject(message.Subject)  # Clean subject line
            
            # Add subject to the compiled list
            compiled_emails.append(subject)
            count += 1  # Count every email

        except Exception as e:
            print(f"Error processing email: {e}")

    return compiled_emails

# Function to clean subject lines by removing prefixes like "RE:" and "FW:"
def clean_subject(subject):
    cleaned_subject = re.sub(r'^\s*(RE:|FW:|bid invite:)\s*', '', subject, flags=re.IGNORECASE)
    return cleaned_subject.strip()

# Function to save the TF-IDF vectorizer to a file
def save_vectorizer(vectorizer, filename='vectorizer.pkl'):
    with open(filename, 'wb') as file:
        pickle.dump(vectorizer, file)
    print(f"Vectorizer saved to {filename}.")

# Main script logic
if __name__ == "__main__":
    # Step 1: Select folder
    selected_folder = select_folder()
    if not selected_folder:
        exit(1)  # Exit if no valid folder is selected

    # Step 2: Ask for maximum number of emails
    try:
        max_emails = int(input("How many recent emails do you want to process? "))
    except ValueError:
        print("Invalid number entered. Defaulting to 10 emails.")
        max_emails = 10  # Default to 10 emails if invalid input

    # Step 3: Get emails from the selected folder
    retrieved_subjects = get_emails_from_folder(selected_folder, max_emails)

    # Step 4: Create and fit the TF-IDF vectorizer
    if retrieved_subjects:
        vectorizer = TfidfVectorizer()
        vectorizer.fit(retrieved_subjects)

        # Step 5: Save the vectorizer to a file
        save_vectorizer(vectorizer)
    else:
        print("No emails found.")
