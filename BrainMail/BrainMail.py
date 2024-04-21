# ---META---
print("----- starting program -----")

# Imports
import os
import time
import logging
import win32com.client
from openai import OpenAI

# Configure logging
logging.basicConfig(filename='logfile.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Master switches for printing and logging
print_switch = True
log_switch = False

# (FUNCTION) Print and log together
def printlog(message):
    # Print the message if print_switch is True
    if print_switch:
        print(message)
        
    # Log the message if log_switch is True
    if log_switch:
        logging.info(message)

# ---PARAMETERS---

# Indicate parameter loading
printlog("--- loading parameters ---")
       
# File path to the txt file containing the recipients
recipients_file_path = 'recipients.txt'

# File path to the txt file containing the API key
api_file_path = 'api.txt'

# File path to the txt file containing the basic instructions
basic_instructions_file_path = 'basic_instructions.txt'

# ---INITIALIZATION---
# (1) List of BrainMail recipients
# Create list of BrainMail recipients and BrainMail recipient properties
bm_recipients = []
bm_recipient_properties = {}

# Read recipients from the txt file and add them to the list created above
with open(recipients_file_path, 'r') as file:
    for line in file:
        # Split the line into recipient email and properties
        items = line.strip().split(',')
        # Extract recipient email address
        email_address = items[0]
        # Extract recipient properties
        properties = items[1:]
        # Store recpients and recipient properties in the dictionary
        bm_recipient_properties[email_address] = properties

# Fill BrainMail recipients list with email addresses from the txt file
bm_recipients = list(bm_recipient_properties.keys())

# List recipients
printlog("List of Recipients:")
recipient_number = 1
for recipient, properties in bm_recipient_properties.items():
    printlog(f"Recipient {recipient_number}: {recipient}, Properties: {properties}")
    recipient_number = recipient_number +1

# (2) API Key
# > (FUNCTION) Read the API key from a txt file
def read_api_key_from_file(api_file_path):
    try:
        # Open the file in read mode
        with open(api_file_path, 'r') as file:
            # Read the content of the file
            api_key = file.read().strip()
            return api_key
    except FileNotFoundError:
        printlog("Text file containing API key not found.")
        return None

# >> (FUNCTION CALL) Define API key variable
api_key = read_api_key_from_file(api_file_path)

# >>> (CHECK) Was the API key successfully read?
if api_key:
    printlog(f"API key found: {api_key}")

    # Connect to the OpenAI API
    client = OpenAI(api_key=api_key)
else:
    printlog("API key not found. Make sure the file api.txt exists.")

# (3) Basic Instructions
# > (FUNCTION) Read the basic instruction from a txt file
def read_basic_instructions_from_file(basic_instructions_file_path):
    try:
        with open(basic_instructions_file_path, 'r') as file:
            # Den Inhalt der Datei lesen und in die Variable basic_instructions setzen
            basic_instructions = file.read().strip()
            return basic_instructions
    except FileNotFoundError:
        printlog("Text file containing basic instructions not found.")
        return None

# >> (FUNCTION CALL) Define basic instructions variable
basic_instructions = read_basic_instructions_from_file(basic_instructions_file_path)

# >>> (CHECK) Were the basic instructions successfully read?
if basic_instructions:
    print("Basic instructions found:", basic_instructions)
else:
    print("Basic instructions not found. Make sure the file basic_instructions.txt exists.")

# (4) Unread Emails
# (FUNCTION) Count the number of unread emails
def count_unread_emails(folder):
    # Filter unread emails
    unread_messages = folder.Items.Restrict("[Unread] = True")

    # Return number of unread emails
    return unread_messages.Count

# (5) OpenAI Response
# (FUNCTION) Get the OpenAI response
def get_openai_response(individual_instructions, email_content):
    # Chat completion function
    full_instructions = "Instructions:" + "" + basic_instructions + " " + individual_instructions + "\n\n" + "The Email: " + email_content
    # printlog("--- full instructions start ---" + "\n\n" + full_instructions)
    chat_completion = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": full_instructions,
            }
        ],
        model="gpt-3.5-turbo", #The GPT model, change with updates
    )
    # printlog("--- full instructons end --- auto reply start ---" + "\n\n" + (chat_completion.choices[0].message.content) + "\n\n" + "--- auto reply end ---")
    return chat_completion.choices[0].message.content

# (6) Auto Reply
# (FUNCTION) Generate the auto reply
def auto_reply():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6) # 6 represents the Inbox
    printlog("--- starting auto reply process ---")

    # Try to retrieve the "BrainMail" folder
    try:
        brainmail_folder = inbox.Folders("BrainMail")
    except Exception as e:
        # If the folder is not found, create it in the Inbox
        brainmail_folder = inbox.Folders.Add("BrainMail")
    
    # Count the number of unread emails before starting the loop
    unread_count = count_unread_emails(inbox)
    logging.info(f"Number of unread emails at start: {unread_count}")
    printlog(f"Number of unread emails at start: {unread_count}")

    # Retrieve all unread emails
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    messages = messages.Restrict("[Unread] = True")

    # Create list to store unread emails
    unread_emails = []
    for message in messages:
        unread_emails.append(message)
        
    # Set message counter to 1
    message_counter = 1    

    # Process all unread emails
    for message in unread_emails:
        
        # Get sender email address and subject of the email
        sender_email = message.SenderEmailAddress
        subject = message.Subject
        # Get email content
        email_content = message.Body  

        # Check if the sender is in the brainmail recipients list
        if sender_email not in bm_recipients:
            # Log and print information about the email
            printlog(f"({message_counter}) Sender: {sender_email}, Subject: {subject}" + "\n" + f"{sender_email} is not a BrainMail recipient! No action performed.")
            # Increase message counter
            message_counter = message_counter + 1  

        else:
            # Printlog information about the email
            printlog(f"({message_counter}) Sender: {sender_email}, Subject: {subject}" + "\n" + f"{sender_email} is a BrainMail recipient! Generating automatic reply.")

            # Extract individual instructions from the dictionary (convert to string, strip characters)
            individual_instructions = str(bm_recipient_properties[sender_email])
            individual_instructions = individual_instructions.strip("[']")
            print(f"These individual instructions were found: {individual_instructions}")

            # Generate automatic reply based on email content
            openai_response = get_openai_response(individual_instructions, email_content)

            # Add OpenAI response to the reply body
            reply_body = f"{openai_response}"
            reply = message.Reply()
            reply.Body = reply_body

            # Send automatic reply
            reply.Send()
                    
            # Mark email as read
            message.Unread = False

            # Move email to the target folder
            message.Move(brainmail_folder)
             
            # Increase message counter
            message_counter = message_counter + 1  

    # End loop after all unread emails are processed
    printlog("--- auto reply completed ---")
    printlog("All unread emails checked, all unread emails by BrainMail recipients processed.")

# Run script once but only if it is not imported into another script
# if __name__ == "__main__":
#     auto_reply()

# Run script every x seconds

# Define wait time in seconds
wait_time_seconds = 30
    
while True:
    # Starting time of current run
    start_time = time.time()

    # Start AutoReply function
    auto_reply()

    # Calculate remainign time until next runstarts
    remaining_time = wait_time_seconds

    # Loop to print remaing time every 30 seconds
    while remaining_time > 0:

        # Print remaining time
        print(f"Next run in {int(remaining_time)} seconds")

        # Wait for 30 seconds or the remaining time,whichever is shorter
        time.sleep(min(10, remaining_time))

        # Update remaining time
        remaining_time -= 10