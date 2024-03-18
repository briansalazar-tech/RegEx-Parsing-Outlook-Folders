import win32com.client, re, datetime, csv
from datetime import date, timedelta

SAVE_DIRECTORY = ""
today = date.today()
date_delta = str(today - timedelta(days=1)) # Time delta checks specific date. Could be yesterday, or any day.

def email_parsing(search, email_body):
    """
    Function searches for digits that appear after the specified search entry. 
    If a number is found, that number is returned, else 'not in email' is returned.
    """

    # Keyword is searched for after numbers
    if search == " starts with a space":
        search_desktop_regex = re.compile(r"(\d)*" + search)
        search_desktop_search = re.search(search_desktop_regex, email_body)
        if search_desktop_search == None:
            search_desktop_value = "Not in email"
        else:
            search_desktop_string = str(search_desktop_search.group())
            search_desktop_value = search_desktop_string.split(search)[0]

    # Keyword is searched for before numbers
    else:
        search_desktop_regex = re.compile(search + r"(\d)*")
        search_desktop_search = re.search(search_desktop_regex, email_body)
        if search_desktop_search == None:
            search_desktop_value = "Not in email"
        else:
            search_desktop_string = str(search_desktop_search.group())
            search_desktop_value = search_desktop_string.split(search)[1]

    return search_desktop_value

### EMAIL SETUP ###
# Outlook Setup
outlook = win32com.client.Dispatch("outlook.application")
mapi = outlook.GetNamespace("MAPI")
root_folder = mapi.Folders["youroutlook@email.com"] # Access email folder
foldername = root_folder.Folders("Folder Name") # Access desired subfolder

emails_in_folder = 0

### ITTERATE THROUHG EMAIL MESSAGES ###
for item in foldername.Items:
    emails_in_folder += 1
    messages = foldername.Items

    message = messages.Item(emails_in_folder)
    subject = message.subject
    date = message.SentOn.strftime("%m/%d/%y")
    body = message.body
    
    # Keywords to serach for in emails body
    key_words = {
        " starts with a space": "",
        "Phrase 1: ": "",
        "Phrase 2: ": "",
        # "Additional searches: ": "", 
    }
    
    # Gets the date portion from SentOn and converts it to a string to compair against time delta.
    emaildate = str(message.SentOn).split()[0]

    if emaildate >= date_delta: # Could alternatively compair before or a certain time range.
        print(f"Match: {subject}")
        
        # Pass keywords through the Regex function
        for search_phrase in key_words:
                #print(f"Key in passed: {search_phrase}")
                key_words[search_phrase] = email_parsing(search=search_phrase, email_body=body)
        
        ### CREATE AND SAVE TO CSV ###
        # Save location for csv - Creates a file if it does not exist and appends each match to the file.
        csv_report_file = open(f"{SAVE_DIRECTORY}FILENAME-{datetime.date.today()}.csv", mode = "a", newline = "")
        csv_writer = csv.writer(csv_report_file, delimiter = ",")
        csv_writer.writerow([
                subject, # Column A
                date, # Column B
                # key_words[" starts with a space"], # Additional  columns...
                ])
        csv_report_file.close()

print(f"Complete. Emails searched: {emails_in_folder}")