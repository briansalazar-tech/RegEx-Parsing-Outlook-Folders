Parsing Folders in the Windows Outlook desktop Application with Regular Expressions

I created these three scripts to gather data from emails that I received daily in my Outlook email. Those reports were all saved to a specified folder using email rules. Using regular expressions, I was able to save data and create reports based on the data that I gathered from those daily email.

For these scripts I saved three separate variations, but they can be further tinkered with to get any information that is useful from your Outlook email.

The first script (emailparsing_all) I created iterates through all the emails in the specified folder using keywords to be found in the email’s body. That data is then appended to a csv file. If the csv does not exist, a new one is created. 

The second script (emailparsing_daterange) essentially performs the same function as the first one but checks to see if the email was sent on or after the specified date. This can be changed to check for before a specified date or even a certain date range.

The third script (emailparsing_singledate) only appends data to a csv from a specified date. This script is useful if you want to schedule it to run daily instead of writing every single email or email for a particular range to a csv.
