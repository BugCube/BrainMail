# BrainMail

## Python Script that uses OpenAI to automatically answer emails in Outlook

This script uses OpenAI to generate individualized answers to incoming emails from addresses that you specify (e.g. friends, students, your mother-in-law...)

The answers are build by promting OpenAI with (1) basic instructions that are of interest for all senders, (1) individual instructions that you specify for each sender, and (3) the original email from the sender.

You will need an OpenAI account and your OpenAI API key for this. Token usage is low.

The script generates a new folder in Outlook named "BrainMail". Emails from recipents on the list will be moved there, marked as read, and an answer will be send.

Step 1: Download the code.

Step 2: Specify auto reply recipient list in the recipients.txt file in the form <email address>,<individual instructions>.

Step 3: Enter your API key in the api.txt file in the form <API key>

Step 4: Enter your basic instructions in the basic_instructions.txt file in the form <basic instructions> (e.g. I am out of office until the end of May.)

Step 5: Consider if you want logging. If yet, set the respective varible to 'True'.

Step 6: Ativate the script. Per default, it runs through your unread Outlook emails once every 30 seconds. If you want the script to run on demand, just meddle with the comments at the bottom.

Do not judge too hard, this is my first attempt at doing something useful.