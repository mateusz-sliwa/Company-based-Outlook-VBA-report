# Company-based-Outlook-VBA-report
In order to automate a process of generating an end-of-day shared mailbox report i've created a VBA Macro that counts how many
processed/unprocessed/breached e-mails there are for a specific date 
(in that case it's the current date of generating this message)
Macro, gets the shared mailbox by the name of it, then iterates through all items that have been recieved today, and sorts them into
specific categories: Total(overall number of e-mail messages), Processed(Replied e-mails), Unprocessed(not replied yet), Breached(Those without a response for longer than 2 days)
Once it has all the data, a notification with relevant informationa appears and an e-mail message template is generated to a recipient prearranged in code itself.
The entire process automates e-mail generation, and prevents desired data from being incorrect due possible employee counting mistakes.
