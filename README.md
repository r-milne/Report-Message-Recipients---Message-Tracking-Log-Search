# Report Message Recipients - Message Tracking Log Search
 Report Message Recipients - Message Tracking Log Search

Purpose of this script to search the Exchange message tracking log to determine who received a copy of a given email. 
This may be useful if an organisation is the target of a phishing or spam campaign.  In such a case you may want to obtain a list of all of  the people who received a copy of that particular email.
 

Output format is a simple CSV which a column for the recipient, and a column for the message subject. 
The CSV can then be used as the input file to drive Search-Mailbox or to email users that they were the victim of a phish and not to open the offending message.
 

Since the message tracking logs are used by the Search-MessageTrackingLog cmdlet, it is assumed that the script is used whilst the logs still contain the necessary data
By default 30 days of tracking log data are retaining, though this may have been modified.
 

If needed, Search-Mailbox could be leveraged as an alternative though this script is not written for that cmdlet. 
 

An empty array is declared that will be used to hold the data gathered during each iteration.

This allows for the additional information to be easily added on, and then either echo it to the screen or export to a CSV file
 

A custom PSObject is used so that we can add data to it from various sources, Get-Mailbox, Get-MailboxStatistics, Get-ADUser etc.
There is no limit to your creativity! 
 

The CSV is created in the $PWD which is the Present Working Directory, i.e. where the script is saved
 

Please refer to this blog post for details:
"to be published>"
