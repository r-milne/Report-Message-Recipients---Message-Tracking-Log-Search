<# 

.SYNOPSIS
	Purpose of this script to search the Exchange message tracking log to determine who received a copy of a given email.  
	
	This may be useful if an organisation is the target of a phishing or spam campaign.  In such a case you may want to obtain a list of all of 
	the people who received a copy of that particular email.

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
   	



.DESCRIPTION
	All transport servers are queried.  This is collection of server objects is piped to the Get-MessageTrackingLog cmdlet so the organisation is searched.  If required a search filter could be used to
	limit the transport instances searched.
	The legacy Get-TransportServer cmdlet us used.  If you only have Exchange 2013 onwards you can use Get-TransportService instead.  

	Adjust the TrackingLongEntries variable to target the correct search information. 
	The Syntax is available 
	https://docs.microsoft.com/en-us/powershell/module/exchange/mail-flow/get-messagetrackinglog?view=exchange-ps 
	  
	An example looking for a particular message ID 
	$TrackingLogEntries = Get-TransportServer | Get-MessageTrackingLog -MessageId BLUPR05MB1873FA96E0FFD976C57ED335DE810@BLUPR05MB1873.namprd05.prod.outlook.com -EventId RECEIVE -ResultSize Unlimited
	
	Date range + sender example 
	$TrackingLogEntries = Get-TransportServer | Get-MessageTrackingLog -Start "03/13/2015 09:00:00" -End "03/15/2015 17:00:00" -Sender "dick@tailspintoys.ca.com"  -ResultSize Unlimited 
	
	  
	$TrackingLogEntries = Get-TransportServer | Get-MessageTrackingLog -MessageSubject "Subject goes here"   -ResultSize Unlimited 
		 
	The MessageSubject parameter filters the message tracking log entries by the value of the message subject. 
	The value of the MessageSubject parameter automatically supports partial matches without using wildcards or special characters. 
	For example, if you specify the MessageSubject value sea, the results include messages with Seattle in the subject. By default, message subjects are stored in the message tracking logs.
	
	Search by sender
	$TrackingLogEntries = Get-TransportServer | Get-MessageTrackingLog -Sender user-1@tailspintoys.ca  -ResultSize Unlimited 



.ASSUMPTIONS
	You are running this from the Exchange Management Shell

	Script is being executed with sufficient permissions to access the server(s) targeted. 

	You can live with the Write-Host cmdlets :) 

	You can add your error handling if you need it.  

	

.VERSION
  
	1.0  1-05-2018 -- Initial script released to the scripting gallery 

	This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
	THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
	INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
	We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, 
	provided that You agree: 
	(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
	(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
	(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
	Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within the Premier Customer Services Description.
	This posting is provided "AS IS" with no warranties, and confers no rights. 

	Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.

#>


# Clean up screen 
Clear-Host

# Define arrays in advance 
$Output = @()  
$MsgRecipients = @() 


# Declare a custom PS object. This is the template that will be copied multiple times. 
# This is used to allow easy manipulation of data from potentially different sources 
# Elements are arbitary names.  You can call them what you want.  Just use the names consistently.... 
$TemplateObject = New-Object PSObject | Select Recipient, MessageSubject,EventID

# Change the search to refine what/where you want to search.  See examples above. 
$TrackingLogEntries = Get-TransportServer | Get-MessageTrackingLog -MessageSubject "Tax Return Information" -EventId DELIVER -ResultSize Unlimited

# Use an array around the $TrackingLogEntries to work out the count.  Needed for later to display the progress bar
$Count = @($TrackingLogEntries).count

# Loop through all of the message tracking log entries which matched the search query 
ForEach($TrackingLogEntry in $TrackingLogEntries)
{

	# Write a handy dandy progress bar to the screen so that we know how far along this is...
	# Increment the counter 
	$Int = $Int + 1
	# Work out the current percentage 
	$Percent = $Int/$Count * 100

	# Write the progress bar out with the necessary verbiage....
	Write-Progress -Activity "Processing tracking log details" -Status "Processing entry $Int of $Count " -PercentComplete $Percent 

	# Note that there may be multiple recipients on a give message tracking log entry.  We need to expand this to be able to properly report.
	# This is an example output -- note the two names on the Recipients. 
	# EventId  Source   Sender                                      Recipients                                  MessageSubject
	# -------  ------   ------                                      ----------                                  --------------
	# RECEIVE  SMTP     User-1@tailspintoys.ca                      {Local-1@tailspintoys.ca, llocal@tailspi... Two Recipients

	# Save the message subject 
	$MsgSubject = $TrackingLogEntry.MessageSubject 
	# Display message subject to screen.  REM out if not reqired/wanted 
	# Write-Host "Processing Message Subject:$MsgSubject" -ForeGroundColor Magenta
	Write-Host

	# Save the message EventID
	$MsgEventID = $TrackingLogEntry.EventID 

	# Assume that there may be one or more recipients per message tracking log entry.  Save the recipients to the pre-defined array $MsgRecipients 
	$MsgRecipients  = $TrackingLogEntry | Select-Object  -ExpandProperty Recipients

	# Loop through the recipients of this particular message tracking log entry and append the message subject and then squirrel this to the output variable 
	ForEach ($MsgRecipient in $MsgRecipients)
	{
		# Make a copy of the TemplateObject.  Then work with the copy...
		$WorkingObject = $TemplateObject | Select-Object * 
		
		$WorkingObject.Recipient      = $MsgRecipient 
		$WorkingObject.MessageSubject = $MsgSubject 
		$WorkingObject.EventID	      = $MsgEventID

    		# Display output to screen.  REM out if not reqired/wanted 
		# $WorkingObject
		

		# Append  current results to final output
    		$Output += $WorkingObject

	}


}

# Output is written to a file in present working directory.  Edit as necessary. 
$Output | Export-Csv -Path $PWD\Output.csv -NoTypeInformation  



