# Outlook Attachments Downloader

This is a CLI program written in C#,  in order to learn the C# synthax, and manipulate the Microsoft Office Interop Outlook library. 

# What does Outlook Attachments Downloader ?

OAD will fetch all your accounts from your Outlook application, and return them with all folders and subfolders. Then you will have the possibility to select several folders and get all attachments in a specified folder.

## Improvements

I have some ideas like download only attachments (exclude pictures included in mails), select or exclude files according to a regex pattern.

## Possibles issues
This is a proof of concept, I only tested this with small inbox folders ( less than 200 attachments), so bugs can occured.

Note : attachments saved or located in OneDrive cannot be fetched.
