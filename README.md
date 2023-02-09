# DEPRECATED - NOT WORKING 

Recent versions of Node have caused issues with security in older Node applications. Given that and given the potential to move away from DevOps, I am deprecating this app - JMH

# Outlook Azure Add-in

The Outlook Azure Add-in is a sideloaded plugin that allows users to create 'Task' or 'Bug' DevOps ticket using a received email.
During the creation of the ticket, users can select whether to include attachments, assign the ticket, and set the title of the ticket. 
The email (with inline images embedded) is passed to the ticket as both a description and an attached html file.  Any other attachments are 
attached to the ticket as a regular attachment.  

## Installation

To install you will need to do the following:

1.  From Outlook, select ***File** and then ***Manage Add-Ins**: 

![Outlook Image](/MicrosoftTeams-image.png)

2. The next step will be to select ***My Add-Ins** and then ***Add Custom Add-In**:

![Outlook Image](/MicrosoftTeams-imageb.png)

You will want to install the [manifest.prod.xml](/dist/manifest.prod.xml) file to activate the add-in.

## Usage

The add-in is based upon Outlook's ReadMessage functions.  Because it's based on reading your Outlook message, the add-in will not activate until a message is selected:

![Closed Add-In](/closed-add-in.png)

Once a message is selected the Add-In becomes active:

![Open Add-in](/outlook-add-in-open.png)

To use the Add-in, just provide the information which the form is asking and a new ***Task** or ***Bug** ticket will be created:

![Add-In](/app-screenshot.png)

## Support
Issues with the app?  Why not just create a ticket using the app?

## Roadmap
There are several changes that I am looking at in terms of future versions:
- Updating the API calls to the Microsoft Graph API for better access to user accounts 
- Rebuilding the app within a C# environment to allow for more customization than exists in the Yeoman/Office.js file
- Including the ability to add an additional description to the main description

## License
License???  What license?  