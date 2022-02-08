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

[Closed Add-In]()
```python
import foobar

# returns 'words'
foobar.pluralize('word')

# returns 'geese'
foobar.pluralize('goose')

# returns 'phenomenon'
foobar.singularize('phenomena')
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
License???  What license?  