# Outlook add-in API requirement set 1.1

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> **Note**: This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set. 

## What's new in 1.1?

Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.

### Change log

- Added [Body](/javascript/api/office_1_1/office.Body) object: Provides methods for adding and updating the content of an item in an Outlook add-in.
- Added [Location](/javascript/api/office_1_1/office.Location) object: Provides methods to get and set the location of a meeting in an Outlook add-in.
- Added [Recipients](/javascript/api/office_1_1/office.Recipients) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.
- Added [Subject](/javascript/api/office_1_1/office.Subject) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
- Added [Time](/javascript/api/office_1_1/office.Time) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.
- Added [Office.context.mailbox.item.addFileAttachmentAsync](/Office-item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.
- Added [Office.context.mailbox.item.addItemAttachmentAsync](/Office-item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.
- Added [Office.context.mailbox.item.removeAttachmentAsync](/Office-item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.
- Added [Office.context.mailbox.item.body](/Office-item.md#body-body): Gets an object that provides methods for manipulating the body of an item.
- Added [Office.context.mailbox.item.bcc](/Office-item.md#bcc-recipients): Gets or sets the recipients on the Bcc (blind carbon copy) line of a message.
- Added [Office.MailboxEnums.RecipientType](/javascript/api/office_1_1/office.mailboxenums.recipienttype): Specifies the type of recipient for an appointment.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)