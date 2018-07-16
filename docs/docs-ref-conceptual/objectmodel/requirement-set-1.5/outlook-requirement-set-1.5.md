# Outlook add-in API requirement set 1.5

The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.

> [!NOTE]
> This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.

## What's new in 1.5?

Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.

- Added support for [pinnable taskpanes](https://docs.microsoft.com/outlook/add-ins/manifests/pinnable-taskpane).
- Added support for calling [REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Added ability to mark an attachment as inline.
- Added ability to close a taskpane or dialog.

### Change log

- Added [Office.context.mailbox.addHandlerAsync](/Office-mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.
- Added [Office.EventType](/javascript/api/office/office.eventtype): Specifies the event associated with an event handler.
- Added [Office.context.mailbox.restUrl](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox?product=outlook&version=v1.5#resturl-string): Gets the URL of the REST endpoint for this email account.
- Modified [Office.context.mailbox.getCallbackTokenAsync](/Office-mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.
- Added [Office.context.ui.closeContainer](/javascript/api/office/office.officeui): 
- Modified [Office.context.mailbox.item.addFileAttachmentAsync](/Office-item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyAllForm](/Office-item.md#displayreplyallformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.
- Modified [Office.context.mailbox.item.displayReplyForm](/Office-item.md#displayreplyformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.

## See also

- [Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook add-in code samples](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Get started](https://docs.microsoft.com/outlook/add-ins/quick-start)