| Class | Fields | Description |
|:---|:---|:---|
|[DragAndDropEventArgs](/.draganddropeventargs)|[dragAndDropEventData](/.draganddropeventargs#outlook-javascript/api/outlook/-draganddropeventargs-draganddropeventdata-member)|Gets the details about the mouse pointer position within an add-in's task pane and the messages or file attachments being dragged and dropped into the task pane.|
||[type](/.draganddropeventargs#outlook-javascript/api/outlook/-draganddropeventargs-type-member)|Gets the type of the event.|
|[DragoverEventData](/.dragovereventdata)|[pageX](/.dragovereventdata#outlook-javascript/api/outlook/-dragovereventdata-pagex-member)|Gets the x-coordinate of the mouse pointer that represents the horizontal position in pixels.|
||[pageY](/.dragovereventdata#outlook-javascript/api/outlook/-dragovereventdata-pagey-member)|Gets the y-coordinate of the mouse pointer that represents the vertical position in pixels.|
||[type](/.dragovereventdata#outlook-javascript/api/outlook/-dragovereventdata-type-member)|Gets the type of drag-and-drop event.|
|[DropEventData](/.dropeventdata)|[dataTransfer](/.dropeventdata#outlook-javascript/api/outlook/-dropeventdata-datatransfer-member)|Gets the messages or file attachments being dragged and dropped into an add-in's task pane.|
||[pageX](/.dropeventdata#outlook-javascript/api/outlook/-dropeventdata-pagex-member)|Gets the x-coordinate of the mouse pointer that represents the horizontal position in pixels.|
||[pageY](/.dropeventdata#outlook-javascript/api/outlook/-dropeventdata-pagey-member)|Gets the y-coordinate of the mouse pointer that represents the vertical position in pixels.|
||[type](/.dropeventdata#outlook-javascript/api/outlook/-dropeventdata-type-member)|Gets the type of drag-and-drop event.|
|[DroppedItemDetails](/.droppeditemdetails)|[fileContent](/.droppeditemdetails#outlook-javascript/api/outlook/-droppeditemdetails-filecontent-member)|Gets the contents of the file being dragged and dropped.|
||[name](/.droppeditemdetails#outlook-javascript/api/outlook/-droppeditemdetails-name-member)|Gets the name of the file being dragged and dropped.|
||[type](/.droppeditemdetails#outlook-javascript/api/outlook/-droppeditemdetails-type-member)|Gets the type of the file being dragged and dropped.|
|[DroppedItems](/.droppeditems)|[files](/.droppeditems#outlook-javascript/api/outlook/-droppeditems-files-member)|Gets an array of the messages or file attachments being dragged and dropped into an add-in's task pane.|
|[Mailbox](/.mailbox)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/.mailbox#outlook-javascript/api/outlook/-mailbox-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/.mailbox#outlook-javascript/api/outlook/-mailbox-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[getCallbackTokenAsync(options: Office.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: Office.AsyncResult<string>) => void)](/.mailbox#outlook-javascript/api/outlook/-mailbox-getcallbacktokenasync-member(1))|Gets a string that contains a token used to call REST APIs or Exchange Web Services (EWS).|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/.mailbox#outlook-javascript/api/outlook/-mailbox-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/.mailbox#outlook-javascript/api/outlook/-mailbox-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[restUrl](/.mailbox#outlook-javascript/api/outlook/-mailbox-resturl-member)|Gets the URL of the REST endpoint for this email account.|
