| Class | Fields | Description |
|:---|:---|:---|
|[Mailbox](/javascript/api/outlook/outlook.mailbox)|[addHandlerAsync(eventType: Office.EventType \| string, handler: (type: Office.EventType) => void, options?: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.mailbox#addhandlerasync-eventtype--handler--type-)|Adds an event handler for a supported event. **Note**: Events are available only with task pane.|
||[getCallbackTokenAsync(options: Office.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: Office.AsyncResult<string>) => void)](/javascript/api/outlook/outlook.mailbox#getcallbacktokenasync-options--isrest--callback--asyncresult-)|Gets a string that contains a token used to call REST APIs or Exchange Web Services (EWS).|
||[isRest](/javascript/api/outlook/outlook.mailbox#isrest)||
||[removeHandlerAsync(eventType: Office.EventType \| string, options?: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.mailbox#removehandlerasync-eventtype--options--callback--asyncresult-)|Removes the event handlers for a supported event type. **Note**: Events are available only with task pane.|
||[restUrl](/javascript/api/outlook/outlook.mailbox#resturl)|Gets the URL of the REST endpoint for this email account.|
