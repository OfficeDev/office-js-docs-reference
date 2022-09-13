| Class | Fields | Description |
|:---|:---|:---|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[onCommentAdded](/javascript/api/word/word.commentcollection#word-word-commentcollection-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.commentcollection#word-word-commentcollection-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.commentcollection#word-word-commentcollection-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.commentcollection#word-word-commentcollection-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.commentcollection#word-word-commentcollection-oncommentselected-member)|Occurs when a comment is selected.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Occurs when the content control is deleted.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Occurs when selection within the content control is changed.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|The event type.|
|[Document](/javascript/api/word/word.document)|[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Occurs when a content control is added.|
|[Field](/javascript/api/word/word.field)|[delete()](/javascript/api/word/word.field#word-word-field-delete-member(1))|Deletes the field.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Sets the picture at the specified level in the list.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
