| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Wraps the Body object with a content control.|
||[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
|[ComboBoxContentControl](/javascript/api/word/word.comboboxcontentcontrol)|[addListItem(displayText: string, value?: string, index?: number)](/javascript/api/word/word.comboboxcontentcontrol#word-word-comboboxcontentcontrol-addlistitem-member(1))|Adds a new list item to this combo box content control and returns a Word.ContentControlListItem object.|
||[deleteAllListItems()](/javascript/api/word/word.comboboxcontentcontrol#word-word-comboboxcontentcontrol-deletealllistitems-member(1))|Deletes all list items in this combo box content control.|
||[listItems](/javascript/api/word/word.comboboxcontentcontrol#word-word-comboboxcontentcontrol-listitems-member)|Gets the collection of list items in the combo box content control.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[comboBoxContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-comboboxcontentcontrol-member)|Specifies the combo box-related data if the content control's type is 'ComboBox'.|
||[dropDownListContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-dropdownlistcontentcontrol-member)|Specifies the dropdown list-related data if the content control's type is 'DropDownList'.|
||[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
||[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlListItem](/javascript/api/word/word.contentcontrollistitem)|[delete()](/javascript/api/word/word.contentcontrollistitem#word-word-contentcontrollistitem-delete-member(1))|Deletes the list item.|
||[displayText](/javascript/api/word/word.contentcontrollistitem#word-word-contentcontrollistitem-displaytext-member)|Specifies the display text of a list item for a dropdown list or combo box content control.|
||[index](/javascript/api/word/word.contentcontrollistitem#word-word-contentcontrollistitem-index-member)|Specifies the index location of a content control list item in the collection of list items.|
||[select()](/javascript/api/word/word.contentcontrollistitem#word-word-contentcontrollistitem-select-member(1))|Selects the list item and sets the text of the content control to the value of the list item.|
||[value](/javascript/api/word/word.contentcontrollistitem#word-word-contentcontrollistitem-value-member)|Specifies the programmatic value of a list item for a dropdown list or combo box content control.|
|[ContentControlListItemCollection](/javascript/api/word/word.contentcontrollistitemcollection)|[getFirst()](/javascript/api/word/word.contentcontrollistitemcollection#word-word-contentcontrollistitemcollection-getfirst-member(1))|Gets the first list item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrollistitemcollection#word-word-contentcontrollistitemcollection-getfirstornullobject-member(1))|Gets the first list item in this collection.|
||[items](/javascript/api/word/word.contentcontrollistitemcollection#word-word-contentcontrollistitemcollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[DropDownListContentControl](/javascript/api/word/word.dropdownlistcontentcontrol)|[addListItem(displayText: string, value?: string, index?: number)](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-addlistitem-member(1))|Adds a new list item to this dropdown list content control and returns a Word.ContentControlListItem object.|
||[deleteAllListItems()](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-deletealllistitems-member(1))|Deletes all list items in this dropdown list content control.|
||[listItems](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-listitems-member)|Gets the collection of list items in the dropdown list content control.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[List](/javascript/api/word/word.list)|[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
|[Paragraph](/javascript/api/word/word.paragraph)|[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Wraps the Paragraph object with a content control.|
||[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Wraps the Range object with a content control.|
||[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
