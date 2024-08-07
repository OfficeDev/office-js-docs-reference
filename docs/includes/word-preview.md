| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[insertContentControl(contentControlType?: Word.ContentControlType.richText \| Word.ContentControlType.plainText \| Word.ContentControlType.checkBox \| Word.ContentControlType.dropDownList \| Word.ContentControlType.comboBox \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox")](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Wraps the Body object with a content control.|
||[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
|[Border](/javascript/api/word/word.border)|[color](/javascript/api/word/word.border#word-word-border-color-member)|Specifies the color for the border.|
||[location](/javascript/api/word/word.border#word-word-border-location-member)|Gets the location of the border.|
||[type](/javascript/api/word/word.border#word-word-border-type-member)|Specifies the border type for the border.|
||[visible](/javascript/api/word/word.border#word-word-border-visible-member)|Specifies whether the border is visible.|
||[width](/javascript/api/word/word.border#word-word-border-width-member)|Specifies the width for the border.|
|[BorderCollection](/javascript/api/word/word.bordercollection)|[getByLocation(borderLocation: Word.BorderLocation.top \| Word.BorderLocation.left \| Word.BorderLocation.bottom \| Word.BorderLocation.right \| Word.BorderLocation.insideHorizontal \| Word.BorderLocation.insideVertical \| "Top" \| "Left" \| "Bottom" \| "Right" \| "InsideHorizontal" \| "InsideVertical")](/javascript/api/word/word.bordercollection#word-word-bordercollection-getbylocation-member(1))|Gets the border that has the specified location.|
||[getFirst()](/javascript/api/word/word.bordercollection#word-word-bordercollection-getfirst-member(1))|Gets the first border in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.bordercollection#word-word-bordercollection-getfirstornullobject-member(1))|Gets the first border in this collection.|
||[getItem(index: number)](/javascript/api/word/word.bordercollection#word-word-bordercollection-getitem-member(1))|Gets a Border object by its index in the collection.|
||[insideBorderColor](/javascript/api/word/word.bordercollection#word-word-bordercollection-insidebordercolor-member)|Specifies the 24-bit color of the inside borders.|
||[insideBorderType](/javascript/api/word/word.bordercollection#word-word-bordercollection-insidebordertype-member)|Specifies the border type of the inside borders.|
||[insideBorderWidth](/javascript/api/word/word.bordercollection#word-word-bordercollection-insideborderwidth-member)|Specifies the width of the inside borders.|
||[items](/javascript/api/word/word.bordercollection#word-word-bordercollection-items-member)|Gets the loaded child items in this collection.|
||[outsideBorderColor](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsidebordercolor-member)|Specifies the 24-bit color of the outside borders.|
||[outsideBorderType](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsidebordertype-member)|Specifies the border type of the outside borders.|
||[outsideBorderWidth](/javascript/api/word/word.bordercollection#word-word-bordercollection-outsideborderwidth-member)|Specifies the width of the outside borders.|
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
|[Document](/javascript/api/word/word.document)|[compare(filePath: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-compare-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
|[DocumentCompareOptions](/javascript/api/word/word.documentcompareoptions)|[addToRecentFiles](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-addtorecentfiles-member)|True adds the document to the list of recently used files on the File menu.|
||[authorName](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-authorname-member)|The reviewer name associated with the differences generated by the comparison.|
||[compareTarget](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-comparetarget-member)|The target document for the comparison.|
||[detectFormatChanges](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-detectformatchanges-member)|True (default) for the comparison to include detection of format changes.|
||[ignoreAllComparisonWarnings](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-ignoreallcomparisonwarnings-member)|True compares the documents without notifying a user of problems.|
||[removeDateAndTime](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removedateandtime-member)|True removes date and time stamp information from tracked changes in the returned Document object.|
||[removePersonalInformation](/javascript/api/word/word.documentcompareoptions#word-word-documentcompareoptions-removepersonalinformation-member)|True removes all user information from comments, revisions, and the properties dialog box in the returned Document object.|
|[DropDownListContentControl](/javascript/api/word/word.dropdownlistcontentcontrol)|[addListItem(displayText: string, value?: string, index?: number)](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-addlistitem-member(1))|Adds a new list item to this dropdown list content control and returns a Word.ContentControlListItem object.|
||[deleteAllListItems()](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-deletealllistitems-member(1))|Deletes all list items in this dropdown list content control.|
||[listItems](/javascript/api/word/word.dropdownlistcontentcontrol#word-word-dropdownlistcontentcontrol-listitems-member)|Gets the collection of list items in the dropdown list content control.|
|[Field](/javascript/api/word/word.field)|[showCodes](/javascript/api/word/word.field#word-word-field-showcodes-member)|Specifies whether the field codes are displayed for the specified field.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Gets the Base64-encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Sets the picture at the specified level in the list.|
|[ListLevel](/javascript/api/word/word.listlevel)|[alignment](/javascript/api/word/word.listlevel#word-word-listlevel-alignment-member)|Specifies the horizontal alignment of the list level.|
||[font](/javascript/api/word/word.listlevel#word-word-listlevel-font-member)|Gets a Font object that represents the character formatting of the specified object.|
||[linkedStyle](/javascript/api/word/word.listlevel#word-word-listlevel-linkedstyle-member)|Specifies the name of the style that's linked to the specified list level object.|
||[numberFormat](/javascript/api/word/word.listlevel#word-word-listlevel-numberformat-member)|Specifies the number format for the specified list level.|
||[numberPosition](/javascript/api/word/word.listlevel#word-word-listlevel-numberposition-member)|Specifies the position (in points) of the number or bullet for the specified list level object.|
||[numberStyle](/javascript/api/word/word.listlevel#word-word-listlevel-numberstyle-member)|Specifies the number style for the list level object.|
||[resetOnHigher](/javascript/api/word/word.listlevel#word-word-listlevel-resetonhigher-member)|Specifies the list level that must appear before the specified list level restarts numbering at 1.|
||[startAt](/javascript/api/word/word.listlevel#word-word-listlevel-startat-member)|Specifies the starting number for the specified list level object.|
||[tabPosition](/javascript/api/word/word.listlevel#word-word-listlevel-tabposition-member)|Specifies the tab position for the specified list level object.|
||[textPosition](/javascript/api/word/word.listlevel#word-word-listlevel-textposition-member)|Specifies the position (in points) for the second line of wrapping text for the specified list level object.|
||[trailingCharacter](/javascript/api/word/word.listlevel#word-word-listlevel-trailingcharacter-member)|Specifies the character inserted after the number for the specified list level.|
|[ListLevelCollection](/javascript/api/word/word.listlevelcollection)|[getFirst()](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-getfirst-member(1))|Gets the first list level in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-getfirstornullobject-member(1))|Gets the first list level in this collection.|
||[items](/javascript/api/word/word.listlevelcollection#word-word-listlevelcollection-items-member)|Gets the loaded child items in this collection.|
|[ListTemplate](/javascript/api/word/word.listtemplate)|[listLevels](/javascript/api/word/word.listtemplate#word-word-listtemplate-listlevels-member)|Gets a ListLevels collection that represents all the levels for the specified ListTemplate.|
||[outlineNumbered](/javascript/api/word/word.listtemplate#word-word-listtemplate-outlinenumbered-member)|Specifies whether the specified ListTemplate object is outline numbered.|
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
|[Shading](/javascript/api/word/word.shading)|[foregroundPatternColor](/javascript/api/word/word.shading#word-word-shading-foregroundpatterncolor-member)|Specifies the color for the foreground of the object.|
||[texture](/javascript/api/word/word.shading#word-word-shading-texture-member)|Specifies the shading texture of the object.|
|[Style](/javascript/api/word/word.style)|[borders](/javascript/api/word/word.style#word-word-style-borders-member)|Specifies a BorderCollection object that represents all the borders for the specified style.|
||[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
||[listTemplate](/javascript/api/word/word.style#word-word-style-listtemplate-member)|Gets a ListTemplate object that represents the list formatting for the specified Style object.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[alignment](/javascript/api/word/word.tablestyle#word-word-tablestyle-alignment-member)|Specifies the table's alignment against the page margin.|
||[allowBreakAcrossPage](/javascript/api/word/word.tablestyle#word-word-tablestyle-allowbreakacrosspage-member)|Specifies whether lines in tables formatted with a specified style break across pages.|
