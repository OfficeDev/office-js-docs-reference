| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[getParagraphById(id: string)](/javascript/api/word/word.body#word-word-body-getparagraphbyid-member(1))|Gets the paragraph by its unique id.|
||[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
||[onParagraphAdded](/javascript/api/word/word.body#word-word-body-onparagraphadded-member)|Occurs when new paragraphs are added.|
||[onParagraphChanged](/javascript/api/word/word.body#word-word-body-onparagraphchanged-member)|Occurs when paragraphs are changed.|
||[onParagraphDeleted](/javascript/api/word/word.body#word-word-body-onparagraphdeleted-member)|Occurs when paragraphs are deleted.|
|[Border](/javascript/api/word/word.border)|[lineColor](/javascript/api/word/word.border#word-word-border-linecolor-member)|Specifies the line color for the border.|
||[lineStyle](/javascript/api/word/word.border#word-word-border-linestyle-member)|Specifies the line style for the border.|
||[lineWidth](/javascript/api/word/word.border#word-word-border-linewidth-member)|Specifies the line width for the border.|
||[positionType](/javascript/api/word/word.border#word-word-border-positiontype-member)|Gets the position type of the border.|
||[visible](/javascript/api/word/word.border#word-word-border-visible-member)|Specifies whether the border is visible.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|The event type.|
||[ids](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-ids-member)|Gets the content control IDs.|
||[source](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-source-member)|The source of the event.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[Field](/javascript/api/word/word.field)|[showCodes](/javascript/api/word/word.field#word-word-field-showcodes-member)|Specifies whether the field codes are displayed for the specified field.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importCustomProperties](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importcustomproperties-member)|Represents whether the custom properties from the source document should be imported.|
||[importCustomXmlParts](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importcustomxmlparts-member)|Represents whether the custom XML parts from the source document should be imported.|
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
|[Paragraph](/javascript/api/word/word.paragraph)|[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
||[uniqueLocalId](/javascript/api/word/word.paragraph#word-word-paragraph-uniquelocalid-member)|Gets a string that represents the paragraph identifier in the current session.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[ids](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-ids-member)|Gets the IDs of the involved paragraphs.|
||[source](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[ids](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-ids-member)|Gets the IDs of the involved paragraphs.|
||[source](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[ids](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-ids-member)|Gets the IDs of the involved paragraphs.|
||[source](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
|[Shading](/javascript/api/word/word.shading)|[backgroundPatternColor](/javascript/api/word/word.shading#word-word-shading-backgroundpatterncolor-member)|Specifies the color for the background of the object.|
||[foregroundPatternColor](/javascript/api/word/word.shading#word-word-shading-foregroundpatterncolor-member)|Specifies the color for the foreground of the object.|
||[texture](/javascript/api/word/word.shading#word-word-shading-texture-member)|Specifies the shading texture of the object.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
||[listTemplate](/javascript/api/word/word.style#word-word-style-listtemplate-member)|Gets a ListTemplate object that represents the list formatting for the specified Style object.|
||[shading](/javascript/api/word/word.style#word-word-style-shading-member)|Gets a Shading object that represents the shading for the specified style.|
||[tableStyle](/javascript/api/word/word.style#word-word-style-tablestyle-member)|Gets a TableStyle object representing Style properties that can be applied to a table.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[alignment](/javascript/api/word/word.tablestyle#word-word-tablestyle-alignment-member)|Specifies the table's alignment against the page margin.|
||[allowBreakAcrossPage](/javascript/api/word/word.tablestyle#word-word-tablestyle-allowbreakacrosspage-member)|Specifies whether lines in tables formatted with a specified style break across pages.|
||[bottomCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-bottomcellmargin-member)|Specifies the amount of space to add between the contents and the bottom borders of the cells.|
||[cellSpacing](/javascript/api/word/word.tablestyle#word-word-tablestyle-cellspacing-member)|Specifies the spacing (in points) between the cells in a table style.|
||[leftCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-leftcellmargin-member)|Specifies the amount of space to add between the contents and the left borders of the cells.|
||[rightCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-rightcellmargin-member)|Specifies the amount of space to add between the contents and the right borders of the cells.|
||[topCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-topcellmargin-member)|Specifies the amount of space to add between the contents and the top borders of the cells.|
