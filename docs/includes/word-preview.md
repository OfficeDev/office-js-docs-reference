| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
||[shapes](/javascript/api/word/word.body#word-word-body-shapes-member)|Gets the collection of shape objects in the body, including both inline and floating shapes.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
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
||[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[Document](/javascript/api/word/word.document)|[compareFromBase64(base64File: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-comparefrombase64-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[InsertShapeOptions](/javascript/api/word/word.insertshapeoptions)|[height](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-height-member)|Represents the height of the shape being inserted.|
||[left](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-left-member)|Represents the left position of the shape being inserted.|
||[top](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-top-member)|Represents the top position of the shape being inserted.|
||[width](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-width-member)|Represents the width of the shape being inserted.|
|[Paragraph](/javascript/api/word/word.paragraph)|[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertgeometricshape-member(1))|Inserts a geometric shape with its anchor at the beginning of the paragraph.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertpicturefrombase64-member(1))|Inserts a floating picture with its anchor at the beginning of the paragraph.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttextbox-member(1))|Inserts a floating text box with its anchor at the beginning of the paragraph.|
||[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
||[shapes](/javascript/api/word/word.paragraph#word-word-paragraph-shapes-member)|Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertgeometricshape-member(1))|Inserts a geometric shape with its anchor at the beginning of the range.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertpicturefrombase64-member(1))|Inserts a floating picture with its anchor at the beginning of the range.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-inserttextbox-member(1))|Inserts a floating text box with its anchor at the beginning of the range.|
||[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
||[shapes](/javascript/api/word/word.range#word-word-range-shapes-member)|Gets the collection of shape objects anchored in the range, including both inline and floating shapes.|
|[Shape](/javascript/api/word/word.shape)|[body](/javascript/api/word/word.shape#word-word-shape-body-member)|Represents the body object of the shape.|
||[delete()](/javascript/api/word/word.shape#word-word-shape-delete-member(1))|Deletes the shape and its content.|
||[geometricShapeType](/javascript/api/word/word.shape#word-word-shape-geometricshapetype-member)|The geometric shape type of the shape.|
||[height](/javascript/api/word/word.shape#word-word-shape-height-member)|The height, in points, of the shape.|
||[id](/javascript/api/word/word.shape#word-word-shape-id-member)|Gets an integer that represents the shape identifier.|
||[isChild](/javascript/api/word/word.shape#word-word-shape-ischild-member)|Check whether this shape is a child of a group shape or a canvas shape.|
||[left](/javascript/api/word/word.shape#word-word-shape-left-member)|The distance, in points, from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[moveHorizontally(distance: number)](/javascript/api/word/word.shape#word-word-shape-movehorizontally-member(1))|Moves the shape horizontally by the number of points.|
||[moveVertically(distance: number)](/javascript/api/word/word.shape#word-word-shape-movevertically-member(1))|Moves the shape vertically by the number of points.|
||[name](/javascript/api/word/word.shape#word-word-shape-name-member)|The name of the shape.|
||[parentGroup](/javascript/api/word/word.shape#word-word-shape-parentgroup-member)|Gets the top-level parent group shape of this child shape.|
||[relativeHorizontalPosition](/javascript/api/word/word.shape#word-word-shape-relativehorizontalposition-member)|The relative horizontal position of the shape.|
||[relativeVerticalPosition](/javascript/api/word/word.shape#word-word-shape-relativeverticalposition-member)|The relative vertical position of the shape.|
||[select(selectMultipleShapes?: boolean)](/javascript/api/word/word.shape#word-word-shape-select-member(1))|Selects the shape.|
||[shapeGroup](/javascript/api/word/word.shape#word-word-shape-shapegroup-member)|Gets the shape group associated with the shape.|
||[top](/javascript/api/word/word.shape#word-word-shape-top-member)|The distance, in points, from the top edge of the shape to the vertical relative position, see Word.RelativeVerticalPosition.|
||[type](/javascript/api/word/word.shape#word-word-shape-type-member)|Gets the shape type.|
||[width](/javascript/api/word/word.shape#word-word-shape-width-member)|The width, in points, of the shape.|
|[ShapeCollection](/javascript/api/word/word.shapecollection)|[getByGeometricTypes(types: Word.GeometricShapeType[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbygeometrictypes-member(1))|Gets the shapes that have the specified geometric types.|
||[getById(id: number)](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyid-member(1))|Gets a shape by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyidornullobject-member(1))|Gets a shape by its identifier.|
||[getByIds(ids: number[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbyids-member(1))|Gets the shapes by the identifiers.|
||[getByNames(names: string[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbynames-member(1))|Gets the shapes that have the specified names.|
||[getByTypes(types: Word.ShapeType[])](/javascript/api/word/word.shapecollection#word-word-shapecollection-getbytypes-member(1))|Gets the shapes that have the specified types.|
||[getFirst()](/javascript/api/word/word.shapecollection#word-word-shapecollection-getfirst-member(1))|Gets the first shape in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.shapecollection#word-word-shapecollection-getfirstornullobject-member(1))|Gets the first shape in this collection.|
||[group()](/javascript/api/word/word.shapecollection#word-word-shapecollection-group-member(1))|Groups floating shapes in this collection, inline shapes will be skipped.|
||[items](/javascript/api/word/word.shapecollection#word-word-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[ShapeGroup](/javascript/api/word/word.shapegroup)|[id](/javascript/api/word/word.shapegroup#word-word-shapegroup-id-member)|Gets an integer that represents the shape group identifier.|
||[shape](/javascript/api/word/word.shapegroup#word-word-shapegroup-shape-member)|Gets the Shape object associated with the group.|
||[shapes](/javascript/api/word/word.shapegroup#word-word-shapegroup-shapes-member)|Gets the collection of Shape objects.|
||[ungroup()](/javascript/api/word/word.shapegroup#word-word-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
