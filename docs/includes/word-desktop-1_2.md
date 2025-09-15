| Class | Fields | Description |
|:---|:---|:---|
|[Body](/.body)|[shapes](/.body#word-javascript/api/word/-body-shapes-member)|Gets the collection of shape objects in the body, including both inline and floating shapes.|
|[Canvas](/.canvas)|[id](/.canvas#word-javascript/api/word/-canvas-id-member)|Gets an integer that represents the canvas identifier.|
||[shape](/.canvas#word-javascript/api/word/-canvas-shape-member)|Gets the Shape object associated with the canvas.|
||[shapes](/.canvas#word-javascript/api/word/-canvas-shapes-member)|Gets the collection of Shape objects.|
|[Document](/.document)|[activeWindow](/.document#word-javascript/api/word/-document-activewindow-member)|Gets the active window for the document.|
||[compareFromBase64(base64File: string, documentCompareOptions?: Word.DocumentCompareOptions)](/.document#word-javascript/api/word/-document-comparefrombase64-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
||[windows](/.document#word-javascript/api/word/-document-windows-member)|Gets the collection of `Word.Window` objects for the document.|
|[Font](/.font)|[hidden](/.font#word-javascript/api/word/-font-hidden-member)|Specifies a value that indicates whether the font is tagged as hidden.|
|[InsertShapeOptions](/.insertshapeoptions)|[height](/.insertshapeoptions#word-javascript/api/word/-insertshapeoptions-height-member)|Represents the height of the shape being inserted.|
||[left](/.insertshapeoptions#word-javascript/api/word/-insertshapeoptions-left-member)|Represents the left position of the shape being inserted.|
||[top](/.insertshapeoptions#word-javascript/api/word/-insertshapeoptions-top-member)|Represents the top position of the shape being inserted.|
||[width](/.insertshapeoptions#word-javascript/api/word/-insertshapeoptions-width-member)|Represents the width of the shape being inserted.|
|[Page](/.page)|[getNext()](/.page#word-javascript/api/word/-page-getnext-member(1))|Gets the next page in the pane.|
||[getNextOrNullObject()](/.page#word-javascript/api/word/-page-getnextornullobject-member(1))|Gets the next page.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/.page#word-javascript/api/word/-page-getrange-member(1))|Gets the whole page, or the starting or ending point of the page, as a range.|
||[height](/.page#word-javascript/api/word/-page-height-member)|Gets the height, in points, of the paper defined in the Page Setup dialog box.|
||[index](/.page#word-javascript/api/word/-page-index-member)|Gets the index of the page.|
||[width](/.page#word-javascript/api/word/-page-width-member)|Gets the width, in points, of the paper defined in the Page Setup dialog box.|
|[PageCollection](/.pagecollection)|[getFirst()](/.pagecollection#word-javascript/api/word/-pagecollection-getfirst-member(1))|Gets the first page in this collection.|
||[getFirstOrNullObject()](/.pagecollection#word-javascript/api/word/-pagecollection-getfirstornullobject-member(1))|Gets the first page in this collection.|
||[items](/.pagecollection#word-javascript/api/word/-pagecollection-items-member)|Gets the loaded child items in this collection.|
|[Pane](/.pane)|[getNext()](/.pane#word-javascript/api/word/-pane-getnext-member(1))|Gets the next pane in the window.|
||[getNextOrNullObject()](/.pane#word-javascript/api/word/-pane-getnextornullobject-member(1))|Gets the next pane.|
||[pages](/.pane#word-javascript/api/word/-pane-pages-member)|Gets the collection of pages in the pane.|
||[pagesEnclosingViewport](/.pane#word-javascript/api/word/-pane-pagesenclosingviewport-member)|Gets the `PageCollection` shown in the viewport of the pane.|
|[PaneCollection](/.panecollection)|[getFirst()](/.panecollection#word-javascript/api/word/-panecollection-getfirst-member(1))|Gets the first pane in this collection.|
||[getFirstOrNullObject()](/.panecollection#word-javascript/api/word/-panecollection-getfirstornullobject-member(1))|Gets the first pane in this collection.|
||[items](/.panecollection#word-javascript/api/word/-panecollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/.paragraph)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/.paragraph#word-javascript/api/word/-paragraph-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/.paragraph#word-javascript/api/word/-paragraph-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/.paragraph#word-javascript/api/word/-paragraph-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/.paragraph#word-javascript/api/word/-paragraph-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.|
||[shapes](/.paragraph#word-javascript/api/word/-paragraph-shapes-member)|Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes.|
|[Range](/.range)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/.range#word-javascript/api/word/-range-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the range.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/.range#word-javascript/api/word/-range-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the range.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/.range#word-javascript/api/word/-range-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the range.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/.range#word-javascript/api/word/-range-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the range.|
||[pages](/.range#word-javascript/api/word/-range-pages-member)|Gets the collection of pages in the range.|
||[shapes](/.range#word-javascript/api/word/-range-shapes-member)|Gets the collection of shape objects anchored in the range, including both inline and floating shapes.|
|[Shape](/.shape)|[allowOverlap](/.shape#word-javascript/api/word/-shape-allowoverlap-member)|Specifies whether a given shape can overlap other shapes.|
||[altTextDescription](/.shape#word-javascript/api/word/-shape-alttextdescription-member)|Specifies a string that represents the alternative text associated with the shape.|
||[body](/.shape#word-javascript/api/word/-shape-body-member)|Represents the body object of the shape.|
||[canvas](/.shape#word-javascript/api/word/-shape-canvas-member)|Gets the canvas associated with the shape.|
||[delete()](/.shape#word-javascript/api/word/-shape-delete-member(1))|Deletes the shape and its content.|
||[fill](/.shape#word-javascript/api/word/-shape-fill-member)|Returns the fill formatting of the shape.|
||[geometricShapeType](/.shape#word-javascript/api/word/-shape-geometricshapetype-member)|The geometric shape type of the shape.|
||[height](/.shape#word-javascript/api/word/-shape-height-member)|The height, in points, of the shape.|
||[heightRelative](/.shape#word-javascript/api/word/-shape-heightrelative-member)|The percentage of shape height to vertical relative size, see Word.RelativeSize.|
||[id](/.shape#word-javascript/api/word/-shape-id-member)|Gets an integer that represents the shape identifier.|
||[isChild](/.shape#word-javascript/api/word/-shape-ischild-member)|Check whether this shape is a child of a group shape or a canvas shape.|
||[left](/.shape#word-javascript/api/word/-shape-left-member)|The distance, in points, from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[leftRelative](/.shape#word-javascript/api/word/-shape-leftrelative-member)|The relative left position as a percentage from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[lockAspectRatio](/.shape#word-javascript/api/word/-shape-lockaspectratio-member)|Specifies if the aspect ratio of this shape is locked.|
||[moveHorizontally(distance: number)](/.shape#word-javascript/api/word/-shape-movehorizontally-member(1))|Moves the shape horizontally by the number of points.|
||[moveVertically(distance: number)](/.shape#word-javascript/api/word/-shape-movevertically-member(1))|Moves the shape vertically by the number of points.|
||[name](/.shape#word-javascript/api/word/-shape-name-member)|The name of the shape.|
||[parentCanvas](/.shape#word-javascript/api/word/-shape-parentcanvas-member)|Gets the top-level parent canvas shape of this child shape.|
||[parentGroup](/.shape#word-javascript/api/word/-shape-parentgroup-member)|Gets the top-level parent group shape of this child shape.|
||[relativeHorizontalPosition](/.shape#word-javascript/api/word/-shape-relativehorizontalposition-member)|The relative horizontal position of the shape.|
||[relativeHorizontalSize](/.shape#word-javascript/api/word/-shape-relativehorizontalsize-member)|The relative horizontal size of the shape.|
||[relativeVerticalPosition](/.shape#word-javascript/api/word/-shape-relativeverticalposition-member)|The relative vertical position of the shape.|
||[relativeVerticalSize](/.shape#word-javascript/api/word/-shape-relativeverticalsize-member)|The relative vertical size of the shape.|
||[rotation](/.shape#word-javascript/api/word/-shape-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Word.ShapeScaleType, scaleFrom?: Word.ShapeScaleFrom)](/.shape#word-javascript/api/word/-shape-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Word.ShapeScaleType, scaleFrom?: Word.ShapeScaleFrom)](/.shape#word-javascript/api/word/-shape-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[select(selectMultipleShapes?: boolean)](/.shape#word-javascript/api/word/-shape-select-member(1))|Selects the shape.|
||[shapeGroup](/.shape#word-javascript/api/word/-shape-shapegroup-member)|Gets the shape group associated with the shape.|
||[textFrame](/.shape#word-javascript/api/word/-shape-textframe-member)|Gets the text frame object of the shape.|
||[textWrap](/.shape#word-javascript/api/word/-shape-textwrap-member)|Returns the text wrap formatting of the shape.|
||[top](/.shape#word-javascript/api/word/-shape-top-member)|The distance, in points, from the top edge of the shape to the vertical relative position (see Word.RelativeVerticalPosition).|
||[topRelative](/.shape#word-javascript/api/word/-shape-toprelative-member)|The relative top position as a percentage from the top edge of the shape to the vertical relative position, see Word.RelativeVerticalPosition.|
||[type](/.shape#word-javascript/api/word/-shape-type-member)|Gets the shape type.|
||[visible](/.shape#word-javascript/api/word/-shape-visible-member)|Specifies if the shape is visible.|
||[width](/.shape#word-javascript/api/word/-shape-width-member)|The width, in points, of the shape.|
||[widthRelative](/.shape#word-javascript/api/word/-shape-widthrelative-member)|The percentage of shape width to horizontal relative size, see Word.RelativeSize.|
|[ShapeCollection](/.shapecollection)|[getByGeometricTypes(types: Word.GeometricShapeType[])](/.shapecollection#word-javascript/api/word/-shapecollection-getbygeometrictypes-member(1))|Gets the shapes that have the specified geometric types.|
||[getById(id: number)](/.shapecollection#word-javascript/api/word/-shapecollection-getbyid-member(1))|Gets a shape by its identifier.|
||[getByIdOrNullObject(id: number)](/.shapecollection#word-javascript/api/word/-shapecollection-getbyidornullobject-member(1))|Gets a shape by its identifier.|
||[getByIds(ids: number[])](/.shapecollection#word-javascript/api/word/-shapecollection-getbyids-member(1))|Gets the shapes by the identifiers.|
||[getByNames(names: string[])](/.shapecollection#word-javascript/api/word/-shapecollection-getbynames-member(1))|Gets the shapes that have the specified names.|
||[getByTypes(types: Word.ShapeType[])](/.shapecollection#word-javascript/api/word/-shapecollection-getbytypes-member(1))|Gets the shapes that have the specified types.|
||[getFirst()](/.shapecollection#word-javascript/api/word/-shapecollection-getfirst-member(1))|Gets the first shape in this collection.|
||[getFirstOrNullObject()](/.shapecollection#word-javascript/api/word/-shapecollection-getfirstornullobject-member(1))|Gets the first shape in this collection.|
||[group()](/.shapecollection#word-javascript/api/word/-shapecollection-group-member(1))|Groups floating shapes in this collection, inline shapes will be skipped.|
||[items](/.shapecollection#word-javascript/api/word/-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[ShapeFill](/.shapefill)|[backgroundColor](/.shapefill#word-javascript/api/word/-shapefill-backgroundcolor-member)|Specifies the shape fill background color.|
||[clear()](/.shapefill#word-javascript/api/word/-shapefill-clear-member(1))|Clears the fill formatting of this shape and set it to `Word.ShapeFillType.NoFill`;|
||[foregroundColor](/.shapefill#word-javascript/api/word/-shapefill-foregroundcolor-member)|Specifies the shape fill foreground color.|
||[setSolidColor(color: string)](/.shapefill#word-javascript/api/word/-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/.shapefill#word-javascript/api/word/-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/.shapefill#word-javascript/api/word/-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeGroup](/.shapegroup)|[id](/.shapegroup#word-javascript/api/word/-shapegroup-id-member)|Gets an integer that represents the shape group identifier.|
||[shape](/.shapegroup#word-javascript/api/word/-shapegroup-shape-member)|Gets the Shape object associated with the group.|
||[shapes](/.shapegroup#word-javascript/api/word/-shapegroup-shapes-member)|Gets the collection of Shape objects.|
||[ungroup()](/.shapegroup#word-javascript/api/word/-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeTextWrap](/.shapetextwrap)|[bottomDistance](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-bottomdistance-member)|Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.|
||[leftDistance](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-leftdistance-member)|Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.|
||[rightDistance](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-rightdistance-member)|Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.|
||[side](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-side-member)|Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.|
||[topDistance](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-topdistance-member)|Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.|
||[type](/.shapetextwrap#word-javascript/api/word/-shapetextwrap-type-member)|Specifies the text wrap type around the shape.|
|[TextFrame](/.textframe)|[autoSizeSetting](/.textframe#word-javascript/api/word/-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/.textframe#word-javascript/api/word/-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[hasText](/.textframe#word-javascript/api/word/-textframe-hastext-member)|Specifies if the text frame contains text.|
||[leftMargin](/.textframe#word-javascript/api/word/-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[noTextRotation](/.textframe#word-javascript/api/word/-textframe-notextrotation-member)|Returns True if text in the text frame shouldn't rotate when the shape is rotated.|
||[orientation](/.textframe#word-javascript/api/word/-textframe-orientation-member)|Represents the angle to which the text is oriented for the text frame.|
||[rightMargin](/.textframe#word-javascript/api/word/-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[topMargin](/.textframe#word-javascript/api/word/-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/.textframe#word-javascript/api/word/-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[wordWrap](/.textframe#word-javascript/api/word/-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[Window](/.window)|[activePane](/.window#word-javascript/api/word/-window-activepane-member)|Gets the active pane in the window.|
||[panes](/.window#word-javascript/api/word/-window-panes-member)|Gets the collection of panes in the window.|
|[WindowCollection](/.windowcollection)|[getFirst()](/.windowcollection#word-javascript/api/word/-windowcollection-getfirst-member(1))|Gets the first window in this collection.|
||[getFirstOrNullObject()](/.windowcollection#word-javascript/api/word/-windowcollection-getfirstornullobject-member(1))|Gets the first window in this collection.|
||[items](/.windowcollection#word-javascript/api/word/-windowcollection-items-member)|Gets the loaded child items in this collection.|
