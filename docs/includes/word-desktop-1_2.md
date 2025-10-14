| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[shapes](/javascript/api/word/word.body#word-word-body-shapes-member)|Gets the collection of `Shape` objects in the body, including both inline and floating shapes.|
|[Canvas](/javascript/api/word/word.canvas)|[id](/javascript/api/word/word.canvas#word-word-canvas-id-member)|Gets an integer that represents the canvas identifier.|
||[shape](/javascript/api/word/word.canvas#word-word-canvas-shape-member)|Gets the `Shape` object associated with the canvas.|
||[shapes](/javascript/api/word/word.canvas#word-word-canvas-shapes-member)|Gets the collection of Word.Shape objects.|
|[Document](/javascript/api/word/word.document)|[activeWindow](/javascript/api/word/word.document#word-word-document-activewindow-member)|Gets the active window for the document.|
||[compareFromBase64(base64File: string, documentCompareOptions?: Word.DocumentCompareOptions)](/javascript/api/word/word.document#word-word-document-comparefrombase64-member(1))|Displays revision marks that indicate where the specified document differs from another document.|
||[windows](/javascript/api/word/word.document#word-word-document-windows-member)|Gets the collection of `Word.Window` objects for the document.|
|[Font](/javascript/api/word/word.font)|[hidden](/javascript/api/word/word.font#word-word-font-hidden-member)|Specifies whether the font is tagged as hidden.|
|[InsertShapeOptions](/javascript/api/word/word.insertshapeoptions)|[height](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-height-member)|If provided, specifies the height of the shape being inserted.|
||[left](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-left-member)|If provided, specifies the left position of the shape being inserted.|
||[top](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-top-member)|If provided, specifies the top position of the shape being inserted.|
||[width](/javascript/api/word/word.insertshapeoptions#word-word-insertshapeoptions-width-member)|If provided, specifies the width of the shape being inserted.|
|[Page](/javascript/api/word/word.page)|[getNext()](/javascript/api/word/word.page#word-word-page-getnext-member(1))|Gets the next page in the pane.|
||[getNextOrNullObject()](/javascript/api/word/word.page#word-word-page-getnextornullobject-member(1))|Gets the next page.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/javascript/api/word/word.page#word-word-page-getrange-member(1))|Gets the whole page, or the starting or ending point of the page, as a range.|
||[height](/javascript/api/word/word.page#word-word-page-height-member)|Gets the height, in points, of the paper defined in the Page Setup dialog box.|
||[index](/javascript/api/word/word.page#word-word-page-index-member)|Gets the index of the page.|
||[width](/javascript/api/word/word.page#word-word-page-width-member)|Gets the width, in points, of the paper defined in the Page Setup dialog box.|
|[PageCollection](/javascript/api/word/word.pagecollection)|[getFirst()](/javascript/api/word/word.pagecollection#word-word-pagecollection-getfirst-member(1))|Gets the first page in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.pagecollection#word-word-pagecollection-getfirstornullobject-member(1))|Gets the first page in this collection.|
||[items](/javascript/api/word/word.pagecollection#word-word-pagecollection-items-member)|Gets the loaded child items in this collection.|
|[Pane](/javascript/api/word/word.pane)|[getNext()](/javascript/api/word/word.pane#word-word-pane-getnext-member(1))|Gets the next pane in the window.|
||[getNextOrNullObject()](/javascript/api/word/word.pane#word-word-pane-getnextornullobject-member(1))|Gets the next pane.|
||[pages](/javascript/api/word/word.pane#word-word-pane-pages-member)|Gets the collection of pages in the pane.|
||[pagesEnclosingViewport](/javascript/api/word/word.pane#word-word-pane-pagesenclosingviewport-member)|Gets the `PageCollection` shown in the viewport of the pane.|
|[PaneCollection](/javascript/api/word/word.panecollection)|[getFirst()](/javascript/api/word/word.panecollection#word-word-panecollection-getfirst-member(1))|Gets the first pane in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.panecollection#word-word-panecollection-getfirstornullobject-member(1))|Gets the first pane in this collection.|
||[items](/javascript/api/word/word.panecollection#word-word-panecollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.|
||[shapes](/javascript/api/word/word.paragraph#word-word-paragraph-shapes-member)|Gets the collection of `Shape` objects anchored in the paragraph, including both inline and floating shapes.|
|[Range](/javascript/api/word/word.range)|[insertCanvas(insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertcanvas-member(1))|Inserts a floating canvas in front of text with its anchor at the beginning of the range.|
||[insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertgeometricshape-member(1))|Inserts a geometric shape in front of text with its anchor at the beginning of the range.|
||[insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-insertpicturefrombase64-member(1))|Inserts a floating picture in front of text with its anchor at the beginning of the range.|
||[insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions)](/javascript/api/word/word.range#word-word-range-inserttextbox-member(1))|Inserts a floating text box in front of text with its anchor at the beginning of the range.|
||[pages](/javascript/api/word/word.range#word-word-range-pages-member)|Gets the collection of pages in the range.|
||[shapes](/javascript/api/word/word.range#word-word-range-shapes-member)|Gets the collection of `Shape` objects anchored in the range, including both inline and floating shapes.|
|[Shape](/javascript/api/word/word.shape)|[allowOverlap](/javascript/api/word/word.shape#word-word-shape-allowoverlap-member)|Specifies whether a given shape can overlap other shapes.|
||[altTextDescription](/javascript/api/word/word.shape#word-word-shape-alttextdescription-member)|Specifies a string that represents the alternative text associated with the shape.|
||[body](/javascript/api/word/word.shape#word-word-shape-body-member)|Represents the `Body` object of the shape.|
||[canvas](/javascript/api/word/word.shape#word-word-shape-canvas-member)|Gets the canvas associated with the shape.|
||[delete()](/javascript/api/word/word.shape#word-word-shape-delete-member(1))|Deletes the shape and its content.|
||[fill](/javascript/api/word/word.shape#word-word-shape-fill-member)|Returns the fill formatting of the shape.|
||[geometricShapeType](/javascript/api/word/word.shape#word-word-shape-geometricshapetype-member)|The geometric shape type of the shape.|
||[height](/javascript/api/word/word.shape#word-word-shape-height-member)|The height, in points, of the shape.|
||[heightRelative](/javascript/api/word/word.shape#word-word-shape-heightrelative-member)|The percentage of shape height to vertical relative size, see Word.RelativeSize.|
||[id](/javascript/api/word/word.shape#word-word-shape-id-member)|Gets an integer that represents the shape identifier.|
||[isChild](/javascript/api/word/word.shape#word-word-shape-ischild-member)|Check whether this shape is a child of a group shape or a canvas shape.|
||[left](/javascript/api/word/word.shape#word-word-shape-left-member)|The distance, in points, from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[leftRelative](/javascript/api/word/word.shape#word-word-shape-leftrelative-member)|The relative left position as a percentage from the left side of the shape to the horizontal relative position, see Word.RelativeHorizontalPosition.|
||[lockAspectRatio](/javascript/api/word/word.shape#word-word-shape-lockaspectratio-member)|Specifies if the aspect ratio of this shape is locked.|
||[moveHorizontally(distance: number)](/javascript/api/word/word.shape#word-word-shape-movehorizontally-member(1))|Moves the shape horizontally by the number of points.|
||[moveVertically(distance: number)](/javascript/api/word/word.shape#word-word-shape-movevertically-member(1))|Moves the shape vertically by the number of points.|
||[name](/javascript/api/word/word.shape#word-word-shape-name-member)|The name of the shape.|
||[parentCanvas](/javascript/api/word/word.shape#word-word-shape-parentcanvas-member)|Gets the top-level parent canvas shape of this child shape.|
||[parentGroup](/javascript/api/word/word.shape#word-word-shape-parentgroup-member)|Gets the top-level parent group shape of this child shape.|
||[relativeHorizontalPosition](/javascript/api/word/word.shape#word-word-shape-relativehorizontalposition-member)|The relative horizontal position of the shape.|
||[relativeHorizontalSize](/javascript/api/word/word.shape#word-word-shape-relativehorizontalsize-member)|The relative horizontal size of the shape.|
||[relativeVerticalPosition](/javascript/api/word/word.shape#word-word-shape-relativeverticalposition-member)|The relative vertical position of the shape.|
||[relativeVerticalSize](/javascript/api/word/word.shape#word-word-shape-relativeverticalsize-member)|The relative vertical size of the shape.|
||[rotation](/javascript/api/word/word.shape#word-word-shape-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Word.ShapeScaleType, scaleFrom?: Word.ShapeScaleFrom)](/javascript/api/word/word.shape#word-word-shape-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Word.ShapeScaleType, scaleFrom?: Word.ShapeScaleFrom)](/javascript/api/word/word.shape#word-word-shape-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[select(selectMultipleShapes?: boolean)](/javascript/api/word/word.shape#word-word-shape-select-member(1))|Selects the shape.|
||[shapeGroup](/javascript/api/word/word.shape#word-word-shape-shapegroup-member)|Gets the shape group associated with the shape.|
||[textFrame](/javascript/api/word/word.shape#word-word-shape-textframe-member)|Gets the `TextFrame` object of the shape.|
||[textWrap](/javascript/api/word/word.shape#word-word-shape-textwrap-member)|Returns the text wrap formatting of the shape.|
||[top](/javascript/api/word/word.shape#word-word-shape-top-member)|The distance, in points, from the top edge of the shape to the vertical relative position (see Word.RelativeVerticalPosition).|
||[topRelative](/javascript/api/word/word.shape#word-word-shape-toprelative-member)|The relative top position as a percentage from the top edge of the shape to the vertical relative position, see Word.RelativeVerticalPosition.|
||[type](/javascript/api/word/word.shape#word-word-shape-type-member)|Gets the shape type.|
||[visible](/javascript/api/word/word.shape#word-word-shape-visible-member)|Specifies if the shape is visible.|
||[width](/javascript/api/word/word.shape#word-word-shape-width-member)|The width, in points, of the shape.|
||[widthRelative](/javascript/api/word/word.shape#word-word-shape-widthrelative-member)|The percentage of shape width to horizontal relative size, see Word.RelativeSize.|
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
|[ShapeFill](/javascript/api/word/word.shapefill)|[backgroundColor](/javascript/api/word/word.shapefill#word-word-shapefill-backgroundcolor-member)|Specifies the shape fill background color.|
||[clear()](/javascript/api/word/word.shapefill#word-word-shapefill-clear-member(1))|Clears the fill formatting of this shape and sets it to `Word.ShapeFillType.noFill`.|
||[foregroundColor](/javascript/api/word/word.shapefill#word-word-shapefill-foregroundcolor-member)|Specifies the shape fill foreground color.|
||[setSolidColor(color: string)](/javascript/api/word/word.shapefill#word-word-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/word/word.shapefill#word-word-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/word/word.shapefill#word-word-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeGroup](/javascript/api/word/word.shapegroup)|[id](/javascript/api/word/word.shapegroup#word-word-shapegroup-id-member)|Gets an integer that represents the shape group identifier.|
||[shape](/javascript/api/word/word.shapegroup#word-word-shapegroup-shape-member)|Gets the `Shape` object associated with the group.|
||[shapes](/javascript/api/word/word.shapegroup#word-word-shapegroup-shapes-member)|Gets the collection of `Shape` objects.|
||[ungroup()](/javascript/api/word/word.shapegroup#word-word-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeTextWrap](/javascript/api/word/word.shapetextwrap)|[bottomDistance](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-bottomdistance-member)|Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.|
||[leftDistance](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-leftdistance-member)|Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.|
||[rightDistance](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-rightdistance-member)|Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.|
||[side](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-side-member)|Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.|
||[topDistance](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-topdistance-member)|Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.|
||[type](/javascript/api/word/word.shapetextwrap#word-word-shapetextwrap-type-member)|Specifies the text wrap type around the shape.|
|[TextFrame](/javascript/api/word/word.textframe)|[autoSizeSetting](/javascript/api/word/word.textframe#word-word-textframe-autosizesetting-member)|Specifies the automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/word/word.textframe#word-word-textframe-bottommargin-member)|Specifies the bottom margin, in points, of the text frame.|
||[hasText](/javascript/api/word/word.textframe#word-word-textframe-hastext-member)|Returns `true` if the text frame contains text, otherwise, `false`.|
||[leftMargin](/javascript/api/word/word.textframe#word-word-textframe-leftmargin-member)|Specifies the left margin, in points, of the text frame.|
||[noTextRotation](/javascript/api/word/word.textframe#word-word-textframe-notextrotation-member)|Specifies whether the text in the text frame shouldn't rotate when the shape is rotated.|
||[orientation](/javascript/api/word/word.textframe#word-word-textframe-orientation-member)|Specifies the angle to which the text is oriented for the text frame.|
||[rightMargin](/javascript/api/word/word.textframe#word-word-textframe-rightmargin-member)|Specifies the right margin, in points, of the text frame.|
||[topMargin](/javascript/api/word/word.textframe#word-word-textframe-topmargin-member)|Specifies the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/word/word.textframe#word-word-textframe-verticalalignment-member)|Specifies the vertical alignment of the text frame.|
||[wordWrap](/javascript/api/word/word.textframe#word-word-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[Window](/javascript/api/word/word.window)|[activePane](/javascript/api/word/word.window#word-word-window-activepane-member)|Gets the active pane in the window.|
||[panes](/javascript/api/word/word.window#word-word-window-panes-member)|Gets the collection of panes in the window.|
|[WindowCollection](/javascript/api/word/word.windowcollection)|[getFirst()](/javascript/api/word/word.windowcollection#word-word-windowcollection-getfirst-member(1))|Gets the first window in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.windowcollection#word-word-windowcollection-getfirstornullobject-member(1))|Gets the first window in this collection.|
||[items](/javascript/api/word/word.windowcollection#word-word-windowcollection-items-member)|Gets the loaded child items in this collection.|
