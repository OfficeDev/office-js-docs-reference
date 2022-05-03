| Class | Fields | Description |
|:---|:---|:---|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#powerpoint-powerpoint-bulletformat-visible-member)|Specifies if the bullets in the paragraph are visible.|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-bulletformat-member)|Represents the bullet format of the paragraph.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#powerpoint-powerpoint-paragraphformat-horizontalalignment-member)|Represents the horizontal alignment of the paragraph.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[getSelectedShapes()](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedshapes-member(1))|Returns the selected shapes in the current slide of the presentation.|
||[getSelectedSlides()](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedslides-member(1))|Returns the selected slides in the current view of the presentation.|
||[getSelectedTextRange()](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedtextrange-member(1))|Returns the selected {@link PowerPoint.TextRange} in the current view of the presentation.|
||[getSelectedTextRangeOrNullObject()](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedtextrangeornullobject-member(1))|Returns the selected {@link PowerPoint.TextRange} in the current view of the presentation.|
||[setSelectedSlides(slideIds: string[])](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-setselectedslides-member(1))|Selects the slides in the current view of the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[fill](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-fill-member)|Returns the fill formatting of this shape.|
||[getParentSlide()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslide-member(1))|Returns the parent {@link PowerPoint.Slide} object that holds this `Shape`.|
||[getParentSlideLayout()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslidelayout-member(1))|Returns the parent {@link PowerPoint.SlideLayout} object that holds this `Shape`.|
||[getParentSlideLayoutOrNullObject()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslidelayoutornullobject-member(1))|Returns the parent {@link PowerPoint.SlideLayout} object that holds this `Shape`.|
||[getParentSlideMaster()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslidemaster-member(1))|Returns the parent {@link PowerPoint.SlideMaster} object that holds this `Shape`.|
||[getParentSlideMasterOrNullObject()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslidemasterornullobject-member(1))|Returns the parent {@link PowerPoint.SlideMaster} object that holds this `Shape`.|
||[getParentSlideOrNullObject()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-getparentslideornullobject-member(1))|Returns the parent {@link PowerPoint.Slide} object that holds this `Shape`.|
||[height](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-height-member)|Specifies the height, in points, of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the slide.|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-lineformat-member)|Returns the line formatting of this shape.|
||[name](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-name-member)|Specifies the name of this shape.|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-textframe-member)|Returns the text frame object of this shape.|
||[top](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the slide.|
||[type](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-type-member)|Returns the type of this shape.|
||[width](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-width-member)|Specifies the width, in points, of the shape.|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-height-member)|Specifies the height, in points, of the shape.|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-left-member)|Specifies the distance, in points, from the left side of the shape to the left side of the slide.|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-top-member)|Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#powerpoint-powerpoint-shapeaddoptions-width-member)|Specifies the width, in points, of the shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgeometricshape-member(1))|Adds a geometric shape to the slide.|
||[addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addline-member(1))|Adds a line to the slide.|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtextbox-member(1))|Adds a text box to the slide with the provided text as the content.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-clear-member(1))|Clears the fill formatting of this shape.|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-foregroundcolor-member)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-color-member)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-italic-member)|Represents the italic status of font.|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-name-member)|Represents font name (e.g., "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-size-member)|Represents font size in points (e.g., 11).|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-underline-member)|Type of underline applied to the font.|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-color-member)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-dashstyle-member)|Represents the dash style of the line.|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-style-member)|Represents the line style of the shape.|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-visible-member)|Specifies if the line formatting of a shape element is visible.|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#powerpoint-powerpoint-shapelineformat-weight-member)|Represents the weight of the line, in points.|
|[ShapeScopedCollection](/javascript/api/powerpoint/powerpoint.shapescopedcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-getcount-member(1))|Gets the number of shapes in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-getitem-member(1))|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-getitemat-member(1))|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-getitemornullobject-member(1))|Gets a shape using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[setSelectedShapes(shapeIds: string[])](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-setselectedshapes-member(1))|Selects the specified shapes.|
|[SlideScopedCollection](/javascript/api/powerpoint/powerpoint.slidescopedcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidescopedcollection#powerpoint-powerpoint-slidescopedcollection-getcount-member(1))|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidescopedcollection#powerpoint-powerpoint-slidescopedcollection-getitem-member(1))|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidescopedcollection#powerpoint-powerpoint-slidescopedcollection-getitemat-member(1))|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidescopedcollection#powerpoint-powerpoint-slidescopedcollection-getitemornullobject-member(1))|Gets a slide using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidescopedcollection#powerpoint-powerpoint-slidescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-deletetext-member(1))|Deletes all the text in the text frame.|
||[getParentShape()](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-getparentshape-member(1))|Returns the parent {@link PowerPoint.Shape} object that holds this `TextFrame`.|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-hastext-member)|Specifies if the text frame contains text.|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-textrange-member)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#powerpoint-powerpoint-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-font-member)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getParentTextFrame()](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-getparenttextframe-member(1))|Returns the parent {@link PowerPoint.TextFrame} object that holds this `TextRange`.|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-getsubstring-member(1))|Returns a `TextRange` object for the substring in the given range.|
||[length](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-length-member)|Gets or sets the length of the range that this `TextRange` represents.|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-paragraphformat-member)|Represents the paragraph format of the text range.|
||[setSelected()](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-setselected-member(1))|Selects this `TextRange` in the current view.|
||[start](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-start-member)|Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.|
||[text](/javascript/api/powerpoint/powerpoint.textrange#powerpoint-powerpoint-textrange-text-member)|Represents the plain text content of the text range.|
