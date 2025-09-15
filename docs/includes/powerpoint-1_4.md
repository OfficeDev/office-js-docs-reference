| Class | Fields | Description |
|:---|:---|:---|
|[BulletFormat](/.bulletformat)|[visible](/.bulletformat#powerpoint-javascript/api/powerpoint/-bulletformat-visible-member)|Specifies if the bullets in the paragraph are visible.|
|[ParagraphFormat](/.paragraphformat)|[bulletFormat](/.paragraphformat#powerpoint-javascript/api/powerpoint/-paragraphformat-bulletformat-member)|Represents the bullet format of the paragraph.|
||[horizontalAlignment](/.paragraphformat#powerpoint-javascript/api/powerpoint/-paragraphformat-horizontalalignment-member)|Represents the horizontal alignment of the paragraph.|
|[Shape](/.shape)|[fill](/.shape#powerpoint-javascript/api/powerpoint/-shape-fill-member)|Returns the fill formatting of this shape.|
||[height](/.shape#powerpoint-javascript/api/powerpoint/-shape-height-member)|Specifies the height, in points, of the shape.|
||[left](/.shape#powerpoint-javascript/api/powerpoint/-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the slide.|
||[lineFormat](/.shape#powerpoint-javascript/api/powerpoint/-shape-lineformat-member)|Returns the line formatting of this shape.|
||[name](/.shape#powerpoint-javascript/api/powerpoint/-shape-name-member)|Specifies the name of this shape.|
||[textFrame](/.shape#powerpoint-javascript/api/powerpoint/-shape-textframe-member)|Returns the PowerPoint.TextFrame object of this `Shape`.|
||[top](/.shape#powerpoint-javascript/api/powerpoint/-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the slide.|
||[type](/.shape#powerpoint-javascript/api/powerpoint/-shape-type-member)|Returns the type of this shape.|
||[width](/.shape#powerpoint-javascript/api/powerpoint/-shape-width-member)|Specifies the width, in points, of the shape.|
|[ShapeAddOptions](/.shapeaddoptions)|[height](/.shapeaddoptions#powerpoint-javascript/api/powerpoint/-shapeaddoptions-height-member)|Specifies the height, in points, of the shape.|
||[left](/.shapeaddoptions#powerpoint-javascript/api/powerpoint/-shapeaddoptions-left-member)|Specifies the distance, in points, from the left side of the shape to the left side of the slide.|
||[top](/.shapeaddoptions#powerpoint-javascript/api/powerpoint/-shapeaddoptions-top-member)|Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.|
||[width](/.shapeaddoptions#powerpoint-javascript/api/powerpoint/-shapeaddoptions-width-member)|Specifies the width, in points, of the shape.|
|[ShapeCollection](/.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-addgeometricshape-member(1))|Adds a geometric shape to the slide.|
||[addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-addline-member(1))|Adds a line to the slide.|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-addtextbox-member(1))|Adds a text box to the slide with the provided text as the content.|
|[ShapeFill](/.shapefill)|[clear()](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-clear-member(1))|Clears the fill formatting of this shape.|
||[foregroundColor](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-foregroundcolor-member)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[setSolidColor(color: string)](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeFont](/.shapefont)|[bold](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-bold-member)|Specifies whether the text in the `TextRange` is set to bold.|
||[color](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-color-member)|Specifies the HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-italic-member)|Specifies whether the text in the `TextRange` is set to italic.|
||[name](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-name-member)|Specifies the font name (e.g., "Calibri").|
||[size](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-size-member)|Specifies the font size in points (e.g., 11).|
||[underline](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-underline-member)|Specifies the type of underline applied to the font.|
|[ShapeLineFormat](/.shapelineformat)|[color](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-color-member)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-dashstyle-member)|Represents the dash style of the line.|
||[style](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-style-member)|Represents the line style of the shape.|
||[transparency](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-visible-member)|Specifies if the line formatting of a shape element is visible.|
||[weight](/.shapelineformat#powerpoint-javascript/api/powerpoint/-shapelineformat-weight-member)|Represents the weight of the line, in points.|
|[TextFrame](/.textframe)|[autoSizeSetting](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-deletetext-member(1))|Deletes all the text in the text frame.|
||[hasText](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-hastext-member)|Specifies if the text frame contains text.|
||[leftMargin](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[rightMargin](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[textRange](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-textrange-member)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[wordWrap](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-wordwrap-member)|Determines whether lines break automatically to fit text inside the shape.|
|[TextRange](/.textrange)|[font](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-font-member)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getSubstring(start: number, length?: number)](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-getsubstring-member(1))|Returns a `TextRange` object for the substring in the given range.|
||[paragraphFormat](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-paragraphformat-member)|Represents the paragraph format of the text range.|
||[text](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-text-member)|Represents the plain text content of the text range.|
