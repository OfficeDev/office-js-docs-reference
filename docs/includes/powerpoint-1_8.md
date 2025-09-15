| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/.binding)|[delete()](/.binding#powerpoint-javascript/api/powerpoint/-binding-delete-member(1))|Deletes the binding.|
||[getShape()](/.binding#powerpoint-javascript/api/powerpoint/-binding-getshape-member(1))|Returns the shape represented by the binding.|
||[id](/.binding#powerpoint-javascript/api/powerpoint/-binding-id-member)|Represents the binding identifier.|
||[type](/.binding#powerpoint-javascript/api/powerpoint/-binding-type-member)|Returns the type of the binding.|
|[BindingCollection](/.bindingcollection)|[add(shape: PowerPoint.Shape, bindingType: PowerPoint.BindingType, id: string)](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-add-member(1))|Adds a new binding to a particular Shape.|
||[addFromSelection(bindingType: PowerPoint.BindingType, id: string)](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-addfromselection-member(1))|Adds a new binding based on the current selection.|
||[getCount()](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-getcount-member(1))|Gets the number of bindings in the collection.|
||[getItem(key: string)](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-getitem-member(1))|Gets a binding object by ID.|
||[getItemAt(index: number)](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-getitemat-member(1))|Gets a binding object based on its position in the items array.|
||[getItemOrNullObject(id: string)](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-getitemornullobject-member(1))|Gets a binding object by ID.|
||[items](/.bindingcollection#powerpoint-javascript/api/powerpoint/-bindingcollection-items-member)|Gets the loaded child items in this collection.|
|[BorderProperties](/.borderproperties)|[color](/.borderproperties#powerpoint-javascript/api/powerpoint/-borderproperties-color-member)|Represents the line color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[dashStyle](/.borderproperties#powerpoint-javascript/api/powerpoint/-borderproperties-dashstyle-member)|Represents the dash style of the line.|
||[transparency](/.borderproperties#powerpoint-javascript/api/powerpoint/-borderproperties-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[weight](/.borderproperties#powerpoint-javascript/api/powerpoint/-borderproperties-weight-member)|Represents the weight of the line, in points.|
|[FillProperties](/.fillproperties)|[color](/.fillproperties#powerpoint-javascript/api/powerpoint/-fillproperties-color-member)|Represents the shape fill color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[transparency](/.fillproperties#powerpoint-javascript/api/powerpoint/-fillproperties-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
|[FontProperties](/.fontproperties)|[allCaps](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-allcaps-member)|Represents whether the font uses all caps, where lowercase letters are shown as capital letters.|
||[bold](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-bold-member)|Represents the bold status of font.|
||[color](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-color-member)|Represents the HTML color in the hexadecimal format (e.g., "#FF0000" represents red) or as a named HTML color value (e.g., "red").|
||[doubleStrikethrough](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-doublestrikethrough-member)|Represents the double-strikethrough status of the font.|
||[italic](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-italic-member)|Represents the italic status of font.|
||[name](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-name-member)|Represents the font name (e.g., "Calibri").|
||[size](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-size-member)|Represents the font size in points (e.g., 11).|
||[smallCaps](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-smallcaps-member)|Represents whether the text uses small caps, where lowercase letters are shown as small capital letters.|
||[strikethrough](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-strikethrough-member)|Represents the strikethrough status of the font.|
||[subscript](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-subscript-member)|Represents the subscript status of the font.|
||[superscript](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-superscript-member)|Represents the superscript status of the font.|
||[underline](/.fontproperties#powerpoint-javascript/api/powerpoint/-fontproperties-underline-member)|Type of underline applied to the font.|
|[PlaceholderFormat](/.placeholderformat)|[containedType](/.placeholderformat#powerpoint-javascript/api/powerpoint/-placeholderformat-containedtype-member)|Gets the type of the shape contained within the placeholder.|
||[type](/.placeholderformat#powerpoint-javascript/api/powerpoint/-placeholderformat-type-member)|Returns the type of this placeholder.|
|[Presentation](/.presentation)|[bindings](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-bindings-member)|Returns a collection of bindings that are associated with the presentation.|
|[Shape](/.shape)|[getTable()](/.shape#powerpoint-javascript/api/powerpoint/-shape-gettable-member(1))|Returns the `Table` object if this shape is a table.|
||[group](/.shape#powerpoint-javascript/api/powerpoint/-shape-group-member)|Returns the `ShapeGroup` associated with the shape.|
||[level](/.shape#powerpoint-javascript/api/powerpoint/-shape-level-member)|Returns the level of the specified shape.|
||[parentGroup](/.shape#powerpoint-javascript/api/powerpoint/-shape-parentgroup-member)|Returns the parent group of this shape.|
||[placeholderFormat](/.shape#powerpoint-javascript/api/powerpoint/-shape-placeholderformat-member)|Returns the properties that apply specifically to this placeholder.|
||[setZOrder(position: PowerPoint.ShapeZOrder)](/.shape#powerpoint-javascript/api/powerpoint/-shape-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[zOrderPosition](/.shape#powerpoint-javascript/api/powerpoint/-shape-zorderposition-member)|Returns the z-order position of the shape, with 0 representing the bottom of the order stack.|
|[ShapeCollection](/.shapecollection)|[addGroup(values: Array<string \| Shape>)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-addgroup-member(1))|Create a shape group for several shapes.|
||[addTable(rowCount: number, columnCount: number, options?: PowerPoint.TableAddOptions)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-addtable-member(1))|Adds a table to the slide.|
|[ShapeFill](/.shapefill)|[setImage(base64EncodedImage: string)](/.shapefill#powerpoint-javascript/api/powerpoint/-shapefill-setimage-member(1))|Sets the fill formatting of the shape to an image.|
|[ShapeFont](/.shapefont)|[allCaps](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-allcaps-member)|Specifies whether the text in the `TextRange` is set to use the **All Caps** attribute which makes lowercase letters appear as uppercase letters.|
||[doubleStrikethrough](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-doublestrikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Double strikethrough** attribute.|
||[smallCaps](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-smallcaps-member)|Specifies whether the text in the `TextRange` is set to use the **Small Caps** attribute which makes lowercase letters appear as small uppercase letters.|
||[strikethrough](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-strikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Strikethrough** attribute.|
||[subscript](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-subscript-member)|Specifies whether the text in the `TextRange` is set to use the **Subscript** attribute.|
||[superscript](/.shapefont#powerpoint-javascript/api/powerpoint/-shapefont-superscript-member)|Specifies whether the text in the `TextRange` is set to use the **Superscript** attribute.|
|[ShapeGroup](/.shapegroup)|[id](/.shapegroup#powerpoint-javascript/api/powerpoint/-shapegroup-id-member)|Gets the unique ID of the shape group.|
||[shape](/.shapegroup#powerpoint-javascript/api/powerpoint/-shapegroup-shape-member)|Returns the `Shape` object associated with the group.|
||[shapes](/.shapegroup#powerpoint-javascript/api/powerpoint/-shapegroup-shapes-member)|Returns the collection of `Shape` objects in the group.|
||[ungroup()](/.shapegroup#powerpoint-javascript/api/powerpoint/-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeScopedCollection](/.shapescopedcollection)|[group()](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-group-member(1))|Groups all shapes in this collection into a single shape.|
|[Slide](/.slide)|[applyLayout(slideLayout: PowerPoint.SlideLayout)](/.slide#powerpoint-javascript/api/powerpoint/-slide-applylayout-member(1))|Applies the specified layout to the slide, changing its design and structure according to the chosen layout.|
||[exportAsBase64()](/.slide#powerpoint-javascript/api/powerpoint/-slide-exportasbase64-member(1))|Exports the slide to its own presentation file, returned as Base64-encoded data.|
||[getImageAsBase64(options?: PowerPoint.SlideGetImageOptions)](/.slide#powerpoint-javascript/api/powerpoint/-slide-getimageasbase64-member(1))|Renders an image of the slide.|
||[index](/.slide#powerpoint-javascript/api/powerpoint/-slide-index-member)|Returns the zero-based index of the slide representing its position in the presentation.|
||[moveTo(slideIndex: number)](/.slide#powerpoint-javascript/api/powerpoint/-slide-moveto-member(1))|Moves the slide to a new position within the presentation.|
|[SlideGetImageOptions](/.slidegetimageoptions)|[height](/.slidegetimageoptions#powerpoint-javascript/api/powerpoint/-slidegetimageoptions-height-member)|The desired height of the resulting image in pixels.|
||[width](/.slidegetimageoptions#powerpoint-javascript/api/powerpoint/-slidegetimageoptions-width-member)|The desired width of the resulting image in pixels.|
|[SlideLayout](/.slidelayout)|[type](/.slidelayout#powerpoint-javascript/api/powerpoint/-slidelayout-type-member)|Returns the type of the slide layout.|
|[Table](/.table)|[columnCount](/.table#powerpoint-javascript/api/powerpoint/-table-columncount-member)|Gets the number of columns in the table.|
||[getCellOrNullObject(rowIndex: number, columnIndex: number)](/.table#powerpoint-javascript/api/powerpoint/-table-getcellornullobject-member(1))|Gets the cell at the specified `rowIndex` and `columnIndex`.|
||[getMergedAreas()](/.table#powerpoint-javascript/api/powerpoint/-table-getmergedareas-member(1))|Gets a collection of cells that represent the merged areas of the table.|
||[getShape()](/.table#powerpoint-javascript/api/powerpoint/-table-getshape-member(1))|Gets the shape object for the table.|
||[rowCount](/.table#powerpoint-javascript/api/powerpoint/-table-rowcount-member)|Gets the number of rows in the table.|
||[values](/.table#powerpoint-javascript/api/powerpoint/-table-values-member)|Gets all of the values in the table.|
|[TableAddOptions](/.tableaddoptions)|[columns](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-columns-member)|If provided, specifies properties for each column in the table.|
||[height](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-height-member)|Specifies the height, in points, of the table.|
||[left](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-left-member)|Specifies the distance, in points, from the left side of the table to the left side of the slide.|
||[mergedAreas](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-mergedareas-member)|If specified, represents an rectangular area where multiple cells appear as a single cell.|
||[rows](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-rows-member)|If provided, specifies properties for each row in the table.|
||[specificCellProperties](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-specificcellproperties-member)|If provided, specifies properties for each cell in the table.|
||[top](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-top-member)|Specifies the distance, in points, from the top edge of the table to the top edge of the slide.|
||[uniformCellProperties](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-uniformcellproperties-member)|Specifies the formatting which applies uniformly to all of the table cells.|
||[values](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-values-member)|If provided, specifies the values for the table.|
||[width](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-width-member)|Specifies the width, in points, of the table.|
|[TableCell](/.tablecell)|[columnCount](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-columncount-member)|Gets the number of table columns this cell spans across.|
||[columnIndex](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-columnindex-member)|Gets the zero-based column index of the cell within the table.|
||[rowCount](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-rowcount-member)|Gets the number of table rows this cell spans across.|
||[rowIndex](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-rowindex-member)|Gets the zero-based row index of the cell within the table.|
||[text](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-text-member)|Specifies the text content of the table cell.|
|[TableCellBorders](/.tablecellborders)|[bottom](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-bottom-member)|Represents the bottom border.|
||[diagonalDown](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-diagonaldown-member)|Represents the diagonal border (top-left to bottom-right).|
||[diagonalUp](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-diagonalup-member)|Represents the diagonal border (bottom-left to top-right).|
||[left](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-left-member)|Represents the left border.|
||[right](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-right-member)|Represents the right border.|
||[top](/.tablecellborders#powerpoint-javascript/api/powerpoint/-tablecellborders-top-member)|Represents the top border.|
|[TableCellCollection](/.tablecellcollection)|[getCount()](/.tablecellcollection#powerpoint-javascript/api/powerpoint/-tablecellcollection-getcount-member(1))|Gets the number of table cells in the collection.|
||[getItemAtOrNullObject(row: number, column: number)](/.tablecellcollection#powerpoint-javascript/api/powerpoint/-tablecellcollection-getitematornullobject-member(1))|Gets the table cell using its zero-based index in the collection.|
||[items](/.tablecellcollection#powerpoint-javascript/api/powerpoint/-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCellMargins](/.tablecellmargins)|[bottom](/.tablecellmargins#powerpoint-javascript/api/powerpoint/-tablecellmargins-bottom-member)|Specifies the bottom margin in points.|
||[left](/.tablecellmargins#powerpoint-javascript/api/powerpoint/-tablecellmargins-left-member)|Specifies the left margin in points.|
||[right](/.tablecellmargins#powerpoint-javascript/api/powerpoint/-tablecellmargins-right-member)|Specifies the right margin in points.|
||[top](/.tablecellmargins#powerpoint-javascript/api/powerpoint/-tablecellmargins-top-member)|Specifies the top margin in points.|
|[TableCellProperties](/.tablecellproperties)|[borders](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-borders-member)|Specifies the border formatting of the table cell.|
||[fill](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-fill-member)|Specifies the fill formatting of the table cell.|
||[font](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-font-member)|Specifies the font formatting of the table cell.|
||[horizontalAlignment](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-horizontalalignment-member)|Represents the horizontal alignment of the table cell.|
||[indentLevel](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-indentlevel-member)|Represents the indent level of the text in the table cell.|
||[margins](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-margins-member)|Specifies the margin settings in the table cell.|
||[text](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-text-member)|Specifies the text content of the table cell.|
||[textRuns](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-textruns-member)|Specifies the contents of the table cell as an array of PowerPoint.TextRun objects.|
||[verticalAlignment](/.tablecellproperties#powerpoint-javascript/api/powerpoint/-tablecellproperties-verticalalignment-member)|Represents the vertical alignment of the table cell.|
|[TableColumnProperties](/.tablecolumnproperties)|[columnWidth](/.tablecolumnproperties#powerpoint-javascript/api/powerpoint/-tablecolumnproperties-columnwidth-member)|Represents the desired width of each column in points, or is undefined.|
|[TableMergedAreaProperties](/.tablemergedareaproperties)|[columnCount](/.tablemergedareaproperties#powerpoint-javascript/api/powerpoint/-tablemergedareaproperties-columncount-member)|Specifies the number of columns for the merged cells area.|
||[columnIndex](/.tablemergedareaproperties#powerpoint-javascript/api/powerpoint/-tablemergedareaproperties-columnindex-member)|Specifies the zero-based index of the column of the top left cell of the merged area.|
||[rowCount](/.tablemergedareaproperties#powerpoint-javascript/api/powerpoint/-tablemergedareaproperties-rowcount-member)|Specifies the number of rows for the merged cells area.|
||[rowIndex](/.tablemergedareaproperties#powerpoint-javascript/api/powerpoint/-tablemergedareaproperties-rowindex-member)|Specifies the zero-based index of the row of the top left cell of the merged area.|
|[TableRowProperties](/.tablerowproperties)|[rowHeight](/.tablerowproperties#powerpoint-javascript/api/powerpoint/-tablerowproperties-rowheight-member)|Represents the desired height of each row in points, or is undefined.|
|[TextRun](/.textrun)|[font](/.textrun#powerpoint-javascript/api/powerpoint/-textrun-font-member)|The font attributes (such as font name, font size, and color) applied to this text run.|
||[text](/.textrun#powerpoint-javascript/api/powerpoint/-textrun-text-member)|The text of this text run.|
