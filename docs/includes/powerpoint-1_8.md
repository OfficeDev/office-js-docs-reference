| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/javascript/api/powerpoint/powerpoint.binding)|[delete()](/javascript/api/powerpoint/powerpoint.binding#powerpoint-powerpoint-binding-delete-member(1))|Deletes the binding.|
||[getShape()](/javascript/api/powerpoint/powerpoint.binding#powerpoint-powerpoint-binding-getshape-member(1))|Returns the shape represented by the binding.|
||[id](/javascript/api/powerpoint/powerpoint.binding#powerpoint-powerpoint-binding-id-member)|Represents the binding identifier.|
||[type](/javascript/api/powerpoint/powerpoint.binding#powerpoint-powerpoint-binding-type-member)|Returns the type of the binding.|
|[BindingCollection](/javascript/api/powerpoint/powerpoint.bindingcollection)|[add(shape: PowerPoint.Shape, bindingType: PowerPoint.BindingType, id: string)](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-add-member(1))|Adds a new binding to a particular Shape.|
||[addFromSelection(bindingType: PowerPoint.BindingType, id: string)](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-addfromselection-member(1))|Adds a new binding based on the current selection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-getcount-member(1))|Gets the number of bindings in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-getitem-member(1))|Gets a binding object by ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-getitemat-member(1))|Gets a binding object based on its position in the items array.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-getitemornullobject-member(1))|Gets a binding object by ID.|
||[items](/javascript/api/powerpoint/powerpoint.bindingcollection#powerpoint-powerpoint-bindingcollection-items-member)|Gets the loaded child items in this collection.|
|[BorderProperties](/javascript/api/powerpoint/powerpoint.borderproperties)|[color](/javascript/api/powerpoint/powerpoint.borderproperties#powerpoint-powerpoint-borderproperties-color-member)|Represents the line color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.borderproperties#powerpoint-powerpoint-borderproperties-dashstyle-member)|Represents the dash style of the line.|
||[transparency](/javascript/api/powerpoint/powerpoint.borderproperties#powerpoint-powerpoint-borderproperties-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[weight](/javascript/api/powerpoint/powerpoint.borderproperties#powerpoint-powerpoint-borderproperties-weight-member)|Represents the weight of the line, in points.|
|[FillProperties](/javascript/api/powerpoint/powerpoint.fillproperties)|[color](/javascript/api/powerpoint/powerpoint.fillproperties#powerpoint-powerpoint-fillproperties-color-member)|Represents the shape fill color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[transparency](/javascript/api/powerpoint/powerpoint.fillproperties#powerpoint-powerpoint-fillproperties-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
|[FontProperties](/javascript/api/powerpoint/powerpoint.fontproperties)|[allCaps](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-allcaps-member)|Represents whether the font uses all caps, where lowercase letters are shown as capital letters.|
||[bold](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-bold-member)|Represents the bold status of font.|
||[color](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-color-member)|Represents the HTML color in the hexadecimal format (e.g., "#FF0000" represents red) or as a named HTML color value (e.g., "red").|
||[doubleStrikethrough](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-doublestrikethrough-member)|Represents the double-strikethrough status of the font.|
||[italic](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-italic-member)|Represents the italic status of font.|
||[name](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-name-member)|Represents the font name (e.g., "Calibri").|
||[size](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-size-member)|Represents the font size in points (e.g., 11).|
||[smallCaps](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-smallcaps-member)|Represents whether the text uses small caps, where lowercase letters are shown as small capital letters.|
||[strikethrough](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-strikethrough-member)|Represents the strikethrough status of the font.|
||[subscript](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-subscript-member)|Represents the subscript status of the font.|
||[superscript](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-superscript-member)|Represents the superscript status of the font.|
||[underline](/javascript/api/powerpoint/powerpoint.fontproperties#powerpoint-powerpoint-fontproperties-underline-member)|Type of underline applied to the font.|
|[PlaceholderFormat](/javascript/api/powerpoint/powerpoint.placeholderformat)|[containedType](/javascript/api/powerpoint/powerpoint.placeholderformat#powerpoint-powerpoint-placeholderformat-containedtype-member)|Gets the type of the shape contained within the placeholder.|
||[type](/javascript/api/powerpoint/powerpoint.placeholderformat#powerpoint-powerpoint-placeholderformat-type-member)|Returns the type of this placeholder.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[bindings](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-bindings-member)|Returns a collection of bindings that are associated with the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[getTable()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-gettable-member(1))|Returns the `Table` object if this shape is a table.|
||[group](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-group-member)|Returns the `ShapeGroup` associated with the shape.|
||[level](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-level-member)|Returns the level of the specified shape.|
||[parentGroup](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-parentgroup-member)|Returns the parent group of this shape.|
||[placeholderFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-placeholderformat-member)|Returns the properties that apply specifically to this placeholder.|
||[setZOrder(position: PowerPoint.ShapeZOrder)](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[zOrderPosition](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-zorderposition-member)|Returns the z-order position of the shape, with 0 representing the bottom of the order stack.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGroup(values: Array<string \| Shape>)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgroup-member(1))|Create a shape group for several shapes.|
||[addTable(rowCount: number, columnCount: number, options?: PowerPoint.TableAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addtable-member(1))|Adds a table to the slide.|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[setImage(base64EncodedImage: string)](/javascript/api/powerpoint/powerpoint.shapefill#powerpoint-powerpoint-shapefill-setimage-member(1))|Sets the fill formatting of the shape to an image.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[allCaps](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-allcaps-member)|Specifies whether the text in the `TextRange` is set to use the **All Caps** attribute which makes lowercase letters appear as uppercase letters.|
||[doubleStrikethrough](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-doublestrikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Double strikethrough** attribute.|
||[smallCaps](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-smallcaps-member)|Specifies whether the text in the `TextRange` is set to use the **Small Caps** attribute which makes lowercase letters appear as small uppercase letters.|
||[strikethrough](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-strikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Strikethrough** attribute.|
||[subscript](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-subscript-member)|Specifies whether the text in the `TextRange` is set to use the **Subscript** attribute.|
||[superscript](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-superscript-member)|Specifies whether the text in the `TextRange` is set to use the **Superscript** attribute.|
|[ShapeGroup](/javascript/api/powerpoint/powerpoint.shapegroup)|[id](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-id-member)|Gets the unique ID of the shape group.|
||[shape](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-shape-member)|Returns the `Shape` object associated with the group.|
||[shapes](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-shapes-member)|Returns the collection of `Shape` objects in the group.|
||[ungroup()](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeScopedCollection](/javascript/api/powerpoint/powerpoint.shapescopedcollection)|[group()](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-group-member(1))|Groups all shapes in this collection into a single shape.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[applyLayout(slideLayout: PowerPoint.SlideLayout)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-applylayout-member(1))|Applies the specified layout to the slide, changing its design and structure according to the chosen layout.|
||[exportAsBase64()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-exportasbase64-member(1))|Exports the slide to its own presentation file, returned as Base64-encoded data.|
||[getImageAsBase64(options?: PowerPoint.SlideGetImageOptions)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-getimageasbase64-member(1))|Renders an image of the slide.|
||[index](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-index-member)|Returns the zero-based index of the slide representing its position in the presentation.|
||[moveTo(slideIndex: number)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-moveto-member(1))|Moves the slide to a new position within the presentation.|
|[SlideGetImageOptions](/javascript/api/powerpoint/powerpoint.slidegetimageoptions)|[height](/javascript/api/powerpoint/powerpoint.slidegetimageoptions#powerpoint-powerpoint-slidegetimageoptions-height-member)|The desired height of the resulting image in pixels.|
||[width](/javascript/api/powerpoint/powerpoint.slidegetimageoptions#powerpoint-powerpoint-slidegetimageoptions-width-member)|The desired width of the resulting image in pixels.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[type](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-type-member)|Returns the type of the slide layout.|
|[Table](/javascript/api/powerpoint/powerpoint.table)|[columnCount](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-columncount-member)|Gets the number of columns in the table.|
||[getCellOrNullObject(rowIndex: number, columnIndex: number)](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-getcellornullobject-member(1))|Gets the cell at the specified `rowIndex` and `columnIndex`.|
||[getMergedAreas()](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-getmergedareas-member(1))|Gets a collection of cells that represent the merged areas of the table.|
||[getShape()](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-getshape-member(1))|Gets the shape object for the table.|
||[rowCount](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-rowcount-member)|Gets the number of rows in the table.|
||[values](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-values-member)|Gets all of the values in the table.|
|[TableAddOptions](/javascript/api/powerpoint/powerpoint.tableaddoptions)|[columns](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-columns-member)|If provided, specifies properties for each column in the table.|
||[height](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-height-member)|Specifies the height, in points, of the table.|
||[left](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-left-member)|Specifies the distance, in points, from the left side of the table to the left side of the slide.|
||[mergedAreas](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-mergedareas-member)|If specified, represents an rectangular area where multiple cells appear as a single cell.|
||[rows](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-rows-member)|If provided, specifies properties for each row in the table.|
||[specificCellProperties](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-specificcellproperties-member)|If provided, specifies properties for each cell in the table.|
||[top](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-top-member)|Specifies the distance, in points, from the top edge of the table to the top edge of the slide.|
||[uniformCellProperties](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-uniformcellproperties-member)|Specifies the formatting which applies uniformly to all of the table cells.|
||[values](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-values-member)|If provided, specifies the values for the table.|
||[width](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-width-member)|Specifies the width, in points, of the table.|
|[TableCell](/javascript/api/powerpoint/powerpoint.tablecell)|[columnCount](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-columncount-member)|Gets the number of table columns this cell spans across.|
||[columnIndex](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-columnindex-member)|Gets the zero-based column index of the cell within the table.|
||[rowCount](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-rowcount-member)|Gets the number of table rows this cell spans across.|
||[rowIndex](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-rowindex-member)|Gets the zero-based row index of the cell within the table.|
||[text](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-text-member)|Specifies the text content of the table cell.|
|[TableCellBorders](/javascript/api/powerpoint/powerpoint.tablecellborders)|[bottom](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-bottom-member)|Represents the bottom border.|
||[diagonalDown](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-diagonaldown-member)|Represents the diagonal border (top-left to bottom-right).|
||[diagonalUp](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-diagonalup-member)|Represents the diagonal border (bottom-left to top-right).|
||[left](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-left-member)|Represents the left border.|
||[right](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-right-member)|Represents the right border.|
||[top](/javascript/api/powerpoint/powerpoint.tablecellborders#powerpoint-powerpoint-tablecellborders-top-member)|Represents the top border.|
|[TableCellCollection](/javascript/api/powerpoint/powerpoint.tablecellcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.tablecellcollection#powerpoint-powerpoint-tablecellcollection-getcount-member(1))|Gets the number of table cells in the collection.|
||[getItemAtOrNullObject(row: number, column: number)](/javascript/api/powerpoint/powerpoint.tablecellcollection#powerpoint-powerpoint-tablecellcollection-getitematornullobject-member(1))|Gets the table cell using its zero-based index in the collection.|
||[items](/javascript/api/powerpoint/powerpoint.tablecellcollection#powerpoint-powerpoint-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCellMargins](/javascript/api/powerpoint/powerpoint.tablecellmargins)|[bottom](/javascript/api/powerpoint/powerpoint.tablecellmargins#powerpoint-powerpoint-tablecellmargins-bottom-member)|Specifies the bottom margin in points.|
||[left](/javascript/api/powerpoint/powerpoint.tablecellmargins#powerpoint-powerpoint-tablecellmargins-left-member)|Specifies the left margin in points.|
||[right](/javascript/api/powerpoint/powerpoint.tablecellmargins#powerpoint-powerpoint-tablecellmargins-right-member)|Specifies the right margin in points.|
||[top](/javascript/api/powerpoint/powerpoint.tablecellmargins#powerpoint-powerpoint-tablecellmargins-top-member)|Specifies the top margin in points.|
|[TableCellProperties](/javascript/api/powerpoint/powerpoint.tablecellproperties)|[borders](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-borders-member)|Specifies the border formatting of the table cell.|
||[fill](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-fill-member)|Specifies the fill formatting of the table cell.|
||[font](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-font-member)|Specifies the font formatting of the table cell.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-horizontalalignment-member)|Represents the horizontal alignment of the table cell.|
||[indentLevel](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-indentlevel-member)|Represents the indent level of the text in the table cell.|
||[margins](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-margins-member)|Specifies the margin settings in the table cell.|
||[text](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-text-member)|Specifies the text content of the table cell.|
||[textRuns](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-textruns-member)|Specifies the contents of the table cell as an array of TextRun objects.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.tablecellproperties#powerpoint-powerpoint-tablecellproperties-verticalalignment-member)|Represents the vertical alignment of the table cell.|
|[TableColumnProperties](/javascript/api/powerpoint/powerpoint.tablecolumnproperties)|[columnWidth](/javascript/api/powerpoint/powerpoint.tablecolumnproperties#powerpoint-powerpoint-tablecolumnproperties-columnwidth-member)|Represents the desired width of each column in points, or is undefined.|
|[TableMergedAreaProperties](/javascript/api/powerpoint/powerpoint.tablemergedareaproperties)|[columnCount](/javascript/api/powerpoint/powerpoint.tablemergedareaproperties#powerpoint-powerpoint-tablemergedareaproperties-columncount-member)|Specifies the number of columns for the merged cells area.|
||[columnIndex](/javascript/api/powerpoint/powerpoint.tablemergedareaproperties#powerpoint-powerpoint-tablemergedareaproperties-columnindex-member)|Specifies the zero-based index of the column of the top left cell of the merged area.|
||[rowCount](/javascript/api/powerpoint/powerpoint.tablemergedareaproperties#powerpoint-powerpoint-tablemergedareaproperties-rowcount-member)|Specifies the number of rows for the merged cells area.|
||[rowIndex](/javascript/api/powerpoint/powerpoint.tablemergedareaproperties#powerpoint-powerpoint-tablemergedareaproperties-rowindex-member)|Specifies the zero-based index of the row of the top left cell of the merged area.|
|[TableRowProperties](/javascript/api/powerpoint/powerpoint.tablerowproperties)|[rowHeight](/javascript/api/powerpoint/powerpoint.tablerowproperties#powerpoint-powerpoint-tablerowproperties-rowheight-member)|Represents the desired height of each row in points, or is undefined.|
|[TextRun](/javascript/api/powerpoint/powerpoint.textrun)|[font](/javascript/api/powerpoint/powerpoint.textrun#powerpoint-powerpoint-textrun-font-member)|The font attributes (such as font name, font size, and color) applied to this text run.|
||[text](/javascript/api/powerpoint/powerpoint.textrun#powerpoint-powerpoint-textrun-text-member)|The text of this text run.|
