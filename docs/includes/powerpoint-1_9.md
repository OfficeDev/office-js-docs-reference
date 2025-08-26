| Class | Fields | Description |
|:---|:---|:---|
|[Border](/javascript/api/powerpoint/powerpoint.border)|[color](/javascript/api/powerpoint/powerpoint.border#powerpoint-powerpoint-border-color-member)|Represents the line color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[dashStyle](/javascript/api/powerpoint/powerpoint.border#powerpoint-powerpoint-border-dashstyle-member)|Represents the dash style of the line.|
||[transparency](/javascript/api/powerpoint/powerpoint.border#powerpoint-powerpoint-border-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[weight](/javascript/api/powerpoint/powerpoint.border#powerpoint-powerpoint-border-weight-member)|Represents the weight of the line, in points.|
|[Borders](/javascript/api/powerpoint/powerpoint.borders)|[bottom](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-bottom-member)|Gets the bottom border.|
||[diagonalDown](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-diagonaldown-member)|Gets the diagonal border (top-left to bottom-right).|
||[diagonalUp](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-diagonalup-member)|Gets the diagonal border (bottom-left to top-right).|
||[left](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-left-member)|Gets the left border.|
||[right](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-right-member)|Gets the right border.|
||[top](/javascript/api/powerpoint/powerpoint.borders#powerpoint-powerpoint-borders-top-member)|Gets the top border.|
|[Margins](/javascript/api/powerpoint/powerpoint.margins)|[bottom](/javascript/api/powerpoint/powerpoint.margins#powerpoint-powerpoint-margins-bottom-member)|Specifies the bottom margin in points.|
||[left](/javascript/api/powerpoint/powerpoint.margins#powerpoint-powerpoint-margins-left-member)|Specifies the left margin in points.|
||[right](/javascript/api/powerpoint/powerpoint.margins#powerpoint-powerpoint-margins-right-member)|Specifies the right margin in points.|
||[top](/javascript/api/powerpoint/powerpoint.margins#powerpoint-powerpoint-margins-top-member)|Specifies the top margin in points.|
|[Table](/javascript/api/powerpoint/powerpoint.table)|[clear(options?: PowerPoint.TableClearOptions)](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-clear-member(1))|Clears table values and formatting.|
||[columns](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-columns-member)|Gets the collection of columns in the table.|
||[mergeCells(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number)](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-mergecells-member(1))|Creates a merged area starting at the cell specified by rowIndex and columnIndex.|
||[rows](/javascript/api/powerpoint/powerpoint.table#powerpoint-powerpoint-table-rows-member)|Gets the collection of rows in the table.|
|[TableAddOptions](/javascript/api/powerpoint/powerpoint.tableaddoptions)|[style](/javascript/api/powerpoint/powerpoint.tableaddoptions#powerpoint-powerpoint-tableaddoptions-style-member)|Specifies value that represents the table style.|
|[TableCell](/javascript/api/powerpoint/powerpoint.tablecell)|[borders](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-borders-member)|Gets the collection of borders for the table cell.|
||[fill](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-fill-member)|Gets the fill color of the table cell.|
||[font](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-font-member)|Gets the font of the table cell.|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-horizontalalignment-member)|Specifies the horizontal alignment of the table cell.|
||[indentLevel](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-indentlevel-member)|Specifies the indent level of the table cell.|
||[margins](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-margins-member)|Gets the set of margins in the table cell.|
||[resize(rowCount: number, columnCount: number)](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-resize-member(1))|Resizes the table cell to span across a specified number of rows and columns.|
||[split(rowCount: number, columnCount: number)](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
||[textRuns](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-textruns-member)|Specifies the contents of the table cell as an array of TextRun objects.|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.tablecell#powerpoint-powerpoint-tablecell-verticalalignment-member)|Specifies the vertical alignment of the text in the table cell.|
|[TableClearOptions](/javascript/api/powerpoint/powerpoint.tableclearoptions)|[all](/javascript/api/powerpoint/powerpoint.tableclearoptions#powerpoint-powerpoint-tableclearoptions-all-member)|Specifies if both values and formatting of the table should be cleared.|
||[format](/javascript/api/powerpoint/powerpoint.tableclearoptions#powerpoint-powerpoint-tableclearoptions-format-member)|Specifies if the formatting of the table should be cleared.|
||[text](/javascript/api/powerpoint/powerpoint.tableclearoptions#powerpoint-powerpoint-tableclearoptions-text-member)|Specifies if the values of the table should be cleared.|
|[TableColumn](/javascript/api/powerpoint/powerpoint.tablecolumn)|[columnIndex](/javascript/api/powerpoint/powerpoint.tablecolumn#powerpoint-powerpoint-tablecolumn-columnindex-member)|Returns the index number of the column within the column collection of the table.|
||[delete()](/javascript/api/powerpoint/powerpoint.tablecolumn#powerpoint-powerpoint-tablecolumn-delete-member(1))|Deletes the column.|
||[width](/javascript/api/powerpoint/powerpoint.tablecolumn#powerpoint-powerpoint-tablecolumn-width-member)|Retrieves the width of the column in points.|
|[TableColumnCollection](/javascript/api/powerpoint/powerpoint.tablecolumncollection)|[add(index?: number \| null \| undefined, count?: number \| undefined)](/javascript/api/powerpoint/powerpoint.tablecolumncollection#powerpoint-powerpoint-tablecolumncollection-add-member(1))|Adds one or more columns to the table.|
||[deleteColumns(columns: PowerPoint.TableColumn[])](/javascript/api/powerpoint/powerpoint.tablecolumncollection#powerpoint-powerpoint-tablecolumncollection-deletecolumns-member(1))|Deletes the specified columns from the collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tablecolumncollection#powerpoint-powerpoint-tablecolumncollection-getcount-member(1))|Gets the number of columns in the collection.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tablecolumncollection#powerpoint-powerpoint-tablecolumncollection-getitemat-member(1))|Gets the column using its zero-based index in the collection.|
||[items](/javascript/api/powerpoint/powerpoint.tablecolumncollection#powerpoint-powerpoint-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/powerpoint/powerpoint.tablerow)|[currentHeight](/javascript/api/powerpoint/powerpoint.tablerow#powerpoint-powerpoint-tablerow-currentheight-member)|Retrieves the current height of the row in points.|
||[delete()](/javascript/api/powerpoint/powerpoint.tablerow#powerpoint-powerpoint-tablerow-delete-member(1))|Deletes the row.|
||[height](/javascript/api/powerpoint/powerpoint.tablerow#powerpoint-powerpoint-tablerow-height-member)|Specifies the height of the row in points.|
||[rowIndex](/javascript/api/powerpoint/powerpoint.tablerow#powerpoint-powerpoint-tablerow-rowindex-member)|Returns the index number of the row within the rows collection of the table.|
|[TableRowCollection](/javascript/api/powerpoint/powerpoint.tablerowcollection)|[add(index?: number \| null \| undefined, count?: number \| undefined)](/javascript/api/powerpoint/powerpoint.tablerowcollection#powerpoint-powerpoint-tablerowcollection-add-member(1))|Adds one or more rows to the table.|
||[deleteRows(rows: PowerPoint.TableRow[])](/javascript/api/powerpoint/powerpoint.tablerowcollection#powerpoint-powerpoint-tablerowcollection-deleterows-member(1))|Deletes the specified rows from the collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tablerowcollection#powerpoint-powerpoint-tablerowcollection-getcount-member(1))|Gets the number of rows in the collection.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tablerowcollection#powerpoint-powerpoint-tablerowcollection-getitemat-member(1))|Gets the row using its zero-based index in the collection.|
||[items](/javascript/api/powerpoint/powerpoint.tablerowcollection#powerpoint-powerpoint-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
