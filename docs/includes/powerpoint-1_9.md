| Class | Fields | Description |
|:---|:---|:---|
|[Border](/.border)|[color](/.border#powerpoint-javascript/api/powerpoint/-border-color-member)|Represents the line color in the hexadecimal format #RRGGBB (e.g., "FFA500") or as a named HTML color value (e.g., "orange").|
||[dashStyle](/.border#powerpoint-javascript/api/powerpoint/-border-dashstyle-member)|Represents the dash style of the line.|
||[transparency](/.border#powerpoint-javascript/api/powerpoint/-border-transparency-member)|Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear).|
||[weight](/.border#powerpoint-javascript/api/powerpoint/-border-weight-member)|Represents the weight of the line, in points.|
|[Borders](/.borders)|[bottom](/.borders#powerpoint-javascript/api/powerpoint/-borders-bottom-member)|Gets the bottom border.|
||[diagonalDown](/.borders#powerpoint-javascript/api/powerpoint/-borders-diagonaldown-member)|Gets the diagonal border (top-left to bottom-right).|
||[diagonalUp](/.borders#powerpoint-javascript/api/powerpoint/-borders-diagonalup-member)|Gets the diagonal border (bottom-left to top-right).|
||[left](/.borders#powerpoint-javascript/api/powerpoint/-borders-left-member)|Gets the left border.|
||[right](/.borders#powerpoint-javascript/api/powerpoint/-borders-right-member)|Gets the right border.|
||[top](/.borders#powerpoint-javascript/api/powerpoint/-borders-top-member)|Gets the top border.|
|[Margins](/.margins)|[bottom](/.margins#powerpoint-javascript/api/powerpoint/-margins-bottom-member)|Specifies the bottom margin in points.|
||[left](/.margins#powerpoint-javascript/api/powerpoint/-margins-left-member)|Specifies the left margin in points.|
||[right](/.margins#powerpoint-javascript/api/powerpoint/-margins-right-member)|Specifies the right margin in points.|
||[top](/.margins#powerpoint-javascript/api/powerpoint/-margins-top-member)|Specifies the top margin in points.|
|[Table](/.table)|[clear(options?: PowerPoint.TableClearOptions)](/.table#powerpoint-javascript/api/powerpoint/-table-clear-member(1))|Clears table values and formatting.|
||[columns](/.table#powerpoint-javascript/api/powerpoint/-table-columns-member)|Gets the collection of columns in the table.|
||[mergeCells(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number)](/.table#powerpoint-javascript/api/powerpoint/-table-mergecells-member(1))|Creates a merged area starting at the cell specified by rowIndex and columnIndex.|
||[rows](/.table#powerpoint-javascript/api/powerpoint/-table-rows-member)|Gets the collection of rows in the table.|
|[TableAddOptions](/.tableaddoptions)|[style](/.tableaddoptions#powerpoint-javascript/api/powerpoint/-tableaddoptions-style-member)|Specifies value that represents the table style.|
|[TableCell](/.tablecell)|[borders](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-borders-member)|Gets the collection of borders for the table cell.|
||[fill](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-fill-member)|Gets the fill color of the table cell.|
||[font](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-font-member)|Gets the font of the table cell.|
||[horizontalAlignment](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-horizontalalignment-member)|Specifies the horizontal alignment of the table cell.|
||[indentLevel](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-indentlevel-member)|Specifies the indent level of the table cell.|
||[margins](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-margins-member)|Gets the set of margins in the table cell.|
||[resize(rowCount: number, columnCount: number)](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-resize-member(1))|Resizes the table cell to span across a specified number of rows and columns.|
||[split(rowCount: number, columnCount: number)](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
||[textRuns](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-textruns-member)|Specifies the contents of the table cell as an array of PowerPoint.TextRun objects.|
||[verticalAlignment](/.tablecell#powerpoint-javascript/api/powerpoint/-tablecell-verticalalignment-member)|Specifies the vertical alignment of the text in the table cell.|
|[TableClearOptions](/.tableclearoptions)|[all](/.tableclearoptions#powerpoint-javascript/api/powerpoint/-tableclearoptions-all-member)|Specifies if both values and formatting of the table should be cleared.|
||[format](/.tableclearoptions#powerpoint-javascript/api/powerpoint/-tableclearoptions-format-member)|Specifies if the formatting of the table should be cleared.|
||[text](/.tableclearoptions#powerpoint-javascript/api/powerpoint/-tableclearoptions-text-member)|Specifies if the values of the table should be cleared.|
|[TableColumn](/.tablecolumn)|[columnIndex](/.tablecolumn#powerpoint-javascript/api/powerpoint/-tablecolumn-columnindex-member)|Returns the index number of the column within the column collection of the table.|
||[delete()](/.tablecolumn#powerpoint-javascript/api/powerpoint/-tablecolumn-delete-member(1))|Deletes the column.|
||[width](/.tablecolumn#powerpoint-javascript/api/powerpoint/-tablecolumn-width-member)|Retrieves the width of the column in points.|
|[TableColumnCollection](/.tablecolumncollection)|[add(index?: number \| null \| undefined, count?: number \| undefined)](/.tablecolumncollection#powerpoint-javascript/api/powerpoint/-tablecolumncollection-add-member(1))|Adds one or more columns to the table.|
||[deleteColumns(columns: PowerPoint.TableColumn[])](/.tablecolumncollection#powerpoint-javascript/api/powerpoint/-tablecolumncollection-deletecolumns-member(1))|Deletes the specified columns from the collection.|
||[getCount()](/.tablecolumncollection#powerpoint-javascript/api/powerpoint/-tablecolumncollection-getcount-member(1))|Gets the number of columns in the collection.|
||[getItemAt(index: number)](/.tablecolumncollection#powerpoint-javascript/api/powerpoint/-tablecolumncollection-getitemat-member(1))|Gets the column using its zero-based index in the collection.|
||[items](/.tablecolumncollection#powerpoint-javascript/api/powerpoint/-tablecolumncollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/.tablerow)|[currentHeight](/.tablerow#powerpoint-javascript/api/powerpoint/-tablerow-currentheight-member)|Retrieves the current height of the row in points.|
||[delete()](/.tablerow#powerpoint-javascript/api/powerpoint/-tablerow-delete-member(1))|Deletes the row.|
||[height](/.tablerow#powerpoint-javascript/api/powerpoint/-tablerow-height-member)|Specifies the height of the row in points.|
||[rowIndex](/.tablerow#powerpoint-javascript/api/powerpoint/-tablerow-rowindex-member)|Returns the index number of the row within the rows collection of the table.|
|[TableRowCollection](/.tablerowcollection)|[add(index?: number \| null \| undefined, count?: number \| undefined)](/.tablerowcollection#powerpoint-javascript/api/powerpoint/-tablerowcollection-add-member(1))|Adds one or more rows to the table.|
||[deleteRows(rows: PowerPoint.TableRow[])](/.tablerowcollection#powerpoint-javascript/api/powerpoint/-tablerowcollection-deleterows-member(1))|Deletes the specified rows from the collection.|
||[getCount()](/.tablerowcollection#powerpoint-javascript/api/powerpoint/-tablerowcollection-getcount-member(1))|Gets the number of rows in the collection.|
||[getItemAt(index: number)](/.tablerowcollection#powerpoint-javascript/api/powerpoint/-tablerowcollection-getitemat-member(1))|Gets the row using its zero-based index in the collection.|
||[items](/.tablerowcollection#powerpoint-javascript/api/powerpoint/-tablerowcollection-items-member)|Gets the loaded child items in this collection.|
