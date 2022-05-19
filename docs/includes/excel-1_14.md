| Class | Fields | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcolumncriteria-member(1))|Clears the column filter criteria of the AutoFilter.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|Clears the filter criteria and sort state of the AutoFilter.|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-deleteshiftdirection-member)|Represents the direction (such as up or to the left) that the remaining cells will shift when a cell or cells are deleted.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-insertshiftdirection-member)|Represents the direction (such as down or to the right) that the existing cells will shift when a new cell or cells are inserted.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatable-member(1))|Gets the data table on the chart.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1))|Gets the data table on the chart.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-format-member)|Represents the format of a chart data table, which includes fill, font, and border format.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showhorizontalborder-member)|Specifies whether to display the horizontal border of the data table.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showlegendkey-member)|Specifies whether to show the legend key of the data table.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showoutlineborder-member)|Specifies whether to display the outline border of the data table.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showverticalborder-member)|Specifies whether to display the vertical border of the data table.|
||[visible](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-visible-member)|Specifies whether to show the data table of the chart.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-border-member)|Represents the border format of chart data table, which includes color, line style, and weight.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-font-member)|Represents the font attributes (such as font name, font size, and color) for the current object.|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|The comment's content.|
||[contentType](/javascript/api/excel/excel.comment#excel-excel-comment-contenttype-member)|Gets the content type of the comment.|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|Gets the creation time of the comment.|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|Deletes the comment and all the connected replies.|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|Gets the cell where this comment is located.|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|Specifies the comment identifier.|
||[mentions](/javascript/api/excel/excel.comment#excel-excel-comment-mentions-member)|Gets the entities (e.g., people) that are mentioned in comments.|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|Represents a collection of reply objects associated with the comment.|
||[resolved](/javascript/api/excel/excel.comment#excel-excel-comment-resolved-member)|The comment thread status.|
||[richContent](/javascript/api/excel/excel.comment#excel-excel-comment-richcontent-member)|Gets the rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#excel-excel-comment-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemornullobject-member(1))|Gets a comment from the collection based on its ID.|
||[onAdded](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member)|Occurs when the comments are added.|
||[onChanged](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member)|Occurs when comments or replies in a comment collection are changed, including when replies are deleted.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member)|Occurs when comments are deleted in the comment collection.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|The comment reply's content.|
||[contentType](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-contenttype-member)|The content type of the reply.|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|Gets the creation time of the comment reply.|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|Gets the parent comment of this reply.|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|Specifies the comment reply identifier.|
||[mentions](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-mentions-member)|The entities (e.g., people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-resolved-member)|The comment reply status.|
||[richContent](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-richcontent-member)|The rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemornullobject-member(1))|Returns a comment reply identified by its ID.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[cellValue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalue-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[cellValueOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-cellvalueornullobject-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[colorScale](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscale-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[colorScaleOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-colorscaleornullobject-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[custom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-custom-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-customornullobject-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databar-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-databarornullobject-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[delete()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-delete-member(1))|Deletes this conditional format.|
||[getRange()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrange-member(1))|Returns the range the conditonal format is applied to.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getrangeornullobject-member(1))|Returns the range to which the conditonal format is applied.|
||[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditonal format is applied.|
||[iconSet](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconset-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[iconSetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-iconsetornullobject-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[id](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-id-member)|The priority of the conditional format in the current `ConditionalFormatCollection`.|
||[preset](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-preset-member)|Returns the preset criteria conditional format.|
||[presetOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-presetornullobject-member)|Returns the preset criteria conditional format.|
||[priority](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-priority-member)|The priority (or index) within the conditional format collection that this conditional format currently exists in.|
||[stopIfTrue](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-stopiftrue-member)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparison-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-textcomparisonornullobject-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottom-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[topBottomOrNullObject](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-topbottomornullobject-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[type](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-type-member)|A type of conditional format.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemornullobject-member(1))|Returns a conditional format identified by its ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemornullobject-member(1))|Gets a shape using its name or ID.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|Represents the width of the arrowhead at the beginning of the specified line.|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|Represents the shape to which the beginning of the specified line is attached.|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|Represents the connection site to which the beginning of a connector is connected.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|Attaches the beginning of the specified connector to a specified shape.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|Attaches the end of the specified connector to a specified shape.|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|Represents the connector type for the line.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|Detaches the beginning of the specified connector from a shape.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|Detaches the end of the specified connector from a shape.|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|Represents the width of the arrowhead at the end of the specified line.|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|Represents the shape to which the end of the specified line is attached.|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|Represents the connection site to which the end of a connector is connected.|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|Specifies the shape identifier.|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|Specifies if the beginning of the specified line is connected to a shape.|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|Specifies if the end of the specified line is connected to a shape.|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|Returns the `Shape` object associated with the line.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#excel-excel-query-error-member)|Gets the query error message from when the query was last refreshed.|
||[loadedTo](/javascript/api/excel/excel.query#excel-excel-query-loadedto-member)|Gets the query loaded to object type.|
||[loadedToDataModel](/javascript/api/excel/excel.query#excel-excel-query-loadedtodatamodel-member)|Specifies if the query loaded to the data model.|
||[name](/javascript/api/excel/excel.query#excel-excel-query-name-member)|Gets the name of the query.|
||[refreshDate](/javascript/api/excel/excel.query#excel-excel-query-refreshdate-member)|Gets the date and time when the query was last refreshed.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#excel-excel-query-rowsloadedcount-member)|Gets the number of rows that were loaded when the query was last refreshed.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getcount-member(1))|Gets the number of queries in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getitem-member(1))|Gets a query from the collection based on its name.|
||[items](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedents of a cell in the same worksheet or in multiple worksheets.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1))|Returns a range object that is the edge cell of the data region that corresponds to the provided direction.|
||[getResizedRange(deltaRows: number, deltaColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getresizedrange-member(1))|Gets a `Range` object similar to the current `Range` object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|Gets a row contained in the range.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.|
||[getRowsAbove(count?: number)](/javascript/api/excel/excel.range#excel-excel-range-getrowsabove-member(1))|Gets a certain number of rows above the current `Range` object.|
||[getRowsBelow(count?: number)](/javascript/api/excel/excel.range#excel-excel-range-getrowsbelow-member(1))|Gets a certain number of rows below the current `Range` object.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|Specifies the alternative description text for a `Shape` object.|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|Specifies the alternative title text for a `Shape` object.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|Returns the number of connection sites on this shape.|
||[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|Copies and pastes a `Shape` object.|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|Removes the shape from the worksheet.|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|Returns the fill formatting of this shape.|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|Returns the geometric shape associated with the shape.|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|Specifies the geometric shape type of this geometric shape.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|Converts the shape to an image and returns the image as a base64-encoded string.|
||[group](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|Returns the shape group associated with the shape.|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|Specifies the height, in points, of the shape.|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|Specifies the shape identifier.|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|Returns the image associated with the shape.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|Moves the shape vertically by the specified number of points.|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|Specifies the level of the specified shape.|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|Returns the line associated with the shape.|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|Returns the line formatting of this shape.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|Specifies if the aspect ratio of this shape is locked.|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|Specifies the name of the shape.|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|Occurs when the shape is activated.|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|Occurs when the shape is deactivated.|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|Specifies the parent group of this shape.|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|Represents how the object is attached to the cells below it.|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|Returns the text frame object of this shape.|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|Returns the type of this shape.|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|Specifies if the shape is visible.|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|Specifies the width, in points, of the shape.|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemornullobject-member(1))|Gets a shape using its name or ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|Gets a style by name.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|Represents the `AutoFilter` object of the table.|
||[clearFilters()](/javascript/api/excel/excel.table#excel-excel-table-clearfilters-member(1))|Clears all the filters currently applied on the table.|
||[columns](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|Represents a collection of all the columns in the table.|
||[convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1))|Converts the table into a normal range of cells.|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|Deletes the table.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|Gets the range object associated with the data body of the table.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|Gets the range object associated with the header row of the table.|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|Gets the range object associated with the entire table.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|Gets the range object associated with the totals row of the table.|
||[highlightFirstColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightfirstcolumn-member)|Specifies if the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightlastcolumn-member)|Specifies if the last column contains special formatting.|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|Returns a value that uniquely identifies the table in a given workbook.|
||[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|Returns a numeric ID.|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|Name of the table.|
||[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|Occurs when data in cells changes on a specific table.|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|Occurs when the selection changes on a specific table.|
||[reapplyFilters()](/javascript/api/excel/excel.table#excel-excel-table-reapplyfilters-member(1))|Reapplies all the filters currently on the table.|
||[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1))|Resize the table to the new range.|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|Represents a collection of all the rows in the table.|
||[showBandedColumns](/javascript/api/excel/excel.table#excel-excel-table-showbandedcolumns-member)|Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.table#excel-excel-table-showbandedrows-member)|Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.table#excel-excel-table-showfilterbutton-member)|Specifies if the filter buttons are visible at the top of each column header.|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|Specifies if the header row is visible.|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|Specifies if the total row is visible.|
||[sort](/javascript/api/excel/excel.table#excel-excel-table-sort-member)|Represents the sorting for the table.|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|Constant value that represents the table style.|
||[worksheet](/javascript/api/excel/excel.table#excel-excel-table-worksheet-member)|The worksheet containing the current table.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitemornullobject-member(1))|Gets a table by name or ID.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|Creates a blank `TableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|Gets the number of table styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|Gets the default table style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|Gets a `TableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|Gets a `TableStyle` by name.|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|Sets the default table style for use in the parent object's scope.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changedirectionstate-member)|Represents a change to the direction that the cells in a worksheet will shift when a cell or cells are deleted or inserted.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-triggersource-member)|Represents the trigger source of the event.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-isprotected-member)|Gets the current protection status of the worksheet.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the protection status is changed.|
