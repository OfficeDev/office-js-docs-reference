| Class | Fields | Description |
|:---|:---|:---|
|[CheckboxCellControl](/javascript/api/excel/excel.checkboxcellcontrol)|[type](/javascript/api/excel/excel.checkboxcellcontrol#excel-excel-checkboxcellcontrol-type-member)|Represents an interactable control inside of a cell.|
|[EmptyCellControl](/javascript/api/excel/excel.emptycellcontrol)|[type](/javascript/api/excel/excel.emptycellcontrol#excel-excel-emptycellcontrol-type-member)||
|[MixedCellControl](/javascript/api/excel/excel.mixedcellcontrol)|[type](/javascript/api/excel/excel.mixedcellcontrol#excel-excel-mixedcellcontrol-type-member)||
|[Note](/javascript/api/excel/excel.note)|[authorName](/javascript/api/excel/excel.note#excel-excel-note-authorname-member)|Gets the author of the note.|
||[content](/javascript/api/excel/excel.note#excel-excel-note-content-member)|Gets or sets the text of the note.|
||[delete()](/javascript/api/excel/excel.note#excel-excel-note-delete-member(1))|Deletes the note.|
||[getLocation()](/javascript/api/excel/excel.note#excel-excel-note-getlocation-member(1))|Gets the cell where this note is located.|
||[height](/javascript/api/excel/excel.note#excel-excel-note-height-member)|Specifies the height of the note.|
||[visible](/javascript/api/excel/excel.note#excel-excel-note-visible-member)|Specifies the visibility of the note.|
||[width](/javascript/api/excel/excel.note#excel-excel-note-width-member)|Specifies the width of the note.|
|[NoteCollection](/javascript/api/excel/excel.notecollection)|[add(cellAddress: Range \| string, content: any)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-add-member(1))|Adds a new note with the given content on the given cell.|
||[getCount()](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getcount-member(1))|Gets the number of notes in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getitem-member(1))|Gets a note by its cell address.|
||[getItemAt(index: number)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getitemat-member(1))|Gets a note from the collection based on its position.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-getitemornullobject-member(1))|Gets a note by its cell address.|
||[items](/javascript/api/excel/excel.notecollection#excel-excel-notecollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[clearOrResetContents()](/javascript/api/excel/excel.range#excel-excel-range-clearorresetcontents-member(1))|Clears the values of the cells in the range, with special consideration given to cells containing controls.|
||[control](/javascript/api/excel/excel.range#excel-excel-range-control-member)|Accesses the cell control applied to this range.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[clearOrResetContents()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clearorresetcontents-member(1))|Clears the values of the cells in the ranges, with special consideration given to cells containing controls.|
||[select()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-select-member(1))|Selects the specified range areas in the Excel UI.|
|[RangeTextRun](/javascript/api/excel/excel.rangetextrun)|[font](/javascript/api/excel/excel.rangetextrun#excel-excel-rangetextrun-font-member)|The font attributes (such as font name, font size, and color) applied to this text run.|
||[text](/javascript/api/excel/excel.rangetextrun#excel-excel-rangetextrun-text-member)|The text of this text run.|
|[UnknownCellControl](/javascript/api/excel/excel.unknowncellcontrol)|[type](/javascript/api/excel/excel.unknowncellcontrol#excel-excel-unknowncellcontrol-type-member)||
|[Workbook](/javascript/api/excel/excel.workbook)|[notes](/javascript/api/excel/excel.workbook#excel-excel-workbook-notes-member)|Returns a collection of all the notes objects in the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[notes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-notes-member)|Returns a collection of all the notes objects in the worksheet.|
