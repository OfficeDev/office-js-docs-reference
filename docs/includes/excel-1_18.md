| Class | Fields | Description |
|:---|:---|:---|
|[CheckboxCellControl](/.checkboxcellcontrol)|[type](/.checkboxcellcontrol#excel-javascript/api/excel/-checkboxcellcontrol-type-member)|Represents an interactable control inside of a cell.|
|[EmptyCellControl](/.emptycellcontrol)|[type](/.emptycellcontrol#excel-javascript/api/excel/-emptycellcontrol-type-member)||
|[MixedCellControl](/.mixedcellcontrol)|[type](/.mixedcellcontrol#excel-javascript/api/excel/-mixedcellcontrol-type-member)||
|[Note](/.note)|[authorName](/.note#excel-javascript/api/excel/-note-authorname-member)|Gets the author of the note.|
||[content](/.note#excel-javascript/api/excel/-note-content-member)|Specifies the text of the note.|
||[delete()](/.note#excel-javascript/api/excel/-note-delete-member(1))|Deletes the note.|
||[getLocation()](/.note#excel-javascript/api/excel/-note-getlocation-member(1))|Gets the cell where this note is located.|
||[height](/.note#excel-javascript/api/excel/-note-height-member)|Specifies the height of the note.|
||[visible](/.note#excel-javascript/api/excel/-note-visible-member)|Specifies the visibility of the note.|
||[width](/.note#excel-javascript/api/excel/-note-width-member)|Specifies the width of the note.|
|[NoteCollection](/.notecollection)|[add(cellAddress: Range \| string, content: any)](/.notecollection#excel-javascript/api/excel/-notecollection-add-member(1))|Adds a new note with the given content on the given cell.|
||[getCount()](/.notecollection#excel-javascript/api/excel/-notecollection-getcount-member(1))|Gets the number of notes in the collection.|
||[getItem(key: string)](/.notecollection#excel-javascript/api/excel/-notecollection-getitem-member(1))|Gets a note by its cell address.|
||[getItemAt(index: number)](/.notecollection#excel-javascript/api/excel/-notecollection-getitemat-member(1))|Gets a note from the collection based on its position.|
||[getItemOrNullObject(key: string)](/.notecollection#excel-javascript/api/excel/-notecollection-getitemornullobject-member(1))|Gets a note by its cell address.|
||[items](/.notecollection#excel-javascript/api/excel/-notecollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/.range)|[clearOrResetContents()](/.range#excel-javascript/api/excel/-range-clearorresetcontents-member(1))|Clears the values of the cells in the range, with special consideration given to cells containing controls.|
||[control](/.range#excel-javascript/api/excel/-range-control-member)|Accesses the cell control applied to this range.|
|[RangeAreas](/.rangeareas)|[clearOrResetContents()](/.rangeareas#excel-javascript/api/excel/-rangeareas-clearorresetcontents-member(1))|Clears the values of the cells in the ranges, with special consideration given to cells containing controls.|
||[select()](/.rangeareas#excel-javascript/api/excel/-rangeareas-select-member(1))|Selects the specified range areas in the Excel UI.|
|[RangeTextRun](/.rangetextrun)|[font](/.rangetextrun#excel-javascript/api/excel/-rangetextrun-font-member)|The font attributes (such as font name, font size, and color) applied to this text run.|
||[text](/.rangetextrun#excel-javascript/api/excel/-rangetextrun-text-member)|The text of this text run.|
|[SettableCellProperties](/.settablecellproperties)|[textRuns](/.settablecellproperties#excel-javascript/api/excel/-settablecellproperties-textruns-member)|Represents the `textRuns` property.|
|[UnknownCellControl](/.unknowncellcontrol)|[type](/.unknowncellcontrol#excel-javascript/api/excel/-unknowncellcontrol-type-member)||
|[Workbook](/.workbook)|[notes](/.workbook#excel-javascript/api/excel/-workbook-notes-member)|Returns a collection of all the notes objects in the workbook.|
|[Worksheet](/.worksheet)|[notes](/.worksheet#excel-javascript/api/excel/-worksheet-notes-member)|Returns a collection of all the notes objects in the worksheet.|
