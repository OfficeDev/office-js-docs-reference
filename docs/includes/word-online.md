| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the content control.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the content control.|
|[Document](/javascript/api/word/word.document)|[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Inserts an endnote.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Inserts a footnote.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|
