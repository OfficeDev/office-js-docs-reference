| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[openDocument(filePath: string)](/javascript/api/word/word.application#word-word-application-opendocument-member(1))|Opens a document and displays it in a new tab or window.|
|[Body](/javascript/api/word/word.body)|[getTrackedChanges()](/javascript/api/word/word.body#word-word-body-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getTrackedChanges()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the content control.|
|[Document](/javascript/api/word/word.document)|[getParagraphByUniqueLocalId(id: string)](/javascript/api/word/word.document#word-word-document-getparagraphbyuniquelocalid-member(1))|Gets the paragraph by its unique local ID.|
||[importStylesFromJson(stylesJson: string, importedStylesConflictBehavior?: Word.ImportedStylesConflictBehavior)](/javascript/api/word/word.document#word-word-document-importstylesfromjson-member(1))|Import styles from a JSON-formatted string.|
||[onParagraphAdded](/javascript/api/word/word.document#word-word-document-onparagraphadded-member)|Occurs when the user adds new paragraphs.|
||[onParagraphChanged](/javascript/api/word/word.document#word-word-document-onparagraphchanged-member)|Occurs when the user changes paragraphs.|
||[onParagraphDeleted](/javascript/api/word/word.document#word-word-document-onparagraphdeleted-member)|Occurs when the user deletes paragraphs.|
|[InsertFileOptions](/javascript/api/word/word.insertfileoptions)|[importCustomProperties](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importcustomproperties-member)|Represents whether the custom properties from the source document should be imported.|
||[importCustomXmlParts](/javascript/api/word/word.insertfileoptions#word-word-insertfileoptions-importcustomxmlparts-member)|Represents whether the custom XML parts from the source document should be imported.|
|[Paragraph](/javascript/api/word/word.paragraph)|[getTrackedChanges()](/javascript/api/word/word.paragraph#word-word-paragraph-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the paragraph.|
||[uniqueLocalId](/javascript/api/word/word.paragraph#word-word-paragraph-uniquelocalid-member)|Gets a string that represents the paragraph identifier in the current session.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[source](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[source](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[source](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[Range](/javascript/api/word/word.range)|[getTrackedChanges()](/javascript/api/word/word.range#word-word-range-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the range.|
|[Shading](/javascript/api/word/word.shading)|[backgroundPatternColor](/javascript/api/word/word.shading#word-word-shading-backgroundpatterncolor-member)|Specifies the color for the background of the object.|
|[Style](/javascript/api/word/word.style)|[shading](/javascript/api/word/word.style#word-word-style-shading-member)|Gets a Shading object that represents the shading for the specified style.|
||[tableStyle](/javascript/api/word/word.style#word-word-style-tablestyle-member)|Gets a TableStyle object representing Style properties that can be applied to a table.|
|[TableStyle](/javascript/api/word/word.tablestyle)|[bottomCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-bottomcellmargin-member)|Specifies the amount of space to add between the contents and the bottom borders of the cells.|
||[cellSpacing](/javascript/api/word/word.tablestyle#word-word-tablestyle-cellspacing-member)|Specifies the spacing (in points) between the cells in a table style.|
||[leftCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-leftcellmargin-member)|Specifies the amount of space to add between the contents and the left borders of the cells.|
||[rightCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-rightcellmargin-member)|Specifies the amount of space to add between the contents and the right borders of the cells.|
||[topCellMargin](/javascript/api/word/word.tablestyle#word-word-tablestyle-topcellmargin-member)|Specifies the amount of space to add between the contents and the top borders of the cells.|
|[TrackedChange](/javascript/api/word/word.trackedchange)|[accept()](/javascript/api/word/word.trackedchange#word-word-trackedchange-accept-member(1))|Accepts the tracked change.|
||[author](/javascript/api/word/word.trackedchange#word-word-trackedchange-author-member)|Gets the author of the tracked change.|
||[date](/javascript/api/word/word.trackedchange#word-word-trackedchange-date-member)|Gets the date of the tracked change.|
||[getNext()](/javascript/api/word/word.trackedchange#word-word-trackedchange-getnext-member(1))|Gets the next tracked change.|
||[getNextOrNullObject()](/javascript/api/word/word.trackedchange#word-word-trackedchange-getnextornullobject-member(1))|Gets the next tracked change.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/javascript/api/word/word.trackedchange#word-word-trackedchange-getrange-member(1))|Gets the range of the tracked change.|
||[reject()](/javascript/api/word/word.trackedchange#word-word-trackedchange-reject-member(1))|Rejects the tracked change.|
||[text](/javascript/api/word/word.trackedchange#word-word-trackedchange-text-member)|Gets the text of the tracked change.|
||[type](/javascript/api/word/word.trackedchange#word-word-trackedchange-type-member)|Gets the type of the tracked change.|
|[TrackedChangeCollection](/javascript/api/word/word.trackedchangecollection)|[acceptAll()](/javascript/api/word/word.trackedchangecollection#word-word-trackedchangecollection-acceptall-member(1))|Accepts all the tracked changes in the collection.|
||[getFirst()](/javascript/api/word/word.trackedchangecollection#word-word-trackedchangecollection-getfirst-member(1))|Gets the first TrackedChange in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.trackedchangecollection#word-word-trackedchangecollection-getfirstornullobject-member(1))|Gets the first TrackedChange in this collection.|
||[items](/javascript/api/word/word.trackedchangecollection#word-word-trackedchangecollection-items-member)|Gets the loaded child items in this collection.|
||[rejectAll()](/javascript/api/word/word.trackedchangecollection#word-word-trackedchangecollection-rejectall-member(1))|Rejects all the tracked changes in the collection.|
