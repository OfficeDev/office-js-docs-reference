| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[openDocument(filePath: string)](/.application#word-javascript/api/word/-application-opendocument-member(1))|Opens a document and displays it in a new tab or window.|
|[Body](/.body)|[getTrackedChanges()](/.body#word-javascript/api/word/-body-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the body.|
|[ContentControl](/.contentcontrol)|[getTrackedChanges()](/.contentcontrol#word-javascript/api/word/-contentcontrol-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the content control.|
|[Document](/.document)|[getParagraphByUniqueLocalId(id: string)](/.document#word-javascript/api/word/-document-getparagraphbyuniquelocalid-member(1))|Gets the paragraph by its unique local ID.|
||[importStylesFromJson(stylesJson: string, importedStylesConflictBehavior?: Word.ImportedStylesConflictBehavior)](/.document#word-javascript/api/word/-document-importstylesfromjson-member(1))|Import styles from a JSON-formatted string.|
||[onParagraphAdded](/.document#word-javascript/api/word/-document-onparagraphadded-member)|Occurs when the user adds new paragraphs.|
||[onParagraphChanged](/.document#word-javascript/api/word/-document-onparagraphchanged-member)|Occurs when the user changes paragraphs.|
||[onParagraphDeleted](/.document#word-javascript/api/word/-document-onparagraphdeleted-member)|Occurs when the user deletes paragraphs.|
|[InsertFileOptions](/.insertfileoptions)|[importCustomProperties](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importcustomproperties-member)|Represents whether the custom properties from the source document should be imported.|
||[importCustomXmlParts](/.insertfileoptions#word-javascript/api/word/-insertfileoptions-importcustomxmlparts-member)|Represents whether the custom XML parts from the source document should be imported.|
|[Paragraph](/.paragraph)|[getTrackedChanges()](/.paragraph#word-javascript/api/word/-paragraph-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the paragraph.|
||[uniqueLocalId](/.paragraph#word-javascript/api/word/-paragraph-uniquelocalid-member)|Gets a string that represents the paragraph identifier in the current session.|
|[ParagraphAddedEventArgs](/.paragraphaddedeventargs)|[source](/.paragraphaddedeventargs#word-javascript/api/word/-paragraphaddedeventargs-source-member)|The source of the event.|
||[type](/.paragraphaddedeventargs#word-javascript/api/word/-paragraphaddedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/.paragraphaddedeventargs#word-javascript/api/word/-paragraphaddedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[ParagraphChangedEventArgs](/.paragraphchangedeventargs)|[source](/.paragraphchangedeventargs#word-javascript/api/word/-paragraphchangedeventargs-source-member)|The source of the event.|
||[type](/.paragraphchangedeventargs#word-javascript/api/word/-paragraphchangedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/.paragraphchangedeventargs#word-javascript/api/word/-paragraphchangedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[ParagraphDeletedEventArgs](/.paragraphdeletedeventargs)|[source](/.paragraphdeletedeventargs#word-javascript/api/word/-paragraphdeletedeventargs-source-member)|The source of the event.|
||[type](/.paragraphdeletedeventargs#word-javascript/api/word/-paragraphdeletedeventargs-type-member)|The event type.|
||[uniqueLocalIds](/.paragraphdeletedeventargs#word-javascript/api/word/-paragraphdeletedeventargs-uniquelocalids-member)|Gets the unique IDs of the involved paragraphs.|
|[Range](/.range)|[getTrackedChanges()](/.range#word-javascript/api/word/-range-gettrackedchanges-member(1))|Gets the collection of the TrackedChange objects in the range.|
|[Shading](/.shading)|[backgroundPatternColor](/.shading#word-javascript/api/word/-shading-backgroundpatterncolor-member)|Specifies the color for the background of the object.|
|[Style](/.style)|[shading](/.style#word-javascript/api/word/-style-shading-member)|Gets a Shading object that represents the shading for the specified style.|
||[tableStyle](/.style#word-javascript/api/word/-style-tablestyle-member)|Gets a TableStyle object representing Style properties that can be applied to a table.|
|[TableStyle](/.tablestyle)|[bottomCellMargin](/.tablestyle#word-javascript/api/word/-tablestyle-bottomcellmargin-member)|Specifies the amount of space to add between the contents and the bottom borders of the cells.|
||[cellSpacing](/.tablestyle#word-javascript/api/word/-tablestyle-cellspacing-member)|Specifies the spacing (in points) between the cells in a table style.|
||[leftCellMargin](/.tablestyle#word-javascript/api/word/-tablestyle-leftcellmargin-member)|Specifies the amount of space to add between the contents and the left borders of the cells.|
||[rightCellMargin](/.tablestyle#word-javascript/api/word/-tablestyle-rightcellmargin-member)|Specifies the amount of space to add between the contents and the right borders of the cells.|
||[topCellMargin](/.tablestyle#word-javascript/api/word/-tablestyle-topcellmargin-member)|Specifies the amount of space to add between the contents and the top borders of the cells.|
|[TrackedChange](/.trackedchange)|[accept()](/.trackedchange#word-javascript/api/word/-trackedchange-accept-member(1))|Accepts the tracked change.|
||[author](/.trackedchange#word-javascript/api/word/-trackedchange-author-member)|Gets the author of the tracked change.|
||[date](/.trackedchange#word-javascript/api/word/-trackedchange-date-member)|Gets the date of the tracked change.|
||[getNext()](/.trackedchange#word-javascript/api/word/-trackedchange-getnext-member(1))|Gets the next tracked change.|
||[getNextOrNullObject()](/.trackedchange#word-javascript/api/word/-trackedchange-getnextornullobject-member(1))|Gets the next tracked change.|
||[getRange(rangeLocation?: Word.RangeLocation.whole \| Word.RangeLocation.start \| Word.RangeLocation.end \| "Whole" \| "Start" \| "End")](/.trackedchange#word-javascript/api/word/-trackedchange-getrange-member(1))|Gets the range of the tracked change.|
||[reject()](/.trackedchange#word-javascript/api/word/-trackedchange-reject-member(1))|Rejects the tracked change.|
||[text](/.trackedchange#word-javascript/api/word/-trackedchange-text-member)|Gets the text of the tracked change.|
||[type](/.trackedchange#word-javascript/api/word/-trackedchange-type-member)|Gets the type of the tracked change.|
|[TrackedChangeCollection](/.trackedchangecollection)|[acceptAll()](/.trackedchangecollection#word-javascript/api/word/-trackedchangecollection-acceptall-member(1))|Accepts all the tracked changes in the collection.|
||[getFirst()](/.trackedchangecollection#word-javascript/api/word/-trackedchangecollection-getfirst-member(1))|Gets the first TrackedChange in this collection.|
||[getFirstOrNullObject()](/.trackedchangecollection#word-javascript/api/word/-trackedchangecollection-getfirstornullobject-member(1))|Gets the first TrackedChange in this collection.|
||[items](/.trackedchangecollection#word-javascript/api/word/-trackedchangecollection-items-member)|Gets the loaded child items in this collection.|
||[rejectAll()](/.trackedchangecollection#word-javascript/api/word/-trackedchangecollection-rejectall-member(1))|Rejects all the tracked changes in the collection.|
