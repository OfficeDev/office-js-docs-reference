| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)||[Body](/javascript/api/word/word.body)|[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the body object.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Occurs when the content control is deleted.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Occurs when selection within the content control is changed.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the content control object.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|The event type.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Gets the full XML content of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Inserts the given XML under the parent element identified by xpath at child position index.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Queries the XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#id)|Gets the ID of the custom XML part.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Gets the namespace URI of the custom XML part.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Gets a custom XML part based on its ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Gets a bookmark's range.|
||[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Gets the custom XML parts in the document.|
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Occurs when a content control is added.|
||[settings](/javascript/api/word/word.document#settings)|Gets the add-in's settings in the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Gets a bookmark's range.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Gets the custom XML parts in the document.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Gets the add-in's settings in the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Sets the picture at the specified level in the list.|
|[Paragraph](/javascript/api/word/word.paragraph)|[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Gets the names all bookmarks in or overlapping the range.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|Inserts a bookmark on the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the range object.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Deletes the setting.|
||[key](/javascript/api/word/word.setting#key)|Gets the key of the setting.|
||[value](/javascript/api/word/word.setting#value)|Gets or sets the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Gets a setting object by its key, which is case-sensitive.|
||[items](/javascript/api/word/word.settingcollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Merges the cells bounded inclusively by a first and last cell.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the table object.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Inserts a content control on the row.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Merges the row into one cell.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the row.|
