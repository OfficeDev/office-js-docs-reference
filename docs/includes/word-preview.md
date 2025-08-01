| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[onCommentAdded](/javascript/api/word/word.body#word-word-body-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.body#word-word-body-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.body#word-word-body-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.body#word-word-body-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.body#word-word-body-oncommentselected-member)|Occurs when a comment is selected.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Break](/javascript/api/word/word.break)|[pageIndex](/javascript/api/word/word.break#word-word-break-pageindex-member)|Returns the page number on which the break occurs.|
||[range](/javascript/api/word/word.break#word-word-break-range-member)|Returns a `Range` object that represents the portion of the document that's contained in the break.|
|[BreakCollection](/javascript/api/word/word.breakcollection)|[items](/javascript/api/word/word.breakcollection#word-word-breakcollection-items-member)|Gets the loaded child items in this collection.|
|[ColorFormat](/javascript/api/word/word.colorformat)|[brightness](/javascript/api/word/word.colorformat#word-word-colorformat-brightness-member)|Specifies the brightness of a specified shape color.|
||[objectThemeColor](/javascript/api/word/word.colorformat#word-word-colorformat-objectthemecolor-member)|Specifies the theme color for a color format.|
||[rgb](/javascript/api/word/word.colorformat#word-word-colorformat-rgb-member)|Specifies the red-green-blue (RGB) value of the specified color.|
||[tintAndShade](/javascript/api/word/word.colorformat#word-word-colorformat-tintandshade-member)|Specifies the lightening or darkening of a specified shape's color.|
||[type](/javascript/api/word/word.colorformat#word-word-colorformat-type-member)|Returns the shape color type.|
|[CommentDetail](/javascript/api/word/word.commentdetail)|[id](/javascript/api/word/word.commentdetail#word-word-commentdetail-id-member)|Represents the ID of this comment.|
||[replyIds](/javascript/api/word/word.commentdetail#word-word-commentdetail-replyids-member)|Represents the IDs of the replies to this comment.|
|[CommentEventArgs](/javascript/api/word/word.commenteventargs)|[changeType](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-changetype-member)|Represents how the comment changed event is triggered.|
||[commentDetails](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-commentdetails-member)|Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.|
||[source](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-source-member)|The source of the event.|
||[type](/javascript/api/word/word.commenteventargs#word-word-commenteventargs-type-member)|The event type.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onCommentAdded](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-oncommentselected-member)|Occurs when a comment is selected.|
||[resetState()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-resetstate-member(1))|Resets the state of the content control.|
||[setState(contentControlState: Word.ContentControlState)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-setstate-member(1))|Sets the state of the content control.|
||[xmlMapping](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-xmlmapping-member)|Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.|
|[ContentControlAddedEventArgs](/javascript/api/word/word.contentcontroladdedeventargs)|[eventType](/javascript/api/word/word.contentcontroladdedeventargs#word-word-contentcontroladdedeventargs-eventtype-member)|The event type.|
|[ContentControlDataChangedEventArgs](/javascript/api/word/word.contentcontroldatachangedeventargs)|[eventType](/javascript/api/word/word.contentcontroldatachangedeventargs#word-word-contentcontroldatachangedeventargs-eventtype-member)|The event type.|
|[ContentControlDeletedEventArgs](/javascript/api/word/word.contentcontroldeletedeventargs)|[eventType](/javascript/api/word/word.contentcontroldeletedeventargs#word-word-contentcontroldeletedeventargs-eventtype-member)|The event type.|
|[ContentControlEnteredEventArgs](/javascript/api/word/word.contentcontrolenteredeventargs)|[eventType](/javascript/api/word/word.contentcontrolenteredeventargs#word-word-contentcontrolenteredeventargs-eventtype-member)|The event type.|
|[ContentControlExitedEventArgs](/javascript/api/word/word.contentcontrolexitedeventargs)|[eventType](/javascript/api/word/word.contentcontrolexitedeventargs#word-word-contentcontrolexitedeventargs-eventtype-member)|The event type.|
|[ContentControlSelectionChangedEventArgs](/javascript/api/word/word.contentcontrolselectionchangedeventargs)|[eventType](/javascript/api/word/word.contentcontrolselectionchangedeventargs#word-word-contentcontrolselectionchangedeventargs-eventtype-member)|The event type.|
|[CustomXmlAddNodeOptions](/javascript/api/word/word.customxmladdnodeoptions)|[name](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-name-member)|If provided, specifies the base name of the element to be added.|
||[namespaceUri](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nextSibling](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nextsibling-member)|If provided, specifies the node which should become the next sibling of the new node.|
||[nodeType](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nodetype-member)|If provided, specifies the type of node to add.|
||[nodeValue](/javascript/api/word/word.customxmladdnodeoptions#word-word-customxmladdnodeoptions-nodevalue-member)|If provided, specifies the value of the added node for those nodes that allow text.|
|[CustomXmlAddSchemaOptions](/javascript/api/word/word.customxmladdschemaoptions)|[alias](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-alias-member)|If provided, specifies the alias of the schema to be added to the collection.|
||[fileName](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-filename-member)|If provided, specifies the location of the schema on a disk.|
||[installForAllUsers](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-installforallusers-member)|If provided, specifies whether, in the case where the schema is being added to the Schema Library, the Schema Library keys should be written to the registry (`HKEY_LOCAL_MACHINE` for all users or `HKEY_CURRENT_USER` for just the current user).|
||[namespaceUri](/javascript/api/word/word.customxmladdschemaoptions#word-word-customxmladdschemaoptions-namespaceuri-member)|If provided, specifies the namespace of the schema to be added to the collection.|
|[CustomXmlAddValidationErrorOptions](/javascript/api/word/word.customxmladdvalidationerroroptions)|[clearedOnUpdate](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-clearedonupdate-member)|If provided, specifies whether the error is to be cleared from the Word.CustomXmlValidationErrorCollection when the XML is corrected and updated.|
||[errorText](/javascript/api/word/word.customxmladdvalidationerroroptions#word-word-customxmladdvalidationerroroptions-errortext-member)|If provided, specifies the descriptive error text.|
|[CustomXmlAppendChildNodeOptions](/javascript/api/word/word.customxmlappendchildnodeoptions)|[name](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-name-member)|If provided, specifies the base name of the element to be appended.|
||[namespaceUri](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be appended.|
||[nodeType](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/javascript/api/word/word.customxmlappendchildnodeoptions#word-word-customxmlappendchildnodeoptions-nodevalue-member)|If provided, specifies the value of the appended node for those nodes that allow text.|
|[CustomXmlInsertNodeBeforeOptions](/javascript/api/word/word.customxmlinsertnodebeforeoptions)|[name](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-name-member)|If provided, specifies the base name of the element to be inserted.|
||[namespaceUri](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-namespaceuri-member)|If provided, specifies the namespace of the element to be inserted.|
||[nextSibling](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
||[nodeType](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nodetype-member)|If provided, specifies the type of node to append.|
||[nodeValue](/javascript/api/word/word.customxmlinsertnodebeforeoptions#word-word-customxmlinsertnodebeforeoptions-nodevalue-member)|If provided, specifies the value of the inserted node for those nodes that allow text.|
|[CustomXmlInsertSubtreeBeforeOptions](/javascript/api/word/word.customxmlinsertsubtreebeforeoptions)|[nextSibling](/javascript/api/word/word.customxmlinsertsubtreebeforeoptions#word-word-customxmlinsertsubtreebeforeoptions-nextsibling-member)|If provided, specifies the context node.|
|[CustomXmlNode](/javascript/api/word/word.customxmlnode)|[appendChildNode(options?: Word.CustomXmlAppendChildNodeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-appendchildnode-member(1))|Appends a single node as the last child under the context element node in the tree.|
||[appendChildSubtree(xml: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-appendchildsubtree-member(1))|Adds a subtree as the last child under the context element node in the tree.|
||[attributes](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-attributes-member)|Gets a `CustomXmlNodeCollection` object representing the attributes of the current element in the current node.|
||[baseName](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-basename-member)|Gets the base name of the node without the namespace prefix, if one exists.|
||[childNodes](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-childnodes-member)|Gets a `CustomXmlNodeCollection` object containing all of the child elements of the current node.|
||[delete()](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-delete-member(1))|Deletes the current node from the tree (including all of its children, if any exist).|
||[firstChild](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-firstchild-member)|Gets a `CustomXmlNode` object corresponding to the first child element of the current node.|
||[hasChildNodes()](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-haschildnodes-member(1))|Specifies if the current element node has child element nodes.|
||[insertNodeBefore(options?: Word.CustomXmlInsertNodeBeforeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-insertnodebefore-member(1))|Inserts a new node just before the context node in the tree.|
||[insertSubtreeBefore(xml: string, options?: Word.CustomXmlInsertSubtreeBeforeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-insertsubtreebefore-member(1))|Inserts the specified subtree into the location just before the context node.|
||[lastChild](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-lastchild-member)|Gets a `CustomXmlNode` object corresponding to the last child element of the current node.|
||[namespaceUri](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-namespaceuri-member)|Gets the unique address identifier for the namespace of the node.|
||[nextSibling](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nextsibling-member)|Gets the next sibling node (element, comment, or processing instruction) of the current node.|
||[nodeType](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nodetype-member)|Gets the type of the current node.|
||[nodeValue](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-nodevalue-member)|Specifies the value of the current node.|
||[ownerPart](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-ownerpart-member)|Gets the object representing the part associated with this node.|
||[parentNode](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-parentnode-member)|Gets the parent element node of the current node.|
||[previousSibling](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-previoussibling-member)|Gets the previous sibling node (element, comment, or processing instruction) of the current node.|
||[removeChild(child: Word.CustomXmlNode)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-removechild-member(1))|Removes the specified child node from the tree.|
||[replaceChildNode(oldNode: Word.CustomXmlNode, options?: Word.CustomXmlReplaceChildNodeOptions)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-replacechildnode-member(1))|Removes the specified child node and replaces it with a different node in the same location.|
||[replaceChildSubtree(xml: string, oldNode: Word.CustomXmlNode)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-replacechildsubtree-member(1))|Removes the specified node and replaces it with a different subtree in the same location.|
||[selectNodes(xPath: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-selectnodes-member(1))|Selects a collection of nodes matching an XPath expression.|
||[selectSingleNode(xPath: string)](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-selectsinglenode-member(1))|Selects a single node from a collection matching an XPath expression.|
||[text](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-text-member)|Specifies the text for the current node.|
||[xml](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-xml-member)|Gets the XML representation of the current node and its children.|
||[xpath](/javascript/api/word/word.customxmlnode#word-word-customxmlnode-xpath-member)|Gets a string with the canonicalized XPath for the current node.|
|[CustomXmlNodeCollection](/javascript/api/word/word.customxmlnodecollection)|[getCount()](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-getitem-member(1))|Returns a `CustomXmlNode` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlnodecollection#word-word-customxmlnodecollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[addNode(parent: Word.CustomXmlNode, options?: Word.CustomXmlAddNodeOptions)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-addnode-member(1))|Adds a node to the XML tree.|
||[builtIn](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-builtin-member)|Gets a value that indicates whether the `CustomXmlPart` is built-in.|
||[documentElement](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-documentelement-member)|Gets the root element of a bound region of data in the document.|
||[errors](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-errors-member)|Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.|
||[loadXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-loadxml-member(1))|Populates the `CustomXmlPart` object from an XML string.|
||[namespaceManager](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespacemanager-member)|Gets the set of namespace prefix mappings used against the current `CustomXmlPart` object.|
||[schemaCollection](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-schemacollection-member)|Specifies a `CustomXmlSchemaCollection` object representing the set of schemas attached to a bound region of data in the document.|
||[selectNodes(xPath: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-selectnodes-member(1))|Selects a collection of nodes from a custom XML part.|
||[selectSingleNode(xPath: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-selectsinglenode-member(1))|Selects a single node within a custom XML part matching an XPath expression.|
||[xml](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-xml-member)|Gets the XML representation of the current `CustomXmlPart` object.|
|[CustomXmlPrefixMapping](/javascript/api/word/word.customxmlprefixmapping)|[namespaceUri](/javascript/api/word/word.customxmlprefixmapping#word-word-customxmlprefixmapping-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlPrefixMapping` object.|
||[prefix](/javascript/api/word/word.customxmlprefixmapping#word-word-customxmlprefixmapping-prefix-member)|Gets the prefix for the `CustomXmlPrefixMapping` object.|
|[CustomXmlPrefixMappingCollection](/javascript/api/word/word.customxmlprefixmappingcollection)|[addNamespace(prefix: string, namespaceUri: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-addnamespace-member(1))|Adds a custom namespace/prefix mapping to use when querying an item.|
||[getCount()](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-getitem-member(1))|Returns a `CustomXmlPrefixMapping` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-items-member)|Gets the loaded child items in this collection.|
||[lookupNamespace(prefix: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-lookupnamespace-member(1))|Gets the namespace corresponding to the specified prefix.|
||[lookupPrefix(namespaceUri: string)](/javascript/api/word/word.customxmlprefixmappingcollection#word-word-customxmlprefixmappingcollection-lookupprefix-member(1))|Gets the prefix corresponding to the specified namespace.|
|[CustomXmlReplaceChildNodeOptions](/javascript/api/word/word.customxmlreplacechildnodeoptions)|[name](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-name-member)|If provided, specifies the base name of the replacement element.|
||[namespaceUri](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-namespaceuri-member)|If provided, specifies the namespace of the replacement element.|
||[nodeType](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-nodetype-member)|If provided, specifies the type of the replacement node.|
||[nodeValue](/javascript/api/word/word.customxmlreplacechildnodeoptions#word-word-customxmlreplacechildnodeoptions-nodevalue-member)|If provided, specifies the value of the replacement node for those nodes that allow text.|
|[CustomXmlSchema](/javascript/api/word/word.customxmlschema)|[delete()](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-delete-member(1))|Deletes this schema from the Word.CustomXmlSchemaCollection object.|
||[location](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-location-member)|Gets the location of the schema on a computer.|
||[namespaceUri](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-namespaceuri-member)|Gets the unique address identifier for the namespace of the `CustomXmlSchema` object.|
||[reload()](/javascript/api/word/word.customxmlschema#word-word-customxmlschema-reload-member(1))|Reloads the schema from a file.|
|[CustomXmlSchemaCollection](/javascript/api/word/word.customxmlschemacollection)|[add(options?: Word.CustomXmlAddSchemaOptions)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-add-member(1))|Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.|
||[addCollection(schemaCollection: Word.CustomXmlSchemaCollection)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-addcollection-member(1))|Adds an existing schema collection to the current schema collection.|
||[getCount()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getitem-member(1))|Returns a `CustomXmlSchema` object that represents the specified item in the collection.|
||[getNamespaceUri()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-getnamespaceuri-member(1))|Returns the number of items in the collection.|
||[items](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-items-member)|Gets the loaded child items in this collection.|
||[validate()](/javascript/api/word/word.customxmlschemacollection#word-word-customxmlschemacollection-validate-member(1))|Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).|
|[CustomXmlValidationError](/javascript/api/word/word.customxmlvalidationerror)|[delete()](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-delete-member(1))|Deletes this `CustomXmlValidationError` object.|
||[errorCode](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-errorcode-member)|Gets an integer representing the validation error in the `CustomXmlValidationError` object.|
||[name](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-name-member)|Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`|
||[node](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-node-member)|Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.|
||[text](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-text-member)|Gets the text in the `CustomXmlValidationError` object.|
||[type](/javascript/api/word/word.customxmlvalidationerror#word-word-customxmlvalidationerror-type-member)|Gets the type of error generated from the `CustomXmlValidationError` object.|
|[CustomXmlValidationErrorCollection](/javascript/api/word/word.customxmlvalidationerrorcollection)|[add(node: Word.CustomXmlNode, errorName: string, options?: Word.CustomXmlAddValidationErrorOptions)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-add-member(1))|Adds a `CustomXmlValidationError` object containing an XML validation error to the `CustomXmlValidationErrorCollection` object.|
||[getCount()](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getcount-member(1))|Returns the number of items in the collection.|
||[getItem(index: number)](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-getitem-member(1))|Returns a `CustomXmlValidationError` object that represents the specified item in the collection.|
||[items](/javascript/api/word/word.customxmlvalidationerrorcollection#word-word-customxmlvalidationerrorcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[hyperlinks](/javascript/api/word/word.document#word-word-document-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.|
|[FillFormat](/javascript/api/word/word.fillformat)|[backgroundColor](/javascript/api/word/word.fillformat#word-word-fillformat-backgroundcolor-member)|Returns a `ColorFormat` object that represents the background color for the fill.|
||[foregroundColor](/javascript/api/word/word.fillformat#word-word-fillformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill.|
||[gradientAngle](/javascript/api/word/word.fillformat#word-word-fillformat-gradientangle-member)|Specifies the angle of the gradient fill.|
||[gradientColorType](/javascript/api/word/word.fillformat#word-word-fillformat-gradientcolortype-member)|Gets the gradient color type.|
||[gradientDegree](/javascript/api/word/word.fillformat#word-word-fillformat-gradientdegree-member)|Returns how dark or light a one-color gradient fill is.|
||[gradientStyle](/javascript/api/word/word.fillformat#word-word-fillformat-gradientstyle-member)|Returns the gradient style for the fill.|
||[gradientVariant](/javascript/api/word/word.fillformat#word-word-fillformat-gradientvariant-member)|Returns the gradient variant for the fill as an integer value from 1 to 4 for most gradient fills.|
||[isVisible](/javascript/api/word/word.fillformat#word-word-fillformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/javascript/api/word/word.fillformat#word-word-fillformat-pattern-member)|Returns a `PatternType` value that represents the pattern applied to the fill or line.|
||[presetGradientType](/javascript/api/word/word.fillformat#word-word-fillformat-presetgradienttype-member)|Returns the preset gradient type for the fill.|
||[presetTexture](/javascript/api/word/word.fillformat#word-word-fillformat-presettexture-member)|Gets the preset texture.|
||[rotateWithObject](/javascript/api/word/word.fillformat#word-word-fillformat-rotatewithobject-member)|Specifies whether the fill rotates with the shape.|
||[setOneColorGradient(style: Word.GradientStyle, variant: number, degree: number)](/javascript/api/word/word.fillformat#word-word-fillformat-setonecolorgradient-member(1))|Sets the fill to a one-color gradient.|
||[setPatterned(pattern: Word.PatternType)](/javascript/api/word/word.fillformat#word-word-fillformat-setpatterned-member(1))|Sets the fill to a pattern.|
||[setPresetGradient(style: Word.GradientStyle, variant: number, presetGradientType: Word.PresetGradientType)](/javascript/api/word/word.fillformat#word-word-fillformat-setpresetgradient-member(1))|Sets the fill to a preset gradient.|
||[setPresetTextured(presetTexture: Word.PresetTexture)](/javascript/api/word/word.fillformat#word-word-fillformat-setpresettextured-member(1))|Sets the fill to a preset texture.|
||[setTwoColorGradient(style: Word.GradientStyle, variant: number)](/javascript/api/word/word.fillformat#word-word-fillformat-settwocolorgradient-member(1))|Sets the fill to a two-color gradient.|
||[solid()](/javascript/api/word/word.fillformat#word-word-fillformat-solid-member(1))|Sets the fill to a uniform color.|
||[textureAlignment](/javascript/api/word/word.fillformat#word-word-fillformat-texturealignment-member)|Specifies the alignment (the origin of the coordinate grid) for the tiling of the texture fill.|
||[textureHorizontalScale](/javascript/api/word/word.fillformat#word-word-fillformat-texturehorizontalscale-member)|Specifies the horizontal scaling factor for the texture fill.|
||[textureName](/javascript/api/word/word.fillformat#word-word-fillformat-texturename-member)|Returns the name of the custom texture file for the fill.|
||[textureOffsetX](/javascript/api/word/word.fillformat#word-word-fillformat-textureoffsetx-member)|Specifies the horizontal offset of the texture from the origin in points.|
||[textureOffsetY](/javascript/api/word/word.fillformat#word-word-fillformat-textureoffsety-member)|Specifies the vertical offset of the texture.|
||[textureTile](/javascript/api/word/word.fillformat#word-word-fillformat-texturetile-member)|Specifies whether the texture is tiled.|
||[textureType](/javascript/api/word/word.fillformat#word-word-fillformat-texturetype-member)|Returns the texture type for the fill.|
||[textureVerticalScale](/javascript/api/word/word.fillformat#word-word-fillformat-textureverticalscale-member)|Specifies the vertical scaling factor for the texture fill as a value between 0.0 and 1.0.|
||[transparency](/javascript/api/word/word.fillformat#word-word-fillformat-transparency-member)|Specifies the degree of transparency of the fill for a shape as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.fillformat#word-word-fillformat-type-member)|Gets the fill format type.|
|[Font](/javascript/api/word/word.font)|[allCaps](/javascript/api/word/word.font#word-word-font-allcaps-member)|Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters.|
||[boldBidirectional](/javascript/api/word/word.font#word-word-font-boldbidirectional-member)|Specifies whether the font is formatted as bold in a right-to-left language document.|
||[colorIndex](/javascript/api/word/word.font#word-word-font-colorindex-member)|Specifies a `ColorIndex` value that represents the color for the font.|
||[colorIndexBidirectional](/javascript/api/word/word.font#word-word-font-colorindexbidirectional-member)|Specifies the color for the `Font` object in a right-to-left language document.|
||[contextualAlternates](/javascript/api/word/word.font#word-word-font-contextualalternates-member)|Specifies whether contextual alternates are enabled for the font.|
||[decreaseFontSize()](/javascript/api/word/word.font#word-word-font-decreasefontsize-member(1))|Decreases the font size to the next available size.|
||[diacriticColor](/javascript/api/word/word.font#word-word-font-diacriticcolor-member)|Specifies the color to be used for diacritics for the `Font` object.|
||[disableCharacterSpaceGrid](/javascript/api/word/word.font#word-word-font-disablecharacterspacegrid-member)|Specifies whether Microsoft Word ignores the number of characters per line for the corresponding `Font` object.|
||[emboss](/javascript/api/word/word.font#word-word-font-emboss-member)|Specifies whether the font is formatted as embossed.|
||[emphasisMark](/javascript/api/word/word.font#word-word-font-emphasismark-member)|Specifies an `EmphasisMark` value that represents the emphasis mark for a character or designated character string.|
||[engrave](/javascript/api/word/word.font#word-word-font-engrave-member)|Specifies whether the font is formatted as engraved.|
||[fill](/javascript/api/word/word.font#word-word-font-fill-member)|Returns a `FillFormat` object that contains fill formatting properties for the font used by the range of text.|
||[glow](/javascript/api/word/word.font#word-word-font-glow-member)|Returns a `GlowFormat` object that represents the glow formatting for the font used by the range of text.|
||[increaseFontSize()](/javascript/api/word/word.font#word-word-font-increasefontsize-member(1))|Increases the font size to the next available size.|
||[italicBidirectional](/javascript/api/word/word.font#word-word-font-italicbidirectional-member)|Specifies whether the font is italicized in a right-to-left language document.|
||[kerning](/javascript/api/word/word.font#word-word-font-kerning-member)|Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.|
||[ligature](/javascript/api/word/word.font#word-word-font-ligature-member)|Specifies the ligature setting for the `Font` object.|
||[line](/javascript/api/word/word.font#word-word-font-line-member)|Returns a `LineFormat` object that specifies the formatting for a line.|
||[nameAscii](/javascript/api/word/word.font#word-word-font-nameascii-member)|Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).|
||[nameBidirectional](/javascript/api/word/word.font#word-word-font-namebidirectional-member)|Specifies the font name in a right-to-left language document.|
||[nameFarEast](/javascript/api/word/word.font#word-word-font-namefareast-member)|Specifies the East Asian font name.|
||[nameOther](/javascript/api/word/word.font#word-word-font-nameother-member)|Specifies the font used for characters with codes from 128 through 255.|
||[numberForm](/javascript/api/word/word.font#word-word-font-numberform-member)|Specifies the number form setting for an OpenType font.|
||[numberSpacing](/javascript/api/word/word.font#word-word-font-numberspacing-member)|Specifies the number spacing setting for the font.|
||[outline](/javascript/api/word/word.font#word-word-font-outline-member)|Specifies if the font is formatted as outlined.|
||[position](/javascript/api/word/word.font#word-word-font-position-member)|Specifies the position of text (in points) relative to the base line.|
||[reflection](/javascript/api/word/word.font#word-word-font-reflection-member)|Returns a `ReflectionFormat` object that represents the reflection formatting for a shape.|
||[reset()](/javascript/api/word/word.font#word-word-font-reset-member(1))|Removes manual character formatting.|
||[scaling](/javascript/api/word/word.font#word-word-font-scaling-member)|Specifies the scaling percentage applied to the font.|
||[setAsTemplateDefault()](/javascript/api/word/word.font#word-word-font-setastemplatedefault-member(1))|Sets the specified font formatting as the default for the active document and all new documents based on the active template.|
||[shadow](/javascript/api/word/word.font#word-word-font-shadow-member)|Specifies if the font is formatted as shadowed.|
||[sizeBidirectional](/javascript/api/word/word.font#word-word-font-sizebidirectional-member)|Specifies the font size in points for right-to-left text.|
||[smallCaps](/javascript/api/word/word.font#word-word-font-smallcaps-member)|Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters.|
||[spacing](/javascript/api/word/word.font#word-word-font-spacing-member)|Specifies the spacing between characters.|
||[stylisticSet](/javascript/api/word/word.font#word-word-font-stylisticset-member)|Specifies the stylistic set for the font.|
||[textColor](/javascript/api/word/word.font#word-word-font-textcolor-member)|Returns a `ColorFormat` object that represents the color for the font.|
||[textShadow](/javascript/api/word/word.font#word-word-font-textshadow-member)|Returns a `ShadowFormat` object that specifies the shadow formatting for the font.|
||[threeDimensionalFormat](/javascript/api/word/word.font#word-word-font-threedimensionalformat-member)|Returns a `ThreeDimensionalFormat` object that contains 3-dimensional (3D) effect formatting properties for the font.|
||[underlineColor](/javascript/api/word/word.font#word-word-font-underlinecolor-member)|Specifies the color of the underline for the `Font` object.|
|[GlowFormat](/javascript/api/word/word.glowformat)|[color](/javascript/api/word/word.glowformat#word-word-glowformat-color-member)|Returns a `ColorFormat` object that represents the color for a glow effect.|
||[radius](/javascript/api/word/word.glowformat#word-word-glowformat-radius-member)|Specifies the length of the radius for a glow effect.|
||[transparency](/javascript/api/word/word.glowformat#word-word-glowformat-transparency-member)|Specifies the degree of transparency for the glow effect as a value between 0.0 (opaque) and 1.0 (clear).|
|[Hyperlink](/javascript/api/word/word.hyperlink)|[addToFavorites()](/javascript/api/word/word.hyperlink#word-word-hyperlink-addtofavorites-member(1))|Creates a shortcut to the document or hyperlink and adds it to the **Favorites** folder.|
||[address](/javascript/api/word/word.hyperlink#word-word-hyperlink-address-member)|Specifies the address (for example, a file name or URL) of the hyperlink.|
||[createNewDocument(fileName: string, editNow: boolean, overwrite: boolean)](/javascript/api/word/word.hyperlink#word-word-hyperlink-createnewdocument-member(1))|Creates a new document linked to the hyperlink.|
||[delete()](/javascript/api/word/word.hyperlink#word-word-hyperlink-delete-member(1))|Deletes the hyperlink.|
||[emailSubject](/javascript/api/word/word.hyperlink#word-word-hyperlink-emailsubject-member)|Specifies the text string for the hyperlink's subject line.|
||[isExtraInfoRequired](/javascript/api/word/word.hyperlink#word-word-hyperlink-isextrainforequired-member)|Returns `true` if extra information is required to resolve the hyperlink.|
||[name](/javascript/api/word/word.hyperlink#word-word-hyperlink-name-member)|Returns the name of the `Hyperlink` object.|
||[range](/javascript/api/word/word.hyperlink#word-word-hyperlink-range-member)|Returns a `Range` object that represents the portion of the document that's contained within the hyperlink.|
||[screenTip](/javascript/api/word/word.hyperlink#word-word-hyperlink-screentip-member)|Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.|
||[subAddress](/javascript/api/word/word.hyperlink#word-word-hyperlink-subaddress-member)|Specifies a named location in the destination of the hyperlink.|
||[target](/javascript/api/word/word.hyperlink#word-word-hyperlink-target-member)|Specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/javascript/api/word/word.hyperlink#word-word-hyperlink-texttodisplay-member)|Specifies the hyperlink's visible text in the document.|
||[type](/javascript/api/word/word.hyperlink#word-word-hyperlink-type-member)|Returns the hyperlink type.|
|[HyperlinkAddOptions](/javascript/api/word/word.hyperlinkaddoptions)|[address](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-address-member)|If provided, specifies the address (e.g., URL or file path) of the hyperlink.|
||[screenTip](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-screentip-member)|If provided, specifies the text that appears as a tooltip.|
||[subAddress](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-subaddress-member)|If provided, specifies the location within the file or document.|
||[target](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-target-member)|If provided, specifies the name of the frame or window in which to load the hyperlink.|
||[textToDisplay](/javascript/api/word/word.hyperlinkaddoptions#word-word-hyperlinkaddoptions-texttodisplay-member)|If provided, specifies the visible text of the hyperlink.|
|[HyperlinkCollection](/javascript/api/word/word.hyperlinkcollection)|[add(anchor: Word.Range, options?: Word.HyperlinkAddOptions)](/javascript/api/word/word.hyperlinkcollection#word-word-hyperlinkcollection-add-member(1))|Returns a `Hyperlink` object that represents a new hyperlink added to a range, selection, or document.|
||[items](/javascript/api/word/word.hyperlinkcollection#word-word-hyperlinkcollection-items-member)|Gets the loaded child items in this collection.|
|[LineFormat](/javascript/api/word/word.lineformat)|[backgroundColor](/javascript/api/word/word.lineformat#word-word-lineformat-backgroundcolor-member)|Gets a `ColorFormat` object that represents the background color for a patterned line.|
||[beginArrowheadLength](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadlength-member)|Specifies the length of the arrowhead at the beginning of the line.|
||[beginArrowheadStyle](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadstyle-member)|Specifies the style of the arrowhead at the beginning of the line.|
||[beginArrowheadWidth](/javascript/api/word/word.lineformat#word-word-lineformat-beginarrowheadwidth-member)|Specifies the width of the arrowhead at the beginning of the line.|
||[dashStyle](/javascript/api/word/word.lineformat#word-word-lineformat-dashstyle-member)|Specifies the dash style for the line.|
||[endArrowheadLength](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadlength-member)|Specifies the length of the arrowhead at the end of the line.|
||[endArrowheadStyle](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadstyle-member)|Specifies the style of the arrowhead at the end of the line.|
||[endArrowheadWidth](/javascript/api/word/word.lineformat#word-word-lineformat-endarrowheadwidth-member)|Specifies the width of the arrowhead at the end of the line.|
||[foregroundColor](/javascript/api/word/word.lineformat#word-word-lineformat-foregroundcolor-member)|Gets a `ColorFormat` object that represents the foreground color for the line.|
||[insetPen](/javascript/api/word/word.lineformat#word-word-lineformat-insetpen-member)|Specifies if to draw lines inside a shape.|
||[isVisible](/javascript/api/word/word.lineformat#word-word-lineformat-isvisible-member)|Specifies if the object, or the formatting applied to it, is visible.|
||[pattern](/javascript/api/word/word.lineformat#word-word-lineformat-pattern-member)|Specifies the pattern applied to the line.|
||[style](/javascript/api/word/word.lineformat#word-word-lineformat-style-member)|Specifies the line format style.|
||[transparency](/javascript/api/word/word.lineformat#word-word-lineformat-transparency-member)|Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).|
||[weight](/javascript/api/word/word.lineformat#word-word-lineformat-weight-member)|Specifies the thickness of the line in points.|
|[ListFormat](/javascript/api/word/word.listformat)|[applyBulletDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applybulletdefault-member(1))|Adds bullets and formatting to the paragraphs in the range.|
||[applyListTemplateWithLevel(listTemplate: Word.ListTemplate, options?: Word.ListTemplateApplyOptions)](/javascript/api/word/word.listformat#word-word-listformat-applylisttemplatewithlevel-member(1))|Applies a list template with a specific level to the paragraphs in the range.|
||[applyNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applynumberdefault-member(1))|Adds numbering and formatting to the paragraphs in the range.|
||[applyOutlineNumberDefault(defaultListBehavior: Word.DefaultListBehavior)](/javascript/api/word/word.listformat#word-word-listformat-applyoutlinenumberdefault-member(1))|Adds outline numbering and formatting to the paragraphs in the range.|
||[canContinuePreviousList(listTemplate: Word.ListTemplate)](/javascript/api/word/word.listformat#word-word-listformat-cancontinuepreviouslist-member(1))|Determines whether the `ListFormat` object can continue a previous list.|
||[convertNumbersToText(numberType: Word.NumberType)](/javascript/api/word/word.listformat#word-word-listformat-convertnumberstotext-member(1))|Converts numbers in the list to plain text.|
||[countNumberedItems(options?: Word.ListFormatCountNumberedItemsOptions)](/javascript/api/word/word.listformat#word-word-listformat-countnumbereditems-member(1))|Counts the numbered items in the list.|
||[isSingleList](/javascript/api/word/word.listformat#word-word-listformat-issinglelist-member)|Indicates whether the `ListFormat` object contains a single list.|
||[isSingleListTemplate](/javascript/api/word/word.listformat#word-word-listformat-issinglelisttemplate-member)|Indicates whether the `ListFormat` object contains a single list template.|
||[list](/javascript/api/word/word.listformat#word-word-listformat-list-member)|Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.|
||[listIndent()](/javascript/api/word/word.listformat#word-word-listformat-listindent-member(1))|Indents the list by one level.|
||[listLevelNumber](/javascript/api/word/word.listformat#word-word-listformat-listlevelnumber-member)|Specifies the list level number for the first paragraph for the `ListFormat` object.|
||[listOutdent()](/javascript/api/word/word.listformat#word-word-listformat-listoutdent-member(1))|Outdents the list by one level.|
||[listString](/javascript/api/word/word.listformat#word-word-listformat-liststring-member)|Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.|
||[listTemplate](/javascript/api/word/word.listformat#word-word-listformat-listtemplate-member)|Gets the list template associated with the `ListFormat` object.|
||[listType](/javascript/api/word/word.listformat#word-word-listformat-listtype-member)|Gets the type of the list for the `ListFormat` object.|
||[listValue](/javascript/api/word/word.listformat#word-word-listformat-listvalue-member)|Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.|
||[removeNumbers(numberType: Word.NumberType)](/javascript/api/word/word.listformat#word-word-listformat-removenumbers-member(1))|Removes numbering from the list.|
|[ListFormatCountNumberedItemsOptions](/javascript/api/word/word.listformatcountnumbereditemsoptions)|[level](/javascript/api/word/word.listformatcountnumbereditemsoptions#word-word-listformatcountnumbereditemsoptions-level-member)|If provided, specifies the level to count.|
||[numberType](/javascript/api/word/word.listformatcountnumbereditemsoptions#word-word-listformatcountnumbereditemsoptions-numbertype-member)|If provided, specifies the type of number to count.|
|[ListTemplateApplyOptions](/javascript/api/word/word.listtemplateapplyoptions)|[applyLevel](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-applylevel-member)|If provided, specifies the level to apply in the list template.|
||[applyTo](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-applyto-member)|If provided, specifies which part of the list to apply the template to.|
||[continuePreviousList](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-continuepreviouslist-member)|If provided, specifies whether to continue the previous list.|
||[defaultListBehavior](/javascript/api/word/word.listtemplateapplyoptions#word-word-listtemplateapplyoptions-defaultlistbehavior-member)|If provided, specifies the default list behavior.|
|[Page](/javascript/api/word/word.page)|[breaks](/javascript/api/word/word.page#word-word-page-breaks-member)|Gets a `BreakCollection` object that represents the breaks on the page.|
|[Paragraph](/javascript/api/word/word.paragraph)|[onCommentAdded](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeleted](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeleted-member)|Occurs when comments are deleted.|
||[onCommentDeselected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.paragraph#word-word-paragraph-oncommentselected-member)|Occurs when a comment is selected.|
|[ParagraphAddedEventArgs](/javascript/api/word/word.paragraphaddedeventargs)|[type](/javascript/api/word/word.paragraphaddedeventargs#word-word-paragraphaddedeventargs-type-member)|The event type.|
|[ParagraphChangedEventArgs](/javascript/api/word/word.paragraphchangedeventargs)|[type](/javascript/api/word/word.paragraphchangedeventargs#word-word-paragraphchangedeventargs-type-member)|The event type.|
|[ParagraphDeletedEventArgs](/javascript/api/word/word.paragraphdeletedeventargs)|[type](/javascript/api/word/word.paragraphdeletedeventargs#word-word-paragraphdeletedeventargs-type-member)|The event type.|
|[Range](/javascript/api/word/word.range)|[hasNoProofing](/javascript/api/word/word.range#word-word-range-hasnoproofing-member)|Specifies the proofing status (spelling and grammar checking) of the range.|
||[hyperlinks](/javascript/api/word/word.range#word-word-range-hyperlinks-member)|Returns a `HyperlinkCollection` object that represents all the hyperlinks in the range.|
||[listFormat](/javascript/api/word/word.range#word-word-range-listformat-member)|Returns a `ListFormat` object that represents all the list formatting characteristics of the range.|
||[onCommentAdded](/javascript/api/word/word.range#word-word-range-oncommentadded-member)|Occurs when new comments are added.|
||[onCommentChanged](/javascript/api/word/word.range#word-word-range-oncommentchanged-member)|Occurs when a comment or its reply is changed.|
||[onCommentDeselected](/javascript/api/word/word.range#word-word-range-oncommentdeselected-member)|Occurs when a comment is deselected.|
||[onCommentSelected](/javascript/api/word/word.range#word-word-range-oncommentselected-member)|Occurs when a comment is selected.|
|[ReflectionFormat](/javascript/api/word/word.reflectionformat)|[blur](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-blur-member)|Specifies the degree of blur effect applied to the `ReflectionFormat` object as a value between 0.0 and 100.0.|
||[offset](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-offset-member)|Specifies the amount of separation, in points, of the reflected image from the shape.|
||[size](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-size-member)|Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.|
||[transparency](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-transparency-member)|Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.reflectionformat#word-word-reflectionformat-type-member)|Specifies a `ReflectionType` value that represents the type and direction of the lighting for a shape reflection.|
|[ShadowFormat](/javascript/api/word/word.shadowformat)|[blur](/javascript/api/word/word.shadowformat#word-word-shadowformat-blur-member)|Specifies the blur level for a shadow format as a value between 0.0 and 100.0.|
||[foregroundColor](/javascript/api/word/word.shadowformat#word-word-shadowformat-foregroundcolor-member)|Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.|
||[incrementOffsetX(increment: number)](/javascript/api/word/word.shadowformat#word-word-shadowformat-incrementoffsetx-member(1))|Changes the horizontal offset of the shadow by the number of points.|
||[incrementOffsetY(increment: number)](/javascript/api/word/word.shadowformat#word-word-shadowformat-incrementoffsety-member(1))|Changes the vertical offset of the shadow by the specified number of points.|
||[isVisible](/javascript/api/word/word.shadowformat#word-word-shadowformat-isvisible-member)|Specifies whether the object or the formatting applied to it is visible.|
||[obscured](/javascript/api/word/word.shadowformat#word-word-shadowformat-obscured-member)|Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill,|
||[offsetX](/javascript/api/word/word.shadowformat#word-word-shadowformat-offsetx-member)|Specifies the horizontal offset (in points) of the shadow from the shape.|
||[offsetY](/javascript/api/word/word.shadowformat#word-word-shadowformat-offsety-member)|Specifies the vertical offset (in points) of the shadow from the shape.|
||[rotateWithShape](/javascript/api/word/word.shadowformat#word-word-shadowformat-rotatewithshape-member)|Specifies whether to rotate the shadow when rotating the shape.|
||[size](/javascript/api/word/word.shadowformat#word-word-shadowformat-size-member)|Specifies the width of the shadow.|
||[style](/javascript/api/word/word.shadowformat#word-word-shadowformat-style-member)|Specifies the type of shadow formatting to apply to a shape.|
||[transparency](/javascript/api/word/word.shadowformat#word-word-shadowformat-transparency-member)|Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).|
||[type](/javascript/api/word/word.shadowformat#word-word-shadowformat-type-member)|Specifies the shape shadow type.|
|[Style](/javascript/api/word/word.style)|[description](/javascript/api/word/word.style#word-word-style-description-member)|Gets the description of the specified style.|
|[ThreeDimensionalFormat](/javascript/api/word/word.threedimensionalformat)|[bevelBottomDepth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottomdepth-member)|Specifies the depth of the bottom bevel.|
||[bevelBottomInset](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottominset-member)|Specifies the inset size for the bottom bevel.|
||[bevelBottomType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-bevelbottomtype-member)|Specifies a `BevelType` value that represents the bevel type for the bottom bevel.|
||[bevelTopDepth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltopdepth-member)|Specifies the depth of the top bevel.|
||[bevelTopInset](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltopinset-member)|Specifies the inset size for the top bevel.|
||[bevelTopType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-beveltoptype-member)|Specifies a `BevelType` value that represents the bevel type for the top bevel.|
||[contourColor](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-contourcolor-member)|Returns a `ColorFormat` object that represents color of the contour of a shape.|
||[contourWidth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-contourwidth-member)|Specifies the width of the contour of a shape.|
||[depth](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-depth-member)|Specifies the depth of the shape's extrusion.|
||[extrusionColor](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-extrusioncolor-member)|Returns a `ColorFormat` object that represents the color of the shape's extrusion.|
||[extrusionColorType](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-extrusioncolortype-member)|Specifies whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion)|
||[fieldOfView](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-fieldofview-member)|Specifies the amount of perspective for a shape.|
||[incrementRotationHorizontal(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationhorizontal-member(1))|Horizontally rotates a shape on the x-axis.|
||[incrementRotationVertical(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationvertical-member(1))|Vertically rotates a shape on the y-axis.|
||[incrementRotationX(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationx-member(1))|Changes the rotation around the x-axis.|
||[incrementRotationY(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationy-member(1))|Changes the rotation around the y-axis.|
||[incrementRotationZ(increment: number)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-incrementrotationz-member(1))|Rotates a shape on the z-axis.|
||[isPerspective](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-isperspective-member)|Specifies `true` if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point,|
||[isVisible](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-isvisible-member)|Specifies if the specified object, or the formatting applied to it, is visible.|
||[lightAngle](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-lightangle-member)|Specifies the angle of the lighting.|
||[presetCamera](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetcamera-member)|Returns a `PresetCamera` value that represents the camera presets.|
||[presetExtrusionDirection](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetextrusiondirection-member)|Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion).|
||[presetLighting](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlighting-member)|Specifies a `LightRigType` value that represents the lighting preset.|
||[presetLightingDirection](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlightingdirection-member)|Specifies the position of the light source relative to the extrusion.|
||[presetLightingSoftness](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetlightingsoftness-member)|Specifies the intensity of the extrusion lighting.|
||[presetMaterial](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetmaterial-member)|Specifies the extrusion surface material.|
||[presetThreeDimensionalFormat](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-presetthreedimensionalformat-member)|Returns the preset extrusion format.|
||[projectText](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-projecttext-member)|Specifies whether text on a shape rotates with shape.|
||[resetRotation()](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-resetrotation-member(1))|Resets the extrusion rotation around the x-axis, y-axis, and z-axis to 0.|
||[rotationX](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationx-member)|Specifies the rotation of the extruded shape around the x-axis in degrees.|
||[rotationY](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationy-member)|Specifies the rotation of the extruded shape around the y-axis in degrees.|
||[rotationZ](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-rotationz-member)|Specifies the z-axis rotation of the camera.|
||[setExtrusionDirection(presetExtrusionDirection: Word.PresetExtrusionDirection)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setextrusiondirection-member(1))|Sets the direction of the extrusion's sweep path.|
||[setPresetCamera(presetCamera: Word.PresetCamera)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setpresetcamera-member(1))|Sets the camera preset for the shape.|
||[setThreeDimensionalFormat(presetThreeDimensionalFormat: Word.PresetThreeDimensionalFormat)](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-setthreedimensionalformat-member(1))|Sets the preset extrusion format.|
||[z](/javascript/api/word/word.threedimensionalformat#word-word-threedimensionalformat-z-member)|Specifies the position on the z-axis for the shape.|
|[XmlMapping](/javascript/api/word/word.xmlmapping)|[customXmlNode](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-customxmlnode-member)|Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.|
||[customXmlPart](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-customxmlpart-member)|Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.|
||[delete()](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-delete-member(1))|Deletes the XML mapping from the parent content control.|
||[isMapped](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-ismapped-member)|Returns whether the content control in the document is mapped to an XML node in the document's XML data store.|
||[prefixMappings](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-prefixmappings-member)|Returns the prefix mappings used to evaluate the XPath for the current XML mapping.|
||[setMapping(xPath: string, options?: Word.XmlSetMappingOptions)](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-setmapping-member(1))|Allows creating or changing the XML mapping on the content control.|
||[setMappingByNode(node: Word.CustomXmlNode)](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-setmappingbynode-member(1))|Allows creating or changing the XML data mapping on the content control.|
||[xpath](/javascript/api/word/word.xmlmapping#word-word-xmlmapping-xpath-member)|Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.|
|[XmlSetMappingOptions](/javascript/api/word/word.xmlsetmappingoptions)|[prefixMapping](/javascript/api/word/word.xmlsetmappingoptions#word-word-xmlsetmappingoptions-prefixmapping-member)|If provided, specifies the prefix mappings to use when querying the expression provided in the `xPath` parameter of the `XmlMapping.setMapping` calling method.|
||[source](/javascript/api/word/word.xmlsetmappingoptions#word-word-xmlsetmappingoptions-source-member)|If provided, specifies the desired custom XML data to map the content control to.|
