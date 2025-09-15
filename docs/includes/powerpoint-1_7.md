| Class | Fields | Description |
|:---|:---|:---|
|[CustomProperty](/.customproperty)|[delete()](/.customproperty#powerpoint-javascript/api/powerpoint/-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/.customproperty#powerpoint-javascript/api/powerpoint/-customproperty-key-member)|The string that uniquely identifies the custom property.|
||[type](/.customproperty#powerpoint-javascript/api/powerpoint/-customproperty-type-member)|The type of the value used for the custom property.|
||[value](/.customproperty#powerpoint-javascript/api/powerpoint/-customproperty-value-member)|The value of the custom property.|
|[CustomPropertyCollection](/.custompropertycollection)|[add(key: string, value: boolean \| Date \| number \| string)](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-add-member(1))|Creates a new `CustomProperty` or updates the property with the given key.|
||[deleteAll()](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-getcount-member(1))|Gets the number of custom properties in the collection.|
||[getItem(key: string)](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-getitem-member(1))|Gets a `CustomProperty` by its key.|
||[getItemOrNullObject(key: string)](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-getitemornullobject-member(1))|Gets a `CustomProperty` by its key.|
||[items](/.custompropertycollection#powerpoint-javascript/api/powerpoint/-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPart](/.customxmlpart)|[delete()](/.customxmlpart#powerpoint-javascript/api/powerpoint/-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[getXml()](/.customxmlpart#powerpoint-javascript/api/powerpoint/-customxmlpart-getxml-member(1))|Gets the XML content of the custom XML part.|
||[id](/.customxmlpart#powerpoint-javascript/api/powerpoint/-customxmlpart-id-member)|The ID of the custom XML part.|
||[namespaceUri](/.customxmlpart#powerpoint-javascript/api/powerpoint/-customxmlpart-namespaceuri-member)|The namespace URI of the custom XML part.|
||[setXml(xml: string)](/.customxmlpart#powerpoint-javascript/api/powerpoint/-customxmlpart-setxml-member(1))|Sets the XML content for the custom XML part.|
|[CustomXmlPartCollection](/.customxmlpartcollection)|[add(xml: string)](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-add-member(1))|Adds a new `CustomXmlPart` to the collection.|
||[getByNamespace(namespaceUri: string)](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-getcount-member(1))|Gets the number of custom XML parts in the collection.|
||[getItem(id: string)](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-getitem-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-getitemornullobject-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[items](/.customxmlpartcollection#powerpoint-javascript/api/powerpoint/-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/.customxmlpartscopedcollection)|[getCount()](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-getcount-member(1))|Gets the number of custom XML parts in this collection.|
||[getItem(id: string)](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-getitem-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getItemOrNullObject(id: string)](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getOnlyItem()](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/.customxmlpartscopedcollection#powerpoint-javascript/api/powerpoint/-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentProperties](/.documentproperties)|[author](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-author-member)|The author of the presentation.|
||[category](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-category-member)|The category of the presentation.|
||[comments](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-comments-member)|The Comments field in the metadata of the presentation.|
||[company](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-company-member)|The company of the presentation.|
||[creationDate](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-creationdate-member)|The creation date of the presentation.|
||[customProperties](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-customproperties-member)|The collection of custom properties of the presentation.|
||[keywords](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-keywords-member)|The keywords of the presentation.|
||[lastAuthor](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-lastauthor-member)|The last author of the presentation.|
||[manager](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-manager-member)|The manager of the presentation.|
||[revisionNumber](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-revisionnumber-member)|The revision number of the presentation.|
||[subject](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-subject-member)|The subject of the presentation.|
||[title](/.documentproperties#powerpoint-javascript/api/powerpoint/-documentproperties-title-member)|The title of the presentation.|
|[Presentation](/.presentation)|[customXmlParts](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-customxmlparts-member)|Returns a collection of custom XML parts that are associated with the presentation.|
||[properties](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-properties-member)|Gets the properties of the presentation.|
|[Shape](/.shape)|[customXmlParts](/.shape#powerpoint-javascript/api/powerpoint/-shape-customxmlparts-member)|Returns a collection of custom XML parts in the shape.|
|[Slide](/.slide)|[customXmlParts](/.slide#powerpoint-javascript/api/powerpoint/-slide-customxmlparts-member)|Returns a collection of custom XML parts in the slide.|
|[SlideLayout](/.slidelayout)|[customXmlParts](/.slidelayout#powerpoint-javascript/api/powerpoint/-slidelayout-customxmlparts-member)|Returns a collection of custom XML parts in the slide layout.|
|[SlideMaster](/.slidemaster)|[customXmlParts](/.slidemaster#powerpoint-javascript/api/powerpoint/-slidemaster-customxmlparts-member)|Returns a collection of custom XML parts in the Slide Master.|
