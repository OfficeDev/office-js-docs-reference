| Class | Fields | Description |
|:---|:---|:---|
|[CustomProperty](/javascript/api/powerpoint/powerpoint.customproperty)|[delete()](/javascript/api/powerpoint/powerpoint.customproperty#powerpoint-powerpoint-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/powerpoint/powerpoint.customproperty#powerpoint-powerpoint-customproperty-key-member)|The string that uniquely identifies the custom property.|
||[type](/javascript/api/powerpoint/powerpoint.customproperty#powerpoint-powerpoint-customproperty-type-member)|The type of the value used for the custom property.|
||[value](/javascript/api/powerpoint/powerpoint.customproperty#powerpoint-powerpoint-customproperty-value-member)|The value of the custom property.|
|[CustomPropertyCollection](/javascript/api/powerpoint/powerpoint.custompropertycollection)|[add(key: string, value: boolean \| Date \| number \| string)](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-add-member(1))|Creates a new `CustomProperty` or updates the property with the given key.|
||[deleteAll()](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-getcount-member(1))|Gets the number of custom properties in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-getitem-member(1))|Gets a `CustomProperty` by its key.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-getitemornullobject-member(1))|Gets a `CustomProperty` by its key.|
||[items](/javascript/api/powerpoint/powerpoint.custompropertycollection#powerpoint-powerpoint-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPart](/javascript/api/powerpoint/powerpoint.customxmlpart)|[delete()](/javascript/api/powerpoint/powerpoint.customxmlpart#powerpoint-powerpoint-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[getXml()](/javascript/api/powerpoint/powerpoint.customxmlpart#powerpoint-powerpoint-customxmlpart-getxml-member(1))|Gets the XML content of the custom XML part.|
||[id](/javascript/api/powerpoint/powerpoint.customxmlpart#powerpoint-powerpoint-customxmlpart-id-member)|The ID of the custom XML part.|
||[namespaceUri](/javascript/api/powerpoint/powerpoint.customxmlpart#powerpoint-powerpoint-customxmlpart-namespaceuri-member)|The namespace URI of the custom XML part.|
||[setXml(xml: string)](/javascript/api/powerpoint/powerpoint.customxmlpart#powerpoint-powerpoint-customxmlpart-setxml-member(1))|Sets the XML content for the custom XML part.|
|[CustomXmlPartCollection](/javascript/api/powerpoint/powerpoint.customxmlpartcollection)|[add(xml: string)](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-add-member(1))|Adds a new `CustomXmlPart` to the collection.|
||[getByNamespace(namespaceUri: string)](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-getcount-member(1))|Gets the number of custom XML parts in the collection.|
||[getItem(id: string)](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-getitem-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-getitemornullobject-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[items](/javascript/api/powerpoint/powerpoint.customxmlpartcollection#powerpoint-powerpoint-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-getcount-member(1))|Gets the number of custom XML parts in this collection.|
||[getItem(id: string)](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-getitem-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a `CustomXmlPart` based on its ID.|
||[getOnlyItem()](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/powerpoint/powerpoint.customxmlpartscopedcollection#powerpoint-powerpoint-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentProperties](/javascript/api/powerpoint/powerpoint.documentproperties)|[author](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-author-member)|The author of the presentation.|
||[category](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-category-member)|The category of the presentation.|
||[comments](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-comments-member)|The Comments field in the metadata of the presentation.|
||[company](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-company-member)|The company of the presentation.|
||[creationDate](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-creationdate-member)|The creation date of the presentation.|
||[customProperties](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-customproperties-member)|The collection of custom properties of the presentation.|
||[keywords](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-keywords-member)|The keywords of the presentation.|
||[lastAuthor](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-lastauthor-member)|The last author of the presentation.|
||[manager](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-manager-member)|The manager of the presentation.|
||[revisionNumber](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-revisionnumber-member)|The revision number of the presentation.|
||[subject](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-subject-member)|The subject of the presentation.|
||[title](/javascript/api/powerpoint/powerpoint.documentproperties#powerpoint-powerpoint-documentproperties-title-member)|The title of the presentation.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[customXmlParts](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-customxmlparts-member)|Returns a collection of custom XML parts that are associated with the presentation.|
||[properties](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-properties-member)|Gets the properties of the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[customXmlParts](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-customxmlparts-member)|Returns a collection of custom XML parts in the shape.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[customXmlParts](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-customxmlparts-member)|Returns a collection of custom XML parts in the slide.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[customXmlParts](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-customxmlparts-member)|Returns a collection of custom XML parts in the slide layout.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[customXmlParts](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-customxmlparts-member)|Returns a collection of custom XML parts in the Slide Master.|
