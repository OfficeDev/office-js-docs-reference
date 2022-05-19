| Class | Fields | Description |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|Specifies the ID of a Slide Layout to be used for the new slide.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|Specifies the ID of a Slide Master to be used for the new slide.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|Returns the collection of `SlideMaster` objects that are in the presentation.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|Returns a collection of tags attached to the presentation.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|Deletes the shape from the shape collection.|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|Gets the unique ID of the shape.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|Returns a collection of tags in the shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|Gets the number of shapes in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|Gets a shape using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[layout](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|Gets the layout of the slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|Returns a collection of shapes in the slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|Gets the `SlideMaster` object that represents the slide's default content.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|Returns a collection of tags in the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|Adds a new slide at the end of the collection.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|Gets the unique ID of the slide layout.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|Gets the name of the slide layout.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|Returns a collection of shapes in the slide layout.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|Gets the number of layouts in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|Gets a layout using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|Gets a layout using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|Gets a layout using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|Gets the loaded child items in this collection.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|Gets the unique ID of the Slide Master.|
||[layouts](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|Gets the collection of layouts provided by the Slide Master for slides.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|Gets the unique name of the Slide Master.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|Returns a collection of shapes in the Slide Master.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|Gets the number of Slide Masters in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|Gets a Slide Master using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|Gets a Slide Master using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|Gets a Slide Master using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|Gets the loaded child items in this collection.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|Gets the unique ID of the tag.|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|Gets the value of the tag.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|Adds a new tag at the end of the collection.|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|Deletes the tag with the given `key` in this collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|Gets the number of tags in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|Gets a tag using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|Gets a tag using its zero-based index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|Gets a tag using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|Gets the loaded child items in this collection.|
