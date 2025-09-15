| Class | Fields | Description |
|:---|:---|:---|
|[AddSlideOptions](/.addslideoptions)|[layoutId](/.addslideoptions#powerpoint-javascript/api/powerpoint/-addslideoptions-layoutid-member)|Specifies the ID of a Slide Layout to be used for the new slide.|
||[slideMasterId](/.addslideoptions#powerpoint-javascript/api/powerpoint/-addslideoptions-slidemasterid-member)|Specifies the ID of a Slide Master to be used for the new slide.|
|[Presentation](/.presentation)|[slideMasters](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-slidemasters-member)|Returns the collection of `SlideMaster` objects that are in the presentation.|
||[tags](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-tags-member)|Returns a collection of tags attached to the presentation.|
|[Shape](/.shape)|[delete()](/.shape#powerpoint-javascript/api/powerpoint/-shape-delete-member(1))|Deletes the shape from the shape collection.|
||[id](/.shape#powerpoint-javascript/api/powerpoint/-shape-id-member)|Gets the unique ID of the shape.|
||[tags](/.shape#powerpoint-javascript/api/powerpoint/-shape-tags-member)|Returns a collection of tags in the shape.|
|[ShapeCollection](/.shapecollection)|[getCount()](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-getcount-member(1))|Gets the number of shapes in the collection.|
||[getItem(key: string)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-getitem-member(1))|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-getitemat-member(1))|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-getitemornullobject-member(1))|Gets a shape using its unique ID.|
||[items](/.shapecollection#powerpoint-javascript/api/powerpoint/-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[Slide](/.slide)|[layout](/.slide#powerpoint-javascript/api/powerpoint/-slide-layout-member)|Gets the layout of the slide.|
||[shapes](/.slide#powerpoint-javascript/api/powerpoint/-slide-shapes-member)|Returns a collection of shapes in the slide.|
||[slideMaster](/.slide#powerpoint-javascript/api/powerpoint/-slide-slidemaster-member)|Gets the `SlideMaster` object that represents the slide's default content.|
||[tags](/.slide#powerpoint-javascript/api/powerpoint/-slide-tags-member)|Returns a collection of tags in the slide.|
|[SlideCollection](/.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-add-member(1))|Adds a new slide at the end of the collection.|
|[SlideLayout](/.slidelayout)|[id](/.slidelayout#powerpoint-javascript/api/powerpoint/-slidelayout-id-member)|Gets the unique ID of the slide layout.|
||[name](/.slidelayout#powerpoint-javascript/api/powerpoint/-slidelayout-name-member)|Gets the name of the slide layout.|
||[shapes](/.slidelayout#powerpoint-javascript/api/powerpoint/-slidelayout-shapes-member)|Returns a collection of shapes in the slide layout.|
|[SlideLayoutCollection](/.slidelayoutcollection)|[getCount()](/.slidelayoutcollection#powerpoint-javascript/api/powerpoint/-slidelayoutcollection-getcount-member(1))|Gets the number of layouts in the collection.|
||[getItem(key: string)](/.slidelayoutcollection#powerpoint-javascript/api/powerpoint/-slidelayoutcollection-getitem-member(1))|Gets a layout using its unique ID.|
||[getItemAt(index: number)](/.slidelayoutcollection#powerpoint-javascript/api/powerpoint/-slidelayoutcollection-getitemat-member(1))|Gets a layout using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.slidelayoutcollection#powerpoint-javascript/api/powerpoint/-slidelayoutcollection-getitemornullobject-member(1))|Gets a layout using its unique ID.|
||[items](/.slidelayoutcollection#powerpoint-javascript/api/powerpoint/-slidelayoutcollection-items-member)|Gets the loaded child items in this collection.|
|[SlideMaster](/.slidemaster)|[id](/.slidemaster#powerpoint-javascript/api/powerpoint/-slidemaster-id-member)|Gets the unique ID of the Slide Master.|
||[layouts](/.slidemaster#powerpoint-javascript/api/powerpoint/-slidemaster-layouts-member)|Gets the collection of layouts provided by the Slide Master for slides.|
||[name](/.slidemaster#powerpoint-javascript/api/powerpoint/-slidemaster-name-member)|Gets the unique name of the Slide Master.|
||[shapes](/.slidemaster#powerpoint-javascript/api/powerpoint/-slidemaster-shapes-member)|Returns a collection of shapes in the Slide Master.|
|[SlideMasterCollection](/.slidemastercollection)|[getCount()](/.slidemastercollection#powerpoint-javascript/api/powerpoint/-slidemastercollection-getcount-member(1))|Gets the number of Slide Masters in the collection.|
||[getItem(key: string)](/.slidemastercollection#powerpoint-javascript/api/powerpoint/-slidemastercollection-getitem-member(1))|Gets a Slide Master using its unique ID.|
||[getItemAt(index: number)](/.slidemastercollection#powerpoint-javascript/api/powerpoint/-slidemastercollection-getitemat-member(1))|Gets a Slide Master using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.slidemastercollection#powerpoint-javascript/api/powerpoint/-slidemastercollection-getitemornullobject-member(1))|Gets a Slide Master using its unique ID.|
||[items](/.slidemastercollection#powerpoint-javascript/api/powerpoint/-slidemastercollection-items-member)|Gets the loaded child items in this collection.|
|[Tag](/.tag)|[key](/.tag#powerpoint-javascript/api/powerpoint/-tag-key-member)|Gets the unique ID of the tag.|
||[value](/.tag#powerpoint-javascript/api/powerpoint/-tag-value-member)|Gets the value of the tag.|
|[TagCollection](/.tagcollection)|[add(key: string, value: string)](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-add-member(1))|Adds a new tag at the end of the collection.|
||[delete(key: string)](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-delete-member(1))|Deletes the tag with the given `key` in this collection.|
||[getCount()](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-getcount-member(1))|Gets the number of tags in the collection.|
||[getItem(key: string)](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-getitem-member(1))|Gets a tag using its unique ID.|
||[getItemAt(index: number)](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-getitemat-member(1))|Gets a tag using its zero-based index in the collection.|
||[getItemOrNullObject(key: string)](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-getitemornullobject-member(1))|Gets a tag using its unique ID.|
||[items](/.tagcollection#powerpoint-javascript/api/powerpoint/-tagcollection-items-member)|Gets the loaded child items in this collection.|
