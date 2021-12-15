| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceSlideIds)|Specifies the slides from the source presentation that will be inserted into the current presentation.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetSlideId)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertSlidesFromBase64_base64File__options_)|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Returns an ordered collection of slides in the presentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete__)|Deletes the slide from the presentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Gets the unique ID of the slide.|
||[layout](/javascript/api/powerpoint/powerpoint.slide#layout)|Gets the layout of the slide.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Returns a collection of shapes in the slide.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Gets the `SlideMaster` object that represents the slide's default content.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Returns a collection of tags in the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Adds a new slide at the end of the collection.|
||[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getCount__)|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItem_key_)|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_)|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getItemOrNullObject_id_)|Gets a slide using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Gets the loaded child items in this collection.|
