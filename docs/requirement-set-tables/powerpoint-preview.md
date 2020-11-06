| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#formatting)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#sourceslideids)|Specifies the slides from the source presentation that will be inserted into the current presentation.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#targetslideid)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#slides)|Returns an ordered collection of slides in the presentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#delete--)|Deletes the slide from the presentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#id)|Gets the unique ID of the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#getcount--)|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitem-key-)|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemat-index-)|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#getitemornullobject-id-)|Gets a slide using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#items)|Gets the loaded child items in this collection.|
