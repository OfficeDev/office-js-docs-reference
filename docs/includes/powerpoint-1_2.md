| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[formatting](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-formatting-member)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-sourceslideids-member)|Specifies the slides from the source presentation that will be inserted into the current presentation.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-targetslideid-member)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slides-member)|Returns an ordered collection of slides in the presentation.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-delete-member(1))|Deletes the slide from the presentation.|
||[id](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-id-member)|Gets the unique ID of the slide.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getcount-member(1))|Gets the number of slides in the collection.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitem-member(1))|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1))|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemornullobject-member(1))|Gets a slide using its unique ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-items-member)|Gets the loaded child items in this collection.|
