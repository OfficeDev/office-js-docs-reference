| Class | Fields | Description |
|:---|:---|:---|
|[InsertSlideOptions](/.insertslideoptions)|[formatting](/.insertslideoptions#powerpoint-javascript/api/powerpoint/-insertslideoptions-formatting-member)|Specifies which formatting to use during slide insertion.|
||[sourceSlideIds](/.insertslideoptions#powerpoint-javascript/api/powerpoint/-insertslideoptions-sourceslideids-member)|Specifies the slides from the source presentation that will be inserted into the current presentation.|
||[targetSlideId](/.insertslideoptions#powerpoint-javascript/api/powerpoint/-insertslideoptions-targetslideid-member)|Specifies where in the presentation the new slides will be inserted.|
|[Presentation](/.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-insertslidesfrombase64-member(1))|Inserts the specified slides from a presentation into the current presentation.|
||[slides](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-slides-member)|Returns an ordered collection of slides in the presentation.|
|[Slide](/.slide)|[delete()](/.slide#powerpoint-javascript/api/powerpoint/-slide-delete-member(1))|Deletes the slide from the presentation.|
||[id](/.slide#powerpoint-javascript/api/powerpoint/-slide-id-member)|Gets the unique ID of the slide.|
|[SlideCollection](/.slidecollection)|[getCount()](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-getcount-member(1))|Gets the number of slides in the collection.|
||[getItem(key: string)](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-getitem-member(1))|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-getitemat-member(1))|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-getitemornullobject-member(1))|Gets a slide using its unique ID.|
||[items](/.slidecollection#powerpoint-javascript/api/powerpoint/-slidecollection-items-member)|Gets the loaded child items in this collection.|
