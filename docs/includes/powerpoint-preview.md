| Class | Fields | Description |
|:---|:---|:---|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[group](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-group-member)|Returns the `ShapeGroup` associated with the shape.|
||[level](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-level-member)|Returns the level of the specified shape.|
||[parentGroup](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-parentgroup-member)|Returns the parent group of this shape.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGroup(values: Array<string \| Shape>)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgroup-member(1))|Create a shape group for several shapes.|
|[ShapeGroup](/javascript/api/powerpoint/powerpoint.shapegroup)|[id](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-id-member)|Gets the unique ID of the shape group.|
||[shape](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-shape-member)|Returns the `Shape` object associated with the group.|
||[shapes](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-shapes-member)|Returns the collection of `Shape` objects in the group.|
||[ungroup()](/javascript/api/powerpoint/powerpoint.shapegroup#powerpoint-powerpoint-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeScopedCollection](/javascript/api/powerpoint/powerpoint.shapescopedcollection)|[group()](/javascript/api/powerpoint/powerpoint.shapescopedcollection#powerpoint-powerpoint-shapescopedcollection-group-member(1))|Groups all shapes in this collection into a single shape.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[applyLayout(slideLayout: PowerPoint.SlideLayout)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-applylayout-member(1))|Applies the specified layout to the slide, changing its design and structure according to the chosen layout.|
||[exportAsBase64()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-exportasbase64-member(1))|Exports the slide to its own presentation file, returned as Base64-encoded data.|
||[getImageAsBase64(options?: PowerPoint.SlideGetImageOptions)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-getimageasbase64-member(1))|Renders an image of the slide.|
||[index](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-index-member)|Returns the zero-based index of the slide representing its position in the presentation.|
||[moveTo(slideIndex: number)](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-moveto-member(1))|Moves the slide to a new position within the presentation.|
|[SlideGetImageOptions](/javascript/api/powerpoint/powerpoint.slidegetimageoptions)|[height](/javascript/api/powerpoint/powerpoint.slidegetimageoptions#powerpoint-powerpoint-slidegetimageoptions-height-member)|The desired height of the resulting image in pixels.|
||[width](/javascript/api/powerpoint/powerpoint.slidegetimageoptions#powerpoint-powerpoint-slidegetimageoptions-width-member)|The desired width of the resulting image in pixels.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[type](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-type-member)|Returns the type of the slide layout.|
