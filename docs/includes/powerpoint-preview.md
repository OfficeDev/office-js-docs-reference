| Class | Fields | Description |
|:---|:---|:---|
|[PlaceholderFormat](/javascript/api/powerpoint/powerpoint.placeholderformat)|[containedType](/javascript/api/powerpoint/powerpoint.placeholderformat#powerpoint-powerpoint-placeholderformat-containedtype-member)|Gets the type of the shape contained within the placeholder.|
||[type](/javascript/api/powerpoint/powerpoint.placeholderformat#powerpoint-powerpoint-placeholderformat-type-member)|Returns the type of this placeholder.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[group](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-group-member)|Returns the `ShapeGroup` associated with the shape.|
||[level](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-level-member)|Returns the level of the specified shape.|
||[parentGroup](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-parentgroup-member)|Returns the parent group of this shape.|
||[placeholderFormat](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-placeholderformat-member)|Returns the properties that apply specifically to this placeholder.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGroup(values: Array<string \| Shape>)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-addgroup-member(1))|Create a shape group for several shapes.|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[allCaps](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-allcaps-member)|Specifies whether the text in the `TextRange` is set to use the **All Caps** attribute which makes lowercase letters appear as uppercase letters.|
||[doubleStrikethrough](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-doublestrikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Double strikethrough** attribute.|
||[smallCaps](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-smallcaps-member)|Specifies whether the text in the `TextRange` is set to use the **Small Caps** attribute which makes lowercase letters appear as small uppercase letters.|
||[strikethrough](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-strikethrough-member)|Specifies whether the text in the `TextRange` is set to use the **Strikethrough** attribute.|
||[subscript](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-subscript-member)|Specifies whether the text in the `TextRange` is set to use the **Subscript** attribute.|
||[superscript](/javascript/api/powerpoint/powerpoint.shapefont#powerpoint-powerpoint-shapefont-superscript-member)|Specifies whether the text in the `TextRange` is set to use the **Superscript** attribute.|
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
