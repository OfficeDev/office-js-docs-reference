| Class | Fields | Description |
|:---|:---|:---|
|[Presentation](/.presentation)|[getSelectedShapes()](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-getselectedshapes-member(1))|Returns the selected shapes in the current slide of the presentation.|
||[getSelectedSlides()](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-getselectedslides-member(1))|Returns the selected slides in the current view of the presentation.|
||[getSelectedTextRange()](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-getselectedtextrange-member(1))|Returns the selected PowerPoint.TextRange in the current view of the presentation.|
||[getSelectedTextRangeOrNullObject()](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-getselectedtextrangeornullobject-member(1))|Returns the selected PowerPoint.TextRange in the current view of the presentation.|
||[id](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-id-member)|Gets the ID of the presentation.|
||[setSelectedSlides(slideIds: string[])](/.presentation#powerpoint-javascript/api/powerpoint/-presentation-setselectedslides-member(1))|Selects the slides in the current view of the presentation.|
|[Shape](/.shape)|[getParentSlide()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslide-member(1))|Returns the parent PowerPoint.Slide object that holds this `Shape`.|
||[getParentSlideLayout()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslidelayout-member(1))|Returns the parent PowerPoint.SlideLayout object that holds this `Shape`.|
||[getParentSlideLayoutOrNullObject()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslidelayoutornullobject-member(1))|Returns the parent PowerPoint.SlideLayout object that holds this `Shape`.|
||[getParentSlideMaster()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslidemaster-member(1))|Returns the parent PowerPoint.SlideMaster object that holds this `Shape`.|
||[getParentSlideMasterOrNullObject()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslidemasterornullobject-member(1))|Returns the parent PowerPoint.SlideMaster object that holds this `Shape`.|
||[getParentSlideOrNullObject()](/.shape#powerpoint-javascript/api/powerpoint/-shape-getparentslideornullobject-member(1))|Returns the parent PowerPoint.Slide object that holds this `Shape`.|
|[ShapeScopedCollection](/.shapescopedcollection)|[getCount()](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-getcount-member(1))|Gets the number of shapes in the collection.|
||[getItem(key: string)](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-getitem-member(1))|Gets a shape using its unique ID.|
||[getItemAt(index: number)](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-getitemat-member(1))|Gets a shape using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-getitemornullobject-member(1))|Gets a shape using its unique ID.|
||[items](/.shapescopedcollection#powerpoint-javascript/api/powerpoint/-shapescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Slide](/.slide)|[setSelectedShapes(shapeIds: string[])](/.slide#powerpoint-javascript/api/powerpoint/-slide-setselectedshapes-member(1))|Selects the specified shapes.|
|[SlideScopedCollection](/.slidescopedcollection)|[getCount()](/.slidescopedcollection#powerpoint-javascript/api/powerpoint/-slidescopedcollection-getcount-member(1))|Gets the number of slides in the collection.|
||[getItem(key: string)](/.slidescopedcollection#powerpoint-javascript/api/powerpoint/-slidescopedcollection-getitem-member(1))|Gets a slide using its unique ID.|
||[getItemAt(index: number)](/.slidescopedcollection#powerpoint-javascript/api/powerpoint/-slidescopedcollection-getitemat-member(1))|Gets a slide using its zero-based index in the collection.|
||[getItemOrNullObject(id: string)](/.slidescopedcollection#powerpoint-javascript/api/powerpoint/-slidescopedcollection-getitemornullobject-member(1))|Gets a slide using its unique ID.|
||[items](/.slidescopedcollection#powerpoint-javascript/api/powerpoint/-slidescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[TextFrame](/.textframe)|[getParentShape()](/.textframe#powerpoint-javascript/api/powerpoint/-textframe-getparentshape-member(1))|Returns the parent PowerPoint.Shape object that holds this `TextFrame`.|
|[TextRange](/.textrange)|[getParentTextFrame()](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-getparenttextframe-member(1))|Returns the parent PowerPoint.TextFrame object that holds this `TextRange`.|
||[length](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-length-member)|Gets or sets the length of the range that this `TextRange` represents.|
||[setSelected()](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-setselected-member(1))|Selects this `TextRange` in the current view.|
||[start](/.textrange#powerpoint-javascript/api/powerpoint/-textrange-start-member)|Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.|
