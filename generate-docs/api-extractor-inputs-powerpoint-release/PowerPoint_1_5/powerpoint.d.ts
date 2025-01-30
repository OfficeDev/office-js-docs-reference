import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * @remarks
     * [Api set: PowerPointApi 1.0]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Create a new instance of the `PowerPoint.Application` object.
         */
        static newObject(context: OfficeExtension.ClientRequestContext): PowerPoint.Application;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Application` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     * @remarks
     * [Api set: PowerPointApi 1.0]
     */
    export class Presentation extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        /**
         * Returns the collection of `SlideMaster` objects that are in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly slideMasters: PowerPoint.SlideMasterCollection;
        /**
         * Returns an ordered collection of slides in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        readonly slides: PowerPoint.SlideCollection;
        /**
         * Returns a collection of tags attached to the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly tags: PowerPoint.TagCollection;
        /**
         * Gets the ID of the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        readonly id: string;
        readonly title: string;
        /**
         * Returns the selected shapes in the current slide of the presentation.
                    If no shapes are selected, an empty collection is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getSelectedShapes(): PowerPoint.ShapeScopedCollection;
        /**
         * Returns the selected slides in the current view of the presentation.
                    The first item in the collection is the active slide that is visible in the editing area.
                    If no slides are selected, an empty collection is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getSelectedSlides(): PowerPoint.SlideScopedCollection;
        /**
         * Returns the selected {@link PowerPoint.TextRange} in the current view of the presentation.
                    Throws an exception if no text is selected.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getSelectedTextRange(): PowerPoint.TextRange;
        /**
         * Returns the selected {@link PowerPoint.TextRange} in the current view of the presentation.
                    If no text is selected, an object with an `isNullObject` property set to `true` is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getSelectedTextRangeOrNullObject(): PowerPoint.TextRange;
        /**
         * Inserts the specified slides from a presentation into the current presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param base64File - The Base64-encoded string representing the source presentation file.
         * @param options - The options that define which slides will be inserted, where the new slides will go, and which presentation's formatting will be used.
         */
        insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions): void;
        /**
         * Selects the slides in the current view of the presentation. Existing slide selection is replaced with the new selection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param slideIds - List of slide IDs to select in the presentation. If the list is empty, selection is cleared.
         */
        setSelectedSlides(slideIds: string[]): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.PresentationLoadOptions): PowerPoint.Presentation;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Presentation;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Presentation;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Presentation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.PresentationData;
    }
    /**
     * Represents the available options when adding a new slide.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export interface AddSlideOptions {
        /**
         * Specifies the ID of a Slide Layout to be used for the new slide.
                    If no `layoutId` is provided, but a `slideMasterId` is provided, then the ID of the first layout from the specified Slide Master will be used.
                    If no `slideMasterId` is provided, but a `layoutId` is provided, then the specified layout needs to be available for the default Slide Master (as specified
                    in the `slideMasterId` description). Otherwise, an error will be thrown.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        layoutId?: string;
        /**
         * Specifies the ID of a Slide Master to be used for the new slide.
                    If no `slideMasterId` is provided, then the previous slide's Slide Master will be used.
                    If there is no previous slide, then the presentation's first Slide Master will be used.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        slideMasterId?: string;
    }
    
    
    
    /**
     * Specifies the type of a shape.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeType {
        /**
         * The given shape's type is unsupported.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        unsupported = "Unsupported",
        /**
         * The shape is an image.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        image = "Image",
        /**
         * The shape is a geometric shape such as rectangle.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        geometricShape = "GeometricShape",
        /**
         * The shape is a group shape which contains sub-shapes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        group = "Group",
        /**
         * The shape is a line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        line = "Line",
        /**
         * The shape is a table.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        table = "Table",
        /**
         * The shape is a callout.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        callout = "Callout",
        /**
         * The shape is a chart.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chart = "Chart",
        /**
         * The shape is a content Office Add-in.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        contentApp = "ContentApp",
        /**
         * The shape is a diagram.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        diagram = "Diagram",
        /**
         * The shape is a freeform object.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        freeform = "Freeform",
        /**
         * The shape is a graphic.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        graphic = "Graphic",
        /**
         * The shape is an ink object.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ink = "Ink",
        /**
         * The shape is a media object.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        media = "Media",
        /**
         * The shape is a 3D model.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        model3D = "Model3D",
        /**
         * The shape is an OLE (Object Linking and Embedding) object.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ole = "Ole",
        /**
         * The shape is a placeholder.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        placeholder = "Placeholder",
        /**
         * The shape is a SmartArt graphic.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        smartArt = "SmartArt",
        /**
         * The shape is a text box.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        textBox = "TextBox",
    }
    
    
    /**
     * Specifies the connector type for line shapes.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ConnectorType {
        /**
         * Straight connector type
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        straight = "Straight",
        /**
         * Elbow connector type
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        elbow = "Elbow",
        /**
         * Curve connector type
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        curve = "Curve",
    }
    /**
     * Specifies the shape type for a `GeometricShape` object.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum GeometricShapeType {
        /**
         * Straight Line from Top-Right Corner to Bottom-Left Corner of the Shape
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        lineInverse = "LineInverse",
        /**
         * Isosceles Triangle
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        triangle = "Triangle",
        /**
         * Right Triangle
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightTriangle = "RightTriangle",
        /**
         * Rectangle
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rectangle = "Rectangle",
        /**
         * Diamond
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        diamond = "Diamond",
        /**
         * Parallelogram
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        parallelogram = "Parallelogram",
        /**
         * Trapezoid
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        trapezoid = "Trapezoid",
        /**
         * Trapezoid which may have Non-Equal Sides
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        nonIsoscelesTrapezoid = "NonIsoscelesTrapezoid",
        /**
         * Pentagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        pentagon = "Pentagon",
        /**
         * Hexagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        hexagon = "Hexagon",
        /**
         * Heptagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        heptagon = "Heptagon",
        /**
         * Octagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        octagon = "Octagon",
        /**
         * Decagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        decagon = "Decagon",
        /**
         * Dodecagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dodecagon = "Dodecagon",
        /**
         * Star: 4 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star4 = "Star4",
        /**
         * Star: 5 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star5 = "Star5",
        /**
         * Star: 6 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star6 = "Star6",
        /**
         * Star: 7 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star7 = "Star7",
        /**
         * Star: 8 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star8 = "Star8",
        /**
         * Star: 10 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star10 = "Star10",
        /**
         * Star: 12 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star12 = "Star12",
        /**
         * Star: 16 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star16 = "Star16",
        /**
         * Star: 24 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star24 = "Star24",
        /**
         * Star: 32 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        star32 = "Star32",
        /**
         * Rectangle: Rounded Corners
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        roundRectangle = "RoundRectangle",
        /**
         * Rectangle: Single Corner Rounded
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        round1Rectangle = "Round1Rectangle",
        /**
         * Rectangle: Top Corners Rounded
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        round2SameRectangle = "Round2SameRectangle",
        /**
         * Rectangle: Diagonal Corners Rounded
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        round2DiagonalRectangle = "Round2DiagonalRectangle",
        /**
         * Rectangle: Top Corners One Rounded and One Snipped
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        snipRoundRectangle = "SnipRoundRectangle",
        /**
         * Rectangle: Single Corner Snipped
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        snip1Rectangle = "Snip1Rectangle",
        /**
         * Rectangle: Top Corners Snipped
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        snip2SameRectangle = "Snip2SameRectangle",
        /**
         * Rectangle: Diagonal Corners Snipped
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        snip2DiagonalRectangle = "Snip2DiagonalRectangle",
        /**
         * Plaque
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        plaque = "Plaque",
        /**
         * Oval
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ellipse = "Ellipse",
        /**
         * Teardrop
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        teardrop = "Teardrop",
        /**
         * Arrow: Pentagon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        homePlate = "HomePlate",
        /**
         * Arrow: Chevron
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chevron = "Chevron",
        /**
         * Partial Circle
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        pieWedge = "PieWedge",
        /**
         * Partial Circle with Adjustable Spanning Area
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        pie = "Pie",
        /**
         * Block Arc
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        blockArc = "BlockArc",
        /**
         * Circle: Hollow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        donut = "Donut",
        /**
         * "Not Allowed" Symbol
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        noSmoking = "NoSmoking",
        /**
         * Arrow: Right
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightArrow = "RightArrow",
        /**
         * Arrow: Left
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftArrow = "LeftArrow",
        /**
         * Arrow: Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        upArrow = "UpArrow",
        /**
         * Arrow: Down
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        downArrow = "DownArrow",
        /**
         * Arrow: Striped Right
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        stripedRightArrow = "StripedRightArrow",
        /**
         * Arrow: Notched Right
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        notchedRightArrow = "NotchedRightArrow",
        /**
         * Arrow: Bent-Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bentUpArrow = "BentUpArrow",
        /**
         * Arrow: Left-Right
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftRightArrow = "LeftRightArrow",
        /**
         * Arrow: Up-Down
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        upDownArrow = "UpDownArrow",
        /**
         * Arrow: Left-Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftUpArrow = "LeftUpArrow",
        /**
         * Arrow: Left-Right-Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftRightUpArrow = "LeftRightUpArrow",
        /**
         * Arrow: Quad
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        quadArrow = "QuadArrow",
        /**
         * Callout: Left Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftArrowCallout = "LeftArrowCallout",
        /**
         * Callout: Right Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightArrowCallout = "RightArrowCallout",
        /**
         * Callout: Up Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        upArrowCallout = "UpArrowCallout",
        /**
         * Callout: Down Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        downArrowCallout = "DownArrowCallout",
        /**
         * Callout: Left-Right Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftRightArrowCallout = "LeftRightArrowCallout",
        /**
         * Callout: Up-Down Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        upDownArrowCallout = "UpDownArrowCallout",
        /**
         * Callout: Quad Arrow
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        quadArrowCallout = "QuadArrowCallout",
        /**
         * Arrow: Bent
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bentArrow = "BentArrow",
        /**
         * Arrow: U-Turn
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        uturnArrow = "UturnArrow",
        /**
         * Arrow: Circular
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        circularArrow = "CircularArrow",
        /**
         * Arrow: Circular with Opposite Arrow Direction
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftCircularArrow = "LeftCircularArrow",
        /**
         * Arrow: Circular with Two Arrows in Both Directions
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftRightCircularArrow = "LeftRightCircularArrow",
        /**
         * Arrow: Curved Right
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        curvedRightArrow = "CurvedRightArrow",
        /**
         * Arrow: Curved Left
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        curvedLeftArrow = "CurvedLeftArrow",
        /**
         * Arrow: Curved Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        curvedUpArrow = "CurvedUpArrow",
        /**
         * Arrow: Curved Down
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        curvedDownArrow = "CurvedDownArrow",
        /**
         * Arrow: Curved Right Arrow with Varying Width
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        swooshArrow = "SwooshArrow",
        /**
         * Cube
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        cube = "Cube",
        /**
         * Cylinder
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        can = "Can",
        /**
         * Lightning Bolt
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        lightningBolt = "LightningBolt",
        /**
         * Heart
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        heart = "Heart",
        /**
         * Sun
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        sun = "Sun",
        /**
         * Moon
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        moon = "Moon",
        /**
         * Smiley Face
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        smileyFace = "SmileyFace",
        /**
         * Explosion: 8 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        irregularSeal1 = "IrregularSeal1",
        /**
         * Explosion: 14 Points
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        irregularSeal2 = "IrregularSeal2",
        /**
         * Rectangle: Folded Corner
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        foldedCorner = "FoldedCorner",
        /**
         * Rectangle: Beveled
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bevel = "Bevel",
        /**
         * Frame
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        frame = "Frame",
        /**
         * Half Frame
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        halfFrame = "HalfFrame",
        /**
         * L-Shape
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        corner = "Corner",
        /**
         * Diagonal Stripe
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        diagonalStripe = "DiagonalStripe",
        /**
         * Chord
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chord = "Chord",
        /**
         * Arc
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        arc = "Arc",
        /**
         * Left Bracket
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftBracket = "LeftBracket",
        /**
         * Right Bracket
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightBracket = "RightBracket",
        /**
         * Left Brace
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftBrace = "LeftBrace",
        /**
         * Right Brace
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightBrace = "RightBrace",
        /**
         * Double Bracket
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bracketPair = "BracketPair",
        /**
         * Double Brace
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bracePair = "BracePair",
        /**
         * Callout: Line with No Border
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        callout1 = "Callout1",
        /**
         * Callout: Bent Line with No Border
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        callout2 = "Callout2",
        /**
         * Callout: Double Bent Line with No Border
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        callout3 = "Callout3",
        /**
         * Callout: Line with Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentCallout1 = "AccentCallout1",
        /**
         * Callout: Bent Line with Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentCallout2 = "AccentCallout2",
        /**
         * Callout: Double Bent Line with Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentCallout3 = "AccentCallout3",
        /**
         * Callout: Line
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        borderCallout1 = "BorderCallout1",
        /**
         * Callout: Bent Line
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        borderCallout2 = "BorderCallout2",
        /**
         * Callout: Double Bent Line
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        borderCallout3 = "BorderCallout3",
        /**
         * Callout: Line with Border and Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentBorderCallout1 = "AccentBorderCallout1",
        /**
         * Callout: Bent Line with Border and Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentBorderCallout2 = "AccentBorderCallout2",
        /**
         * Callout: Double Bent Line with Border and Accent Bar
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        accentBorderCallout3 = "AccentBorderCallout3",
        /**
         * Speech Bubble: Rectangle
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wedgeRectCallout = "WedgeRectCallout",
        /**
         * Speech Bubble: Rectangle with Corners Rounded
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wedgeRRectCallout = "WedgeRRectCallout",
        /**
         * Speech Bubble: Oval
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wedgeEllipseCallout = "WedgeEllipseCallout",
        /**
         * Thought Bubble: Cloud
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        cloudCallout = "CloudCallout",
        /**
         * Cloud
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        cloud = "Cloud",
        /**
         * Ribbon: Tilted Down
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ribbon = "Ribbon",
        /**
         * Ribbon: Tilted Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ribbon2 = "Ribbon2",
        /**
         * Ribbon: Curved and Tilted Down
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ellipseRibbon = "EllipseRibbon",
        /**
         * Ribbon: Curved and Tilted Up
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        ellipseRibbon2 = "EllipseRibbon2",
        /**
         * Ribbon: Straight with Both Left and Right Arrows
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftRightRibbon = "LeftRightRibbon",
        /**
         * Scroll: Vertical
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        verticalScroll = "VerticalScroll",
        /**
         * Scroll: Horizontal
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        horizontalScroll = "HorizontalScroll",
        /**
         * Wave
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wave = "Wave",
        /**
         * Double Wave
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        doubleWave = "DoubleWave",
        /**
         * Cross
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        plus = "Plus",
        /**
         * Flowchart: Process
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartProcess = "FlowChartProcess",
        /**
         * Flowchart: Decision
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartDecision = "FlowChartDecision",
        /**
         * Flowchart: Data
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartInputOutput = "FlowChartInputOutput",
        /**
         * Flowchart: Predefined Process
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartPredefinedProcess = "FlowChartPredefinedProcess",
        /**
         * Flowchart: Internal Storage
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartInternalStorage = "FlowChartInternalStorage",
        /**
         * Flowchart: Document
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartDocument = "FlowChartDocument",
        /**
         * Flowchart: Multidocument
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartMultidocument = "FlowChartMultidocument",
        /**
         * Flowchart: Terminator
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartTerminator = "FlowChartTerminator",
        /**
         * Flowchart: Preparation
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartPreparation = "FlowChartPreparation",
        /**
         * Flowchart: Manual Input
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartManualInput = "FlowChartManualInput",
        /**
         * Flowchart: Manual Operation
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartManualOperation = "FlowChartManualOperation",
        /**
         * Flowchart: Connector
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartConnector = "FlowChartConnector",
        /**
         * Flowchart: Card
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartPunchedCard = "FlowChartPunchedCard",
        /**
         * Flowchart: Punched Tape
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartPunchedTape = "FlowChartPunchedTape",
        /**
         * Flowchart: Summing Junction
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartSummingJunction = "FlowChartSummingJunction",
        /**
         * Flowchart: Or
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartOr = "FlowChartOr",
        /**
         * Flowchart: Collate
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartCollate = "FlowChartCollate",
        /**
         * Flowchart: Sort
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartSort = "FlowChartSort",
        /**
         * Flowchart: Extract
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartExtract = "FlowChartExtract",
        /**
         * Flowchart: Merge
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartMerge = "FlowChartMerge",
        /**
         * FlowChart: Offline Storage
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartOfflineStorage = "FlowChartOfflineStorage",
        /**
         * Flowchart: Stored Data
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartOnlineStorage = "FlowChartOnlineStorage",
        /**
         * Flowchart: Sequential Access Storage
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartMagneticTape = "FlowChartMagneticTape",
        /**
         * Flowchart: Magnetic Disk
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartMagneticDisk = "FlowChartMagneticDisk",
        /**
         * Flowchart: Direct Access Storage
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartMagneticDrum = "FlowChartMagneticDrum",
        /**
         * Flowchart: Display
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartDisplay = "FlowChartDisplay",
        /**
         * Flowchart: Delay
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartDelay = "FlowChartDelay",
        /**
         * Flowchart: Alternate Process
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartAlternateProcess = "FlowChartAlternateProcess",
        /**
         * Flowchart: Off-page Connector
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        flowChartOffpageConnector = "FlowChartOffpageConnector",
        /**
         * Action Button: Blank
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonBlank = "ActionButtonBlank",
        /**
         * Action Button: Go Home
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonHome = "ActionButtonHome",
        /**
         * Action Button: Help
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonHelp = "ActionButtonHelp",
        /**
         * Action Button: Get Information
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonInformation = "ActionButtonInformation",
        /**
         * Action Button: Go Forward or Next
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonForwardNext = "ActionButtonForwardNext",
        /**
         * Action Button: Go Back or Previous
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonBackPrevious = "ActionButtonBackPrevious",
        /**
         * Action Button: Go to End
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonEnd = "ActionButtonEnd",
        /**
         * Action Button: Go to Beginning
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonBeginning = "ActionButtonBeginning",
        /**
         * Action Button: Return
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonReturn = "ActionButtonReturn",
        /**
         * Action Button: Document
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonDocument = "ActionButtonDocument",
        /**
         * Action Button: Sound
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonSound = "ActionButtonSound",
        /**
         * Action Button: Video
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        actionButtonMovie = "ActionButtonMovie",
        /**
         * Gear: A Gear with Six Teeth
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        gear6 = "Gear6",
        /**
         * Gear: A Gear with Nine Teeth
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        gear9 = "Gear9",
        /**
         * Funnel
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        funnel = "Funnel",
        /**
         * Plus Sign
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathPlus = "MathPlus",
        /**
         * Minus Sign
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathMinus = "MathMinus",
        /**
         * Multiplication Sign
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathMultiply = "MathMultiply",
        /**
         * Division Sign
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathDivide = "MathDivide",
        /**
         * Equals
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathEqual = "MathEqual",
        /**
         * Not Equal
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        mathNotEqual = "MathNotEqual",
        /**
         * Four Right Triangles that Define a Rectangular Shape
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        cornerTabs = "CornerTabs",
        /**
         * Four Small Squares that Define a Rectangular Shape.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        squareTabs = "SquareTabs",
        /**
         * Four Quarter Circles that Define a Rectangular Shape.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        plaqueTabs = "PlaqueTabs",
        /**
         * A Rectangle Divided into Four Parts Along Diagonal Lines.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chartX = "ChartX",
        /**
         * A Rectangle Divided into Six Parts Along a Vertical Line and Diagonal Lines.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chartStar = "ChartStar",
        /**
         * A Rectangle Divided Vertically and Horizontally into Four Quarters.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        chartPlus = "ChartPlus",
    }
    /**
     * Represents the available options when adding shapes.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export interface ShapeAddOptions {
        /**
         * Specifies the height, in points, of the shape.
                    When not provided, a default value will be used.
                    Throws an `InvalidArgument` exception when set with a negative value.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        height?: number;
        /**
         * Specifies the distance, in points, from the left side of the shape to the left side of the slide.
                    When not provided, a default value will be used.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        left?: number;
        /**
         * Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.
                    When not provided, a default value will be used.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        top?: number;
        /**
         * Specifies the width, in points, of the shape.
                    When not provided, a default value will be used.
                    Throws an `InvalidArgument` exception when set with a negative value.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        width?: number;
    }
    /**
     * Specifies the dash style for a line.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeLineDashStyle {
        /**
         * The dash line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dash = "Dash",
        /**
         * The dash-dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashDot = "DashDot",
        /**
         * The dash-dot-dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashDotDot = "DashDotDot",
        /**
         * The long dash line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        longDash = "LongDash",
        /**
         * The long dash-dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        longDashDot = "LongDashDot",
        /**
         * The round dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        roundDot = "RoundDot",
        /**
         * The solid line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        solid = "Solid",
        /**
         * The square dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        squareDot = "SquareDot",
        /**
         * The long dash-dot-dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        longDashDotDot = "LongDashDotDot",
        /**
         * The system dash line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        systemDash = "SystemDash",
        /**
         * The system dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        systemDot = "SystemDot",
        /**
         * The system dash-dot line pattern.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        systemDashDot = "SystemDashDot",
    }
    /**
     * Represents the horizontal alignment of the {@link PowerPoint.TextFrame} in a {@link PowerPoint.Shape}.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ParagraphHorizontalAlignment {
        /**
         * Align text to the left margin.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        left = "Left",
        /**
         * Align text in the center.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        center = "Center",
        /**
         * Align text to the right margin.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        right = "Right",
        /**
         * Align text so that it is justified across the whole line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        justify = "Justify",
        /**
         * Specifies the alignment or adjustment of kashida length in Arabic text.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        justifyLow = "JustifyLow",
        /**
         * Distributes the text words across an entire text line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        distributed = "Distributed",
        /**
         * Distributes Thai text specially, because each character is treated as a word.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        thaiDistributed = "ThaiDistributed",
    }
    /**
     * Specifies a shape's fill type.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeFillType {
        /**
         * Specifies that the shape should have no fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        noFill = "NoFill",
        /**
         * Specifies that the shape should have regular solid fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        solid = "Solid",
        /**
         * Specifies that the shape should have gradient fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        gradient = "Gradient",
        /**
         * Specifies that the shape should have pattern fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        pattern = "Pattern",
        /**
         * Specifies that the shape should have picture or texture fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        pictureAndTexture = "PictureAndTexture",
        /**
         * Specifies that the shape should have slide background fill.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        slideBackground = "SlideBackground",
    }
    /**
     * Represents the fill formatting of a shape object.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class ShapeFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        foregroundColor: string;
        /**
         * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        transparency: number;
        /**
         * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly type: PowerPoint.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "PictureAndTexture" | "SlideBackground";
        /**
         * Clears the fill formatting of this shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        clear(): void;
        /**
         * Sets the fill formatting of the shape to a uniform color. This changes the fill type to `Solid`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param color - A string that specifies the fill color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeFillLoadOptions): PowerPoint.ShapeFill;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ShapeFill;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.ShapeFill;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `PowerPoint.ShapeFill` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeFillData;
    }
    /**
     * The type of underline applied to a font.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeFontUnderlineStyle {
        /**
         * No underlining.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        none = "None",
        /**
         * Regular single line underlining.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        single = "Single",
        /**
         * Underlining of text with double lines.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        double = "Double",
        /**
         * Underlining of text with a thick line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        heavy = "Heavy",
        /**
         * Underlining of text with a dotted line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dotted = "Dotted",
        /**
         * Underlining of text with a thick, dotted line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dottedHeavy = "DottedHeavy",
        /**
         * Underlining of text with a line containing dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dash = "Dash",
        /**
         * Underlining of text with a thick line containing dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashHeavy = "DashHeavy",
        /**
         * Underlining of text with a line containing long dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashLong = "DashLong",
        /**
         * Underlining of text with a thick line containing long dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashLongHeavy = "DashLongHeavy",
        /**
         * Underlining of text with a line containing dots and dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dotDash = "DotDash",
        /**
         * Underlining of text with a thick line containing dots and dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dotDashHeavy = "DotDashHeavy",
        /**
         * Underlining of text with a line containing double dots and dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dotDotDash = "DotDotDash",
        /**
         * Underlining of text with a thick line containing double dots and dashes.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dotDotDashHeavy = "DotDotDashHeavy",
        /**
         * Underlining of text with a wavy line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wavy = "Wavy",
        /**
         * Underlining of text with a thick, wavy line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wavyHeavy = "WavyHeavy",
        /**
         * Underlining of text with double wavy lines.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wavyDouble = "WavyDouble",
    }
    /**
     * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class ShapeFont extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies the bold status of font. Returns `null` if the `TextRange` contains both bold and non-bold text fragments.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bold: boolean | null;
        /**
         * Specifies the HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` contains text fragments with different colors.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        color: string | null;
        /**
         * Specifies the italic status of font. Returns 'null' if the 'TextRange' contains both italic and non-italic text fragments.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        italic: boolean | null;
        /**
         * Specifies the font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name. Returns `null` if the `TextRange` contains text fragments with different font names.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        name: string | null;
        /**
         * Specifies the font size in points (e.g., 11). Returns `null` if the `TextRange` contains text fragments with different font sizes.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        size: number | null;
        /**
         * Specifies the type of underline applied to the font. Returns `null` if the `TextRange` contains text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        underline: PowerPoint.ShapeFontUnderlineStyle | "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble" | null;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeFontLoadOptions): PowerPoint.ShapeFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ShapeFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.ShapeFont;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `PowerPoint.ShapeFont` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeFontData;
    }
    /**
     * Represents the vertical alignment of a {@link PowerPoint.TextFrame} in a {@link PowerPoint.Shape}.
                If one the centered options are selected, the contents of the `TextFrame` will be centered horizontally within the `Shape` as a group.
                To change the horizontal alignment of a text, see {@link PowerPoint.ParagraphFormat} and {@link PowerPoint.ParagraphHorizontalAlignment}.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum TextVerticalAlignment {
        /**
         * Specifies that the `TextFrame` should be top aligned to the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        top = "Top",
        /**
         * Specifies that the `TextFrame` should be center aligned to the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        middle = "Middle",
        /**
         * Specifies that the `TextFrame` should be bottom aligned to the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bottom = "Bottom",
        /**
         * Specifies that the `TextFrame` should be top aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        topCentered = "TopCentered",
        /**
         * Specifies that the `TextFrame` should be center aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        middleCentered = "MiddleCentered",
        /**
         * Specifies that the `TextFrame` should be bottom aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bottomCentered = "BottomCentered",
    }
    /**
     * Represents the collection of shapes.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class ShapeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Shape[];
        /**
         * Adds a geometric shape to the slide. Returns a `Shape` object that represents the new shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param geometricShapeType - Specifies the type of the geometric shape. See {@link PowerPoint.GeometricShapeType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape.
         * @returns The newly inserted shape.
         */
        addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a geometric shape to the slide. Returns a `Shape` object that represents the new shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param geometricShapeTypeString - Specifies the type of the geometric shape. See {@link PowerPoint.GeometricShapeType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape.
         * @returns The newly inserted shape.
         */
        addGeometricShape(geometricShapeTypeString: "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus", options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a line to the slide. Returns a `Shape` object that represents the new line.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param connectorType - Specifies the connector type of the line. If not provided, `straight` connector type will be used. See {@link PowerPoint.ConnectorType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape object that contains the line.
         * @returns The newly inserted shape.
         */
        addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a line to the slide. Returns a `Shape` object that represents the new line.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param connectorTypeString - Specifies the connector type of the line. If not provided, `straight` connector type will be used. See {@link PowerPoint.ConnectorType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape object that contains the line.
         * @returns The newly inserted shape.
         */
        addLine(connectorTypeString?: "Straight" | "Elbow" | "Curve", options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a text box to the slide with the provided text as the content. Returns a `Shape` object that represents the new text box.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param text - Specifies the text that will be shown in the created text box.
         * @param options - An optional parameter to specify the additional options such as the position of the text box.
         * @returns The newly inserted shape.
         */
        addTextBox(text: string, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Gets the number of shapes in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         * @returns The number of shapes in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a shape using its unique ID. An error is thrown if the shape does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The ID of the shape.
         * @returns The shape with the unique ID. If such a shape does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Shape;
        /**
         * Gets a shape using its zero-based index in the collection. An error is thrown if the index is out of range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param index - The index of the shape in the collection.
         * @returns The shape at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Shape;
        /**
         * Gets a shape using its unique ID. If such a shape does not exist, an object with an `isNullObject` property set to true is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param id - The ID of the shape.
         * @returns The shape with the unique ID. If such a shape does not exist, an object with an `isNullObject` property set to true is returned.
         */
        getItemOrNullObject(id: string): PowerPoint.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.ShapeCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.ShapeCollectionData;
    }
    /**
     * Represents the layout of a slide.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class SlideLayout extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Returns a collection of shapes in the slide layout.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         * Gets the unique ID of the slide layout.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly id: string;
        /**
         * Gets the name of the slide layout.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly name: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideLayoutLoadOptions): PowerPoint.SlideLayout;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideLayout;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.SlideLayout;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideLayout` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideLayoutData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.SlideLayoutData;
    }
    /**
     * Represents the collection of layouts provided by the Slide Master for slides.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class SlideLayoutCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.SlideLayout[];
        /**
         * Gets the number of layouts in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         * @returns The number of layouts in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a layout using its unique ID.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The ID of the layout.
         * @returns The layout with the unique ID. If such a layout does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.SlideLayout;
        /**
         * Gets a layout using its zero-based index in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param index - The index of the layout in the collection.
         * @returns The layout at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.SlideLayout;
        /**
         * Gets a layout using its unique ID.  If such a layout does not exist, an object with an `isNullObject` property set to true is returned. For further information,
                    see {@link https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param id - The ID of the layout.
         * @returns The layout with the unique ID.
         */
        getItemOrNullObject(id: string): PowerPoint.SlideLayout;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideLayoutCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.SlideLayoutCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideLayoutCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.SlideLayoutCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideLayoutCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideLayoutCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.SlideLayoutCollectionData;
    }
    /**
     * Represents the Slide Master of a slide.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class SlideMaster extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Gets the collection of layouts provided by the Slide Master for slides.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly layouts: PowerPoint.SlideLayoutCollection;
        /**
         * Returns a collection of shapes in the Slide Master.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         * Gets the unique ID of the Slide Master.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly id: string;
        /**
         * Gets the unique name of the Slide Master.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly name: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideMasterLoadOptions): PowerPoint.SlideMaster;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideMaster;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.SlideMaster;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideMaster` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideMasterData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.SlideMasterData;
    }
    /**
     * Represents a single tag in the slide.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class Tag extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly key: string;
        /**
         * Gets the value of the tag.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        value: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.TagLoadOptions): PowerPoint.Tag;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Tag;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Tag;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Tag` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TagData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.TagData;
    }
    /**
     * Represents the collection of tags.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class TagCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Tag[];
        /**
         * Adds a new tag at the end of the collection. If the `key` already exists in the collection, the value of the existing tag will be replaced with the given `value`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The unique ID of a tag, which is unique within this `TagCollection`. 'key' parameter is case-insensitive, but it is always capitalized when saved in the document.
         * @param value - The value of the tag.
         */
        add(key: string, value: string): void;
        /**
         * Deletes the tag with the given `key` in this collection. Does nothing if the `key` does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The unique ID of a tag, which is unique within this `TagCollection`. `key` parameter is case-insensitive.
         */
        delete(key: string): void;
        /**
         * Gets the number of tags in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         * @returns The number of tags in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a tag using its unique ID. An error is thrown if the tag does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The ID of the tag.
         * @returns The tag with the unique ID. If such a tag does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Tag;
        /**
         * Gets a tag using its zero-based index in the collection. An error is thrown if the index is out of range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param index - The index of the tag in the collection.
         * @returns The tag at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Tag;
        /**
         * Gets a tag using its unique ID. If such a tag does not exist, an object with an `isNullObject` property set to true is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The ID of the tag.
         * @returns The tag with the unique ID. If such a tag does not exist, an object with an `isNullObject` property set to true is returned.
         */
        getItemOrNullObject(key: string): PowerPoint.Tag;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.TagCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.TagCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.TagCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.TagCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.TagCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TagCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.TagCollectionData;
    }
    /**
     * Represents a single slide of a presentation.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    export class Slide extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        /**
         * Gets the layout of the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly layout: PowerPoint.SlideLayout;
        /**
         * Returns a collection of shapes in the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         * Gets the `SlideMaster` object that represents the slide's default content.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly slideMaster: PowerPoint.SlideMaster;
        /**
         * Returns a collection of tags in the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly tags: PowerPoint.TagCollection;
        /**
         * Gets the unique ID of the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        readonly id: string;
        /**
         * Deletes the slide from the presentation. Does nothing if the slide does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        delete(): void;
        /**
         * Selects the specified shapes. Existing shape selection is replaced with the new selection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param shapeIds - List of shape IDs to select in the slide. If the list is empty, the selection is cleared.
         */
        setSelectedShapes(shapeIds: string[]): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideLoadOptions): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Slide;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Slide` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.SlideData;
    }
    /**
     * Represents a collection of shapes.
     *
     * @remarks
     * [Api set: PowerPointApi 1.5]
     */
    export class ShapeScopedCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Shape[];
        /**
         * Gets the number of shapes in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         * @returns The number of shapes in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a shape using its unique ID. An error is thrown if the shape does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param key - The ID of the shape.
         * @returns The shape with the unique ID. If such a shape does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Shape;
        /**
         * Gets a shape using its zero-based index in the collection. An error is thrown if the index is out of range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param index - The index of the shape in the collection.
         * @returns The shape at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Shape;
        /**
         * Gets a shape using its unique ID. If such a shape does not exist, an object with an `isNullObject` property set to true is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param id - The ID of the shape.
         * @returns The shape with the unique ID. If such a shape does not exist, an object with an `isNullObject` property set to true is returned.
         */
        getItemOrNullObject(id: string): PowerPoint.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeScopedCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.ShapeScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ShapeScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.ShapeScopedCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.ShapeScopedCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeScopedCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.ShapeScopedCollectionData;
    }
    /**
     * Specifies the style for a line.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeLineStyle {
        /**
         * Single line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        single = "Single",
        /**
         * Thick line with a thin line on each side.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        thickBetweenThin = "ThickBetweenThin",
        /**
         * Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        thickThin = "ThickThin",
        /**
         * Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        thinThick = "ThinThick",
        /**
         * Two thin lines.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        thinThin = "ThinThin",
    }
    /**
     * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class ShapeLineFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        color: string;
        /**
         * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        dashStyle: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
        /**
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        style: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
        /**
         * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        transparency: number;
        /**
         * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        visible: boolean;
        /**
         * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        weight: number;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions): PowerPoint.ShapeLineFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ShapeLineFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.ShapeLineFormat;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.ShapeLineFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeLineFormatData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.ShapeLineFormatData;
    }
    /**
     * Determines the type of automatic sizing allowed.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    enum ShapeAutoSize {
        /**
         * No autosizing.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        autoSizeNone = "AutoSizeNone",
        /**
         * The text is adjusted to fit the shape.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        autoSizeTextToFitShape = "AutoSizeTextToFitShape",
        /**
         * The shape is adjusted to fit the text.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        autoSizeShapeToFitText = "AutoSizeShapeToFitText",
        /**
         * A combination of automatic sizing schemes are used.
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        autoSizeMixed = "AutoSizeMixed",
    }
    /**
     * Represents the bullet formatting properties of a text that is attached to the {@link PowerPoint.ParagraphFormat}.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class BulletFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        visible: boolean;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.BulletFormatLoadOptions): PowerPoint.BulletFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.BulletFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.BulletFormat;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.BulletFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.BulletFormatData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.BulletFormatData;
    }
    /**
     * Represents the paragraph formatting properties of a text that is attached to the {@link PowerPoint.TextRange}.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class ParagraphFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the bullet format of the paragraph. See {@link PowerPoint.BulletFormat} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly bulletFormat: PowerPoint.BulletFormat;
        /**
         * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ParagraphFormatLoadOptions): PowerPoint.ParagraphFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.ParagraphFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.ParagraphFormat;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.ParagraphFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ParagraphFormatData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.ParagraphFormatData;
    }
    /**
     * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class TextRange extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Returns a `ShapeFont` object that represents the font attributes for the text range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly font: PowerPoint.ShapeFont;
        /**
         * Represents the paragraph format of the text range. See {@link PowerPoint.ParagraphFormat} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly paragraphFormat: PowerPoint.ParagraphFormat;
        /**
         * Gets or sets the length of the range that this `TextRange` represents.
                    Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the available text from the starting point.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        length: number;
        /**
         * Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.
                    Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the text.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        start: number;
        /**
         * Represents the plain text content of the text range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        text: string;
        /**
         * Returns the parent {@link PowerPoint.TextFrame} object that holds this `TextRange`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentTextFrame(): PowerPoint.TextFrame;
        /**
         * Returns a `TextRange` object for the substring in the given range.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         *
         * @param start - The zero-based index of the first character to get from the text range.
         * @param length - Optional. The number of characters to be returned in the new text range. If length is omitted, all the characters from start to the end of the text range's last paragraph will be returned.
         */
        getSubstring(start: number, length?: number): PowerPoint.TextRange;
        /**
         * Selects this `TextRange` in the current view.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        setSelected(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.TextRangeLoadOptions): PowerPoint.TextRange;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.TextRange;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.TextRange;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.TextRange` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TextRangeData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.TextRangeData;
    }
    /**
     * Represents the text frame of a shape object.
     *
     * @remarks
     * [Api set: PowerPointApi 1.4]
     */
    export class TextFrame extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See {@link PowerPoint.TextRange} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly textRange: PowerPoint.TextRange;
        /**
         * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        autoSizeSetting: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
        /**
         * Represents the bottom margin, in points, of the text frame.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        bottomMargin: number;
        /**
         * Specifies if the text frame contains text.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly hasText: boolean;
        /**
         * Represents the left margin, in points, of the text frame.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        leftMargin: number;
        /**
         * Represents the right margin, in points, of the text frame.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        rightMargin: number;
        /**
         * Represents the top margin, in points, of the text frame.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        topMargin: number;
        /**
         * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        verticalAlignment: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
        /**
         * Determines whether lines break automatically to fit text inside the shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        wordWrap: boolean;
        /**
         * Deletes all the text in the text frame.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        deleteText(): void;
        /**
         * Returns the parent {@link PowerPoint.Shape} object that holds this `TextFrame`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentShape(): PowerPoint.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.TextFrameLoadOptions): PowerPoint.TextFrame;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.TextFrame;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.TextFrame;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.TextFrame` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TextFrameData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.TextFrameData;
    }
    /**
     * Represents a single shape in the slide.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class Shape extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Returns the fill formatting of this shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly fill: PowerPoint.ShapeFill;
        /**
         * Returns the line formatting of this shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly lineFormat: PowerPoint.ShapeLineFormat;
        /**
         * Returns a collection of tags in the shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly tags: PowerPoint.TagCollection;
        /**
         * Returns the text frame object of this shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly textFrame: PowerPoint.TextFrame;
        /**
         * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        height: number;
        /**
         * Gets the unique ID of the shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        readonly id: string;
        /**
         * The distance, in points, from the left side of the shape to the left side of the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        left: number;
        /**
         * Specifies the name of this shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        name: string;
        /**
         * The distance, in points, from the top edge of the shape to the top edge of the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        top: number;
        /**
         * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        readonly type: PowerPoint.ShapeType | "Unsupported" | "Image" | "GeometricShape" | "Group" | "Line" | "Table" | "Callout" | "Chart" | "ContentApp" | "Diagram" | "Freeform" | "Graphic" | "Ink" | "Media" | "Model3D" | "Ole" | "Placeholder" | "SmartArt" | "TextBox";
        /**
         * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        width: number;
        /**
         * Deletes the shape from the shape collection. Does nothing if the shape does not exist.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        delete(): void;
        /**
         * Returns the parent {@link PowerPoint.Slide} object that holds this `Shape`. Throws an exception if this shape does not belong to a `Slide`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlide(): PowerPoint.Slide;
        /**
         * Returns the parent {@link PowerPoint.SlideLayout} object that holds this `Shape`. Throws an exception if this shape does not belong to a `SlideLayout`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlideLayout(): PowerPoint.SlideLayout;
        /**
         * Returns the parent {@link PowerPoint.SlideLayout} object that holds this `Shape`. If this shape does not belong to a `SlideLayout`, an object with an `isNullObject` property set to `true` is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlideLayoutOrNullObject(): PowerPoint.SlideLayout;
        /**
         * Returns the parent {@link PowerPoint.SlideMaster} object that holds this `Shape`. Throws an exception if this shape does not belong to a `SlideMaster`.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlideMaster(): PowerPoint.SlideMaster;
        /**
         * Returns the parent {@link PowerPoint.SlideMaster} object that holds this `Shape`. If this shape does not belong to a `SlideMaster`, an object with an `isNullObject` property set to `true` is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlideMasterOrNullObject(): PowerPoint.SlideMaster;
        /**
         * Returns the parent {@link PowerPoint.Slide} object that holds this `Shape`. If this shape does not belong to a `Slide`, an object with an `isNullObject` property set to `true` is returned.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        getParentSlideOrNullObject(): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.ShapeLoadOptions): PowerPoint.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Shape;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Shape` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.ShapeData;
    }
    
    
    
    
    /**
     * Specifies the formatting options for when slides are inserted.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    enum InsertSlideFormatting {
        /**
         * Copy the source theme into the target presentation and use that theme.
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        keepSourceFormatting = "KeepSourceFormatting",
        /**
         * Use the existing theme in the target presentation.
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        useDestinationTheme = "UseDestinationTheme",
    }
    /**
     * Represents the available options when inserting slides.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    export interface InsertSlideOptions {
        /**
         * Specifies which formatting to use during slide insertion.
                    The default option is to use "KeepSourceFormatting".
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        formatting?: PowerPoint.InsertSlideFormatting | "KeepSourceFormatting" | "UseDestinationTheme";
        /**
         * Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.
                    The order of these slides is preserved during the insertion.
                    If any of the source slides are not found, or if the IDs are invalid, the operation throws a `SlideNotFound` exception and no slides will be inserted.
                    All of the source slides will be inserted when `sourceSlideIds` is not provided (this is the default behavior).
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        sourceSlideIds?: string[];
        /**
         * Specifies where in the presentation the new slides will be inserted. The new slides will be inserted after the slide with the given slide ID.
                    If `targetSlideId` is not provided, the slides will be inserted at the beginning of the presentation.
                    If `targetSlideId` is invalid or if it is pointing to a non-existing slide, the operation throws a `SlideNotFound` exception and no slides will be inserted.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        targetSlideId?: string;
    }
    /**
     * Represents the collection of slides in the presentation.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    export class SlideCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Slide[];
        /**
         * Adds a new slide at the end of the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param options - The options that define the theme of the new slide.
         */
        add(options?: PowerPoint.AddSlideOptions): void;
        /**
         * Gets the number of slides in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         * @returns The number of slides in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a slide using its unique ID.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param key - The ID of the slide.
         * @returns The slide with the unique ID. If such a slide does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Slide;
        /**
         * Gets a slide using its zero-based index in the collection. Slides are stored in the same order as they
                    are shown in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param index - The index of the slide in the collection.
         * @returns The slide at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Slide;
        /**
         * Gets a slide using its unique ID. If such a slide does not exist, an object with an `isNullObject` property set to true is returned. For further information, see
                    {@link https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param id - The ID of the slide.
         * @returns The slide with the unique ID.
         */
        getItemOrNullObject(id: string): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.SlideCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.SlideCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.SlideCollectionData;
    }
    /**
     * Represents a collection of slides in the presentation.
     *
     * @remarks
     * [Api set: PowerPointApi 1.5]
     */
    export class SlideScopedCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Slide[];
        /**
         * Gets the number of slides in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         * @returns The number of slides in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a slide using its unique ID.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param key - The ID of the slide.
         * @returns The slide with the unique ID. If such a slide does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Slide;
        /**
         * Gets a slide using its zero-based index in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param index - The index of the slide in the collection.
         * @returns The slide at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Slide;
        /**
         * Gets a slide using its unique ID. If such a slide does not exist, an object with an `isNullObject` property set to true is returned. For further information, see
                    {@link https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         *
         * @param id - The ID of the slide.
         * @returns The slide with the unique ID.
         */
        getItemOrNullObject(id: string): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideScopedCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.SlideScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideScopedCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.SlideScopedCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideScopedCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideScopedCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.SlideScopedCollectionData;
    }
    /**
     * Represents the collection of Slide Masters in the presentation.
     *
     * @remarks
     * [Api set: PowerPointApi 1.3]
     */
    export class SlideMasterCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.SlideMaster[];
        /**
         * Gets the number of Slide Masters in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         * @returns The number of Slide Masters in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a Slide Master using its unique ID.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param key - The ID of the Slide Master.
         * @returns The Slide Master with the unique ID. If such a Slide Master does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.SlideMaster;
        /**
         * Gets a Slide Master using its zero-based index in the collection.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param index - The index of the Slide Master in the collection.
         * @returns The Slide Master at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.SlideMaster;
        /**
         * Gets a Slide Master using its unique ID. If such a Slide Master does not exist, an object with an `isNullObject` property set to true is returned.
                    For further information, see {@link https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}."
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         *
         * @param id - The ID of the Slide Master.
         * @returns The Slide Master with the unique ID.
         */
        getItemOrNullObject(id: string): PowerPoint.SlideMaster;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideMasterCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.SlideMasterCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideMasterCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.SlideMasterCollection;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.SlideMasterCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideMasterCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
         */
        toJSON(): PowerPoint.Interfaces.SlideMasterCollectionData;
    }
    enum ErrorCodes {
        generalException = "GeneralException",
    }
    export namespace Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        export interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
        /** An interface for updating data on the `CustomXmlPartScopedCollection` object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `CustomXmlPartCollection` object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `Hyperlink` object, for use in `hyperlink.set({ ... })`. */
        export interface HyperlinkUpdateData {
            
            
        }
        /** An interface for updating data on the `HyperlinkCollection` object, for use in `hyperlinkCollection.set({ ... })`. */
        export interface HyperlinkCollectionUpdateData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            /**
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            foregroundColor?: string;
            /**
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: number;
        }
        /** An interface for updating data on the `ShapeFont` object, for use in `shapeFont.set({ ... })`. */
        export interface ShapeFontUpdateData {
            /**
             * Specifies the bold status of font. Returns `null` if the `TextRange` contains both bold and non-bold text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bold?: boolean | null;
            /**
             * Specifies the HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` contains text fragments with different colors.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: string | null;
            /**
             * Specifies the italic status of font. Returns 'null' if the 'TextRange' contains both italic and non-italic text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            italic?: boolean | null;
            /**
             * Specifies the font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name. Returns `null` if the `TextRange` contains text fragments with different font names.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: string | null;
            /**
             * Specifies the font size in points (e.g., 11). Returns `null` if the `TextRange` contains text fragments with different font sizes.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            size?: number | null;
            /**
             * Specifies the type of underline applied to the font. Returns `null` if the `TextRange` contains text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            underline?: ShapeFontUnderlineStyle | null;
        }
        /** An interface for updating data on the `ShapeCollection` object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `SlideLayoutCollection` object, for use in `slideLayoutCollection.set({ ... })`. */
        export interface SlideLayoutCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface for updating data on the `Tag` object, for use in `tag.set({ ... })`. */
        export interface TagUpdateData {
            /**
             * Gets the value of the tag.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            value?: string;
        }
        /** An interface for updating data on the `TagCollection` object, for use in `tagCollection.set({ ... })`. */
        export interface TagCollectionUpdateData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface for updating data on the `ShapeScopedCollection` object, for use in `shapeScopedCollection.set({ ... })`. */
        export interface ShapeScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `ShapeLineFormat` object, for use in `shapeLineFormat.set({ ... })`. */
        export interface ShapeLineFormatUpdateData {
            /**
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: string;
            /**
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            dashStyle?: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
            /**
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            style?: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
            /**
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: number;
            /**
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
            /**
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            weight?: number;
        }
        /** An interface for updating data on the `BulletFormat` object, for use in `bulletFormat.set({ ... })`. */
        export interface BulletFormatUpdateData {
            /**
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the `ParagraphFormat` object, for use in `paragraphFormat.set({ ... })`. */
        export interface ParagraphFormatUpdateData {
            /**
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            horizontalAlignment?: PowerPoint.ParagraphHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
        }
        /** An interface for updating data on the `TextRange` object, for use in `textRange.set({ ... })`. */
        export interface TextRangeUpdateData {
            /**
             * Gets or sets the length of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the available text from the starting point.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            length?: number;
            /**
             * Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the text.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            start?: number;
            /**
             * Represents the plain text content of the text range.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            text?: string;
        }
        /** An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            /**
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            autoSizeSetting?: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
            /**
             * Represents the bottom margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bottomMargin?: number;
            /**
             * Represents the left margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            leftMargin?: number;
            /**
             * Represents the right margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            rightMargin?: number;
            /**
             * Represents the top margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            topMargin?: number;
            /**
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            verticalAlignment?: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
            /**
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            wordWrap?: boolean;
        }
        /** An interface for updating data on the `Shape` object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            /**
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            height?: number;
            /**
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            left?: number;
            /**
             * Specifies the name of this shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: string;
            /**
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            top?: number;
            /**
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            width?: number;
        }
        /** An interface for updating data on the `CustomProperty` object, for use in `customProperty.set({ ... })`. */
        export interface CustomPropertyUpdateData {
            
        }
        /** An interface for updating data on the `CustomPropertyCollection` object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the `DocumentProperties` object, for use in `documentProperties.set({ ... })`. */
        export interface DocumentPropertiesUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `SlideCollection` object, for use in `slideCollection.set({ ... })`. */
        export interface SlideCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface for updating data on the `SlideScopedCollection` object, for use in `slideScopedCollection.set({ ... })`. */
        export interface SlideScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface for updating data on the `SlideMasterCollection` object, for use in `slideMasterCollection.set({ ... })`. */
        export interface SlideMasterCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /** An interface describing the data returned by calling `presentation.toJSON()`. */
        export interface PresentationData {
            /**
             * Gets the ID of the presentation.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            id?: string;
            title?: string;
        }
        /** An interface describing the data returned by calling `customXmlPart.toJSON()`. */
        export interface CustomXmlPartData {
            
            
        }
        /** An interface describing the data returned by calling `customXmlPartScopedCollection.toJSON()`. */
        export interface CustomXmlPartScopedCollectionData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `customXmlPartCollection.toJSON()`. */
        export interface CustomXmlPartCollectionData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `hyperlink.toJSON()`. */
        export interface HyperlinkData {
            
            
        }
        /** An interface describing the data returned by calling `hyperlinkCollection.toJSON()`. */
        export interface HyperlinkCollectionData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            /**
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            foregroundColor?: string;
            /**
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: number;
            /**
             * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: PowerPoint.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "PictureAndTexture" | "SlideBackground";
        }
        /** An interface describing the data returned by calling `shapeFont.toJSON()`. */
        export interface ShapeFontData {
            /**
             * Specifies the bold status of font. Returns `null` if the `TextRange` contains both bold and non-bold text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bold?: boolean | null;
            /**
             * Specifies the HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` contains text fragments with different colors.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: string | null;
            /**
             * Specifies the italic status of font. Returns 'null' if the 'TextRange' contains both italic and non-italic text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            italic?: boolean | null;
            /**
             * Specifies the font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name. Returns `null` if the `TextRange` contains text fragments with different font names.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: string | null;
            /**
             * Specifies the font size in points (e.g., 11). Returns `null` if the `TextRange` contains text fragments with different font sizes.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            size?: number | null;
            /**
             * Specifies the type of underline applied to the font. Returns `null` if the `TextRange` contains text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            underline?: ShapeFontUnderlineStyle | null;
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `slideLayout.toJSON()`. */
        export interface SlideLayoutData {
            /**
             * Gets the unique ID of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: string;
            /**
             * Gets the name of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling `slideLayoutCollection.toJSON()`. */
        export interface SlideLayoutCollectionData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface describing the data returned by calling `slideMaster.toJSON()`. */
        export interface SlideMasterData {
            /**
             * Gets the unique ID of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: string;
            /**
             * Gets the unique name of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling `tag.toJSON()`. */
        export interface TagData {
            /**
             * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            key?: string;
            /**
             * Gets the value of the tag.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            value?: string;
        }
        /** An interface describing the data returned by calling `tagCollection.toJSON()`. */
        export interface TagCollectionData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface describing the data returned by calling `slide.toJSON()`. */
        export interface SlideData {
            /**
             * Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling `shapeScopedCollection.toJSON()`. */
        export interface ShapeScopedCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `shapeLineFormat.toJSON()`. */
        export interface ShapeLineFormatData {
            /**
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: string;
            /**
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            dashStyle?: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
            /**
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            style?: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
            /**
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: number;
            /**
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
            /**
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            weight?: number;
        }
        /** An interface describing the data returned by calling `bulletFormat.toJSON()`. */
        export interface BulletFormatData {
            /**
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `paragraphFormat.toJSON()`. */
        export interface ParagraphFormatData {
            /**
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            horizontalAlignment?: PowerPoint.ParagraphHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
        }
        /** An interface describing the data returned by calling `textRange.toJSON()`. */
        export interface TextRangeData {
            /**
             * Gets or sets the length of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the available text from the starting point.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            length?: number;
            /**
             * Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the text.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            start?: number;
            /**
             * Represents the plain text content of the text range.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            /**
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            autoSizeSetting?: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
            /**
             * Represents the bottom margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bottomMargin?: number;
            /**
             * Specifies if the text frame contains text.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            hasText?: boolean;
            /**
             * Represents the left margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            leftMargin?: number;
            /**
             * Represents the right margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            rightMargin?: number;
            /**
             * Represents the top margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            topMargin?: number;
            /**
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            verticalAlignment?: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
            /**
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            wordWrap?: boolean;
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            /**
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            height?: number;
            /**
             * Gets the unique ID of the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: string;
            /**
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            left?: number;
            /**
             * Specifies the name of this shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: string;
            /**
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            top?: number;
            /**
             * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: PowerPoint.ShapeType | "Unsupported" | "Image" | "GeometricShape" | "Group" | "Line" | "Table" | "Callout" | "Chart" | "ContentApp" | "Diagram" | "Freeform" | "Graphic" | "Ink" | "Media" | "Model3D" | "Ole" | "Placeholder" | "SmartArt" | "TextBox";
            /**
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `customProperty.toJSON()`. */
        export interface CustomPropertyData {
            
            
            
        }
        /** An interface describing the data returned by calling `customPropertyCollection.toJSON()`. */
        export interface CustomPropertyCollectionData {
            items?: PowerPoint.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling `documentProperties.toJSON()`. */
        export interface DocumentPropertiesData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `slideCollection.toJSON()`. */
        export interface SlideCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `slideScopedCollection.toJSON()`. */
        export interface SlideScopedCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `slideMasterCollection.toJSON()`. */
        export interface SlideMasterCollectionData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /**
         * @remarks
         * [Api set: PowerPointApi 1.0]
         */
        export interface PresentationLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
             * Gets the ID of the presentation.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            id?: boolean;
            title?: boolean;
        }
        
        
        
        
        
        /**
         * Represents the fill formatting of a shape object.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface ShapeFillLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            foregroundColor?: boolean;
            /**
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: boolean;
            /**
             * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: boolean;
        }
        /**
         * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface ShapeFontLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Specifies the bold status of font. Returns `null` if the `TextRange` contains both bold and non-bold text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bold?: boolean;
            /**
             * Specifies the HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` contains text fragments with different colors.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: boolean;
            /**
             * Specifies the italic status of font. Returns 'null' if the 'TextRange' contains both italic and non-italic text fragments.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            italic?: boolean;
            /**
             * Specifies the font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name. Returns `null` if the `TextRange` contains text fragments with different font names.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: boolean;
            /**
             * Specifies the font size in points (e.g., 11). Returns `null` if the `TextRange` contains text fragments with different font sizes.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            size?: boolean;
            /**
             * Specifies the type of underline applied to the font. Returns `null` if the `TextRange` contains text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            underline?: boolean;
        }
        /**
         * Represents the collection of shapes.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface ShapeCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Returns the fill formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            fill?: PowerPoint.Interfaces.ShapeFillLoadOptions;
            /**
            * For EACH ITEM in the collection: Returns the line formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            lineFormat?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions;
            /**
            * For EACH ITEM in the collection: Returns the text frame object of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            textFrame?: PowerPoint.Interfaces.TextFrameLoadOptions;
            /**
             * For EACH ITEM in the collection: Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            height?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            left?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the name of this shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: boolean;
            /**
             * For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            top?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            width?: boolean;
        }
        /**
         * Represents the layout of a slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface SlideLayoutLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the unique ID of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * Gets the name of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: boolean;
        }
        /**
         * Represents the collection of layouts provided by the Slide Master for slides.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface SlideLayoutCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the name of the slide layout.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: boolean;
        }
        /**
         * Represents the Slide Master of a slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface SlideMasterLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the unique ID of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * Gets the unique name of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: boolean;
        }
        /**
         * Represents a single tag in the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface TagLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            key?: boolean;
            /**
             * Gets the value of the tag.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            value?: boolean;
        }
        /**
         * Represents the collection of tags.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface TagCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            key?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the value of the tag.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            value?: boolean;
        }
        /**
         * Represents a single slide of a presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the layout of the slide.
            *
            * @remarks
            * [Api set: PowerPointApi 1.3]
            */
            layout?: PowerPoint.Interfaces.SlideLayoutLoadOptions;
            /**
            * Gets the `SlideMaster` object that represents the slide's default content.
            *
            * @remarks
            * [Api set: PowerPointApi 1.3]
            */
            slideMaster?: PowerPoint.Interfaces.SlideMasterLoadOptions;
            /**
             * Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         * Represents a collection of shapes.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        export interface ShapeScopedCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Returns the fill formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.5]
            */
            fill?: PowerPoint.Interfaces.ShapeFillLoadOptions;
            /**
            * For EACH ITEM in the collection: Returns the line formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.5]
            */
            lineFormat?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions;
            /**
            * For EACH ITEM in the collection: Returns the text frame object of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.5]
            */
            textFrame?: PowerPoint.Interfaces.TextFrameLoadOptions;
            /**
             * For EACH ITEM in the collection: Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            height?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            left?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the name of this shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: boolean;
            /**
             * For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            top?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            width?: boolean;
        }
        /**
         * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface ShapeLineFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            color?: boolean;
            /**
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            dashStyle?: boolean;
            /**
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            style?: boolean;
            /**
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            transparency?: boolean;
            /**
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
            /**
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            weight?: boolean;
        }
        /**
         * Represents the bullet formatting properties of a text that is attached to the {@link PowerPoint.ParagraphFormat}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface BulletFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            visible?: boolean;
        }
        /**
         * Represents the paragraph formatting properties of a text that is attached to the {@link PowerPoint.TextRange}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface ParagraphFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the bullet format of the paragraph. See {@link PowerPoint.BulletFormat} for details.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            bulletFormat?: PowerPoint.Interfaces.BulletFormatLoadOptions;
            /**
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            horizontalAlignment?: boolean;
        }
        /**
         * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface TextRangeLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Returns a `ShapeFont` object that represents the font attributes for the text range.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            font?: PowerPoint.Interfaces.ShapeFontLoadOptions;
            /**
            * Represents the paragraph format of the text range. See {@link PowerPoint.ParagraphFormat} for details.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            paragraphFormat?: PowerPoint.Interfaces.ParagraphFormatLoadOptions;
            /**
             * Gets or sets the length of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the available text from the starting point.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            length?: boolean;
            /**
             * Gets or sets zero-based index, relative to the parent text frame, for the starting position of the range that this `TextRange` represents.
                        Throws an `InvalidArgument` exception when set with a negative value or if the value is greater than the length of the text.
             *
             * @remarks
             * [Api set: PowerPointApi 1.5]
             */
            start?: boolean;
            /**
             * Represents the plain text content of the text range.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            text?: boolean;
        }
        /**
         * Represents the text frame of a shape object.
         *
         * @remarks
         * [Api set: PowerPointApi 1.4]
         */
        export interface TextFrameLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See {@link PowerPoint.TextRange} for details.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            textRange?: PowerPoint.Interfaces.TextRangeLoadOptions;
            /**
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            autoSizeSetting?: boolean;
            /**
             * Represents the bottom margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            bottomMargin?: boolean;
            /**
             * Specifies if the text frame contains text.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            hasText?: boolean;
            /**
             * Represents the left margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            leftMargin?: boolean;
            /**
             * Represents the right margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            rightMargin?: boolean;
            /**
             * Represents the top margin, in points, of the text frame.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            topMargin?: boolean;
            /**
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            verticalAlignment?: boolean;
            /**
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            wordWrap?: boolean;
        }
        /**
         * Represents a single shape in the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface ShapeLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Returns the fill formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            fill?: PowerPoint.Interfaces.ShapeFillLoadOptions;
            /**
            * Returns the line formatting of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            lineFormat?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions;
            /**
            * Returns the text frame object of this shape.
            *
            * @remarks
            * [Api set: PowerPointApi 1.4]
            */
            textFrame?: PowerPoint.Interfaces.TextFrameLoadOptions;
            /**
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            height?: boolean;
            /**
             * Gets the unique ID of the shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            left?: boolean;
            /**
             * Specifies the name of this shape.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            name?: boolean;
            /**
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            top?: boolean;
            /**
             * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            type?: boolean;
            /**
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * @remarks
             * [Api set: PowerPointApi 1.4]
             */
            width?: boolean;
        }
        
        
        
        /**
         * Represents the collection of slides in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the layout of the slide.
            *
            * @remarks
            * [Api set: PowerPointApi 1.3]
            */
            layout?: PowerPoint.Interfaces.SlideLayoutLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the `SlideMaster` object that represents the slide's default content.
            *
            * @remarks
            * [Api set: PowerPointApi 1.3]
            */
            slideMaster?: PowerPoint.Interfaces.SlideMasterLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         * Represents a collection of slides in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.5]
         */
        export interface SlideScopedCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the layout of the slide.
            *
            * @remarks
            * [Api set: PowerPointApi 1.5]
            */
            layout?: PowerPoint.Interfaces.SlideLayoutLoadOptions;
            /**
            * For EACH ITEM in the collection: Gets the `SlideMaster` object that represents the slide's default content.
            *
            * @remarks
            * [Api set: PowerPointApi 1.5]
            */
            slideMaster?: PowerPoint.Interfaces.SlideMasterLoadOptions;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         * Represents the collection of Slide Masters in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.3]
         */
        export interface SlideMasterCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the unique name of the Slide Master.
             *
             * @remarks
             * [Api set: PowerPointApi 1.3]
             */
            name?: boolean;
        }
    }
}
export declare namespace PowerPoint {
    /**
     * The RequestContext object facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the request context is required to get access to the PowerPoint object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string);
        readonly presentation: Presentation;
        readonly application: Application;
    }
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}
export declare namespace PowerPoint {
    /**
     * Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a Base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The Base64-encoded .pptx file. The default value is null.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}


////////////////////////////////////////////////////////////////
///////////////////// End PowerPoint APIs //////////////////////
////////////////////////////////////////////////////////////////