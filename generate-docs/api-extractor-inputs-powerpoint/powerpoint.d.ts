import { OfficeExtension } from "../api-extractor-inputs-office/office"
import { Office as Outlook} from "../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * [Api set: PowerPointApi 1.0]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Create a new instance of PowerPoint.Application object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): PowerPoint.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     * [Api set: PowerPointApi 1.0]
     */
    export class Presentation extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns the collection of `SlideMaster` objects that are in the presentation.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly slideMasters: PowerPoint.SlideMasterCollection;
        /**
         *
         * Returns an ordered collection of slides in the presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        readonly slides: PowerPoint.SlideCollection;
        /**
         *
         * Returns a collection of tags attached to the presentation.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly tags: PowerPoint.TagCollection;
        readonly title: string;
        /**
         * Inserts the specified slides from a presentation into the current presentation.
         *
         * [Api set: PowerPointApi 1.2]
         *
         * @param base64File - The base64-encoded string representing the source presentation file.
         * @param options - The options that define which slides will be inserted, where the new slides will go, and which presentation's formatting will be used.
         */
        insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions): void;
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Presentation object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.PresentationData;
    }
    /**
     *
     * Represents the available options when adding a new slide.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export interface AddSlideOptions {
        /**
         *
         * Specifies the ID of a Slide Layout to be used for the new slide.
                    If no `layoutId` is provided, but a `slideMasterId` is provided, then the ID of the first layout from the specified Slide Master will be used.
                    If no `slideMasterId` is provided, but a `layoutId` is provided, then the specified layout needs to be available for the default Slide Master (as specified
                    in the `slideMasterId` description). Otherwise, an error will be thrown.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        layoutId?: string;
        /**
         *
         * Specifies the ID of a Slide Master to be used for the new slide.
                    If no `slideMasterId` is provided, then the previous slide's Slide Master will be used.
                    If there is no previous slide, then the presentation's first Slide Master will be used.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        slideMasterId?: string;
    }
    /**
     *
     * Represents the bullet formatting properties of a text that is attached to the {@link PowerPoint.ParagraphFormat}.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class BulletFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.BulletFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.BulletFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.BulletFormatData;
    }
    /**
     *
     * Specifies the connector type for line shapes.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ConnectorType {
        /**
         * Straight connector type
         *
         */
        straight = "Straight",
        /**
         * Elbow connector type
         *
         */
        elbow = "Elbow",
        /**
         * Curve connector type
         *
         */
        curve = "Curve",
    }
    /**
     *
     * Specifies the shape type for a `GeometricShape` object.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum GeometricShapeType {
        /**
         * Straight Line from Top-Right Corner to Bottom-Left Corner of the Shape
         *
         */
        lineInverse = "LineInverse",
        /**
         * Isosceles Triangle
         *
         */
        triangle = "Triangle",
        /**
         * Right Triangle
         *
         */
        rightTriangle = "RightTriangle",
        /**
         * Rectangle
         *
         */
        rectangle = "Rectangle",
        /**
         * Diamond
         *
         */
        diamond = "Diamond",
        /**
         * Parallelogram
         *
         */
        parallelogram = "Parallelogram",
        /**
         * Trapezoid
         *
         */
        trapezoid = "Trapezoid",
        /**
         * Trapezoid which may have Non-Equal Sides
         *
         */
        nonIsoscelesTrapezoid = "NonIsoscelesTrapezoid",
        /**
         * Pentagon
         *
         */
        pentagon = "Pentagon",
        /**
         * Hexagon
         *
         */
        hexagon = "Hexagon",
        /**
         * Heptagon
         *
         */
        heptagon = "Heptagon",
        /**
         * Octagon
         *
         */
        octagon = "Octagon",
        /**
         * Decagon
         *
         */
        decagon = "Decagon",
        /**
         * Dodecagon
         *
         */
        dodecagon = "Dodecagon",
        /**
         * Star: 4 Points
         *
         */
        star4 = "Star4",
        /**
         * Star: 5 Points
         *
         */
        star5 = "Star5",
        /**
         * Star: 6 Points
         *
         */
        star6 = "Star6",
        /**
         * Star: 7 Points
         *
         */
        star7 = "Star7",
        /**
         * Star: 8 Points
         *
         */
        star8 = "Star8",
        /**
         * Star: 10 Points
         *
         */
        star10 = "Star10",
        /**
         * Star: 12 Points
         *
         */
        star12 = "Star12",
        /**
         * Star: 16 Points
         *
         */
        star16 = "Star16",
        /**
         * Star: 24 Points
         *
         */
        star24 = "Star24",
        /**
         * Star: 32 Points
         *
         */
        star32 = "Star32",
        /**
         * Rectangle: Rounded Corners
         *
         */
        roundRectangle = "RoundRectangle",
        /**
         * Rectangle: Single Corner Rounded
         *
         */
        round1Rectangle = "Round1Rectangle",
        /**
         * Rectangle: Top Corners Rounded
         *
         */
        round2SameRectangle = "Round2SameRectangle",
        /**
         * Rectangle: Diagonal Corners Rounded
         *
         */
        round2DiagonalRectangle = "Round2DiagonalRectangle",
        /**
         * Rectangle: Top Corners One Rounded and One Snipped
         *
         */
        snipRoundRectangle = "SnipRoundRectangle",
        /**
         * Rectangle: Single Corner Snipped
         *
         */
        snip1Rectangle = "Snip1Rectangle",
        /**
         * Rectangle: Top Corners Snipped
         *
         */
        snip2SameRectangle = "Snip2SameRectangle",
        /**
         * Rectangle: Diagonal Corners Snipped
         *
         */
        snip2DiagonalRectangle = "Snip2DiagonalRectangle",
        /**
         * Plaque
         *
         */
        plaque = "Plaque",
        /**
         * Oval
         *
         */
        ellipse = "Ellipse",
        /**
         * Teardrop
         *
         */
        teardrop = "Teardrop",
        /**
         * Arrow: Pentagon
         *
         */
        homePlate = "HomePlate",
        /**
         * Arrow: Chevron
         *
         */
        chevron = "Chevron",
        /**
         * Partial Circle
         *
         */
        pieWedge = "PieWedge",
        /**
         * Partial Circle with Adjustable Spanning Area
         *
         */
        pie = "Pie",
        /**
         * Block Arc
         *
         */
        blockArc = "BlockArc",
        /**
         * Circle: Hollow
         *
         */
        donut = "Donut",
        /**
         * "Not Allowed" Symbol
         *
         */
        noSmoking = "NoSmoking",
        /**
         * Arrow: Right
         *
         */
        rightArrow = "RightArrow",
        /**
         * Arrow: Left
         *
         */
        leftArrow = "LeftArrow",
        /**
         * Arrow: Up
         *
         */
        upArrow = "UpArrow",
        /**
         * Arrow: Down
         *
         */
        downArrow = "DownArrow",
        /**
         * Arrow: Striped Right
         *
         */
        stripedRightArrow = "StripedRightArrow",
        /**
         * Arrow: Notched Right
         *
         */
        notchedRightArrow = "NotchedRightArrow",
        /**
         * Arrow: Bent-Up
         *
         */
        bentUpArrow = "BentUpArrow",
        /**
         * Arrow: Left-Right
         *
         */
        leftRightArrow = "LeftRightArrow",
        /**
         * Arrow: Up-Down
         *
         */
        upDownArrow = "UpDownArrow",
        /**
         * Arrow: Left-Up
         *
         */
        leftUpArrow = "LeftUpArrow",
        /**
         * Arrow: Left-Right-Up
         *
         */
        leftRightUpArrow = "LeftRightUpArrow",
        /**
         * Arrow: Quad
         *
         */
        quadArrow = "QuadArrow",
        /**
         * Callout: Left Arrow
         *
         */
        leftArrowCallout = "LeftArrowCallout",
        /**
         * Callout: Right Arrow
         *
         */
        rightArrowCallout = "RightArrowCallout",
        /**
         * Callout: Up Arrow
         *
         */
        upArrowCallout = "UpArrowCallout",
        /**
         * Callout: Down Arrow
         *
         */
        downArrowCallout = "DownArrowCallout",
        /**
         * Callout: Left-Right Arrow
         *
         */
        leftRightArrowCallout = "LeftRightArrowCallout",
        /**
         * Callout: Up-Down Arrow
         *
         */
        upDownArrowCallout = "UpDownArrowCallout",
        /**
         * Callout: Quad Arrow
         *
         */
        quadArrowCallout = "QuadArrowCallout",
        /**
         * Arrow: Bent
         *
         */
        bentArrow = "BentArrow",
        /**
         * Arrow: U-Turn
         *
         */
        uturnArrow = "UturnArrow",
        /**
         * Arrow: Circular
         *
         */
        circularArrow = "CircularArrow",
        /**
         * Arrow: Circular with Opposite Arrow Direction
         *
         */
        leftCircularArrow = "LeftCircularArrow",
        /**
         * Arrow: Circular with Two Arrows in Both Directions
         *
         */
        leftRightCircularArrow = "LeftRightCircularArrow",
        /**
         * Arrow: Curved Right
         *
         */
        curvedRightArrow = "CurvedRightArrow",
        /**
         * Arrow: Curved Left
         *
         */
        curvedLeftArrow = "CurvedLeftArrow",
        /**
         * Arrow: Curved Up
         *
         */
        curvedUpArrow = "CurvedUpArrow",
        /**
         * Arrow: Curved Down
         *
         */
        curvedDownArrow = "CurvedDownArrow",
        /**
         * Arrow: Curved Right Arrow with Varying Width
         *
         */
        swooshArrow = "SwooshArrow",
        /**
         * Cube
         *
         */
        cube = "Cube",
        /**
         * Cylinder
         *
         */
        can = "Can",
        /**
         * Lightning Bolt
         *
         */
        lightningBolt = "LightningBolt",
        /**
         * Heart
         *
         */
        heart = "Heart",
        /**
         * Sun
         *
         */
        sun = "Sun",
        /**
         * Moon
         *
         */
        moon = "Moon",
        /**
         * Smiley Face
         *
         */
        smileyFace = "SmileyFace",
        /**
         * Explosion: 8 Points
         *
         */
        irregularSeal1 = "IrregularSeal1",
        /**
         * Explosion: 14 Points
         *
         */
        irregularSeal2 = "IrregularSeal2",
        /**
         * Rectangle: Folded Corner
         *
         */
        foldedCorner = "FoldedCorner",
        /**
         * Rectangle: Beveled
         *
         */
        bevel = "Bevel",
        /**
         * Frame
         *
         */
        frame = "Frame",
        /**
         * Half Frame
         *
         */
        halfFrame = "HalfFrame",
        /**
         * L-Shape
         *
         */
        corner = "Corner",
        /**
         * Diagonal Stripe
         *
         */
        diagonalStripe = "DiagonalStripe",
        /**
         * Chord
         *
         */
        chord = "Chord",
        /**
         * Arc
         *
         */
        arc = "Arc",
        /**
         * Left Bracket
         *
         */
        leftBracket = "LeftBracket",
        /**
         * Right Bracket
         *
         */
        rightBracket = "RightBracket",
        /**
         * Left Brace
         *
         */
        leftBrace = "LeftBrace",
        /**
         * Right Brace
         *
         */
        rightBrace = "RightBrace",
        /**
         * Double Bracket
         *
         */
        bracketPair = "BracketPair",
        /**
         * Double Brace
         *
         */
        bracePair = "BracePair",
        /**
         * Callout: Line with No Border
         *
         */
        callout1 = "Callout1",
        /**
         * Callout: Bent Line with No Border
         *
         */
        callout2 = "Callout2",
        /**
         * Callout: Double Bent Line with No Border
         *
         */
        callout3 = "Callout3",
        /**
         * Callout: Line with Accent Bar
         *
         */
        accentCallout1 = "AccentCallout1",
        /**
         * Callout: Bent Line with Accent Bar
         *
         */
        accentCallout2 = "AccentCallout2",
        /**
         * Callout: Double Bent Line with Accent Bar
         *
         */
        accentCallout3 = "AccentCallout3",
        /**
         * Callout: Line
         *
         */
        borderCallout1 = "BorderCallout1",
        /**
         * Callout: Bent Line
         *
         */
        borderCallout2 = "BorderCallout2",
        /**
         * Callout: Double Bent Line
         *
         */
        borderCallout3 = "BorderCallout3",
        /**
         * Callout: Line with Border and Accent Bar
         *
         */
        accentBorderCallout1 = "AccentBorderCallout1",
        /**
         * Callout: Bent Line with Border and Accent Bar
         *
         */
        accentBorderCallout2 = "AccentBorderCallout2",
        /**
         * Callout: Double Bent Line with Border and Accent Bar
         *
         */
        accentBorderCallout3 = "AccentBorderCallout3",
        /**
         * Speech Bubble: Rectangle
         *
         */
        wedgeRectCallout = "WedgeRectCallout",
        /**
         * Speech Bubble: Rectangle with Corners Rounded
         *
         */
        wedgeRRectCallout = "WedgeRRectCallout",
        /**
         * Speech Bubble: Oval
         *
         */
        wedgeEllipseCallout = "WedgeEllipseCallout",
        /**
         * Thought Bubble: Cloud
         *
         */
        cloudCallout = "CloudCallout",
        /**
         * Cloud
         *
         */
        cloud = "Cloud",
        /**
         * Ribbon: Tilted Down
         *
         */
        ribbon = "Ribbon",
        /**
         * Ribbon: Tilted Up
         *
         */
        ribbon2 = "Ribbon2",
        /**
         * Ribbon: Curved and Tilted Down
         *
         */
        ellipseRibbon = "EllipseRibbon",
        /**
         * Ribbon: Curved and Tilted Up
         *
         */
        ellipseRibbon2 = "EllipseRibbon2",
        /**
         * Ribbon: Straight with Both Left and Right Arrows
         *
         */
        leftRightRibbon = "LeftRightRibbon",
        /**
         * Scroll: Vertical
         *
         */
        verticalScroll = "VerticalScroll",
        /**
         * Scroll: Horizontal
         *
         */
        horizontalScroll = "HorizontalScroll",
        /**
         * Wave
         *
         */
        wave = "Wave",
        /**
         * Double Wave
         *
         */
        doubleWave = "DoubleWave",
        /**
         * Cross
         *
         */
        plus = "Plus",
        /**
         * Flowchart: Process
         *
         */
        flowChartProcess = "FlowChartProcess",
        /**
         * Flowchart: Decision
         *
         */
        flowChartDecision = "FlowChartDecision",
        /**
         * Flowchart: Data
         *
         */
        flowChartInputOutput = "FlowChartInputOutput",
        /**
         * Flowchart: Predefined Process
         *
         */
        flowChartPredefinedProcess = "FlowChartPredefinedProcess",
        /**
         * Flowchart: Internal Storage
         *
         */
        flowChartInternalStorage = "FlowChartInternalStorage",
        /**
         * Flowchart: Document
         *
         */
        flowChartDocument = "FlowChartDocument",
        /**
         * Flowchart: Multidocument
         *
         */
        flowChartMultidocument = "FlowChartMultidocument",
        /**
         * Flowchart: Terminator
         *
         */
        flowChartTerminator = "FlowChartTerminator",
        /**
         * Flowchart: Preparation
         *
         */
        flowChartPreparation = "FlowChartPreparation",
        /**
         * Flowchart: Manual Input
         *
         */
        flowChartManualInput = "FlowChartManualInput",
        /**
         * Flowchart: Manual Operation
         *
         */
        flowChartManualOperation = "FlowChartManualOperation",
        /**
         * Flowchart: Connector
         *
         */
        flowChartConnector = "FlowChartConnector",
        /**
         * Flowchart: Card
         *
         */
        flowChartPunchedCard = "FlowChartPunchedCard",
        /**
         * Flowchart: Punched Tape
         *
         */
        flowChartPunchedTape = "FlowChartPunchedTape",
        /**
         * Flowchart: Summing Junction
         *
         */
        flowChartSummingJunction = "FlowChartSummingJunction",
        /**
         * Flowchart: Or
         *
         */
        flowChartOr = "FlowChartOr",
        /**
         * Flowchart: Collate
         *
         */
        flowChartCollate = "FlowChartCollate",
        /**
         * Flowchart: Sort
         *
         */
        flowChartSort = "FlowChartSort",
        /**
         * Flowchart: Extract
         *
         */
        flowChartExtract = "FlowChartExtract",
        /**
         * Flowchart: Merge
         *
         */
        flowChartMerge = "FlowChartMerge",
        /**
         * FlowChart: Offline Storage
         *
         */
        flowChartOfflineStorage = "FlowChartOfflineStorage",
        /**
         * Flowchart: Stored Data
         *
         */
        flowChartOnlineStorage = "FlowChartOnlineStorage",
        /**
         * Flowchart: Sequential Access Storage
         *
         */
        flowChartMagneticTape = "FlowChartMagneticTape",
        /**
         * Flowchart: Magnetic Disk
         *
         */
        flowChartMagneticDisk = "FlowChartMagneticDisk",
        /**
         * Flowchart: Direct Access Storage
         *
         */
        flowChartMagneticDrum = "FlowChartMagneticDrum",
        /**
         * Flowchart: Display
         *
         */
        flowChartDisplay = "FlowChartDisplay",
        /**
         * Flowchart: Delay
         *
         */
        flowChartDelay = "FlowChartDelay",
        /**
         * Flowchart: Alternate Process
         *
         */
        flowChartAlternateProcess = "FlowChartAlternateProcess",
        /**
         * Flowchart: Off-page Connector
         *
         */
        flowChartOffpageConnector = "FlowChartOffpageConnector",
        /**
         * Action Button: Blank
         *
         */
        actionButtonBlank = "ActionButtonBlank",
        /**
         * Action Button: Go Home
         *
         */
        actionButtonHome = "ActionButtonHome",
        /**
         * Action Button: Help
         *
         */
        actionButtonHelp = "ActionButtonHelp",
        /**
         * Action Button: Get Information
         *
         */
        actionButtonInformation = "ActionButtonInformation",
        /**
         * Action Button: Go Forward or Next
         *
         */
        actionButtonForwardNext = "ActionButtonForwardNext",
        /**
         * Action Button: Go Back or Previous
         *
         */
        actionButtonBackPrevious = "ActionButtonBackPrevious",
        /**
         * Action Button: Go to End
         *
         */
        actionButtonEnd = "ActionButtonEnd",
        /**
         * Action Button: Go to Beginning
         *
         */
        actionButtonBeginning = "ActionButtonBeginning",
        /**
         * Action Button: Return
         *
         */
        actionButtonReturn = "ActionButtonReturn",
        /**
         * Action Button: Document
         *
         */
        actionButtonDocument = "ActionButtonDocument",
        /**
         * Action Button: Sound
         *
         */
        actionButtonSound = "ActionButtonSound",
        /**
         * Action Button: Video
         *
         */
        actionButtonMovie = "ActionButtonMovie",
        /**
         * Gear: A Gear with Six Teeth
         *
         */
        gear6 = "Gear6",
        /**
         * Gear: A Gear with Nine Teeth
         *
         */
        gear9 = "Gear9",
        /**
         * Funnel
         *
         */
        funnel = "Funnel",
        /**
         * Plus Sign
         *
         */
        mathPlus = "MathPlus",
        /**
         * Minus Sign
         *
         */
        mathMinus = "MathMinus",
        /**
         * Multiplication Sign
         *
         */
        mathMultiply = "MathMultiply",
        /**
         * Division Sign
         *
         */
        mathDivide = "MathDivide",
        /**
         * Equals
         *
         */
        mathEqual = "MathEqual",
        /**
         * Not Equal
         *
         */
        mathNotEqual = "MathNotEqual",
        /**
         * Four Right Triangles that Define a Rectangular Shape
         *
         */
        cornerTabs = "CornerTabs",
        /**
         * Four Small Squares that Define a Rectangular Shape.
         *
         */
        squareTabs = "SquareTabs",
        /**
         * Four Quarter Circles that Define a Rectangular Shape.
         *
         */
        plaqueTabs = "PlaqueTabs",
        /**
         * A Rectangle Divided into Four Parts Along Diagonal Lines.
         *
         */
        chartX = "ChartX",
        /**
         * A Rectangle Divided into Six Parts Along a Vertical Line and Diagonal Lines.
         *
         */
        chartStar = "ChartStar",
        /**
         * A Rectangle Divided Vertically and Horizontally into Four Quarters.
         *
         */
        chartPlus = "ChartPlus",
    }
    /**
     *
     * Specifies the formatting options for when slides are inserted.
     *
     * [Api set: PowerPointApi 1.2]
     */
    enum InsertSlideFormatting {
        /**
         * Copy the source theme into the target presentation and use that theme.
         *
         */
        keepSourceFormatting = "KeepSourceFormatting",
        /**
         * Use the existing theme in the target presentation.
         *
         */
        useDestinationTheme = "UseDestinationTheme",
    }
    /**
     *
     * Represents the available options when inserting slides.
     *
     * [Api set: PowerPointApi 1.2]
     */
    export interface InsertSlideOptions {
        /**
         *
         * Specifies which formatting to use during slide insertion.
                    The default option is to use "KeepSourceFormatting".
         *
         * [Api set: PowerPointApi 1.2]
         */
        formatting?: PowerPoint.InsertSlideFormatting | "KeepSourceFormatting" | "UseDestinationTheme";
        /**
         *
         * Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.
                    The order of these slides is preserved during the insertion.
                    If any of the source slides are not found, or if the IDs are invalid, the operation throws a `SlideNotFound` exception and no slides will be inserted.
                    All of the source slides will be inserted when `sourceSlideIds` is not provided (this is the default behavior).
         *
         * [Api set: PowerPointApi 1.2]
         */
        sourceSlideIds?: string[];
        /**
         *
         * Specifies where in the presentation the new slides will be inserted. The new slides will be inserted after the slide with the given slide ID.
                    If `targetSlideId` is not provided, the slides will be inserted at the beginning of the presentation.
                    If `targetSlideId` is invalid or if it is pointing to a non-existing slide, the operation throws a `SlideNotFound` exception and no slides will be inserted.
         *
         * [Api set: PowerPointApi 1.2]
         */
        targetSlideId?: string;
    }
    /**
     *
     * Represents the horizontal alignment of the {@link PowerPoint.TextFrame} in a {@link PowerPoint.Shape}.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ParagraphHorizontalAlignment {
        /**
         * Align text to the left margin.
         *
         */
        left = "Left",
        /**
         * Align text in the center.
         *
         */
        center = "Center",
        /**
         * Align text to the right margin.
         *
         */
        right = "Right",
        /**
         * Align text so that it is justified across the whole line.
         *
         */
        justify = "Justify",
        /**
         * Specifies the alignment or adjustment of kashida length in Arabic text.
         *
         */
        justifyLow = "JustifyLow",
        /**
         * Distributes the text words across an entire text line.
         *
         */
        distributed = "Distributed",
        /**
         * Distributes Thai text specially, because each character is treated as a word.
         *
         */
        thaiDistributed = "ThaiDistributed",
    }
    /**
     *
     * Represents the paragraph formatting properties of a text that is attached to the {@link PowerPoint.TextRange}.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class ParagraphFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents the bullet format of the paragraph. See {@link PowerPoint.BulletFormat} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly bulletFormat: PowerPoint.BulletFormat;
        /**
         *
         * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.ParagraphFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ParagraphFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ParagraphFormatData;
    }
    /**
     *
     * Specifies a shape's fill type.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeFillType {
        /**
         * Specifies that the shape should have no fill.
         *
         */
        noFill = "NoFill",
        /**
         * Specifies that the shape should have regular solid fill.
         *
         */
        solid = "Solid",
        /**
         * Specifies that the shape should have gradient fill.
         *
         */
        gradient = "Gradient",
        /**
         * Specifies that the shape should have pattern fill.
         *
         */
        pattern = "Pattern",
        /**
         * Specifies that the shape should have picture or texture fill.
         *
         */
        pictureAndTexture = "PictureAndTexture",
        /**
         * Specifies that the shape should have slide background fill.
         *
         */
        slideBackground = "SlideBackground",
    }
    /**
     *
     * Represents the fill formatting of a shape object.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class ShapeFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        foregroundColor: string;
        /**
         *
         * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        transparency: number;
        /**
         *
         * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly type: PowerPoint.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "PictureAndTexture" | "SlideBackground";
        /**
         * Clears the fill formatting of this shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        clear(): void;
        /**
         * Sets the fill formatting of the shape to a uniform color. This changes the fill type to `Solid`.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.ShapeFill object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeFillData;
    }
    /**
     *
     * Specifies the style for a line.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeLineStyle {
        /**
         * Single line.
         *
         */
        single = "Single",
        /**
         * Thick line with a thin line on each side.
         *
         */
        thickBetweenThin = "ThickBetweenThin",
        /**
         * Thick line next to thin line. For horizontal lines, the thick line is above the thin line. For vertical lines, the thick line is to the left of the thin line.
         *
         */
        thickThin = "ThickThin",
        /**
         * Thick line next to thin line. For horizontal lines, the thick line is below the thin line. For vertical lines, the thick line is to the right of the thin line.
         *
         */
        thinThick = "ThinThick",
        /**
         * Two thin lines.
         *
         */
        thinThin = "ThinThin",
    }
    /**
     *
     * Specifies the dash style for a line.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeLineDashStyle {
        /**
         * The dash line pattern
         *
         */
        dash = "Dash",
        /**
         * The dash-dot line pattern
         *
         */
        dashDot = "DashDot",
        /**
         * The dash-dot-dot line pattern
         *
         */
        dashDotDot = "DashDotDot",
        /**
         * The long dash line pattern
         *
         */
        longDash = "LongDash",
        /**
         * The long dash-dot line pattern
         *
         */
        longDashDot = "LongDashDot",
        /**
         * The round dot line pattern
         *
         */
        roundDot = "RoundDot",
        /**
         * The solid line pattern
         *
         */
        solid = "Solid",
        /**
         * The square dot line pattern
         *
         */
        squareDot = "SquareDot",
        /**
         * The long dash-dot-dot line pattern
         *
         */
        longDashDotDot = "LongDashDotDot",
        /**
         * The system dash line pattern
         *
         */
        systemDash = "SystemDash",
        /**
         * The system dot line pattern
         *
         */
        systemDot = "SystemDot",
        /**
         * The system dash-dot line pattern
         *
         */
        systemDashDot = "SystemDashDot",
    }
    /**
     *
     * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class ShapeLineFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        color: string;
        /**
         *
         * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        dashStyle: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
        /**
         *
         * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        style: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
        /**
         *
         * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        transparency: number;
        /**
         *
         * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        visible: boolean;
        /**
         *
         * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.ShapeLineFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeLineFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeLineFormatData;
    }
    /**
     *
     * Specifies the type of a shape.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeType {
        /**
         * The given shape's type is unsupported.
         *
         */
        unsupported = "Unsupported",
        /**
         * The shape is an image
         *
         */
        image = "Image",
        /**
         * The shape is a geometric shape such as rectangle
         *
         */
        geometricShape = "GeometricShape",
        /**
         * The shape is a group shape which contains sub-shapes
         *
         */
        group = "Group",
        /**
         * The shape is a line
         *
         */
        line = "Line",
    }
    /**
     *
     * Represents a single tag in the slide.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class Tag extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly key: string;
        /**
         *
         * Gets the value of the tag.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Tag object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TagData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.TagData;
    }
    /**
     *
     * Represents the collection of tags.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class TagCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Tag[];
        /**
         * Adds a new tag at the end of the collection. If the `key` already exists in the collection, the value of the existing tag will be replaced with the given `value`.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The unique ID of a tag, which is unique within this `TagCollection`. 'key' parameter is case-insensitive, but it is always capitalized when saved in the document.
         * @param value - The value of the tag.
         */
        add(key: string, value: string): void;
        /**
         * Deletes the tag with the given `key` in this collection. Does nothing if the `key` does not exist.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The unique ID of a tag, which is unique within this `TagCollection`. `key` parameter is case-insensitive.
         */
        delete(key: string): void;
        /**
         * Gets the number of tags in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         * @returns The number of tags in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a tag using its unique ID. An error is thrown if the tag does not exist.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The ID of the tag.
         * @returns The tag with the unique ID. If such a tag does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Tag;
        /**
         * Gets a tag using its zero-based index in the collection. An error is thrown if the index is out of range.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param index - The index of the tag in the collection.
         * @returns The tag at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Tag;
        /**
         * Gets a tag using its unique ID. If such a tag does not exist, an object with an `isNullObject` property set to true is returned.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.TagCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TagCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.TagCollectionData;
    }
    /**
     *
     * Determines the type of automatic sizing allowed.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeAutoSize {
        /**
         * No autosizing.
         *
         */
        autoSizeNone = "AutoSizeNone",
        /**
         * The text is adjusted to fit the shape.
         *
         */
        autoSizeTextToFitShape = "AutoSizeTextToFitShape",
        /**
         * The shape is adjusted to fit the text.
         *
         */
        autoSizeShapeToFitText = "AutoSizeShapeToFitText",
        /**
         * A combination of automatic sizing schemes are used.
         *
         */
        autoSizeMixed = "AutoSizeMixed",
    }
    /**
     *
     * Represents the vertical alignment of a {@link PowerPoint.TextFrame} in a {@link PowerPoint.Shape}.
                If one the centered options are selected, the contents of the `TextFrame` will be centered horizontally within the `Shape` as a group.
                To change the horizontal alignment of a text, see {@link PowerPoint.ParagraphFormat} and {@link PowerPoint.ParagraphHorizontalAlignment }.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum TextVerticalAlignment {
        /**
         * Specifies that the `TextFrame` should be top aligned to the `Shape`.
         *
         */
        top = "Top",
        /**
         * Specifies that the `TextFrame` should be center aligned to the `Shape`.
         *
         */
        middle = "Middle",
        /**
         * Specifies that the `TextFrame` should be bottom aligned to the `Shape`.
         *
         */
        bottom = "Bottom",
        /**
         * Specifies that the `TextFrame` should be top aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         *
         */
        topCentered = "TopCentered",
        /**
         * Specifies that the `TextFrame` should be center aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         *
         */
        middleCentered = "MiddleCentered",
        /**
         * Specifies that the `TextFrame` should be bottom aligned vertically to the `Shape`. Contents of the `TextFrame` will be centered horizontally within the `Shape`.
         *
         */
        bottomCentered = "BottomCentered",
    }
    /**
     *
     * The type of underline applied to a font.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    enum ShapeFontUnderlineStyle {
        /**
         * No underlining.
         *
         */
        none = "None",
        /**
         * Regular single line underlining.
         *
         */
        single = "Single",
        /**
         * Underlining of text with double lines.
         *
         */
        double = "Double",
        /**
         * Underlining of text with a thick line.
         *
         */
        heavy = "Heavy",
        /**
         * Underlining of text with a dotted line.
         *
         */
        dotted = "Dotted",
        /**
         * Underlining of text with a thick, dotted line.
         *
         */
        dottedHeavy = "DottedHeavy",
        /**
         * Underlining of text with a line containing dashes.
         *
         */
        dash = "Dash",
        /**
         * Underlining of text with a thick line containing dashes.
         *
         */
        dashHeavy = "DashHeavy",
        /**
         * Underlining of text with a line containing long dashes.
         *
         */
        dashLong = "DashLong",
        /**
         * Underlining of text with a thick line containing long dashes.
         *
         */
        dashLongHeavy = "DashLongHeavy",
        /**
         * Underlining of text with a line containing dots and dashes.
         *
         */
        dotDash = "DotDash",
        /**
         * Underlining of text with a thick line containing dots and dashes.
         *
         */
        dotDashHeavy = "DotDashHeavy",
        /**
         * Underlining of text with a line containing double dots and dashes.
         *
         */
        dotDotDash = "DotDotDash",
        /**
         * Underlining of text with a thick line containing double dots and dashes.
         *
         */
        dotDotDashHeavy = "DotDotDashHeavy",
        /**
         * Underlining of text with a wavy line.
         *
         */
        wavy = "Wavy",
        /**
         * Underlining of text with a thick, wavy line.
         *
         */
        wavyHeavy = "WavyHeavy",
        /**
         * Underlining of text with double wavy lines.
         *
         */
        wavyDouble = "WavyDouble",
    }
    /**
     *
     * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class ShapeFont extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        bold: boolean;
        /**
         *
         * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        color: string;
        /**
         *
         * Represents the italic status of font. Returns 'null' if the 'TextRange' includes both italic and non-italic text fragments.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        italic: boolean;
        /**
         *
         * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        name: string;
        /**
         *
         * Represents font size in points (e.g., 11). Returns null if the TextRange includes text fragments with different font sizes.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        size: number;
        /**
         *
         * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        underline: PowerPoint.ShapeFontUnderlineStyle | "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble";
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.ShapeFont object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeFontData;
    }
    /**
     *
     * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class TextRange extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns a `ShapeFont` object that represents the font attributes for the text range.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly font: PowerPoint.ShapeFont;
        /**
         *
         * Represents the paragraph format of the text range. See {@link PowerPoint.ParagraphFormat} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly paragraphFormat: PowerPoint.ParagraphFormat;
        /**
         *
         * Represents the plain text content of the text range.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        text: string;
        /**
         * Returns a `TextRange` object for the substring in the given range.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param start - The zero-based index of the first character to get from the text range.
         * @param length - Optional. The number of characters to be returned in the new text range. If length is omitted, all the characters from start to the end of the text range's last paragraph will be returned.
         */
        getSubstring(start: number, length?: number): PowerPoint.TextRange;
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.TextRange object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TextRangeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.TextRangeData;
    }
    /**
     *
     * Represents the text frame of a shape object.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class TextFrame extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See {@link PowerPoint.TextRange} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly textRange: PowerPoint.TextRange;
        /**
         *
         * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        autoSizeSetting: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
        /**
         *
         * Represents the bottom margin, in points, of the text frame.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        bottomMargin: number;
        /**
         *
         * Specifies if the text frame contains text.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly hasText: boolean;
        /**
         *
         * Represents the left margin, in points, of the text frame.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        leftMargin: number;
        /**
         *
         * Represents the right margin, in points, of the text frame.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        rightMargin: number;
        /**
         *
         * Represents the top margin, in points, of the text frame.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        topMargin: number;
        /**
         *
         * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        verticalAlignment: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
        /**
         *
         * Determines whether lines break automatically to fit text inside the shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        wordWrap: boolean;
        /**
         * Deletes all the text in the text frame.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        deleteText(): void;
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.TextFrame object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.TextFrameData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.TextFrameData;
    }
    /**
     *
     * Represents a single shape in the slide.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class Shape extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns the fill formatting of this shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly fill: PowerPoint.ShapeFill;
        /**
         *
         * Returns the line formatting of this shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly lineFormat: PowerPoint.ShapeLineFormat;
        /**
         *
         * Returns a collection of tags in the shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly tags: PowerPoint.TagCollection;
        /**
         *
         * Returns the text frame object of this shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly textFrame: PowerPoint.TextFrame;
        /**
         *
         * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        height: number;
        /**
         *
         * Gets the unique ID of the shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly id: string;
        /**
         *
         * The distance, in points, from the left side of the shape to the left side of the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        left: number;
        /**
         *
         * Specifies the name of this shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        name: string;
        /**
         *
         * The distance, in points, from the top edge of the shape to the top edge of the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        top: number;
        /**
         *
         * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly type: PowerPoint.ShapeType | "Unsupported" | "Image" | "GeometricShape" | "Group" | "Line";
        /**
         *
         * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        width: number;
        /**
         * Deletes the shape from the shape collection. Does nothing if the shape does not exist.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        delete(): void;
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Shape object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.ShapeData;
    }
    /**
     *
     * Represents the available options when adding shapes.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export interface ShapeAddOptions {
        /**
         *
         * Specifies the height, in points, of the shape.
                    When not provided, a default value will be used.
                    Throws an `InvalidArgument` exception when set with a negative value.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        height?: number;
        /**
         *
         * Specifies the distance, in points, from the left side of the shape to the left side of the slide.
                    When not provided, a default value will be used.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        left?: number;
        /**
         *
         * Specifies the distance, in points, from the top edge of the shape to the top edge of the slide.
                    When not provided, a default value will be used.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        top?: number;
        /**
         *
         * Specifies the width, in points, of the shape.
                    When not provided, a default value will be used.
                    Throws an `InvalidArgument` exception when set with a negative value.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        width?: number;
    }
    /**
     *
     * Represents the collection of shapes.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class ShapeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Shape[];
        /**
         * Adds a geometric shape to the slide. Returns a `Shape` object that represents the new shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param geometricShapeType - Specifies the type of the geometric shape. See {@link PowerPoint.GeometricShapeType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape.
         * @returns The newly inserted shape.
         */
        addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a geometric shape to the slide. Returns a `Shape` object that represents the new shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param geometricShapeTypeString - Specifies the type of the geometric shape. See {@link PowerPoint.GeometricShapeType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape.
         * @returns The newly inserted shape.
         */
        addGeometricShape(geometricShapeTypeString: "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus", options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a line to the slide. Returns a `Shape` object that represents the new line.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param connectorType - Specifies the connector type of the line. If not provided, `straight` connector type will be used. See {@link PowerPoint.ConnectorType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape object that contains the line.
         * @returns The newly inserted shape.
         */
        addLine(connectorType?: PowerPoint.ConnectorType, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a line to the slide. Returns a `Shape` object that represents the new line.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param connectorTypeString - Specifies the connector type of the line. If not provided, `straight` connector type will be used. See {@link PowerPoint.ConnectorType} for details.
         * @param options - An optional parameter to specify the additional options such as the position of the shape object that contains the line.
         * @returns The newly inserted shape.
         */
        addLine(connectorTypeString?: "Straight" | "Elbow" | "Curve", options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Adds a text box to the slide with the provided text as the content. Returns a `Shape` object that represents the new text box.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param text - Specifies the text that will be shown in the created text box.
         * @param options - An optional parameter to specify the additional options such as the position of the text box.
         * @returns The newly inserted shape.
         */
        addTextBox(text: string, options?: PowerPoint.ShapeAddOptions): PowerPoint.Shape;
        /**
         * Gets the number of shapes in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         * @returns The number of shapes in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a shape using its unique ID. An error is thrown if the shape does not exist.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The ID of the shape.
         * @returns The shape with the unique ID. If such a shape does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Shape;
        /**
         * Gets a shape using its zero-based index in the collection. An error is thrown if the index is out of range.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param index - The index of the shape in the collection.
         * @returns The shape at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Shape;
        /**
         * Gets a shape using its unique ID. If such a shape does not exist, an object with an `isNullObject` property set to true is returned.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.ShapeCollectionData;
    }
    /**
     *
     * Represents the layout of a slide.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class SlideLayout extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns a collection of shapes in the slide layout.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         *
         * Gets the unique ID of the slide layout.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly id: string;
        /**
         *
         * Gets the name of the slide layout.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.SlideLayout object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideLayoutData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.SlideLayoutData;
    }
    /**
     *
     * Represents the collection of layouts provided by the Slide Master for slides.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class SlideLayoutCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.SlideLayout[];
        /**
         * Gets the number of layouts in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         * @returns The number of layouts in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a layout using its unique ID.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The ID of the layout.
         * @returns The layout with the unique ID. If such a layout does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.SlideLayout;
        /**
         * Gets a layout using its zero-based index in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param index - The index of the layout in the collection.
         * @returns The layout at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.SlideLayout;
        /**
         * Gets a layout using its unique ID.  If such a layout does not exist, an object with an `isNullObject` property set to true is returned. For further information,
                    see {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.SlideLayoutCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideLayoutCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.SlideLayoutCollectionData;
    }
    /**
     *
     * Represents the Slide Master of a slide.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class SlideMaster extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Gets the collection of layouts provided by the Slide Master for slides.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly layouts: PowerPoint.SlideLayoutCollection;
        /**
         *
         * Returns a collection of shapes in the Slide Master.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         *
         * Gets the unique ID of the Slide Master.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly id: string;
        /**
         *
         * Gets the unique name of the Slide Master.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.SlideMaster object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideMasterData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.SlideMasterData;
    }
    /**
     *
     * Represents a single slide of a presentation.
     *
     * [Api set: PowerPointApi 1.2]
     */
    export class Slide extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Gets the layout of the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly layout: PowerPoint.SlideLayout;
        /**
         *
         * Returns a collection of shapes in the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly shapes: PowerPoint.ShapeCollection;
        /**
         *
         * Gets the `SlideMaster` object that represents the slide's default content.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly slideMaster: PowerPoint.SlideMaster;
        /**
         *
         * Returns a collection of tags in the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        readonly tags: PowerPoint.TagCollection;
        /**
         *
         * Gets the unique ID of the slide.
         *
         * [Api set: PowerPointApi 1.2]
         */
        readonly id: string;
        /**
         * Deletes the slide from the presentation. Does nothing if the slide does not exist.
         *
         * [Api set: PowerPointApi 1.2]
         */
        delete(): void;
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Slide object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.SlideData;
    }
    /**
     *
     * Represents the collection of slides in the presentation.
     *
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
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param options - The options that define the theme of the new slide.
         */
        add(options?: PowerPoint.AddSlideOptions): void;
        /**
         * Gets the number of slides in the collection.
         *
         * [Api set: PowerPointApi 1.2]
         * @returns The number of slides in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a slide using its unique ID.
         *
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
         * [Api set: PowerPointApi 1.2]
         *
         * @param index - The index of the slide in the collection.
         * @returns The slide at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Slide;
        /**
         * Gets a slide using its unique ID. If such a slide does not exist, an object with an `isNullObject` property set to true is returned. For further information,
                    see {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods
                    and properties}.
         *
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.SlideCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.SlideCollectionData;
    }
    /**
     *
     * Represents the collection of Slide Masters in the presentation.
     *
     * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
     * @beta
     */
    export class SlideMasterCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.SlideMaster[];
        /**
         * Gets the number of Slide Masters in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         * @returns The number of Slide Masters in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a Slide Master using its unique ID.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param key - The ID of the Slide Master.
         * @returns The Slide Master with the unique ID. If such a Slide Master does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.SlideMaster;
        /**
         * Gets a Slide Master using its zero-based index in the collection.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         *
         * @param index - The index of the Slide Master in the collection.
         * @returns The Slide Master at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.SlideMaster;
        /**
         * Gets a Slide Master using its unique ID. If such a Slide Master does not exist, an object with an `isNullObject` property set to true is returned.
                    For further information, see {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}."
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.SlideMasterCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideMasterCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.SlideMasterCollectionData;
    }
    enum ErrorCodes {
        generalException = "GeneralException",
    }
    export module Interfaces {
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
        /** An interface for updating data on the BulletFormat object, for use in `bulletFormat.set({ ... })`. */
        export interface BulletFormatUpdateData {
            /**
             *
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
        }
        /** An interface for updating data on the ParagraphFormat object, for use in `paragraphFormat.set({ ... })`. */
        export interface ParagraphFormatUpdateData {
            /**
             *
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            horizontalAlignment?: PowerPoint.ParagraphHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
        }
        /** An interface for updating data on the ShapeFill object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            /**
             *
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            foregroundColor?: string;
            /**
             *
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: number;
        }
        /** An interface for updating data on the ShapeLineFormat object, for use in `shapeLineFormat.set({ ... })`. */
        export interface ShapeLineFormatUpdateData {
            /**
             *
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: string;
            /**
             *
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            dashStyle?: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
            /**
             *
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            style?: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
            /**
             *
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: number;
            /**
             *
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
            /**
             *
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            weight?: number;
        }
        /** An interface for updating data on the Tag object, for use in `tag.set({ ... })`. */
        export interface TagUpdateData {
            /**
             *
             * Gets the value of the tag.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: string;
        }
        /** An interface for updating data on the TagCollection object, for use in `tagCollection.set({ ... })`. */
        export interface TagCollectionUpdateData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface for updating data on the ShapeFont object, for use in `shapeFont.set({ ... })`. */
        export interface ShapeFontUpdateData {
            /**
             *
             * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bold?: boolean;
            /**
             *
             * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: string;
            /**
             *
             * Represents the italic status of font. Returns 'null' if the 'TextRange' includes both italic and non-italic text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            italic?: boolean;
            /**
             *
             * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: string;
            /**
             *
             * Represents font size in points (e.g., 11). Returns null if the TextRange includes text fragments with different font sizes.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            size?: number;
            /**
             *
             * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            underline?: PowerPoint.ShapeFontUnderlineStyle | "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble";
        }
        /** An interface for updating data on the TextRange object, for use in `textRange.set({ ... })`. */
        export interface TextRangeUpdateData {
            /**
             *
             * Represents the plain text content of the text range.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            text?: string;
        }
        /** An interface for updating data on the TextFrame object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            /**
             *
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            autoSizeSetting?: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
            /**
             *
             * Represents the bottom margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bottomMargin?: number;
            /**
             *
             * Represents the left margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            leftMargin?: number;
            /**
             *
             * Represents the right margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            rightMargin?: number;
            /**
             *
             * Represents the top margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            topMargin?: number;
            /**
             *
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            verticalAlignment?: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
            /**
             *
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            wordWrap?: boolean;
        }
        /** An interface for updating data on the Shape object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            /**
             *
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            height?: number;
            /**
             *
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            left?: number;
            /**
             *
             * Specifies the name of this shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: string;
            /**
             *
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            top?: number;
            /**
             *
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            width?: number;
        }
        /** An interface for updating data on the ShapeCollection object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the SlideLayoutCollection object, for use in `slideLayoutCollection.set({ ... })`. */
        export interface SlideLayoutCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface for updating data on the SlideCollection object, for use in `slideCollection.set({ ... })`. */
        export interface SlideCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface for updating data on the SlideMasterCollection object, for use in `slideMasterCollection.set({ ... })`. */
        export interface SlideMasterCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /** An interface describing the data returned by calling `presentation.toJSON()`. */
        export interface PresentationData {
            title?: string;
        }
        /** An interface describing the data returned by calling `bulletFormat.toJSON()`. */
        export interface BulletFormatData {
            /**
             *
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `paragraphFormat.toJSON()`. */
        export interface ParagraphFormatData {
            /**
             *
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            horizontalAlignment?: PowerPoint.ParagraphHorizontalAlignment | "Left" | "Center" | "Right" | "Justify" | "JustifyLow" | "Distributed" | "ThaiDistributed";
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            /**
             *
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            foregroundColor?: string;
            /**
             *
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: number;
            /**
             *
             * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            type?: PowerPoint.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "PictureAndTexture" | "SlideBackground";
        }
        /** An interface describing the data returned by calling `shapeLineFormat.toJSON()`. */
        export interface ShapeLineFormatData {
            /**
             *
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: string;
            /**
             *
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            dashStyle?: PowerPoint.ShapeLineDashStyle | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "RoundDot" | "Solid" | "SquareDot" | "LongDashDotDot" | "SystemDash" | "SystemDot" | "SystemDashDot";
            /**
             *
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            style?: PowerPoint.ShapeLineStyle | "Single" | "ThickBetweenThin" | "ThickThin" | "ThinThick" | "ThinThin";
            /**
             *
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: number;
            /**
             *
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
            /**
             *
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            weight?: number;
        }
        /** An interface describing the data returned by calling `tag.toJSON()`. */
        export interface TagData {
            /**
             *
             * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: string;
            /**
             *
             * Gets the value of the tag.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: string;
        }
        /** An interface describing the data returned by calling `tagCollection.toJSON()`. */
        export interface TagCollectionData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface describing the data returned by calling `shapeFont.toJSON()`. */
        export interface ShapeFontData {
            /**
             *
             * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bold?: boolean;
            /**
             *
             * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: string;
            /**
             *
             * Represents the italic status of font. Returns 'null' if the 'TextRange' includes both italic and non-italic text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            italic?: boolean;
            /**
             *
             * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: string;
            /**
             *
             * Represents font size in points (e.g., 11). Returns null if the TextRange includes text fragments with different font sizes.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            size?: number;
            /**
             *
             * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            underline?: PowerPoint.ShapeFontUnderlineStyle | "None" | "Single" | "Double" | "Heavy" | "Dotted" | "DottedHeavy" | "Dash" | "DashHeavy" | "DashLong" | "DashLongHeavy" | "DotDash" | "DotDashHeavy" | "DotDotDash" | "DotDotDashHeavy" | "Wavy" | "WavyHeavy" | "WavyDouble";
        }
        /** An interface describing the data returned by calling `textRange.toJSON()`. */
        export interface TextRangeData {
            /**
             *
             * Represents the plain text content of the text range.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            text?: string;
        }
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            /**
             *
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            autoSizeSetting?: PowerPoint.ShapeAutoSize | "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText" | "AutoSizeMixed";
            /**
             *
             * Represents the bottom margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bottomMargin?: number;
            /**
             *
             * Specifies if the text frame contains text.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            hasText?: boolean;
            /**
             *
             * Represents the left margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            leftMargin?: number;
            /**
             *
             * Represents the right margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            rightMargin?: number;
            /**
             *
             * Represents the top margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            topMargin?: number;
            /**
             *
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            verticalAlignment?: PowerPoint.TextVerticalAlignment | "Top" | "Middle" | "Bottom" | "TopCentered" | "MiddleCentered" | "BottomCentered";
            /**
             *
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            wordWrap?: boolean;
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            /**
             *
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            height?: number;
            /**
             *
             * Gets the unique ID of the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: string;
            /**
             *
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            left?: number;
            /**
             *
             * Specifies the name of this shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: string;
            /**
             *
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            top?: number;
            /**
             *
             * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            type?: PowerPoint.ShapeType | "Unsupported" | "Image" | "GeometricShape" | "Group" | "Line";
            /**
             *
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `slideLayout.toJSON()`. */
        export interface SlideLayoutData {
            /**
             *
             * Gets the unique ID of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: string;
            /**
             *
             * Gets the name of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
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
             *
             * Gets the unique ID of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: string;
            /**
             *
             * Gets the unique name of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: string;
        }
        /** An interface describing the data returned by calling `slide.toJSON()`. */
        export interface SlideData {
            /**
             *
             * Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling `slideCollection.toJSON()`. */
        export interface SlideCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `slideMasterCollection.toJSON()`. */
        export interface SlideMasterCollectionData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /**
         * [Api set: PowerPointApi 1.0]
         */
        export interface PresentationLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            title?: boolean;
        }
        /**
         *
         * Represents the bullet formatting properties of a text that is attached to the {@link PowerPoint.ParagraphFormat}.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface BulletFormatLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Specifies if the bullets in the paragraph are visible. Returns 'null' if the 'TextRange' includes text fragments with different bullet visibility values.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
        }
        /**
         *
         * Represents the paragraph formatting properties of a text that is attached to the {@link PowerPoint.TextRange}.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ParagraphFormatLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Represents the bullet format of the paragraph. See {@link PowerPoint.BulletFormat} for details.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            bulletFormat?: PowerPoint.Interfaces.BulletFormatLoadOptions;
            /**
             *
             * Represents the horizontal alignment of the paragraph. Returns 'null' if the 'TextRange' includes text fragments with different horizontal alignment values. See {@link PowerPoint.ParagraphHorizontalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            horizontalAlignment?: boolean;
        }
        /**
         *
         * Represents the fill formatting of a shape object.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ShapeFillLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            foregroundColor?: boolean;
            /**
             *
             * Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns `null` if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: boolean;
            /**
             *
             * Returns the fill type of the shape. See {@link PowerPoint.ShapeFillType} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            type?: boolean;
        }
        /**
         *
         * Represents the line formatting for the shape object. For images and geometric shapes, line formatting represents the border of the shape.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ShapeLineFormatLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: boolean;
            /**
             *
             * Represents the dash style of the line. Returns null when the line is not visible or there are inconsistent dash styles. See PowerPoint.ShapeLineDashStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            dashStyle?: boolean;
            /**
             *
             * Represents the line style of the shape. Returns null when the line is not visible or there are inconsistent styles. See PowerPoint.ShapeLineStyle for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            style?: boolean;
            /**
             *
             * Specifies the transparency percentage of the line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has inconsistent transparencies.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            transparency?: boolean;
            /**
             *
             * Specifies if the line formatting of a shape element is visible. Returns `null` when the shape has inconsistent visibilities.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            visible?: boolean;
            /**
             *
             * Represents the weight of the line, in points. Returns `null` when the line is not visible or there are inconsistent line weights.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            weight?: boolean;
        }
        /**
         *
         * Represents a single tag in the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface TagLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: boolean;
            /**
             *
             * Gets the value of the tag.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: boolean;
        }
        /**
         *
         * Represents the collection of tags.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface TagCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the tag. The `key` is unique within the owning `TagCollection` and always stored as uppercase letters within the document.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            key?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the value of the tag.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            value?: boolean;
        }
        /**
         *
         * Represents the font attributes, such as font name, font size, and color, for a shape's TextRange object.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ShapeFontLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Represents the bold status of font. Returns `null` if the `TextRange` includes both bold and non-bold text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bold?: boolean;
            /**
             *
             * HTML color code representation of the text color (e.g., "#FF0000" represents red). Returns `null` if the `TextRange` includes text fragments with different colors.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            color?: boolean;
            /**
             *
             * Represents the italic status of font. Returns 'null' if the 'TextRange' includes both italic and non-italic text fragments.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            italic?: boolean;
            /**
             *
             * Represents font name (e.g., "Calibri"). If the text is a Complex Script or East Asian language, this is the corresponding font name; otherwise it is the Latin font name.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
            /**
             *
             * Represents font size in points (e.g., 11). Returns null if the TextRange includes text fragments with different font sizes.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            size?: boolean;
            /**
             *
             * Type of underline applied to the font. Returns `null` if the `TextRange` includes text fragments with different underline styles. See {@link PowerPoint.ShapeFontUnderlineStyle} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            underline?: boolean;
        }
        /**
         *
         * Contains the text that is attached to a shape, in addition to properties and methods for manipulating the text.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface TextRangeLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Returns a `ShapeFont` object that represents the font attributes for the text range.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            font?: PowerPoint.Interfaces.ShapeFontLoadOptions;
            /**
            *
            * Represents the paragraph format of the text range. See {@link PowerPoint.ParagraphFormat} for details.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            paragraphFormat?: PowerPoint.Interfaces.ParagraphFormatLoadOptions;
            /**
             *
             * Represents the plain text content of the text range.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            text?: boolean;
        }
        /**
         *
         * Represents the text frame of a shape object.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface TextFrameLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text. See {@link PowerPoint.TextRange} for details.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            textRange?: PowerPoint.Interfaces.TextRangeLoadOptions;
            /**
             *
             * The automatic sizing settings for the text frame. A text frame can be set to automatically fit the text to the text frame, to automatically fit the text frame to the text, or not perform any automatic sizing.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            autoSizeSetting?: boolean;
            /**
             *
             * Represents the bottom margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            bottomMargin?: boolean;
            /**
             *
             * Specifies if the text frame contains text.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            hasText?: boolean;
            /**
             *
             * Represents the left margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            leftMargin?: boolean;
            /**
             *
             * Represents the right margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            rightMargin?: boolean;
            /**
             *
             * Represents the top margin, in points, of the text frame.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            topMargin?: boolean;
            /**
             *
             * Represents the vertical alignment of the text frame. See {@link PowerPoint.TextVerticalAlignment} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            verticalAlignment?: boolean;
            /**
             *
             * Determines whether lines break automatically to fit text inside the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            wordWrap?: boolean;
        }
        /**
         *
         * Represents a single shape in the slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ShapeLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Returns the fill formatting of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            fill?: PowerPoint.Interfaces.ShapeFillLoadOptions;
            /**
            *
            * Returns the line formatting of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            lineFormat?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions;
            /**
            *
            * Returns the text frame object of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            textFrame?: PowerPoint.Interfaces.TextFrameLoadOptions;
            /**
             *
             * Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            height?: boolean;
            /**
             *
             * Gets the unique ID of the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            left?: boolean;
            /**
             *
             * Specifies the name of this shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
            /**
             *
             * The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            top?: boolean;
            /**
             *
             * Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            type?: boolean;
            /**
             *
             * Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            width?: boolean;
        }
        /**
         *
         * Represents the collection of shapes.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface ShapeCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Returns the fill formatting of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            fill?: PowerPoint.Interfaces.ShapeFillLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Returns the line formatting of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            lineFormat?: PowerPoint.Interfaces.ShapeLineFormatLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Returns the text frame object of this shape.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            textFrame?: PowerPoint.Interfaces.TextFrameLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Specifies the height, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            height?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The distance, in points, from the left side of the shape to the left side of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            left?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Specifies the name of this shape.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The distance, in points, from the top edge of the shape to the top edge of the slide.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            top?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the type of this shape. See {@link PowerPoint.ShapeType} for details.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            type?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Specifies the width, in points, of the shape. Throws an `InvalidArgument` exception when set with a negative value.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            width?: boolean;
        }
        /**
         *
         * Represents the layout of a slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SlideLayoutLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Gets the unique ID of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * Gets the name of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
        }
        /**
         *
         * Represents the collection of layouts provided by the Slide Master for slides.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SlideLayoutCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the name of the slide layout.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
        }
        /**
         *
         * Represents the Slide Master of a slide.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SlideMasterLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Gets the unique ID of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * Gets the unique name of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            name?: boolean;
        }
        /**
         *
         * Represents a single slide of a presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Gets the layout of the slide.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            layout?: PowerPoint.Interfaces.SlideLayoutLoadOptions;
            /**
            *
            * Gets the `SlideMaster` object that represents the slide's default content.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            slideMaster?: PowerPoint.Interfaces.SlideMasterLoadOptions;
            /**
             *
             * Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         *
         * Represents the collection of slides in the presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the layout of the slide.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            layout?: PowerPoint.Interfaces.SlideLayoutLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the `SlideMaster` object that represents the slide's default content.
            *
            * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
            * @beta
            */
            slideMaster?: PowerPoint.Interfaces.SlideMasterLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         *
         * Represents the collection of Slide Masters in the presentation.
         *
         * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
         * @beta
         */
        export interface SlideMasterCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique name of the Slide Master.
             *
             * [Api set: PowerPointApi BETA (PREVIEW ONLY)]
             * @beta
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
    export function run<T>(batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
}
export declare namespace PowerPoint {
    /**
     * Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The base64-encoded .pptx file. The default value is null.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}


////////////////////////////////////////////////////////////////
///////////////////// End PowerPoint APIs //////////////////////
////////////////////////////////////////////////////////////////