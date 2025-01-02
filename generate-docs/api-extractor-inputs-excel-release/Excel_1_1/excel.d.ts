import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Excel APIs ///////////////////////
////////////////////////////////////////////////////////////////



export declare namespace Excel {
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    export interface Session {
    }
    /**
     * The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string | Session);
        readonly workbook: Workbook;
        readonly application: Application;
        
    }
    export interface RunOptions extends OfficeExtension.RunOptions<Session> {
        /**
         * Determines whether Excel will delay the batch request until the user exits cell edit mode.
         *
         * When false, if the user is in cell edit when the batch request is processed by the host, the batch will automatically fail.
         * When true, the batch request will be executed immediately if the user is not in cell edit mode, but if the user is in cell edit mode the batch request will be delayed until the user exits cell edit mode.
         */
        delayForCellEdit?: boolean;
    }
    /**
     * Executes a batch script that performs actions on the Excel object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
    * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
    * @param options - The additional options for this Excel.run which specify previous objects, whether to delay the request for cell edit, session info, etc.
    * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
    */
    export function run<T>(options: Excel.RunOptions, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     *
     * @param context - A previously-created object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(context: OfficeExtension.ClientRequestContext, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    export function postprocessBindingDescriptor(response: any): any;
    export function getDataCommonPostprocess(response: any, callArgs: any): any;
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Represents the Excel application that manages the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        /**
         * Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for get, 1.8 for set]
         */
        calculationMode: Excel.CalculationMode | "Automatic" | "AutomaticExceptTables" | "Manual";
        
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ApplicationUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Application): void;
        /**
         * Recalculate all currently opened workbooks in Excel.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param calculationType - Specifies the calculation type to use. See `Excel.CalculationType` for details.
         */
        calculate(calculationType: Excel.CalculationType): void;
        /**
         * Recalculate all currently opened workbooks in Excel.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param calculationTypeString - Specifies the calculation type to use. See `Excel.CalculationType` for details.
         */
        calculate(calculationTypeString: "Recalculate" | "Full" | "FullRebuild"): void;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ApplicationLoadOptions): Excel.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Application` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ApplicationData;
    }
    
    /**
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, and ranges.
                To learn more about the workbook object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks | Work with workbooks using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Workbook extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the Excel application instance that contains this workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly application: Excel.Application;
        /**
         * Represents a collection of bindings that are part of the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly bindings: Excel.BindingCollection;
        
        
        
        
        
        /**
         * Represents a collection of workbook-scoped named items (named ranges and constants).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly names: Excel.NamedItemCollection;
        
        
        
        
        
        
        
        
        
        
        /**
         * Represents a collection of tables associated with the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly tables: Excel.TableCollection;
        
        /**
         * Represents a collection of worksheets associated with the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly worksheets: Excel.WorksheetCollection;
        
        
        
        
        
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.WorkbookUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Workbook): void;
        
        
        
        
        
        
        
        
        
        /**
         * Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getSelectedRange(): Excel.Range;
        
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.WorkbookLoadOptions): Excel.Workbook;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Workbook;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Workbook;
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Workbook` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.WorkbookData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.WorkbookData;
    }
    
    
    /**
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
                To learn more about the worksheet object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets | Work with worksheets using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Worksheet extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Returns a collection of charts that are part of the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly charts: Excel.ChartCollection;
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Collection of tables that are part of the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly tables: Excel.TableCollection;
        
        
        /**
         * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly id: string;
        /**
         * The display name of the worksheet. The name must be fewer than 32 characters.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         * The zero-based position of the worksheet within the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        position: number;
        
        
        
        
        
        
        /**
         * The visibility of the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
         */
        visibility: Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.WorksheetUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Worksheet): void;
        /**
         * Activate the worksheet in the Excel UI.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        activate(): void;
        
        
        
        /**
         * Deletes the worksheet from the workbook. Note that if the worksheet's visibility is set to "VeryHidden", the delete operation will fail with an `InvalidOperation` exception. You should first change its visibility to hidden or visible before deleting it.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        
        
        /**
         * Gets the `Range` object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param row - The row number of the cell to be retrieved. Zero-indexed.
         * @param column - The column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Excel.Range;
        
        
        
        
        /**
         * Gets the `Range` object, representing a single rectangular block of cells, specified by the address or name.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param address - Optional. The string representing the address or name of the range. For example, "A1:B2". If not specified, the entire worksheet range is returned.
         */
        getRange(address?: string): Excel.Range;
        
        
        
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.WorksheetLoadOptions): Excel.Worksheet;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Worksheet;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Worksheet;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Worksheet` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.WorksheetData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.WorksheetData;
    }
    /**
     * Represents a collection of worksheet objects that are part of the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class WorksheetCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Worksheet[];
        /**
         * Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call `.activate()` on it.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param name - Optional. The name of the worksheet to be added. If specified, the name should be unique. If not specified, Excel determines the name of the new worksheet.
         */
        add(name?: string): Excel.Worksheet;
        /**
         * Gets the currently active worksheet in the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getActiveWorksheet(): Excel.Worksheet;
        
        
        /**
         * Gets a worksheet object using its name or ID.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param key - The name or ID of the worksheet.
         */
        getItem(key: string): Excel.Worksheet;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.WorksheetCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.WorksheetCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.WorksheetCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.WorksheetCollection;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.WorksheetCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.WorksheetCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.WorksheetCollectionData;
    }
    
    
    
    
    /**
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, or a block of cells.
                To learn more about how ranges are used throughout the API, start with {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-core-concepts#ranges | Ranges in the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Range extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        /**
         * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.RangeFormat;
        
        /**
         * The worksheet containing the current range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly worksheet: Excel.Worksheet;
        /**
         * Specifies the range reference in A1-style. Address value contains the sheet reference (e.g., "Sheet1!A1:B4").
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly address: string;
        /**
         * Represents the range reference for the specified range in the language of the user.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly addressLocal: string;
        /**
         * Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly cellCount: number;
        /**
         * Specifies the total number of columns in the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly columnCount: number;
        
        /**
         * Specifies the column number of the first cell in the range. Zero-indexed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly columnIndex: number;
        /**
         * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        formulas: any[][];
        /**
         * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        formulasLocal: any[][];
        
        
        
        
        
        
        
        
        
        /**
         * Represents Excel's number format code for the given range. For more information about Excel number formatting, see {@link https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68 | Number format codes}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        numberFormat: any[][];
        
        
        /**
         * Returns the total number of rows in the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly rowCount: number;
        
        /**
         * Returns the row number of the first cell in the range. Zero-indexed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly rowIndex: number;
        
        
        /**
         * Text values of the specified range. The text value will not depend on the cell width. The number sign (#) substitution that happens in the Excel UI will not affect the text value returned by the API.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly text: string[][];
        
        /**
         * Specifies the type of data in each cell.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly valueTypes: Excel.RangeValueType[][];
        /**
         * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                    If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        values: any[][];
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Range): void;
        
        
        
        /**
         * Clear range values and formatting, such as fill and border.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param applyTo - Optional. Determines the type of clear action. See `Excel.ClearApplyTo` for details.
         */
        clear(applyTo?: Excel.ClearApplyTo): void;
        /**
         * Clear range values and formatting, such as fill and border.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param applyToString - Optional. Determines the type of clear action. See `Excel.ClearApplyTo` for details.
         */
        clear(applyToString?: "All" | "Formats" | "Contents" | "Hyperlinks" | "RemoveHyperlinks"): void;
        
        
        
        
        /**
         * Deletes the cells associated with the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param shift - Specifies which way to shift the cells. See `Excel.DeleteShiftDirection` for details.
         */
        delete(shift: Excel.DeleteShiftDirection): void;
        /**
         * Deletes the cells associated with the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param shiftString - Specifies which way to shift the cells. See `Excel.DeleteShiftDirection` for details.
         */
        delete(shiftString: "Up" | "Left"): void;
        
        
        
        
        /**
         * Gets the smallest range object that encompasses the given ranges. For example, the `GetBoundingRect` of "B2:C5" and "D10:E15" is "B2:E15".
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param anotherRange - The range object, address, or range name.
         */
        getBoundingRect(anotherRange: Range | string): Excel.Range;
        /**
         * Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range, so long as it stays within the worksheet grid. The returned cell is located relative to the top left cell of the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param row - Row number of the cell to be retrieved. Zero-indexed.
         * @param column - Column number of the cell to be retrieved. Zero-indexed.
         */
        getCell(row: number, column: number): Excel.Range;
        
        /**
         * Gets a column contained in the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param column - Column number of the range to be retrieved. Zero-indexed.
         */
        getColumn(column: number): Excel.Range;
        
        
        
        
        
        
        /**
         * Gets an object that represents the entire column of the range (for example, if the current range represents cells "B4:E11", its `getEntireColumn` is a range that represents columns "B:E").
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getEntireColumn(): Excel.Range;
        /**
         * Gets an object that represents the entire row of the range (for example, if the current range represents cells "B4:E11", its `GetEntireRow` is a range that represents rows "4:11").
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getEntireRow(): Excel.Range;
        
        
        
        /**
         * Gets the range object that represents the rectangular intersection of the given ranges.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param anotherRange - The range object or range address that will be used to determine the intersection of ranges.
         */
        getIntersection(anotherRange: Range | string): Excel.Range;
        
        /**
         * Gets the last cell within the range. For example, the last cell of "B2:D5" is "D5".
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getLastCell(): Excel.Range;
        /**
         * Gets the last column within the range. For example, the last column of "B2:D5" is "D2:D5".
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getLastColumn(): Excel.Range;
        /**
         * Gets the last row within the range. For example, the last row of "B2:D5" is "B5:D5".
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getLastRow(): Excel.Range;
        
        /**
         * Gets an object which represents a range that's offset from the specified range. The dimension of the returned range will match this range. If the resulting range is forced outside the bounds of the worksheet grid, an error will be thrown.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param rowOffset - The number of rows (positive, negative, or 0) by which the range is to be offset. Positive values are offset downward, and negative values are offset upward.
         * @param columnOffset - The number of columns (positive, negative, or 0) by which the range is to be offset. Positive values are offset to the right, and negative values are offset to the left.
         */
        getOffsetRange(rowOffset: number, columnOffset: number): Excel.Range;
        
        
        
        
        
        /**
         * Gets a row contained in the range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param row - Row number of the range to be retrieved. Zero-indexed.
         */
        getRow(row: number): Excel.Range;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new `Range` object at the now blank space.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param shift - Specifies which way to shift the cells. See `Excel.InsertShiftDirection` for details.
         */
        insert(shift: Excel.InsertShiftDirection): Excel.Range;
        /**
         * Inserts a cell or a range of cells into the worksheet in place of this range, and shifts the other cells to make space. Returns a new `Range` object at the now blank space.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param shiftString - Specifies which way to shift the cells. See `Excel.InsertShiftDirection` for details.
         */
        insert(shiftString: "Down" | "Right"): Excel.Range;
        
        
        
        
        /**
         * Selects the specified range in the Excel UI.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        select(): void;
        
        
        
        
        
        
        
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeLoadOptions): Excel.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Range;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created.
         */
        track(): Excel.Range;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Excel.Range;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Range` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * A collection of all the `NamedItem` objects that are part of the workbook or worksheet, depending on how it was reached.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class NamedItemCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.NamedItem[];
        
        
        
        /**
         * Gets a `NamedItem` object using its name.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param name - Nameditem name.
         */
        getItem(name: string): Excel.NamedItem;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.NamedItemCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.NamedItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.NamedItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.NamedItemCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.NamedItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.NamedItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.NamedItemCollectionData;
    }
    /**
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class NamedItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
        
        /**
         * The name of the object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly name: string;
        
        /**
         * Specifies the type of the value returned by the name's formula. See `Excel.NamedItemType` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for String,Integer,Double,Boolean,Range,Error; 1.7 for Array]
         */
        readonly type: Excel.NamedItemType | "String" | "Integer" | "Double" | "Boolean" | "Range" | "Error" | "Array";
        /**
         * Represents the value computed by the name's formula. For a named range, it will return the range address.
                    This API returns the #VALUE! error in the Excel UI if it refers to a user-defined function.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly value: any;
        
        
        /**
         * Specifies if the object is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.NamedItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.NamedItem): void;
        
        /**
         * Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.NamedItemLoadOptions): Excel.NamedItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.NamedItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.NamedItem;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.NamedItem` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.NamedItemData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.NamedItemData;
    }
    
    /**
     * Represents an Office.js binding that is defined in the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Binding extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the binding identifier.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly id: string;
        /**
         * Returns the type of the binding. See `Excel.BindingType` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly type: Excel.BindingType | "Range" | "Table" | "Text";
        
        /**
         * Returns the range represented by the binding. Will throw an error if the binding is not of the correct type.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         * Returns the table represented by the binding. Will throw an error if the binding is not of the correct type.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getTable(): Excel.Table;
        /**
         * Returns the text represented by the binding. Will throw an error if the binding is not of the correct type.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getText(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.BindingLoadOptions): Excel.Binding;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Binding;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Binding;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Binding` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.BindingData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.BindingData;
    }
    /**
     * Represents the collection of all the binding objects that are part of the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class BindingCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Binding[];
        /**
         * Returns the number of bindings in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        
        
        
        
        
        
        /**
         * Gets a binding object by ID.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param id - ID of the binding object to be retrieved.
         */
        getItem(id: string): Excel.Binding;
        /**
         * Gets a binding object based on its position in the items array.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Binding;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.BindingCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.BindingCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.BindingCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.BindingCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.BindingCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.BindingCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.BindingCollectionData;
    }
    /**
     * Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Table[];
        /**
         * Returns the number of tables in the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Creates a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param address - A `Range` object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used. [Api set: ExcelApi 1.1 / 1.3.  Prior to ExcelApi 1.3, this parameter must be a string. Starting with Excel Api 1.3, this parameter may be a Range object or a string.]
         * @param hasHeaders - A boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e., when this property set to `false`), Excel will automatically generate a header and shift the data down by one row.
         */
        add(address: Range | string, hasHeaders: boolean): Excel.Table;
        
        /**
         * Gets a table by name or ID.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param key - Name or ID of the table to be retrieved.
         */
        getItem(key: string): Excel.Table;
        /**
         * Gets a table based on its position in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Table;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.TableCollection;
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.TableCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.TableCollectionData;
    }
    
    /**
     * Represents an Excel table.
                To learn more about the table object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables | Work with tables using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Table extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents a collection of all the columns in the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly columns: Excel.TableColumnCollection;
        /**
         * Represents a collection of all the rows in the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly rows: Excel.TableRowCollection;
        
        
        
        
        /**
         * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly id: string;
        
        /**
         * Name of the table.
                    
                     The set name of the table must follow the guidelines specified in the {@link https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114 | Rename an Excel table} article.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        
        
        
        /**
         * Specifies if the header row is visible. This value can be set to show or remove the header row.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showHeaders: boolean;
        /**
         * Specifies if the total row is visible. This value can be set to show or remove the total row.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showTotals: boolean;
        /**
         * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        style: string;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Table): void;
        
        
        /**
         * Deletes the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         * Gets the range object associated with the data body of the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getDataBodyRange(): Excel.Range;
        /**
         * Gets the range object associated with the header row of the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getHeaderRowRange(): Excel.Range;
        /**
         * Gets the range object associated with the entire table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         * Gets the range object associated with the totals row of the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getTotalRowRange(): Excel.Range;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableLoadOptions): Excel.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Table;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Table` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.TableData;
    }
    /**
     * Represents a collection of all the columns that are part of the table.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableColumnCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.TableColumn[];
        /**
         * Returns the number of columns in the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Adds a new column to the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 requires an index smaller than the total column count; 1.4 allows index to be optional (null or -1) and will append a column at the end; 1.4 allows name parameter at creation time.]
         *
         * @param index - Optional. Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.
         * @param values - Optional. A 2D array of unformatted values of the table column.
         * @param name - Optional. Specifies the name of the new column. If `null`, the default name will be used.
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number, name?: string): Excel.TableColumn;
        
        
        /**
         * Gets a column object by name or ID.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param key - Column name or ID.
         */
        getItem(key: number | string): Excel.TableColumn;
        /**
         * Gets a column based on its position in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.TableColumn;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableColumnCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.TableColumnCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableColumnCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.TableColumnCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.TableColumnCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableColumnCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.TableColumnCollectionData;
    }
    /**
     * Represents a column in a table.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableColumn extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Returns a unique key that identifies the column within the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly id: number;
        /**
         * Returns the index number of the column within the columns collection of the table. Zero-indexed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly index: number;
        /**
         * Specifies the name of the table column.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
         */
        name: string;
        /**
         * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                    If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        values: any[][];
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableColumnUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.TableColumn): void;
        /**
         * Deletes the column from the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         * Gets the range object associated with the data body of the column.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getDataBodyRange(): Excel.Range;
        /**
         * Gets the range object associated with the header row of the column.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getHeaderRowRange(): Excel.Range;
        /**
         * Gets the range object associated with the entire column.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         * Gets the range object associated with the totals row of the column.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getTotalRowRange(): Excel.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableColumnLoadOptions): Excel.TableColumn;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableColumn;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.TableColumn;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.TableColumn` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableColumnData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.TableColumnData;
    }
    /**
     * Represents a collection of all the rows that are part of the table.
                
                 Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                 a `TableRow` object represents the physical location of the table row, but not the data.
                 That is, if the data is sorted or if new rows are added, a table row will continue
                 to point at the index for which it was created.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.TableRow[];
        /**
         * Returns the number of rows in the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Adds one or more rows to the table. The return object will be the top of the newly added row(s).
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for adding a single row; 1.4 allows adding of multiple rows; 1.15 for adding `alwaysInsert` parameter.]
         *
         * @param index - Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values - Optional. A 2D array of unformatted values of the table row.
         * @param alwaysInsert - Optional. Specifies whether the new rows will be inserted into the table when new rows are added. If `true`, the new rows will be inserted into the table. If `false`, the new rows will be added below the table. Default is `true`.
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number, alwaysInsert?: boolean): Excel.TableRow;
        
        
        
        
        /**
         * Gets a row based on its position in the collection.
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableRowCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.TableRowCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableRowCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.TableRowCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.TableRowCollectionData;
    }
    /**
     * Represents a row in a table.
                
                 Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                 a `TableRow` object represents the physical location of the table row, but not the data.
                 That is, if the data is sorted or if new rows are added, a table row will continue
                 to point at the index for which it was created.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Returns the index number of the row within the rows collection of the table. Zero-indexed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly index: number;
        /**
         * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                    If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        values: any[][];
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableRowUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.TableRow): void;
        /**
         * Deletes the row from the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         * Returns the range object associated with the entire row.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableRowLoadOptions): Excel.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.TableRow;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.TableRow` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.TableRowData;
    }
    
    
    
    
    
    
    
    
    
    /**
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Collection of border objects that apply to the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly borders: Excel.RangeBorderCollection;
        /**
         * Returns the fill object defined on the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.RangeFill;
        /**
         * Returns the font object defined on the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.RangeFont;
        
        
        
        /**
         * Represents the horizontal alignment for the specified object. See `Excel.HorizontalAlignment` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        horizontalAlignment: Excel.HorizontalAlignment | "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed";
        
        
        
        
        
        
        
        /**
         * Represents the vertical alignment for the specified object. See `Excel.VerticalAlignment` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        verticalAlignment: Excel.VerticalAlignment | "Top" | "Center" | "Bottom" | "Justify" | "Distributed";
        /**
         * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        wrapText: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.RangeFormat): void;
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeFormatLoadOptions): Excel.RangeFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.RangeFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.RangeFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.RangeFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeFormatData;
    }
    
    /**
     * Represents the background of a range object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeFillUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.RangeFill): void;
        /**
         * Resets the range background.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeFillLoadOptions): Excel.RangeFill;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.RangeFill;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.RangeFill;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.RangeFill` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeFillData;
    }
    /**
     * Represents the border of an object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeBorder extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly sideIndex: Excel.BorderIndex | "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp";
        /**
         * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        style: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
        
        /**
         * Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        weight: Excel.BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeBorderUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.RangeBorder): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeBorderLoadOptions): Excel.RangeBorder;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.RangeBorder;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.RangeBorder;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.RangeBorder` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeBorderData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeBorderData;
    }
    /**
     * Represents the border objects that make up the range border.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeBorderCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.RangeBorder[];
        /**
         * Number of border objects in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        /**
         * Gets a border object using its name.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the border object to be retrieved. See `Excel.BorderIndex` for details.
         */
        getItem(index: Excel.BorderIndex): Excel.RangeBorder;
        /**
         * Gets a border object using its name.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param indexString - Index value of the border object to be retrieved. See `Excel.BorderIndex` for details.
         */
        getItem(indexString: "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp"): Excel.RangeBorder;
        /**
         * Gets a border object using its index.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.RangeBorder;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeBorderCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.RangeBorderCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.RangeBorderCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.RangeBorderCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.RangeBorderCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeBorderCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.RangeBorderCollectionData;
    }
    /**
     * This object represents the font attributes (font name, font size, color, etc.) for an object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeFont extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the bold status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Specifies the italic status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         * Font size.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        
        
        
        
        /**
         * Type of underline applied to the font. See `Excel.RangeUnderlineStyle` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        underline: Excel.RangeUnderlineStyle | "None" | "Single" | "Double" | "SingleAccountant" | "DoubleAccountant";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.RangeFontUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.RangeFont): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.RangeFontLoadOptions): Excel.RangeFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.RangeFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.RangeFont;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.RangeFont` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeFontData;
    }
    /**
     * A collection of all the chart objects on a worksheet.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Chart[];
        /**
         * Returns the number of charts in the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Creates a new chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param type - Represents the type of a chart. See `Excel.ChartType` for details.
         * @param sourceData - The `Range` object corresponding to the source data.
         * @param seriesBy - Optional. Specifies the way columns or rows are used as data series on the chart. See `Excel.ChartSeriesBy` for details.
         */
        add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy): Excel.Chart;
        /**
         * Creates a new chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param typeString - Represents the type of a chart. See `Excel.ChartType` for details.
         * @param sourceData - The `Range` object corresponding to the source data.
         * @param seriesByString - Optional. Specifies the way columns or rows are used as data series on the chart. See `Excel.ChartSeriesBy` for details.
         */
        add(typeString: "Invalid" | "ColumnClustered" | "ColumnStacked" | "ColumnStacked100" | "3DColumnClustered" | "3DColumnStacked" | "3DColumnStacked100" | "BarClustered" | "BarStacked" | "BarStacked100" | "3DBarClustered" | "3DBarStacked" | "3DBarStacked100" | "LineStacked" | "LineStacked100" | "LineMarkers" | "LineMarkersStacked" | "LineMarkersStacked100" | "PieOfPie" | "PieExploded" | "3DPieExploded" | "BarOfPie" | "XYScatterSmooth" | "XYScatterSmoothNoMarkers" | "XYScatterLines" | "XYScatterLinesNoMarkers" | "AreaStacked" | "AreaStacked100" | "3DAreaStacked" | "3DAreaStacked100" | "DoughnutExploded" | "RadarMarkers" | "RadarFilled" | "Surface" | "SurfaceWireframe" | "SurfaceTopView" | "SurfaceTopViewWireframe" | "Bubble" | "Bubble3DEffect" | "StockHLC" | "StockOHLC" | "StockVHLC" | "StockVOHLC" | "CylinderColClustered" | "CylinderColStacked" | "CylinderColStacked100" | "CylinderBarClustered" | "CylinderBarStacked" | "CylinderBarStacked100" | "CylinderCol" | "ConeColClustered" | "ConeColStacked" | "ConeColStacked100" | "ConeBarClustered" | "ConeBarStacked" | "ConeBarStacked100" | "ConeCol" | "PyramidColClustered" | "PyramidColStacked" | "PyramidColStacked100" | "PyramidBarClustered" | "PyramidBarStacked" | "PyramidBarStacked100" | "PyramidCol" | "3DColumn" | "Line" | "3DLine" | "3DPie" | "Pie" | "XYScatter" | "3DArea" | "Area" | "Doughnut" | "Radar" | "Histogram" | "Boxwhisker" | "Pareto" | "RegionMap" | "Treemap" | "Waterfall" | "Sunburst" | "Funnel", sourceData: Range, seriesByString?: "Auto" | "Columns" | "Rows"): Excel.Chart;
        
        /**
         * Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param name - Name of the chart to be retrieved.
         */
        getItem(name: string): Excel.Chart;
        /**
         * Gets a chart based on its position in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.Chart;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.ChartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.ChartCollection;
        
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.ChartCollectionData;
    }
    /**
     * Represents a chart object in a workbook.
                To learn more about the chart object model, see {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts | Work with charts using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Chart extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents chart axes.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly axes: Excel.ChartAxes;
        /**
         * Represents the data labels on the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly dataLabels: Excel.ChartDataLabels;
        /**
         * Encapsulates the format properties for the chart area.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAreaFormat;
        /**
         * Represents the legend for the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly legend: Excel.ChartLegend;
        
        
        /**
         * Represents either a single series or collection of series in the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly series: Excel.ChartSeriesCollection;
        /**
         * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly title: Excel.ChartTitle;
        
        
        
        
        /**
         * Specifies the height, in points, of the chart object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        height: number;
        
        /**
         * The distance, in points, from the left side of the chart to the worksheet origin.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left: number;
        /**
         * Specifies the name of a chart object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        
        
        
        
        
        
        /**
         * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top: number;
        /**
         * Specifies the width, in points, of the chart object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        width: number;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.Chart): void;
        
        /**
         * Deletes the chart object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        
        
        
        
        /**
         * Resets the source data for the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param sourceData - The range object corresponding to the source data.
         * @param seriesBy - Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, and Columns. See `Excel.ChartSeriesBy` for details.
         */
        setData(sourceData: Range, seriesBy?: Excel.ChartSeriesBy): void;
        /**
         * Resets the source data for the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param sourceData - The range object corresponding to the source data.
         * @param seriesByString - Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, and Columns. See `Excel.ChartSeriesBy` for details.
         */
        setData(sourceData: Range, seriesByString?: "Auto" | "Columns" | "Rows"): void;
        /**
         * Positions the chart relative to cells on the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param startCell - The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's right-to-left display settings.
         * @param endCell - Optional. The end cell. If specified, the chart's width and height will be set to fully cover up this cell/range.
         */
        setPosition(startCell: Range | string, endCell?: Range | string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartLoadOptions): Excel.Chart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.Chart;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.Chart;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.Chart` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartData;
    }
    
    /**
     * Encapsulates the format properties for the overall chart area.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAreaFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents the fill format of an object, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents the font attributes (font name, font size, color, etc.) for the current object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAreaFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAreaFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAreaFormatLoadOptions): Excel.ChartAreaFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAreaFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAreaFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAreaFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAreaFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAreaFormatData;
    }
    /**
     * Represents a collection of chart series.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartSeriesCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.ChartSeries[];
        /**
         * Returns the number of series in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        
        /**
         * Retrieves a series based on its position in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.ChartSeries;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartSeriesCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.ChartSeriesCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartSeriesCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.ChartSeriesCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartSeriesCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartSeriesCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.ChartSeriesCollectionData;
    }
    /**
     * Represents a series in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartSeries extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        /**
         * Represents the formatting of a chart series, which includes fill and line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartSeriesFormat;
        
        /**
         * Returns a collection of all points in the series.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly points: Excel.ChartPointsCollection;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartSeriesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartSeries): void;
        
        
        
        
        
        
        
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartSeriesLoadOptions): Excel.ChartSeries;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartSeries;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartSeries;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartSeries` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartSeriesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartSeriesData;
    }
    /**
     * Encapsulates the format properties for the chart series
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartSeriesFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the fill format of a chart series, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly line: Excel.ChartLineFormat;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartSeriesFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartSeriesFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartSeriesFormatLoadOptions): Excel.ChartSeriesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartSeriesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartSeriesFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartSeriesFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartSeriesFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartSeriesFormatData;
    }
    /**
     * A collection of all the chart points within a series inside a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartPointsCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.ChartPoint[];
        /**
         * Returns the number of chart points in the series.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        /**
         * Retrieve a point based on its position within the series.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): Excel.ChartPoint;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartPointsCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.ChartPointsCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartPointsCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.ChartPointsCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartPointsCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartPointsCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.ChartPointsCollectionData;
    }
    /**
     * Represents a point of a series in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartPoint extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Encapsulates the format properties chart point.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartPointFormat;
        
        
        
        
        
        /**
         * Returns the value of a chart point.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly value: any;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartPointUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartPoint): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartPointLoadOptions): Excel.ChartPoint;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartPoint;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartPoint;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartPoint` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartPointData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartPointData;
    }
    /**
     * Represents the formatting object for chart points.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartPointFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents the fill format of a chart, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartPointFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartPointFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartPointFormatLoadOptions): Excel.ChartPointFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartPointFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartPointFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartPointFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartPointFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartPointFormatData;
    }
    /**
     * Represents the chart axes.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxes extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the category axis in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly categoryAxis: Excel.ChartAxis;
        /**
         * Represents the series axis of a 3-D chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly seriesAxis: Excel.ChartAxis;
        /**
         * Represents the value axis in an axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly valueAxis: Excel.ChartAxis;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxes): void;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAxesLoadOptions): Excel.ChartAxes;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAxes;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAxes;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAxes` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxesData;
    }
    /**
     * Represents a single axis in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxis extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart object, which includes line and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAxisFormat;
        /**
         * Returns an object that represents the major gridlines for the specified axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly majorGridlines: Excel.ChartGridlines;
        /**
         * Returns an object that represents the minor gridlines for the specified axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly minorGridlines: Excel.ChartGridlines;
        /**
         * Represents the axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly title: Excel.ChartAxisTitle;
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string. The returned value is always a number.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        majorUnit: any;
        /**
         * Represents the maximum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        maximum: any;
        /**
         * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        minimum: any;
        
        
        /**
         * Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        minorUnit: any;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxisUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxis): void;
        
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAxisLoadOptions): Excel.ChartAxis;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAxis;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAxis;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAxis` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisData;
    }
    /**
     * Encapsulates the format properties for the chart axis.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxisFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        /**
         * Specifies chart line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly line: Excel.ChartLineFormat;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxisFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxisFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAxisFormatLoadOptions): Excel.ChartAxisFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAxisFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAxisFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAxisFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisFormatData;
    }
    /**
     * Represents the title of a chart axis.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxisTitle extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies the formatting of the chart axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAxisTitleFormat;
        /**
         * Specifies the axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        text: string;
        
        /**
         * Specifies if the axis title is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxisTitleUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxisTitle): void;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAxisTitleLoadOptions): Excel.ChartAxisTitle;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAxisTitle;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAxisTitle;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAxisTitle` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisTitleData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisTitleData;
    }
    /**
     * Represents the chart axis title formatting.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        /**
         * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxisTitleFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxisTitleFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartAxisTitleFormatLoadOptions): Excel.ChartAxisTitleFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartAxisTitleFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartAxisTitleFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartAxisTitleFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisTitleFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisTitleFormatData;
    }
    /**
     * Represents a collection of all the data labels on a chart point.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartDataLabels extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies the format of chart data labels, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartDataLabelFormat;
        
        
        
        
        /**
         * Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        position: Excel.ChartDataLabelPosition | "Invalid" | "None" | "Center" | "InsideEnd" | "InsideBase" | "OutsideEnd" | "Left" | "Right" | "Top" | "Bottom" | "BestFit" | "Callout";
        /**
         * String representing the separator used for the data labels on a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        separator: string;
        /**
         * Specifies if the data label bubble size is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showBubbleSize: boolean;
        /**
         * Specifies if the data label category name is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showCategoryName: boolean;
        /**
         * Specifies if the data label legend key is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showLegendKey: boolean;
        /**
         * Specifies if the data label percentage is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showPercentage: boolean;
        /**
         * Specifies if the data label series name is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showSeriesName: boolean;
        /**
         * Specifies if the data label value is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        showValue: boolean;
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartDataLabelsUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartDataLabels): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartDataLabelsLoadOptions): Excel.ChartDataLabels;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartDataLabels;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartDataLabels;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartDataLabels` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartDataLabelsData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartDataLabelsData;
    }
    
    /**
     * Encapsulates the format properties for the chart data labels.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartDataLabelFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents the fill format of the current chart data label.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents the font attributes (such as font name, font size, and color) for a chart data label.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartDataLabelFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartDataLabelFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartDataLabelFormatLoadOptions): Excel.ChartDataLabelFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartDataLabelFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartDataLabelFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartDataLabelFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartDataLabelFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartDataLabelFormatData;
    }
    
    
    
    
    /**
     * Represents major or minor gridlines on a chart axis.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartGridlines extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of chart gridlines.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartGridlinesFormat;
        /**
         * Specifies if the axis gridlines are visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartGridlinesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartGridlines): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartGridlinesLoadOptions): Excel.ChartGridlines;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartGridlines;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartGridlines;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartGridlines` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartGridlinesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartGridlinesData;
    }
    /**
     * Encapsulates the format properties for chart gridlines.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartGridlinesFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents chart line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly line: Excel.ChartLineFormat;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartGridlinesFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartGridlinesFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartGridlinesFormatLoadOptions): Excel.ChartGridlinesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartGridlinesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartGridlinesFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartGridlinesFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartGridlinesFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartGridlinesFormatData;
    }
    /**
     * Represents the legend in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLegend extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart legend, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartLegendFormat;
        
        
        
        /**
         * Specifies if the chart legend should overlap with the main body of the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        overlay: boolean;
        /**
         * Specifies the position of the legend on the chart. See `Excel.ChartLegendPosition` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        position: Excel.ChartLegendPosition | "Invalid" | "Top" | "Bottom" | "Left" | "Right" | "Corner" | "Custom";
        
        
        /**
         * Specifies if the chart legend is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartLegendUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartLegend): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartLegendLoadOptions): Excel.ChartLegend;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartLegend;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartLegend;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartLegend` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLegendData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLegendData;
    }
    
    
    /**
     * Encapsulates the format properties of a chart legend.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLegendFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents the fill format of an object, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents the font attributes such as font name, font size, and color of a chart legend.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartLegendFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartLegendFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartLegendFormatLoadOptions): Excel.ChartLegendFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartLegendFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartLegendFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartLegendFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLegendFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLegendFormatData;
    }
    
    /**
     * Represents a chart title object of a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartTitle extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart title, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartTitleFormat;
        
        
        
        /**
         * Specifies if the chart title will overlay the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        overlay: boolean;
        
        
        /**
         * Specifies the chart's title text.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        text: string;
        
        
        
        /**
         * Specifies if the chart title is visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartTitleUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartTitle): void;
        
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartTitleLoadOptions): Excel.ChartTitle;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartTitle;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartTitle;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartTitle` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartTitleData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartTitleData;
    }
    
    /**
     * Provides access to the formatting options for a chart title.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartTitleFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
         * Represents the fill format of an object, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents the font attributes (such as font name, font size, and color) for an object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.ChartFont;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartTitleFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartTitleFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartTitleFormatLoadOptions): Excel.ChartTitleFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartTitleFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartTitleFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartTitleFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartTitleFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartTitleFormatData;
    }
    /**
     * Represents the fill formatting for a chart element.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Clears the fill color of a chart element.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        
        /**
         * Sets the fill formatting of a chart element to a uniform color.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param color - HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartFill` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    
    
    
    /**
     * Encapsulates the formatting options for line elements.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLineFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of lines in the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        
        
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartLineFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartLineFormat): void;
        /**
         * Clears the line format of a chart element.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartLineFormatLoadOptions): Excel.ChartLineFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartLineFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartLineFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartLineFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLineFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLineFormatData;
    }
    /**
     * This object represents the font attributes (such as font name, font size, and color) for a chart object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartFont extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the bold status of font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Represents the italic status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         * Font name (e.g., "Calibri")
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         * Size of the font (e.g., 11)
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        /**
         * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        underline: Excel.ChartUnderlineStyle | "None" | "Single";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartFontUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartFont): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartFontLoadOptions): Excel.ChartFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartFont;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Excel.ChartFont` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartFontData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartDataLabelPosition {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        center = "Center",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideEnd = "InsideEnd",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideBase = "InsideBase",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        outsideEnd = "OutsideEnd",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bestFit = "BestFit",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        callout = "Callout"
    }
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartLegendPosition {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        corner = "Corner",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        custom = "Custom"
    }
    
    
    
    
    
    /**
     * Specifies whether the series are by rows or by columns. In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartSeriesBy {
        /**
         * In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        auto = "Auto",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columns = "Columns",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        rows = "Rows"
    }
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnClustered = "ColumnClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnStacked = "ColumnStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnStacked100 = "ColumnStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnClustered = "3DColumnClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnStacked = "3DColumnStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnStacked100 = "3DColumnStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barClustered = "BarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barStacked = "BarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barStacked100 = "BarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarClustered = "3DBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarStacked = "3DBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarStacked100 = "3DBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineStacked = "LineStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineStacked100 = "LineStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkers = "LineMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkersStacked = "LineMarkersStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkersStacked100 = "LineMarkersStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pieOfPie = "PieOfPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pieExploded = "PieExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DPieExploded = "3DPieExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barOfPie = "BarOfPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterSmooth = "XYScatterSmooth",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterSmoothNoMarkers = "XYScatterSmoothNoMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterLines = "XYScatterLines",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterLinesNoMarkers = "XYScatterLinesNoMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        areaStacked = "AreaStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        areaStacked100 = "AreaStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DAreaStacked = "3DAreaStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DAreaStacked100 = "3DAreaStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        doughnutExploded = "DoughnutExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radarMarkers = "RadarMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radarFilled = "RadarFilled",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surface = "Surface",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceWireframe = "SurfaceWireframe",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceTopView = "SurfaceTopView",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceTopViewWireframe = "SurfaceTopViewWireframe",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bubble = "Bubble",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bubble3DEffect = "Bubble3DEffect",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockHLC = "StockHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockOHLC = "StockOHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockVHLC = "StockVHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockVOHLC = "StockVOHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColClustered = "CylinderColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColStacked = "CylinderColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColStacked100 = "CylinderColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarClustered = "CylinderBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarStacked = "CylinderBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarStacked100 = "CylinderBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderCol = "CylinderCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColClustered = "ConeColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColStacked = "ConeColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColStacked100 = "ConeColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarClustered = "ConeBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarStacked = "ConeBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarStacked100 = "ConeBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneCol = "ConeCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColClustered = "PyramidColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColStacked = "PyramidColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColStacked100 = "PyramidColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarClustered = "PyramidBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarStacked = "PyramidBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarStacked100 = "PyramidBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidCol = "PyramidCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumn = "3DColumn",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        line = "Line",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DLine = "3DLine",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DPie = "3DPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pie = "Pie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatter = "XYScatter",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DArea = "3DArea",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        area = "Area",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        doughnut = "Doughnut",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radar = "Radar",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        histogram = "Histogram",
                                                            }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartUnderlineStyle {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        single = "Single"
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum BindingType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        range = "Range",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        table = "Table",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        text = "Text"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum BorderIndex {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        edgeTop = "EdgeTop",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        edgeBottom = "EdgeBottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        edgeLeft = "EdgeLeft",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        edgeRight = "EdgeRight",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideVertical = "InsideVertical",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideHorizontal = "InsideHorizontal",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        diagonalDown = "DiagonalDown",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        diagonalUp = "DiagonalUp"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum BorderLineStyle {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        continuous = "Continuous",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        dash = "Dash",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        dashDot = "DashDot",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        dashDotDot = "DashDotDot",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        dot = "Dot",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        slantDashDot = "SlantDashDot"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum BorderWeight {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        hairline = "Hairline",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        thin = "Thin",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        medium = "Medium",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        thick = "Thick"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum CalculationMode {
        /**
         * The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        automatic = "Automatic",
        /**
         * Calculates new formula results every time the relevant data is changed, unless the formula is in a data table.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        automaticExceptTables = "AutomaticExceptTables",
        /**
         * Calculations only occur when the user or add-in requests them.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        manual = "Manual"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum CalculationType {
        /**
         * Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        recalculate = "Recalculate",
        /**
         * This will mark all cells as dirty and then recalculate them.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        full = "Full",
        /**
         * This will rebuild the full dependency chain, mark all cells as dirty and then recalculate them.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        fullRebuild = "FullRebuild"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ClearApplyTo {
        /**
         * Clears everything in the range.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        all = "All",
        /**
         * Clears all formatting for the range, leaving values intact.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        formats = "Formats",
        /**
         * Clears the contents of the range, leaving formatting intact.
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        contents = "Contents",
                    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum DeleteShiftDirection {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        up = "Up",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left"
    }
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum HorizontalAlignment {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        general = "General",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        center = "Center",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        fill = "Fill",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        justify = "Justify",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        centerAcrossSelection = "CenterAcrossSelection",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        distributed = "Distributed"
    }
    
    
    /**
     * Determines the direction in which existing cells will be shifted to accommodate what is being inserted.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum InsertShiftDirection {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        down = "Down",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right"
    }
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1 for String,Integer,Double,Boolean,Range,Error; 1.7 for Array]
     */
    enum NamedItemType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        string = "String",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        integer = "Integer",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        boolean = "Boolean",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        range = "Range",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        error = "Error",
            }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum RangeUnderlineStyle {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        single = "Single",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        singleAccountant = "SingleAccountant",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        doubleAccountant = "DoubleAccountant"
    }
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum SheetVisibility {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible = "Visible",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        hidden = "Hidden",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        veryHidden = "VeryHidden"
    }
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum RangeValueType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        unknown = "Unknown",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        empty = "Empty",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        string = "String",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        integer = "Integer",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        boolean = "Boolean",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        error = "Error",
            }
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum VerticalAlignment {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        center = "Center",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        justify = "Justify",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        distributed = "Distributed"
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        apiNotFound = "ApiNotFound",
        conflict = "Conflict",
        emptyChartSeries = "EmptyChartSeries",
        filteredRangeConflict = "FilteredRangeConflict",
        formulaLengthExceedsLimit = "FormulaLengthExceedsLimit",
        generalException = "GeneralException",
        inactiveWorkbook = "InactiveWorkbook",
        insertDeleteConflict = "InsertDeleteConflict",
        invalidArgument = "InvalidArgument",
        invalidBinding = "InvalidBinding",
        invalidOperation = "InvalidOperation",
        invalidReference = "InvalidReference",
        invalidSelection = "InvalidSelection",
        itemAlreadyExists = "ItemAlreadyExists",
        itemNotFound = "ItemNotFound",
        mergedRangeConflict = "MergedRangeConflict",
        nonBlankCellOffSheet = "NonBlankCellOffSheet",
        notImplemented = "NotImplemented",
        openWorkbookLinksBlocked = "OpenWorkbookLinksBlocked",
        operationCellsExceedLimit = "OperationCellsExceedLimit",
        pivotTableRangeConflict = "PivotTableRangeConflict",
        powerQueryRefreshResourceChallenge = "PowerQueryRefreshResourceChallenge",
        rangeExceedsLimit = "RangeExceedsLimit",
        refreshWorkbookLinksBlocked = "RefreshWorkbookLinksBlocked",
        requestAborted = "RequestAborted",
        responsePayloadSizeLimitExceeded = "ResponsePayloadSizeLimitExceeded",
        unsupportedFeature = "UnsupportedFeature",
        unsupportedFillType = "UnsupportedFillType",
        unsupportedOperation = "UnsupportedOperation",
        unsupportedSheet = "UnsupportedSheet",
        invalidOperationInCellEditMode = "InvalidOperationInCellEditMode"
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
        /** An interface for updating data on the `AllowEditRange` object, for use in `allowEditRange.set({ ... })`. */
        export interface AllowEditRangeUpdateData {
            
            
        }
        /** An interface for updating data on the `AllowEditRangeCollection` object, for use in `allowEditRangeCollection.set({ ... })`. */
        export interface AllowEditRangeCollectionUpdateData {
            items?: Excel.Interfaces.AllowEditRangeData[];
        }
        /** An interface for updating data on the `QueryCollection` object, for use in `queryCollection.set({ ... })`. */
        export interface QueryCollectionUpdateData {
            items?: Excel.Interfaces.QueryData[];
        }
        /** An interface for updating data on the `LinkedWorkbookCollection` object, for use in `linkedWorkbookCollection.set({ ... })`. */
        export interface LinkedWorkbookCollectionUpdateData {
            
            items?: Excel.Interfaces.LinkedWorkbookData[];
        }
        /** An interface for updating data on the `Runtime` object, for use in `runtime.set({ ... })`. */
        export interface RuntimeUpdateData {
            
        }
        /** An interface for updating data on the `Application` object, for use in `application.set({ ... })`. */
        export interface ApplicationUpdateData {
            
            /**
             * Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for get, 1.8 for set]
             */
            calculationMode?: Excel.CalculationMode | "Automatic" | "AutomaticExceptTables" | "Manual";
        }
        /** An interface for updating data on the `IterativeCalculation` object, for use in `iterativeCalculation.set({ ... })`. */
        export interface IterativeCalculationUpdateData {
            
            
            
        }
        /** An interface for updating data on the `Workbook` object, for use in `workbook.set({ ... })`. */
        export interface WorkbookUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `Worksheet` object, for use in `worksheet.set({ ... })`. */
        export interface WorksheetUpdateData {
            
            
            /**
             * The display name of the worksheet. The name must be fewer than 32 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * The zero-based position of the worksheet within the workbook.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: number;
            
            
            
            
            /**
             * The visibility of the worksheet.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
             */
            visibility?: Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden";
        }
        /** An interface for updating data on the `WorksheetCollection` object, for use in `worksheetCollection.set({ ... })`. */
        export interface WorksheetCollectionUpdateData {
            items?: Excel.Interfaces.WorksheetData[];
        }
        /** An interface for updating data on the `Range` object, for use in `range.set({ ... })`. */
        export interface RangeUpdateData {
            
            /**
            * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.RangeFormatUpdateData;
            
            /**
             * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulas?: any[][];
            /**
             * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulasLocal?: any[][];
            
            
            /**
             * Represents Excel's number format code for the given range. For more information about Excel number formatting, see {@link https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68 | Number format codes}.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            numberFormat?: any[][];
            
            
            
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
        }
        /** An interface for updating data on the `RangeAreas` object, for use in `rangeAreas.set({ ... })`. */
        export interface RangeAreasUpdateData {
            
            
            
        }
        /** An interface for updating data on the `RangeView` object, for use in `rangeView.set({ ... })`. */
        export interface RangeViewUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `RangeViewCollection` object, for use in `rangeViewCollection.set({ ... })`. */
        export interface RangeViewCollectionUpdateData {
            items?: Excel.Interfaces.RangeViewData[];
        }
        /** An interface for updating data on the `SettingCollection` object, for use in `settingCollection.set({ ... })`. */
        export interface SettingCollectionUpdateData {
            items?: Excel.Interfaces.SettingData[];
        }
        /** An interface for updating data on the `Setting` object, for use in `setting.set({ ... })`. */
        export interface SettingUpdateData {
            
        }
        /** An interface for updating data on the `NamedItemCollection` object, for use in `namedItemCollection.set({ ... })`. */
        export interface NamedItemCollectionUpdateData {
            items?: Excel.Interfaces.NamedItemData[];
        }
        /** An interface for updating data on the `NamedItem` object, for use in `namedItem.set({ ... })`. */
        export interface NamedItemUpdateData {
            
            
            /**
             * Specifies if the object is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the `BindingCollection` object, for use in `bindingCollection.set({ ... })`. */
        export interface BindingCollectionUpdateData {
            items?: Excel.Interfaces.BindingData[];
        }
        /** An interface for updating data on the `TableCollection` object, for use in `tableCollection.set({ ... })`. */
        export interface TableCollectionUpdateData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface for updating data on the `TableScopedCollection` object, for use in `tableScopedCollection.set({ ... })`. */
        export interface TableScopedCollectionUpdateData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface for updating data on the `Table` object, for use in `table.set({ ... })`. */
        export interface TableUpdateData {
            
            
            /**
             * Name of the table.
                        
                         The set name of the table must follow the guidelines specified in the {@link https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114 | Rename an Excel table} article.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            /**
             * Specifies if the header row is visible. This value can be set to show or remove the header row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showHeaders?: boolean;
            /**
             * Specifies if the total row is visible. This value can be set to show or remove the total row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showTotals?: boolean;
            /**
             * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: string;
        }
        /** An interface for updating data on the `TableColumnCollection` object, for use in `tableColumnCollection.set({ ... })`. */
        export interface TableColumnCollectionUpdateData {
            items?: Excel.Interfaces.TableColumnData[];
        }
        /** An interface for updating data on the `TableColumn` object, for use in `tableColumn.set({ ... })`. */
        export interface TableColumnUpdateData {
            /**
             * Specifies the name of the table column.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
             */
            name?: string;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
        }
        /** An interface for updating data on the `TableRowCollection` object, for use in `tableRowCollection.set({ ... })`. */
        export interface TableRowCollectionUpdateData {
            items?: Excel.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the `TableRow` object, for use in `tableRow.set({ ... })`. */
        export interface TableRowUpdateData {
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
        }
        /** An interface for updating data on the `DataValidation` object, for use in `dataValidation.set({ ... })`. */
        export interface DataValidationUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `RangeFormat` object, for use in `rangeFormat.set({ ... })`. */
        export interface RangeFormatUpdateData {
            /**
            * Collection of border objects that apply to the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            borders?: Excel.Interfaces.RangeBorderCollectionUpdateData;
            /**
            * Returns the fill object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            fill?: Excel.Interfaces.RangeFillUpdateData;
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontUpdateData;
            
            
            
            /**
             * Represents the horizontal alignment for the specified object. See `Excel.HorizontalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            horizontalAlignment?: Excel.HorizontalAlignment | "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed";
            
            
            
            
            
            
            
            /**
             * Represents the vertical alignment for the specified object. See `Excel.VerticalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            verticalAlignment?: Excel.VerticalAlignment | "Top" | "Center" | "Bottom" | "Justify" | "Distributed";
            /**
             * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            wrapText?: boolean;
        }
        /** An interface for updating data on the `FormatProtection` object, for use in `formatProtection.set({ ... })`. */
        export interface FormatProtectionUpdateData {
            
            
        }
        /** An interface for updating data on the `RangeFill` object, for use in `rangeFill.set({ ... })`. */
        export interface RangeFillUpdateData {
            /**
             * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            
            
            
            
        }
        /** An interface for updating data on the `RangeBorder` object, for use in `rangeBorder.set({ ... })`. */
        export interface RangeBorderUpdateData {
            /**
             * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
            
            /**
             * Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            weight?: Excel.BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";
        }
        /** An interface for updating data on the `RangeBorderCollection` object, for use in `rangeBorderCollection.set({ ... })`. */
        export interface RangeBorderCollectionUpdateData {
            
            items?: Excel.Interfaces.RangeBorderData[];
        }
        /** An interface for updating data on the `RangeFont` object, for use in `rangeFont.set({ ... })`. */
        export interface RangeFontUpdateData {
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            
            
            
            
            /**
             * Type of underline applied to the font. See `Excel.RangeUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.RangeUnderlineStyle | "None" | "Single" | "Double" | "SingleAccountant" | "DoubleAccountant";
        }
        /** An interface for updating data on the `ChartCollection` object, for use in `chartCollection.set({ ... })`. */
        export interface ChartCollectionUpdateData {
            items?: Excel.Interfaces.ChartData[];
        }
        /** An interface for updating data on the `Chart` object, for use in `chart.set({ ... })`. */
        export interface ChartUpdateData {
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesUpdateData;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsUpdateData;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatUpdateData;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendUpdateData;
            
            
            /**
            * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartTitleUpdateData;
            
            
            
            /**
             * Specifies the height, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            height?: number;
            /**
             * The distance, in points, from the left side of the chart to the worksheet origin.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            left?: number;
            /**
             * Specifies the name of a chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            /**
             * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            top?: number;
            /**
             * Specifies the width, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the `ChartPivotOptions` object, for use in `chartPivotOptions.set({ ... })`. */
        export interface ChartPivotOptionsUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ChartAreaFormat` object, for use in `chartAreaFormat.set({ ... })`. */
        export interface ChartAreaFormatUpdateData {
            
            /**
            * Represents the font attributes (font name, font size, color, etc.) for the current object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
            
            
        }
        /** An interface for updating data on the `ChartSeriesCollection` object, for use in `chartSeriesCollection.set({ ... })`. */
        export interface ChartSeriesCollectionUpdateData {
            items?: Excel.Interfaces.ChartSeriesData[];
        }
        /** An interface for updating data on the `ChartSeries` object, for use in `chartSeries.set({ ... })`. */
        export interface ChartSeriesUpdateData {
            
            
            
            /**
            * Represents the formatting of a chart series, which includes fill and line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartSeriesFormatUpdateData;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartSeriesFormat` object, for use in `chartSeriesFormat.set({ ... })`. */
        export interface ChartSeriesFormatUpdateData {
            /**
            * Represents line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatUpdateData;
        }
        /** An interface for updating data on the `ChartPointsCollection` object, for use in `chartPointsCollection.set({ ... })`. */
        export interface ChartPointsCollectionUpdateData {
            items?: Excel.Interfaces.ChartPointData[];
        }
        /** An interface for updating data on the `ChartPoint` object, for use in `chartPoint.set({ ... })`. */
        export interface ChartPointUpdateData {
            
            /**
            * Encapsulates the format properties chart point.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartPointFormatUpdateData;
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartPointFormat` object, for use in `chartPointFormat.set({ ... })`. */
        export interface ChartPointFormatUpdateData {
            
        }
        /** An interface for updating data on the `ChartAxes` object, for use in `chartAxes.set({ ... })`. */
        export interface ChartAxesUpdateData {
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisUpdateData;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisUpdateData;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisUpdateData;
        }
        /** An interface for updating data on the `ChartAxis` object, for use in `chartAxis.set({ ... })`. */
        export interface ChartAxisUpdateData {
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatUpdateData;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesUpdateData;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesUpdateData;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleUpdateData;
            
            
            
            
            
            
            
            
            
            /**
             * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string. The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            majorUnit?: any;
            /**
             * Represents the maximum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            maximum?: any;
            /**
             * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minimum?: any;
            
            
            /**
             * Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minorUnit?: any;
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartAxisFormat` object, for use in `chartAxisFormat.set({ ... })`. */
        export interface ChartAxisFormatUpdateData {
            /**
            * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
            /**
            * Specifies chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatUpdateData;
        }
        /** An interface for updating data on the `ChartAxisTitle` object, for use in `chartAxisTitle.set({ ... })`. */
        export interface ChartAxisTitleUpdateData {
            /**
            * Specifies the formatting of the chart axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisTitleFormatUpdateData;
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            /**
             * Specifies if the axis title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the `ChartAxisTitleFormat` object, for use in `chartAxisTitleFormat.set({ ... })`. */
        export interface ChartAxisTitleFormatUpdateData {
            
            /**
            * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
        }
        /** An interface for updating data on the `ChartDataLabels` object, for use in `chartDataLabels.set({ ... })`. */
        export interface ChartDataLabelsUpdateData {
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatUpdateData;
            
            
            
            
            /**
             * Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: Excel.ChartDataLabelPosition | "Invalid" | "None" | "Center" | "InsideEnd" | "InsideBase" | "OutsideEnd" | "Left" | "Right" | "Top" | "Bottom" | "BestFit" | "Callout";
            /**
             * String representing the separator used for the data labels on a chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            separator?: string;
            /**
             * Specifies if the data label bubble size is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showBubbleSize?: boolean;
            /**
             * Specifies if the data label category name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showCategoryName?: boolean;
            /**
             * Specifies if the data label legend key is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showLegendKey?: boolean;
            /**
             * Specifies if the data label percentage is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showPercentage?: boolean;
            /**
             * Specifies if the data label series name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showSeriesName?: boolean;
            /**
             * Specifies if the data label value is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showValue?: boolean;
            
            
        }
        /** An interface for updating data on the `ChartDataLabel` object, for use in `chartDataLabel.set({ ... })`. */
        export interface ChartDataLabelUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartDataLabelFormat` object, for use in `chartDataLabelFormat.set({ ... })`. */
        export interface ChartDataLabelFormatUpdateData {
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for a chart data label.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
        }
        /** An interface for updating data on the `ChartDataTable` object, for use in `chartDataTable.set({ ... })`. */
        export interface ChartDataTableUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartDataTableFormat` object, for use in `chartDataTableFormat.set({ ... })`. */
        export interface ChartDataTableFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `ChartErrorBars` object, for use in `chartErrorBars.set({ ... })`. */
        export interface ChartErrorBarsUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartErrorBarsFormat` object, for use in `chartErrorBarsFormat.set({ ... })`. */
        export interface ChartErrorBarsFormatUpdateData {
            
        }
        /** An interface for updating data on the `ChartGridlines` object, for use in `chartGridlines.set({ ... })`. */
        export interface ChartGridlinesUpdateData {
            /**
            * Represents the formatting of chart gridlines.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartGridlinesFormatUpdateData;
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the `ChartGridlinesFormat` object, for use in `chartGridlinesFormat.set({ ... })`. */
        export interface ChartGridlinesFormatUpdateData {
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatUpdateData;
        }
        /** An interface for updating data on the `ChartLegend` object, for use in `chartLegend.set({ ... })`. */
        export interface ChartLegendUpdateData {
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatUpdateData;
            
            
            /**
             * Specifies if the chart legend should overlap with the main body of the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            /**
             * Specifies the position of the legend on the chart. See `Excel.ChartLegendPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: Excel.ChartLegendPosition | "Invalid" | "Top" | "Bottom" | "Left" | "Right" | "Corner" | "Custom";
            
            
            /**
             * Specifies if the chart legend is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
            
        }
        /** An interface for updating data on the `ChartLegendEntry` object, for use in `chartLegendEntry.set({ ... })`. */
        export interface ChartLegendEntryUpdateData {
            
        }
        /** An interface for updating data on the `ChartLegendEntryCollection` object, for use in `chartLegendEntryCollection.set({ ... })`. */
        export interface ChartLegendEntryCollectionUpdateData {
            items?: Excel.Interfaces.ChartLegendEntryData[];
        }
        /** An interface for updating data on the `ChartLegendFormat` object, for use in `chartLegendFormat.set({ ... })`. */
        export interface ChartLegendFormatUpdateData {
            
            /**
            * Represents the font attributes such as font name, font size, and color of a chart legend.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
        }
        /** An interface for updating data on the `ChartMapOptions` object, for use in `chartMapOptions.set({ ... })`. */
        export interface ChartMapOptionsUpdateData {
            
            
            
        }
        /** An interface for updating data on the `ChartTitle` object, for use in `chartTitle.set({ ... })`. */
        export interface ChartTitleUpdateData {
            /**
            * Represents the formatting of a chart title, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartTitleFormatUpdateData;
            
            
            /**
             * Specifies if the chart title will overlay the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            
            
            /**
             * Specifies the chart's title text.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            
            
            /**
             * Specifies if the chart title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the `ChartFormatString` object, for use in `chartFormatString.set({ ... })`. */
        export interface ChartFormatStringUpdateData {
            
        }
        /** An interface for updating data on the `ChartTitleFormat` object, for use in `chartTitleFormat.set({ ... })`. */
        export interface ChartTitleFormatUpdateData {
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for an object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
        }
        /** An interface for updating data on the `ChartBorder` object, for use in `chartBorder.set({ ... })`. */
        export interface ChartBorderUpdateData {
            
            
            
        }
        /** An interface for updating data on the `ChartBinOptions` object, for use in `chartBinOptions.set({ ... })`. */
        export interface ChartBinOptionsUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartBoxwhiskerOptions` object, for use in `chartBoxwhiskerOptions.set({ ... })`. */
        export interface ChartBoxwhiskerOptionsUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartLineFormat` object, for use in `chartLineFormat.set({ ... })`. */
        export interface ChartLineFormatUpdateData {
            /**
             * HTML color code representing the color of lines in the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            
            
        }
        /** An interface for updating data on the `ChartFont` object, for use in `chartFont.set({ ... })`. */
        export interface ChartFontUpdateData {
            /**
             * Represents the bold status of font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.ChartUnderlineStyle | "None" | "Single";
        }
        /** An interface for updating data on the `ChartTrendline` object, for use in `chartTrendline.set({ ... })`. */
        export interface ChartTrendlineUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartTrendlineCollection` object, for use in `chartTrendlineCollection.set({ ... })`. */
        export interface ChartTrendlineCollectionUpdateData {
            items?: Excel.Interfaces.ChartTrendlineData[];
        }
        /** An interface for updating data on the `ChartTrendlineFormat` object, for use in `chartTrendlineFormat.set({ ... })`. */
        export interface ChartTrendlineFormatUpdateData {
            
        }
        /** An interface for updating data on the `ChartTrendlineLabel` object, for use in `chartTrendlineLabel.set({ ... })`. */
        export interface ChartTrendlineLabelUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartTrendlineLabelFormat` object, for use in `chartTrendlineLabelFormat.set({ ... })`. */
        export interface ChartTrendlineLabelFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `ChartPlotArea` object, for use in `chartPlotArea.set({ ... })`. */
        export interface ChartPlotAreaUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ChartPlotAreaFormat` object, for use in `chartPlotAreaFormat.set({ ... })`. */
        export interface ChartPlotAreaFormatUpdateData {
            
        }
        /** An interface for updating data on the `CustomXmlPartScopedCollection` object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `CustomXmlPartCollection` object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `PivotTableScopedCollection` object, for use in `pivotTableScopedCollection.set({ ... })`. */
        export interface PivotTableScopedCollectionUpdateData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface for updating data on the `PivotTableCollection` object, for use in `pivotTableCollection.set({ ... })`. */
        export interface PivotTableCollectionUpdateData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface for updating data on the `PivotTable` object, for use in `pivotTable.set({ ... })`. */
        export interface PivotTableUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `PivotLayout` object, for use in `pivotLayout.set({ ... })`. */
        export interface PivotLayoutUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `PivotHierarchyCollection` object, for use in `pivotHierarchyCollection.set({ ... })`. */
        export interface PivotHierarchyCollectionUpdateData {
            items?: Excel.Interfaces.PivotHierarchyData[];
        }
        /** An interface for updating data on the `PivotHierarchy` object, for use in `pivotHierarchy.set({ ... })`. */
        export interface PivotHierarchyUpdateData {
            
        }
        /** An interface for updating data on the `RowColumnPivotHierarchyCollection` object, for use in `rowColumnPivotHierarchyCollection.set({ ... })`. */
        export interface RowColumnPivotHierarchyCollectionUpdateData {
            items?: Excel.Interfaces.RowColumnPivotHierarchyData[];
        }
        /** An interface for updating data on the `RowColumnPivotHierarchy` object, for use in `rowColumnPivotHierarchy.set({ ... })`. */
        export interface RowColumnPivotHierarchyUpdateData {
            
            
        }
        /** An interface for updating data on the `FilterPivotHierarchyCollection` object, for use in `filterPivotHierarchyCollection.set({ ... })`. */
        export interface FilterPivotHierarchyCollectionUpdateData {
            items?: Excel.Interfaces.FilterPivotHierarchyData[];
        }
        /** An interface for updating data on the `FilterPivotHierarchy` object, for use in `filterPivotHierarchy.set({ ... })`. */
        export interface FilterPivotHierarchyUpdateData {
            
            
            
        }
        /** An interface for updating data on the `DataPivotHierarchyCollection` object, for use in `dataPivotHierarchyCollection.set({ ... })`. */
        export interface DataPivotHierarchyCollectionUpdateData {
            items?: Excel.Interfaces.DataPivotHierarchyData[];
        }
        /** An interface for updating data on the `DataPivotHierarchy` object, for use in `dataPivotHierarchy.set({ ... })`. */
        export interface DataPivotHierarchyUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `PivotFieldCollection` object, for use in `pivotFieldCollection.set({ ... })`. */
        export interface PivotFieldCollectionUpdateData {
            items?: Excel.Interfaces.PivotFieldData[];
        }
        /** An interface for updating data on the `PivotField` object, for use in `pivotField.set({ ... })`. */
        export interface PivotFieldUpdateData {
            
            
            
        }
        /** An interface for updating data on the `PivotItemCollection` object, for use in `pivotItemCollection.set({ ... })`. */
        export interface PivotItemCollectionUpdateData {
            items?: Excel.Interfaces.PivotItemData[];
        }
        /** An interface for updating data on the `PivotItem` object, for use in `pivotItem.set({ ... })`. */
        export interface PivotItemUpdateData {
            
            
            
        }
        /** An interface for updating data on the `WorksheetCustomProperty` object, for use in `worksheetCustomProperty.set({ ... })`. */
        export interface WorksheetCustomPropertyUpdateData {
            
        }
        /** An interface for updating data on the `WorksheetCustomPropertyCollection` object, for use in `worksheetCustomPropertyCollection.set({ ... })`. */
        export interface WorksheetCustomPropertyCollectionUpdateData {
            items?: Excel.Interfaces.WorksheetCustomPropertyData[];
        }
        /** An interface for updating data on the `DocumentProperties` object, for use in `documentProperties.set({ ... })`. */
        export interface DocumentPropertiesUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `CustomProperty` object, for use in `customProperty.set({ ... })`. */
        export interface CustomPropertyUpdateData {
            
        }
        /** An interface for updating data on the `CustomPropertyCollection` object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: Excel.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the `ConditionalFormatCollection` object, for use in `conditionalFormatCollection.set({ ... })`. */
        export interface ConditionalFormatCollectionUpdateData {
            items?: Excel.Interfaces.ConditionalFormatData[];
        }
        /** An interface for updating data on the `ConditionalFormat` object, for use in `conditionalFormat.set({ ... })`. */
        export interface ConditionalFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `DataBarConditionalFormat` object, for use in `dataBarConditionalFormat.set({ ... })`. */
        export interface DataBarConditionalFormatUpdateData {
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ConditionalDataBarPositiveFormat` object, for use in `conditionalDataBarPositiveFormat.set({ ... })`. */
        export interface ConditionalDataBarPositiveFormatUpdateData {
            
            
            
        }
        /** An interface for updating data on the `ConditionalDataBarNegativeFormat` object, for use in `conditionalDataBarNegativeFormat.set({ ... })`. */
        export interface ConditionalDataBarNegativeFormatUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `CustomConditionalFormat` object, for use in `customConditionalFormat.set({ ... })`. */
        export interface CustomConditionalFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `ConditionalFormatRule` object, for use in `conditionalFormatRule.set({ ... })`. */
        export interface ConditionalFormatRuleUpdateData {
            
            
            
        }
        /** An interface for updating data on the `IconSetConditionalFormat` object, for use in `iconSetConditionalFormat.set({ ... })`. */
        export interface IconSetConditionalFormatUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ColorScaleConditionalFormat` object, for use in `colorScaleConditionalFormat.set({ ... })`. */
        export interface ColorScaleConditionalFormatUpdateData {
            
        }
        /** An interface for updating data on the `TopBottomConditionalFormat` object, for use in `topBottomConditionalFormat.set({ ... })`. */
        export interface TopBottomConditionalFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `PresetCriteriaConditionalFormat` object, for use in `presetCriteriaConditionalFormat.set({ ... })`. */
        export interface PresetCriteriaConditionalFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `TextConditionalFormat` object, for use in `textConditionalFormat.set({ ... })`. */
        export interface TextConditionalFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `CellValueConditionalFormat` object, for use in `cellValueConditionalFormat.set({ ... })`. */
        export interface CellValueConditionalFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `ConditionalRangeFormat` object, for use in `conditionalRangeFormat.set({ ... })`. */
        export interface ConditionalRangeFormatUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ConditionalRangeFont` object, for use in `conditionalRangeFont.set({ ... })`. */
        export interface ConditionalRangeFontUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `ConditionalRangeFill` object, for use in `conditionalRangeFill.set({ ... })`. */
        export interface ConditionalRangeFillUpdateData {
            
        }
        /** An interface for updating data on the `ConditionalRangeBorder` object, for use in `conditionalRangeBorder.set({ ... })`. */
        export interface ConditionalRangeBorderUpdateData {
            
            
        }
        /** An interface for updating data on the `ConditionalRangeBorderCollection` object, for use in `conditionalRangeBorderCollection.set({ ... })`. */
        export interface ConditionalRangeBorderCollectionUpdateData {
            
            
            
            
            items?: Excel.Interfaces.ConditionalRangeBorderData[];
        }
        /** An interface for updating data on the `Style` object, for use in `style.set({ ... })`. */
        export interface StyleUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `StyleCollection` object, for use in `styleCollection.set({ ... })`. */
        export interface StyleCollectionUpdateData {
            items?: Excel.Interfaces.StyleData[];
        }
        /** An interface for updating data on the `TableStyleCollection` object, for use in `tableStyleCollection.set({ ... })`. */
        export interface TableStyleCollectionUpdateData {
            items?: Excel.Interfaces.TableStyleData[];
        }
        /** An interface for updating data on the `TableStyle` object, for use in `tableStyle.set({ ... })`. */
        export interface TableStyleUpdateData {
            
        }
        /** An interface for updating data on the `PivotTableStyleCollection` object, for use in `pivotTableStyleCollection.set({ ... })`. */
        export interface PivotTableStyleCollectionUpdateData {
            items?: Excel.Interfaces.PivotTableStyleData[];
        }
        /** An interface for updating data on the `PivotTableStyle` object, for use in `pivotTableStyle.set({ ... })`. */
        export interface PivotTableStyleUpdateData {
            
        }
        /** An interface for updating data on the `SlicerStyleCollection` object, for use in `slicerStyleCollection.set({ ... })`. */
        export interface SlicerStyleCollectionUpdateData {
            items?: Excel.Interfaces.SlicerStyleData[];
        }
        /** An interface for updating data on the `SlicerStyle` object, for use in `slicerStyle.set({ ... })`. */
        export interface SlicerStyleUpdateData {
            
        }
        /** An interface for updating data on the `TimelineStyleCollection` object, for use in `timelineStyleCollection.set({ ... })`. */
        export interface TimelineStyleCollectionUpdateData {
            items?: Excel.Interfaces.TimelineStyleData[];
        }
        /** An interface for updating data on the `TimelineStyle` object, for use in `timelineStyle.set({ ... })`. */
        export interface TimelineStyleUpdateData {
            
        }
        /** An interface for updating data on the `PageLayout` object, for use in `pageLayout.set({ ... })`. */
        export interface PageLayoutUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `HeaderFooter` object, for use in `headerFooter.set({ ... })`. */
        export interface HeaderFooterUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `HeaderFooterGroup` object, for use in `headerFooterGroup.set({ ... })`. */
        export interface HeaderFooterGroupUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `PageBreakCollection` object, for use in `pageBreakCollection.set({ ... })`. */
        export interface PageBreakCollectionUpdateData {
            items?: Excel.Interfaces.PageBreakData[];
        }
        /** An interface for updating data on the `RangeCollection` object, for use in `rangeCollection.set({ ... })`. */
        export interface RangeCollectionUpdateData {
            items?: Excel.Interfaces.RangeData[];
        }
        /** An interface for updating data on the `RangeAreasCollection` object, for use in `rangeAreasCollection.set({ ... })`. */
        export interface RangeAreasCollectionUpdateData {
            items?: Excel.Interfaces.RangeAreasData[];
        }
        /** An interface for updating data on the `CommentCollection` object, for use in `commentCollection.set({ ... })`. */
        export interface CommentCollectionUpdateData {
            items?: Excel.Interfaces.CommentData[];
        }
        /** An interface for updating data on the `Comment` object, for use in `comment.set({ ... })`. */
        export interface CommentUpdateData {
            
            
        }
        /** An interface for updating data on the `CommentReplyCollection` object, for use in `commentReplyCollection.set({ ... })`. */
        export interface CommentReplyCollectionUpdateData {
            items?: Excel.Interfaces.CommentReplyData[];
        }
        /** An interface for updating data on the `CommentReply` object, for use in `commentReply.set({ ... })`. */
        export interface CommentReplyUpdateData {
            
        }
        /** An interface for updating data on the `ShapeCollection` object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: Excel.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `Shape` object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `GroupShapeCollection` object, for use in `groupShapeCollection.set({ ... })`. */
        export interface GroupShapeCollectionUpdateData {
            items?: Excel.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `Line` object, for use in `line.set({ ... })`. */
        export interface LineUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            
            
        }
        /** An interface for updating data on the `ShapeLineFormat` object, for use in `shapeLineFormat.set({ ... })`. */
        export interface ShapeLineFormatUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TextRange` object, for use in `textRange.set({ ... })`. */
        export interface TextRangeUpdateData {
            
            
        }
        /** An interface for updating data on the `ShapeFont` object, for use in `shapeFont.set({ ... })`. */
        export interface ShapeFontUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Slicer` object, for use in `slicer.set({ ... })`. */
        export interface SlicerUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `SlicerCollection` object, for use in `slicerCollection.set({ ... })`. */
        export interface SlicerCollectionUpdateData {
            items?: Excel.Interfaces.SlicerData[];
        }
        /** An interface for updating data on the `SlicerItem` object, for use in `slicerItem.set({ ... })`. */
        export interface SlicerItemUpdateData {
            
        }
        /** An interface for updating data on the `SlicerItemCollection` object, for use in `slicerItemCollection.set({ ... })`. */
        export interface SlicerItemCollectionUpdateData {
            items?: Excel.Interfaces.SlicerItemData[];
        }
        /** An interface for updating data on the `NamedSheetView` object, for use in `namedSheetView.set({ ... })`. */
        export interface NamedSheetViewUpdateData {
            
        }
        /** An interface for updating data on the `NamedSheetViewCollection` object, for use in `namedSheetViewCollection.set({ ... })`. */
        export interface NamedSheetViewCollectionUpdateData {
            items?: Excel.Interfaces.NamedSheetViewData[];
        }
        /** An interface describing the data returned by calling `allowEditRange.toJSON()`. */
        export interface AllowEditRangeData {
            
            
            
        }
        /** An interface describing the data returned by calling `allowEditRangeCollection.toJSON()`. */
        export interface AllowEditRangeCollectionData {
            items?: Excel.Interfaces.AllowEditRangeData[];
        }
        /** An interface describing the data returned by calling `query.toJSON()`. */
        export interface QueryData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `queryCollection.toJSON()`. */
        export interface QueryCollectionData {
            items?: Excel.Interfaces.QueryData[];
        }
        /** An interface describing the data returned by calling `linkedWorkbook.toJSON()`. */
        export interface LinkedWorkbookData {
            
        }
        /** An interface describing the data returned by calling `linkedWorkbookCollection.toJSON()`. */
        export interface LinkedWorkbookCollectionData {
            items?: Excel.Interfaces.LinkedWorkbookData[];
        }
        /** An interface describing the data returned by calling `runtime.toJSON()`. */
        export interface RuntimeData {
            
        }
        /** An interface describing the data returned by calling `application.toJSON()`. */
        export interface ApplicationData {
            
            
            
            /**
             * Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for get, 1.8 for set]
             */
            calculationMode?: Excel.CalculationMode | "Automatic" | "AutomaticExceptTables" | "Manual";
            
            
            
            
        }
        /** An interface describing the data returned by calling `iterativeCalculation.toJSON()`. */
        export interface IterativeCalculationData {
            
            
            
        }
        /** An interface describing the data returned by calling `workbook.toJSON()`. */
        export interface WorkbookData {
            /**
            * Represents a collection of bindings that are part of the workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            bindings?: Excel.Interfaces.BindingData[];
            
            
            /**
            * Represents a collection of workbook-scoped named items (named ranges and constants).
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            names?: Excel.Interfaces.NamedItemData[];
            
            
            
            
            
            
            
            
            
            /**
            * Represents a collection of tables associated with the workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            tables?: Excel.Interfaces.TableData[];
            
            /**
            * Represents a collection of worksheets associated with the workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            worksheets?: Excel.Interfaces.WorksheetData[];
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `workbookProtection.toJSON()`. */
        export interface WorkbookProtectionData {
            
        }
        /** An interface describing the data returned by calling `workbookCreated.toJSON()`. */
        export interface WorkbookCreatedData {
        }
        /** An interface describing the data returned by calling `worksheet.toJSON()`. */
        export interface WorksheetData {
            
            /**
            * Returns a collection of charts that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            charts?: Excel.Interfaces.ChartData[];
            
            
            
            
            
            
            
            
            
            /**
            * Collection of tables that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            tables?: Excel.Interfaces.TableData[];
            
            
            /**
             * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: string;
            /**
             * The display name of the worksheet. The name must be fewer than 32 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * The zero-based position of the worksheet within the workbook.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: number;
            
            
            
            
            
            
            /**
             * The visibility of the worksheet.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
             */
            visibility?: Excel.SheetVisibility | "Visible" | "Hidden" | "VeryHidden";
        }
        /** An interface describing the data returned by calling `worksheetCollection.toJSON()`. */
        export interface WorksheetCollectionData {
            items?: Excel.Interfaces.WorksheetData[];
        }
        /** An interface describing the data returned by calling `worksheetProtection.toJSON()`. */
        export interface WorksheetProtectionData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `range.toJSON()`. */
        export interface RangeData {
            
            
            /**
            * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.RangeFormatData;
            /**
             * Specifies the range reference in A1-style. Address value contains the sheet reference (e.g., "Sheet1!A1:B4").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            address?: string;
            /**
             * Represents the range reference for the specified range in the language of the user.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            addressLocal?: string;
            /**
             * Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            cellCount?: number;
            /**
             * Specifies the total number of columns in the range.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            columnCount?: number;
            
            /**
             * Specifies the column number of the first cell in the range. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            columnIndex?: number;
            /**
             * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulas?: any[][];
            /**
             * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulasLocal?: any[][];
            
            
            
            
            
            
            
            
            
            /**
             * Represents Excel's number format code for the given range. For more information about Excel number formatting, see {@link https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68 | Number format codes}.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            numberFormat?: any[][];
            
            
            /**
             * Returns the total number of rows in the range.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            rowCount?: number;
            
            /**
             * Returns the row number of the first cell in the range. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            rowIndex?: number;
            
            
            /**
             * Text values of the specified range. The text value will not depend on the cell width. The number sign (#) substitution that happens in the Excel UI will not affect the text value returned by the API.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string[][];
            
            /**
             * Specifies the type of data in each cell.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            valueTypes?: Excel.RangeValueType[][];
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
            
        }
        /** An interface describing the data returned by calling `rangeAreas.toJSON()`. */
        export interface RangeAreasData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `workbookRangeAreas.toJSON()`. */
        export interface WorkbookRangeAreasData {
            
            
            
        }
        /** An interface describing the data returned by calling `rangeView.toJSON()`. */
        export interface RangeViewData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `rangeViewCollection.toJSON()`. */
        export interface RangeViewCollectionData {
            items?: Excel.Interfaces.RangeViewData[];
        }
        /** An interface describing the data returned by calling `settingCollection.toJSON()`. */
        export interface SettingCollectionData {
            items?: Excel.Interfaces.SettingData[];
        }
        /** An interface describing the data returned by calling `setting.toJSON()`. */
        export interface SettingData {
            
            
        }
        /** An interface describing the data returned by calling `namedItemCollection.toJSON()`. */
        export interface NamedItemCollectionData {
            items?: Excel.Interfaces.NamedItemData[];
        }
        /** An interface describing the data returned by calling `namedItem.toJSON()`. */
        export interface NamedItemData {
            
            
            
            /**
             * The name of the object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            /**
             * Specifies the type of the value returned by the name's formula. See `Excel.NamedItemType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for String,Integer,Double,Boolean,Range,Error; 1.7 for Array]
             */
            type?: Excel.NamedItemType | "String" | "Integer" | "Double" | "Boolean" | "Range" | "Error" | "Array";
            /**
             * Represents the value computed by the name's formula. For a named range, it will return the range address.
                        This API returns the #VALUE! error in the Excel UI if it refers to a user-defined function.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: any;
            
            
            /**
             * Specifies if the object is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `namedItemArrayValues.toJSON()`. */
        export interface NamedItemArrayValuesData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `binding.toJSON()`. */
        export interface BindingData {
            /**
             * Represents the binding identifier.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: string;
            /**
             * Returns the type of the binding. See `Excel.BindingType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            type?: Excel.BindingType | "Range" | "Table" | "Text";
        }
        /** An interface describing the data returned by calling `bindingCollection.toJSON()`. */
        export interface BindingCollectionData {
            items?: Excel.Interfaces.BindingData[];
        }
        /** An interface describing the data returned by calling `tableCollection.toJSON()`. */
        export interface TableCollectionData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `tableScopedCollection.toJSON()`. */
        export interface TableScopedCollectionData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `table.toJSON()`. */
        export interface TableData {
            
            /**
            * Represents a collection of all the columns in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            columns?: Excel.Interfaces.TableColumnData[];
            /**
            * Represents a collection of all the rows in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            rows?: Excel.Interfaces.TableRowData[];
            
            
            
            /**
             * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: string;
            
            /**
             * Name of the table.
                        
                         The set name of the table must follow the guidelines specified in the {@link https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114 | Rename an Excel table} article.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            /**
             * Specifies if the header row is visible. This value can be set to show or remove the header row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showHeaders?: boolean;
            /**
             * Specifies if the total row is visible. This value can be set to show or remove the total row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showTotals?: boolean;
            /**
             * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: string;
        }
        /** An interface describing the data returned by calling `tableColumnCollection.toJSON()`. */
        export interface TableColumnCollectionData {
            items?: Excel.Interfaces.TableColumnData[];
        }
        /** An interface describing the data returned by calling `tableColumn.toJSON()`. */
        export interface TableColumnData {
            
            /**
             * Returns a unique key that identifies the column within the table.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: number;
            /**
             * Returns the index number of the column within the columns collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: number;
            /**
             * Specifies the name of the table column.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
             */
            name?: string;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
        }
        /** An interface describing the data returned by calling `tableRowCollection.toJSON()`. */
        export interface TableRowCollectionData {
            items?: Excel.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling `tableRow.toJSON()`. */
        export interface TableRowData {
            /**
             * Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: number;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
            
            
        }
        /** An interface describing the data returned by calling `dataValidation.toJSON()`. */
        export interface DataValidationData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `removeDuplicatesResult.toJSON()`. */
        export interface RemoveDuplicatesResultData {
            
            
        }
        /** An interface describing the data returned by calling `rangeFormat.toJSON()`. */
        export interface RangeFormatData {
            /**
            * Collection of border objects that apply to the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            borders?: Excel.Interfaces.RangeBorderData[];
            /**
            * Returns the fill object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            fill?: Excel.Interfaces.RangeFillData;
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontData;
            
            
            
            /**
             * Represents the horizontal alignment for the specified object. See `Excel.HorizontalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            horizontalAlignment?: Excel.HorizontalAlignment | "General" | "Left" | "Center" | "Right" | "Fill" | "Justify" | "CenterAcrossSelection" | "Distributed";
            
            
            
            
            
            
            
            /**
             * Represents the vertical alignment for the specified object. See `Excel.VerticalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            verticalAlignment?: Excel.VerticalAlignment | "Top" | "Center" | "Bottom" | "Justify" | "Distributed";
            /**
             * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            wrapText?: boolean;
        }
        /** An interface describing the data returned by calling `formatProtection.toJSON()`. */
        export interface FormatProtectionData {
            
            
        }
        /** An interface describing the data returned by calling `rangeFill.toJSON()`. */
        export interface RangeFillData {
            /**
             * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            
            
            
            
        }
        /** An interface describing the data returned by calling `rangeBorder.toJSON()`. */
        export interface RangeBorderData {
            /**
             * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: Excel.BorderIndex | "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp";
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
            
            /**
             * Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            weight?: Excel.BorderWeight | "Hairline" | "Thin" | "Medium" | "Thick";
        }
        /** An interface describing the data returned by calling `rangeBorderCollection.toJSON()`. */
        export interface RangeBorderCollectionData {
            items?: Excel.Interfaces.RangeBorderData[];
        }
        /** An interface describing the data returned by calling `rangeFont.toJSON()`. */
        export interface RangeFontData {
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            
            
            
            
            /**
             * Type of underline applied to the font. See `Excel.RangeUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.RangeUnderlineStyle | "None" | "Single" | "Double" | "SingleAccountant" | "DoubleAccountant";
        }
        /** An interface describing the data returned by calling `chartCollection.toJSON()`. */
        export interface ChartCollectionData {
            items?: Excel.Interfaces.ChartData[];
        }
        /** An interface describing the data returned by calling `chart.toJSON()`. */
        export interface ChartData {
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesData;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsData;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatData;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendData;
            
            
            /**
            * Represents either a single series or collection of series in the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            series?: Excel.Interfaces.ChartSeriesData[];
            /**
            * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartTitleData;
            
            
            
            /**
             * Specifies the height, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            height?: number;
            
            /**
             * The distance, in points, from the left side of the chart to the worksheet origin.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            left?: number;
            /**
             * Specifies the name of a chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            /**
             * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            top?: number;
            /**
             * Specifies the width, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling `chartPivotOptions.toJSON()`. */
        export interface ChartPivotOptionsData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartAreaFormat.toJSON()`. */
        export interface ChartAreaFormatData {
            
            /**
            * Represents the font attributes (font name, font size, color, etc.) for the current object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
            
            
        }
        /** An interface describing the data returned by calling `chartSeriesCollection.toJSON()`. */
        export interface ChartSeriesCollectionData {
            items?: Excel.Interfaces.ChartSeriesData[];
        }
        /** An interface describing the data returned by calling `chartSeries.toJSON()`. */
        export interface ChartSeriesData {
            
            
            
            /**
            * Represents the formatting of a chart series, which includes fill and line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartSeriesFormatData;
            
            /**
            * Returns a collection of all points in the series.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            points?: Excel.Interfaces.ChartPointData[];
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartSeriesFormat.toJSON()`. */
        export interface ChartSeriesFormatData {
            /**
            * Represents line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatData;
        }
        /** An interface describing the data returned by calling `chartPointsCollection.toJSON()`. */
        export interface ChartPointsCollectionData {
            items?: Excel.Interfaces.ChartPointData[];
        }
        /** An interface describing the data returned by calling `chartPoint.toJSON()`. */
        export interface ChartPointData {
            
            /**
            * Encapsulates the format properties chart point.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartPointFormatData;
            
            
            
            
            
            /**
             * Returns the value of a chart point.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: any;
        }
        /** An interface describing the data returned by calling `chartPointFormat.toJSON()`. */
        export interface ChartPointFormatData {
            
        }
        /** An interface describing the data returned by calling `chartAxes.toJSON()`. */
        export interface ChartAxesData {
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisData;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisData;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisData;
        }
        /** An interface describing the data returned by calling `chartAxis.toJSON()`. */
        export interface ChartAxisData {
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatData;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesData;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesData;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleData;
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string. The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            majorUnit?: any;
            /**
             * Represents the maximum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            maximum?: any;
            /**
             * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minimum?: any;
            
            
            /**
             * Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minorUnit?: any;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartAxisFormat.toJSON()`. */
        export interface ChartAxisFormatData {
            /**
            * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
            /**
            * Specifies chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatData;
        }
        /** An interface describing the data returned by calling `chartAxisTitle.toJSON()`. */
        export interface ChartAxisTitleData {
            /**
            * Specifies the formatting of the chart axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisTitleFormatData;
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            /**
             * Specifies if the axis title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `chartAxisTitleFormat.toJSON()`. */
        export interface ChartAxisTitleFormatData {
            
            /**
            * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
        }
        /** An interface describing the data returned by calling `chartDataLabels.toJSON()`. */
        export interface ChartDataLabelsData {
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatData;
            
            
            
            
            /**
             * Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: Excel.ChartDataLabelPosition | "Invalid" | "None" | "Center" | "InsideEnd" | "InsideBase" | "OutsideEnd" | "Left" | "Right" | "Top" | "Bottom" | "BestFit" | "Callout";
            /**
             * String representing the separator used for the data labels on a chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            separator?: string;
            /**
             * Specifies if the data label bubble size is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showBubbleSize?: boolean;
            /**
             * Specifies if the data label category name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showCategoryName?: boolean;
            /**
             * Specifies if the data label legend key is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showLegendKey?: boolean;
            /**
             * Specifies if the data label percentage is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showPercentage?: boolean;
            /**
             * Specifies if the data label series name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showSeriesName?: boolean;
            /**
             * Specifies if the data label value is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showValue?: boolean;
            
            
        }
        /** An interface describing the data returned by calling `chartDataLabel.toJSON()`. */
        export interface ChartDataLabelData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartDataLabelFormat.toJSON()`. */
        export interface ChartDataLabelFormatData {
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for a chart data label.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
        }
        /** An interface describing the data returned by calling `chartDataTable.toJSON()`. */
        export interface ChartDataTableData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartDataTableFormat.toJSON()`. */
        export interface ChartDataTableFormatData {
            
            
        }
        /** An interface describing the data returned by calling `chartErrorBars.toJSON()`. */
        export interface ChartErrorBarsData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartErrorBarsFormat.toJSON()`. */
        export interface ChartErrorBarsFormatData {
            
        }
        /** An interface describing the data returned by calling `chartGridlines.toJSON()`. */
        export interface ChartGridlinesData {
            /**
            * Represents the formatting of chart gridlines.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartGridlinesFormatData;
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `chartGridlinesFormat.toJSON()`. */
        export interface ChartGridlinesFormatData {
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatData;
        }
        /** An interface describing the data returned by calling `chartLegend.toJSON()`. */
        export interface ChartLegendData {
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatData;
            
            
            
            /**
             * Specifies if the chart legend should overlap with the main body of the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            /**
             * Specifies the position of the legend on the chart. See `Excel.ChartLegendPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: Excel.ChartLegendPosition | "Invalid" | "Top" | "Bottom" | "Left" | "Right" | "Corner" | "Custom";
            
            
            /**
             * Specifies if the chart legend is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
            
        }
        /** An interface describing the data returned by calling `chartLegendEntry.toJSON()`. */
        export interface ChartLegendEntryData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartLegendEntryCollection.toJSON()`. */
        export interface ChartLegendEntryCollectionData {
            items?: Excel.Interfaces.ChartLegendEntryData[];
        }
        /** An interface describing the data returned by calling `chartLegendFormat.toJSON()`. */
        export interface ChartLegendFormatData {
            
            /**
            * Represents the font attributes such as font name, font size, and color of a chart legend.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
        }
        /** An interface describing the data returned by calling `chartMapOptions.toJSON()`. */
        export interface ChartMapOptionsData {
            
            
            
        }
        /** An interface describing the data returned by calling `chartTitle.toJSON()`. */
        export interface ChartTitleData {
            /**
            * Represents the formatting of a chart title, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartTitleFormatData;
            
            
            
            /**
             * Specifies if the chart title will overlay the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            
            
            /**
             * Specifies the chart's title text.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            
            
            /**
             * Specifies if the chart title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
            
        }
        /** An interface describing the data returned by calling `chartFormatString.toJSON()`. */
        export interface ChartFormatStringData {
            
        }
        /** An interface describing the data returned by calling `chartTitleFormat.toJSON()`. */
        export interface ChartTitleFormatData {
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for an object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
        }
        /** An interface describing the data returned by calling `chartBorder.toJSON()`. */
        export interface ChartBorderData {
            
            
            
        }
        /** An interface describing the data returned by calling `chartBinOptions.toJSON()`. */
        export interface ChartBinOptionsData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartBoxwhiskerOptions.toJSON()`. */
        export interface ChartBoxwhiskerOptionsData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartLineFormat.toJSON()`. */
        export interface ChartLineFormatData {
            /**
             * HTML color code representing the color of lines in the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            
            
        }
        /** An interface describing the data returned by calling `chartFont.toJSON()`. */
        export interface ChartFontData {
            /**
             * Represents the bold status of font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.ChartUnderlineStyle | "None" | "Single";
        }
        /** An interface describing the data returned by calling `chartTrendline.toJSON()`. */
        export interface ChartTrendlineData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartTrendlineCollection.toJSON()`. */
        export interface ChartTrendlineCollectionData {
            items?: Excel.Interfaces.ChartTrendlineData[];
        }
        /** An interface describing the data returned by calling `chartTrendlineFormat.toJSON()`. */
        export interface ChartTrendlineFormatData {
            
        }
        /** An interface describing the data returned by calling `chartTrendlineLabel.toJSON()`. */
        export interface ChartTrendlineLabelData {
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartTrendlineLabelFormat.toJSON()`. */
        export interface ChartTrendlineLabelFormatData {
            
            
        }
        /** An interface describing the data returned by calling `chartPlotArea.toJSON()`. */
        export interface ChartPlotAreaData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `chartPlotAreaFormat.toJSON()`. */
        export interface ChartPlotAreaFormatData {
            
        }
        /** An interface describing the data returned by calling `tableSort.toJSON()`. */
        export interface TableSortData {
            
            
            
        }
        /** An interface describing the data returned by calling `filter.toJSON()`. */
        export interface FilterData {
            
        }
        /** An interface describing the data returned by calling `autoFilter.toJSON()`. */
        export interface AutoFilterData {
            
            
            
        }
        /** An interface describing the data returned by calling `cultureInfo.toJSON()`. */
        export interface CultureInfoData {
            
            
            
        }
        /** An interface describing the data returned by calling `numberFormatInfo.toJSON()`. */
        export interface NumberFormatInfoData {
            
            
            
        }
        /** An interface describing the data returned by calling `datetimeFormatInfo.toJSON()`. */
        export interface DatetimeFormatInfoData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `customXmlPartScopedCollection.toJSON()`. */
        export interface CustomXmlPartScopedCollectionData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `customXmlPartCollection.toJSON()`. */
        export interface CustomXmlPartCollectionData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `customXmlPart.toJSON()`. */
        export interface CustomXmlPartData {
            
            
        }
        /** An interface describing the data returned by calling `pivotTableScopedCollection.toJSON()`. */
        export interface PivotTableScopedCollectionData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface describing the data returned by calling `pivotTableCollection.toJSON()`. */
        export interface PivotTableCollectionData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface describing the data returned by calling `pivotTable.toJSON()`. */
        export interface PivotTableData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pivotLayout.toJSON()`. */
        export interface PivotLayoutData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pivotHierarchyCollection.toJSON()`. */
        export interface PivotHierarchyCollectionData {
            items?: Excel.Interfaces.PivotHierarchyData[];
        }
        /** An interface describing the data returned by calling `pivotHierarchy.toJSON()`. */
        export interface PivotHierarchyData {
            
            
            
        }
        /** An interface describing the data returned by calling `rowColumnPivotHierarchyCollection.toJSON()`. */
        export interface RowColumnPivotHierarchyCollectionData {
            items?: Excel.Interfaces.RowColumnPivotHierarchyData[];
        }
        /** An interface describing the data returned by calling `rowColumnPivotHierarchy.toJSON()`. */
        export interface RowColumnPivotHierarchyData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `filterPivotHierarchyCollection.toJSON()`. */
        export interface FilterPivotHierarchyCollectionData {
            items?: Excel.Interfaces.FilterPivotHierarchyData[];
        }
        /** An interface describing the data returned by calling `filterPivotHierarchy.toJSON()`. */
        export interface FilterPivotHierarchyData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `dataPivotHierarchyCollection.toJSON()`. */
        export interface DataPivotHierarchyCollectionData {
            items?: Excel.Interfaces.DataPivotHierarchyData[];
        }
        /** An interface describing the data returned by calling `dataPivotHierarchy.toJSON()`. */
        export interface DataPivotHierarchyData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pivotFieldCollection.toJSON()`. */
        export interface PivotFieldCollectionData {
            items?: Excel.Interfaces.PivotFieldData[];
        }
        /** An interface describing the data returned by calling `pivotField.toJSON()`. */
        export interface PivotFieldData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pivotItemCollection.toJSON()`. */
        export interface PivotItemCollectionData {
            items?: Excel.Interfaces.PivotItemData[];
        }
        /** An interface describing the data returned by calling `pivotItem.toJSON()`. */
        export interface PivotItemData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `worksheetCustomProperty.toJSON()`. */
        export interface WorksheetCustomPropertyData {
            
            
        }
        /** An interface describing the data returned by calling `worksheetCustomPropertyCollection.toJSON()`. */
        export interface WorksheetCustomPropertyCollectionData {
            items?: Excel.Interfaces.WorksheetCustomPropertyData[];
        }
        /** An interface describing the data returned by calling `documentProperties.toJSON()`. */
        export interface DocumentPropertiesData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `customProperty.toJSON()`. */
        export interface CustomPropertyData {
            
            
            
        }
        /** An interface describing the data returned by calling `customPropertyCollection.toJSON()`. */
        export interface CustomPropertyCollectionData {
            items?: Excel.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling `conditionalFormatCollection.toJSON()`. */
        export interface ConditionalFormatCollectionData {
            items?: Excel.Interfaces.ConditionalFormatData[];
        }
        /** An interface describing the data returned by calling `conditionalFormat.toJSON()`. */
        export interface ConditionalFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `dataBarConditionalFormat.toJSON()`. */
        export interface DataBarConditionalFormatData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalDataBarPositiveFormat.toJSON()`. */
        export interface ConditionalDataBarPositiveFormatData {
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalDataBarNegativeFormat.toJSON()`. */
        export interface ConditionalDataBarNegativeFormatData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `customConditionalFormat.toJSON()`. */
        export interface CustomConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `conditionalFormatRule.toJSON()`. */
        export interface ConditionalFormatRuleData {
            
            
            
        }
        /** An interface describing the data returned by calling `iconSetConditionalFormat.toJSON()`. */
        export interface IconSetConditionalFormatData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `colorScaleConditionalFormat.toJSON()`. */
        export interface ColorScaleConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `topBottomConditionalFormat.toJSON()`. */
        export interface TopBottomConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `presetCriteriaConditionalFormat.toJSON()`. */
        export interface PresetCriteriaConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `textConditionalFormat.toJSON()`. */
        export interface TextConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `cellValueConditionalFormat.toJSON()`. */
        export interface CellValueConditionalFormatData {
            
            
        }
        /** An interface describing the data returned by calling `conditionalRangeFormat.toJSON()`. */
        export interface ConditionalRangeFormatData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalRangeFont.toJSON()`. */
        export interface ConditionalRangeFontData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalRangeFill.toJSON()`. */
        export interface ConditionalRangeFillData {
            
        }
        /** An interface describing the data returned by calling `conditionalRangeBorder.toJSON()`. */
        export interface ConditionalRangeBorderData {
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalRangeBorderCollection.toJSON()`. */
        export interface ConditionalRangeBorderCollectionData {
            items?: Excel.Interfaces.ConditionalRangeBorderData[];
        }
        /** An interface describing the data returned by calling `style.toJSON()`. */
        export interface StyleData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `styleCollection.toJSON()`. */
        export interface StyleCollectionData {
            items?: Excel.Interfaces.StyleData[];
        }
        /** An interface describing the data returned by calling `tableStyleCollection.toJSON()`. */
        export interface TableStyleCollectionData {
            items?: Excel.Interfaces.TableStyleData[];
        }
        /** An interface describing the data returned by calling `tableStyle.toJSON()`. */
        export interface TableStyleData {
            
            
        }
        /** An interface describing the data returned by calling `pivotTableStyleCollection.toJSON()`. */
        export interface PivotTableStyleCollectionData {
            items?: Excel.Interfaces.PivotTableStyleData[];
        }
        /** An interface describing the data returned by calling `pivotTableStyle.toJSON()`. */
        export interface PivotTableStyleData {
            
            
        }
        /** An interface describing the data returned by calling `slicerStyleCollection.toJSON()`. */
        export interface SlicerStyleCollectionData {
            items?: Excel.Interfaces.SlicerStyleData[];
        }
        /** An interface describing the data returned by calling `slicerStyle.toJSON()`. */
        export interface SlicerStyleData {
            
            
        }
        /** An interface describing the data returned by calling `timelineStyleCollection.toJSON()`. */
        export interface TimelineStyleCollectionData {
            items?: Excel.Interfaces.TimelineStyleData[];
        }
        /** An interface describing the data returned by calling `timelineStyle.toJSON()`. */
        export interface TimelineStyleData {
            
            
        }
        /** An interface describing the data returned by calling `pageLayout.toJSON()`. */
        export interface PageLayoutData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `headerFooter.toJSON()`. */
        export interface HeaderFooterData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `headerFooterGroup.toJSON()`. */
        export interface HeaderFooterGroupData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pageBreak.toJSON()`. */
        export interface PageBreakData {
            
            
        }
        /** An interface describing the data returned by calling `pageBreakCollection.toJSON()`. */
        export interface PageBreakCollectionData {
            items?: Excel.Interfaces.PageBreakData[];
        }
        /** An interface describing the data returned by calling `rangeCollection.toJSON()`. */
        export interface RangeCollectionData {
            items?: Excel.Interfaces.RangeData[];
        }
        /** An interface describing the data returned by calling `rangeAreasCollection.toJSON()`. */
        export interface RangeAreasCollectionData {
            items?: Excel.Interfaces.RangeAreasData[];
        }
        /** An interface describing the data returned by calling `commentCollection.toJSON()`. */
        export interface CommentCollectionData {
            items?: Excel.Interfaces.CommentData[];
        }
        /** An interface describing the data returned by calling `comment.toJSON()`. */
        export interface CommentData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `commentReplyCollection.toJSON()`. */
        export interface CommentReplyCollectionData {
            items?: Excel.Interfaces.CommentReplyData[];
        }
        /** An interface describing the data returned by calling `commentReply.toJSON()`. */
        export interface CommentReplyData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: Excel.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `geometricShape.toJSON()`. */
        export interface GeometricShapeData {
            
        }
        /** An interface describing the data returned by calling `image.toJSON()`. */
        export interface ImageData {
            
            
        }
        /** An interface describing the data returned by calling `shapeGroup.toJSON()`. */
        export interface ShapeGroupData {
            
            
        }
        /** An interface describing the data returned by calling `groupShapeCollection.toJSON()`. */
        export interface GroupShapeCollectionData {
            items?: Excel.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `line.toJSON()`. */
        export interface LineData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            
            
            
        }
        /** An interface describing the data returned by calling `shapeLineFormat.toJSON()`. */
        export interface ShapeLineFormatData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `textRange.toJSON()`. */
        export interface TextRangeData {
            
            
        }
        /** An interface describing the data returned by calling `shapeFont.toJSON()`. */
        export interface ShapeFontData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `slicer.toJSON()`. */
        export interface SlicerData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `slicerCollection.toJSON()`. */
        export interface SlicerCollectionData {
            items?: Excel.Interfaces.SlicerData[];
        }
        /** An interface describing the data returned by calling `slicerItem.toJSON()`. */
        export interface SlicerItemData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `slicerItemCollection.toJSON()`. */
        export interface SlicerItemCollectionData {
            items?: Excel.Interfaces.SlicerItemData[];
        }
        /** An interface describing the data returned by calling `namedSheetView.toJSON()`. */
        export interface NamedSheetViewData {
            
        }
        /** An interface describing the data returned by calling `namedSheetViewCollection.toJSON()`. */
        export interface NamedSheetViewCollectionData {
            items?: Excel.Interfaces.NamedSheetViewData[];
        }
        /** An interface describing the data returned by calling `functionResult.toJSON()`. */
        export interface FunctionResultData<T> {
            
            
        }
        
        
        
        
        
        
        
        /**
         * Represents the Excel application that manages the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ApplicationLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
             * Returns the calculation mode used in the workbook, as defined by the constants in `Excel.CalculationMode`. Possible values are: `Automatic`, where Excel controls recalculation; `AutomaticExceptTables`, where Excel controls recalculation but ignores changes in tables; `Manual`, where calculation is done when the user requests it.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for get, 1.8 for set]
             */
            calculationMode?: boolean;
            
            
            
            
        }
        
        /**
         * Workbook is the top level object which contains related workbook objects such as worksheets, tables, and ranges.
                    To learn more about the workbook object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks | Work with workbooks using the Excel JavaScript API}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface WorkbookLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the Excel application instance that contains this workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            application?: Excel.Interfaces.ApplicationLoadOptions;
            /**
            * Represents a collection of bindings that are part of the workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            bindings?: Excel.Interfaces.BindingCollectionLoadOptions;
            
            
            
            /**
            * Represents a collection of tables associated with the workbook.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            tables?: Excel.Interfaces.TableCollectionLoadOptions;
            
            
            
            
            
            
            
            
        }
        
        /**
         * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
                    To learn more about the worksheet object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets | Work with worksheets using the Excel JavaScript API}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface WorksheetLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Returns a collection of charts that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            charts?: Excel.Interfaces.ChartCollectionLoadOptions;
            
            
            /**
            * Collection of tables that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            tables?: Excel.Interfaces.TableCollectionLoadOptions;
            
            /**
             * Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * The display name of the worksheet. The name must be fewer than 32 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * The zero-based position of the worksheet within the workbook.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: boolean;
            
            
            
            
            
            
            /**
             * The visibility of the worksheet.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
             */
            visibility?: boolean;
        }
        /**
         * Represents a collection of worksheet objects that are part of the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface WorksheetCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * For EACH ITEM in the collection: Returns a collection of charts that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            charts?: Excel.Interfaces.ChartCollectionLoadOptions;
            
            
            /**
            * For EACH ITEM in the collection: Collection of tables that are part of the worksheet.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            tables?: Excel.Interfaces.TableCollectionLoadOptions;
            
            /**
             * For EACH ITEM in the collection: Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: The display name of the worksheet. The name must be fewer than 32 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * For EACH ITEM in the collection: The zero-based position of the worksheet within the workbook.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: boolean;
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: The visibility of the worksheet.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for reading visibility; 1.2 for setting it.]
             */
            visibility?: boolean;
        }
        
        /**
         * Range represents a set of one or more contiguous cells such as a cell, a row, a column, or a block of cells.
                    To learn more about how ranges are used throughout the API, start with {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-core-concepts#ranges | Ranges in the Excel JavaScript API}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.RangeFormatLoadOptions;
            /**
            * The worksheet containing the current range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            worksheet?: Excel.Interfaces.WorksheetLoadOptions;
            /**
             * Specifies the range reference in A1-style. Address value contains the sheet reference (e.g., "Sheet1!A1:B4").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            address?: boolean;
            /**
             * Represents the range reference for the specified range in the language of the user.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            addressLocal?: boolean;
            /**
             * Specifies the number of cells in the range. This API will return -1 if the cell count exceeds 2^31-1 (2,147,483,647).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            cellCount?: boolean;
            /**
             * Specifies the total number of columns in the range.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            columnCount?: boolean;
            
            /**
             * Specifies the column number of the first cell in the range. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            columnIndex?: boolean;
            /**
             * Represents the formula in A1-style notation. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulas?: boolean;
            /**
             * Represents the formula in A1-style notation, in the user's language and number-formatting locale. For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German. If a cell has no formula, its value is returned instead.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            formulasLocal?: boolean;
            
            
            
            
            
            
            
            
            
            /**
             * Represents Excel's number format code for the given range. For more information about Excel number formatting, see {@link https://support.microsoft.com/office/5026bbd6-04bc-48cd-bf33-80f18b4eae68 | Number format codes}.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            numberFormat?: boolean;
            
            
            /**
             * Returns the total number of rows in the range.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            rowCount?: boolean;
            
            /**
             * Returns the row number of the first cell in the range. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            rowIndex?: boolean;
            
            
            /**
             * Text values of the specified range. The text value will not depend on the cell width. The number sign (#) substitution that happens in the Excel UI will not affect the text value returned by the API.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: boolean;
            
            /**
             * Specifies the type of data in each cell.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            valueTypes?: boolean;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
            
            
            
        }
        
        
        
        
        
        
        /**
         * A collection of all the `NamedItem` objects that are part of the workbook or worksheet, depending on how it was reached.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface NamedItemCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: The name of the object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            /**
             * For EACH ITEM in the collection: Specifies the type of the value returned by the name's formula. See `Excel.NamedItemType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for String,Integer,Double,Boolean,Range,Error; 1.7 for Array]
             */
            type?: boolean;
            /**
             * For EACH ITEM in the collection: Represents the value computed by the name's formula. For a named range, it will return the range address.
                        This API returns the #VALUE! error in the Excel UI if it refers to a user-defined function.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: boolean;
            
            
            /**
             * For EACH ITEM in the collection: Specifies if the object is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /**
         * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface NamedItemLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            
            
            /**
             * The name of the object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            /**
             * Specifies the type of the value returned by the name's formula. See `Excel.NamedItemType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for String,Integer,Double,Boolean,Range,Error; 1.7 for Array]
             */
            type?: boolean;
            /**
             * Represents the value computed by the name's formula. For a named range, it will return the range address.
                        This API returns the #VALUE! error in the Excel UI if it refers to a user-defined function.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: boolean;
            
            
            /**
             * Specifies if the object is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        
        /**
         * Represents an Office.js binding that is defined in the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface BindingLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Represents the binding identifier.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * Returns the type of the binding. See `Excel.BindingType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Represents the collection of all the binding objects that are part of the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface BindingCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Represents the binding identifier.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the type of the binding. See `Excel.BindingType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * For EACH ITEM in the collection: Represents a collection of all the columns in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            columns?: Excel.Interfaces.TableColumnCollectionLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents a collection of all the rows in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            rows?: Excel.Interfaces.TableRowCollectionLoadOptions;
            
            
            
            
            /**
             * For EACH ITEM in the collection: Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            
            /**
             * For EACH ITEM in the collection: Name of the table.
                        
                         The set name of the table must follow the guidelines specified in the {@link https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114 | Rename an Excel table} article.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies if the header row is visible. This value can be set to show or remove the header row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showHeaders?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies if the total row is visible. This value can be set to show or remove the total row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showTotals?: boolean;
            /**
             * For EACH ITEM in the collection: Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
        }
        
        /**
         * Represents an Excel table.
                    To learn more about the table object model, read {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables | Work with tables using the Excel JavaScript API}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Represents a collection of all the columns in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            columns?: Excel.Interfaces.TableColumnCollectionLoadOptions;
            /**
            * Represents a collection of all the rows in the table.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            rows?: Excel.Interfaces.TableRowCollectionLoadOptions;
            
            
            
            
            /**
             * Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            
            /**
             * Name of the table.
                        
                         The set name of the table must follow the guidelines specified in the {@link https://support.microsoft.com/office/fbf49a4f-82a3-43eb-8ba2-44d21233b114 | Rename an Excel table} article.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            /**
             * Specifies if the header row is visible. This value can be set to show or remove the header row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showHeaders?: boolean;
            /**
             * Specifies if the total row is visible. This value can be set to show or remove the total row.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showTotals?: boolean;
            /**
             * Constant value that represents the table style. Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11". A custom user-defined style present in the workbook can also be specified.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
        }
        /**
         * Represents a collection of all the columns that are part of the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableColumnCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
             * For EACH ITEM in the collection: Returns a unique key that identifies the column within the table.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the index number of the column within the columns collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the name of the table column.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
             */
            name?: boolean;
            /**
             * For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
            
            
        }
        /**
         * Represents a column in a table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableColumnLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
             * Returns a unique key that identifies the column within the table.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * Returns the index number of the column within the columns collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * Specifies the name of the table column.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
             */
            name?: boolean;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
            
            
        }
        /**
         * Represents a collection of all the rows that are part of the table.
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableRowCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
            
            
        }
        /**
         * Represents a row in a table.
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableRowLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
            
            
        }
        
        
        /**
         * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Collection of border objects that apply to the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            borders?: Excel.Interfaces.RangeBorderCollectionLoadOptions;
            /**
            * Returns the fill object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            fill?: Excel.Interfaces.RangeFillLoadOptions;
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontLoadOptions;
            
            
            
            /**
             * Represents the horizontal alignment for the specified object. See `Excel.HorizontalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            horizontalAlignment?: boolean;
            
            
            
            
            
            
            
            /**
             * Represents the vertical alignment for the specified object. See `Excel.VerticalAlignment` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            verticalAlignment?: boolean;
            /**
             * Specifies if Excel wraps the text in the object. A `null` value indicates that the entire range doesn't have a uniform wrap setting
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            wrapText?: boolean;
        }
        
        /**
         * Represents the background of a range object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeFillLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            
            
            
            
        }
        /**
         * Represents the border of an object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeBorderLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: boolean;
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
            
            /**
             * Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            weight?: boolean;
        }
        /**
         * Represents the border objects that make up the range border.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeBorderCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * For EACH ITEM in the collection: HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * For EACH ITEM in the collection: Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: boolean;
            /**
             * For EACH ITEM in the collection: One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
            
            /**
             * For EACH ITEM in the collection: Specifies the weight of the border around a range. See `Excel.BorderWeight` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            weight?: boolean;
        }
        /**
         * This object represents the font attributes (font name, font size, color, etc.) for an object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeFontLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: boolean;
            
            
            
            
            /**
             * Type of underline applied to the font. See `Excel.RangeUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: boolean;
        }
        /**
         * A collection of all the chart objects on a worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsLoadOptions;
            /**
            * For EACH ITEM in the collection: Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendLoadOptions;
            
            
            /**
            * For EACH ITEM in the collection: Represents either a single series or collection of series in the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            series?: Excel.Interfaces.ChartSeriesCollectionLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartTitleLoadOptions;
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the height, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            height?: boolean;
            
            /**
             * For EACH ITEM in the collection: The distance, in points, from the left side of the chart to the worksheet origin.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            left?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the name of a chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            top?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the width, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            width?: boolean;
        }
        /**
         * Represents a chart object in a workbook.
                    To learn more about the chart object model, see {@link https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts | Work with charts using the Excel JavaScript API}.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesLoadOptions;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsLoadOptions;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatLoadOptions;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendLoadOptions;
            
            
            /**
            * Represents either a single series or collection of series in the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            series?: Excel.Interfaces.ChartSeriesCollectionLoadOptions;
            /**
            * Represents the title of the specified chart, including the text, visibility, position, and formatting of the title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartTitleLoadOptions;
            
            
            
            
            /**
             * Specifies the height, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            height?: boolean;
            
            /**
             * The distance, in points, from the left side of the chart to the worksheet origin.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            left?: boolean;
            /**
             * Specifies the name of a chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            
            
            
            /**
             * Specifies the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            top?: boolean;
            /**
             * Specifies the width, in points, of the chart object.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            width?: boolean;
        }
        
        /**
         * Encapsulates the format properties for the overall chart area.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAreaFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Represents the font attributes (font name, font size, color, etc.) for the current object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
            
            
        }
        /**
         * Represents a collection of chart series.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartSeriesCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
            * For EACH ITEM in the collection: Represents the formatting of a chart series, which includes fill and line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartSeriesFormatLoadOptions;
            
            /**
            * For EACH ITEM in the collection: Returns a collection of all points in the series.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            points?: Excel.Interfaces.ChartPointsCollectionLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            
            
            
            
            
            
            
            
        }
        /**
         * Represents a series in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartSeriesLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
            * Represents the formatting of a chart series, which includes fill and line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartSeriesFormatLoadOptions;
            
            /**
            * Returns a collection of all points in the series.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            points?: Excel.Interfaces.ChartPointsCollectionLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the name of a series in a chart. The name's length should not be greater than 255 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            
            
            
            
            
            
            
            
            
            
            
        }
        /**
         * Encapsulates the format properties for the chart series
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartSeriesFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatLoadOptions;
        }
        /**
         * A collection of all the chart points within a series inside a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartPointsCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * For EACH ITEM in the collection: Encapsulates the format properties chart point.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartPointFormatLoadOptions;
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Returns the value of a chart point.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: boolean;
        }
        /**
         * Represents a point of a series in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartPointLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Encapsulates the format properties chart point.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartPointFormatLoadOptions;
            
            
            
            
            
            /**
             * Returns the value of a chart point.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            value?: boolean;
        }
        /**
         * Represents the formatting object for chart points.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartPointFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
        }
        /**
         * Represents the chart axes.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxesLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisLoadOptions;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisLoadOptions;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisLoadOptions;
        }
        /**
         * Represents a single axis in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxisLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatLoadOptions;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesLoadOptions;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesLoadOptions;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Represents the interval between two major tick marks. Can be set to a numeric value or an empty string. The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            majorUnit?: boolean;
            /**
             * Represents the maximum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            maximum?: boolean;
            /**
             * Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minimum?: boolean;
            
            
            /**
             * Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            minorUnit?: boolean;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /**
         * Encapsulates the format properties for the chart axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxisFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
            /**
            * Specifies chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatLoadOptions;
        }
        /**
         * Represents the title of a chart axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxisTitleLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Specifies the formatting of the chart axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisTitleFormatLoadOptions;
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: boolean;
            
            /**
             * Specifies if the axis title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /**
         * Represents the chart axis title formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxisTitleFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
        }
        /**
         * Represents a collection of all the data labels on a chart point.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartDataLabelsLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatLoadOptions;
            
            
            
            
            /**
             * Value that represents the position of the data label. See `Excel.ChartDataLabelPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: boolean;
            /**
             * String representing the separator used for the data labels on a chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            separator?: boolean;
            /**
             * Specifies if the data label bubble size is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showBubbleSize?: boolean;
            /**
             * Specifies if the data label category name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showCategoryName?: boolean;
            /**
             * Specifies if the data label legend key is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showLegendKey?: boolean;
            /**
             * Specifies if the data label percentage is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showPercentage?: boolean;
            /**
             * Specifies if the data label series name is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showSeriesName?: boolean;
            /**
             * Specifies if the data label value is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            showValue?: boolean;
            
            
        }
        
        /**
         * Encapsulates the format properties for the chart data labels.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartDataLabelFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for a chart data label.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
        }
        
        
        
        
        /**
         * Represents major or minor gridlines on a chart axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartGridlinesLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of chart gridlines.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartGridlinesFormatLoadOptions;
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /**
         * Encapsulates the format properties for chart gridlines.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartGridlinesFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatLoadOptions;
        }
        /**
         * Represents the legend in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartLegendLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatLoadOptions;
            
            
            /**
             * Specifies if the chart legend should overlap with the main body of the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            /**
             * Specifies the position of the legend on the chart. See `Excel.ChartLegendPosition` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            position?: boolean;
            
            
            /**
             * Specifies if the chart legend is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
            
        }
        
        
        /**
         * Encapsulates the format properties of a chart legend.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartLegendFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Represents the font attributes such as font name, font size, and color of a chart legend.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
        }
        
        /**
         * Represents a chart title object of a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartTitleLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of a chart title, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartTitleFormatLoadOptions;
            
            
            
            /**
             * Specifies if the chart title will overlay the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            overlay?: boolean;
            
            
            /**
             * Specifies the chart's title text.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: boolean;
            
            
            
            /**
             * Specifies if the chart title is visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
            
        }
        
        /**
         * Provides access to the formatting options for a chart title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartTitleFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            /**
            * Represents the font attributes (such as font name, font size, and color) for an object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontLoadOptions;
        }
        
        
        
        /**
         * Encapsulates the formatting options for line elements.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartLineFormatLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * HTML color code representing the color of lines in the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            
            
        }
        /**
         * This object represents the font attributes (such as font name, font size, and color) for a chart object.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartFontLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Represents the bold status of font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: boolean;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: boolean;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    }
}


////////////////////////////////////////////////////////////////
//////////////////////// End Excel APIs ////////////////////////
////////////////////////////////////////////////////////////////