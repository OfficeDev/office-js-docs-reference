////////////////////////////////////////////////////////////////
//////////////////// Begin Office namespace ////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export var Preview: {
        startCustomFunctions(): Promise<void>;
    }

    export var Promise: PromiseConstructor;
    export var context: Context;
    /**
     * This method is called after the Office API was loaded.
     * @param reason - Indicates how the app was initialized
     */
    export function initialize(reason: InitializationReason): void;
    /**
    * Ensures that the Office JavaScript APIs are ready to be called by the add-in. If the framework hasn't initialized yet, the callback or promise will wait until the Office host is ready to accept API calls.
    * Note that though this API is intended to be used inside an Office add-in, it can also be used outside the add-in.  In that case, once Office.js determines that it is running outside of an Office host application, it will call the callback and resolve the promise with "null" for both the host and platform.
    * @param callback - An optional callback method, that will receive the host and platform info. Alternatively, rather than use a callback, an add-in may simply wait for the Promise returned by the function to resolve.
    * @returns A Promise that contains the host and platform info, once initialization is completed.
    */
    export function onReady(callback?: (info: { host: HostType, platform: PlatformType }) => any): Promise<{ host: HostType, platform: PlatformType }>;
    /**
     * Indicates if the large namespace for objects will be used or not.
     * @param useShortNamespace - Indicates if 'true' that the short namespace will be used
     */
    export function useShortNamespace(useShortNamespace: boolean): void;
    // Enumerations
    export enum AsyncResultStatus {
        /**
         * Operation succeeded
         */
        Succeeded,
        /**
         * Operation failed, check error object
         */
        Failed
    }
    export enum InitializationReason {
        /**
         * Indicates the app was just inserted in the document
         */
        Inserted,
        /**
         * Indicates if the extension already existed in the document
         */
        DocumentOpened
    }
    export enum HostType {
        /**
         * Host is Word
         */
        Word,
        /**
         * Host is Excel
         */
        Excel,
        /**
         * Host is PowerPoint
         */
        PowerPoint,
        /**
         * Host is Outlook
         */
        Outlook,
        /**
         * Host is OneNote
         */
        OneNote,
        /**
         * Host is Project
         */
        Project,
        /**
         * Host is Access
         */
        Access
    }
    export enum PlatformType {
        /**
         * Platform is PC
         */
        PC,
        /**
         * Platform is Web
         */
        OfficeOnline,
        /**
         * Platform is Mac
         */
        Mac,
        /**
         * Platform is iOS
         */
        iOS,
        /**
         * Platform is Android
         */
        Android,
        /**
         * Platform is Winrt
         */
        Universal
    }
    // Objects
        /**
		* An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.
		* 
		* **Hosts**: Access, Excel, Outlook, PowerPoint, Project, Word
        * **Last changed in**: 1.1
        * 
        * @remarks
        * When the function you pass to the callback parameter of an "Async" method executes, it receives an AsyncResult object that you can access from the callback function's only parameter.
		*/
        export interface AsyncResult {
        /**
        * Gets the user-defined item passed to the optional `asyncContext` parameter of the invoked method in the same state as it was passed in. This the user-defined item (which can be of any JavaScript type: String, Number, Boolean, Object, Array, Null, or Undefined) passed to the optional `asyncContext` parameter of the invoked method. Returns Undefined, if you didn't pass anything to the asyncContext parameter.
        * 
        * **Hosts**: Access, Excel, Outlook, PowerPoint, Project, Word
        * **Last changed in**: 1.1
        */
        asyncContext: any;
        /**
        * Gets the [AsyncResultStatus](asyncresultstatus.md) of the asynchronous operation.
        * 
        * **Hosts**: Access, Excel, Outlook, PowerPoint, Project, Word
        * **Last changed in**: 1.1
        */
        status: AsyncResultStatus;
        /**
        * Gets an [Error](error.md) object that provides a description of the error, if any error occurred.
        * 
        * **Hosts**: Access, Excel, Outlook, PowerPoint, Project, Word
        * **Last changed in**: 1.1
        */
        error: Error;
        /**
        * Gets the payload or content of this asynchronous operation, if any.
        * Note: What the value property returns for a particular "Async" method varies depending on the purpose and context of that method. To determine what is returned by the value property for an "Async" method, refer to the "Callback value" section of the method's topic. For a complete listing of the "Async" methods, see the Remarks section of the AsyncResult object topic.
        * 
        * **Hosts**: Access, Excel, Outlook, PowerPoint, Project, Word
        * **Last changed in**: 1.1
        * 
        * @remarks
        * You access the AsyncResult object in the function passed as the argument to the callback parameter of an "Async" method, such as the `getSelectedDataAsync` and `setSelectedDataAsync` methods of the Document object.
        */
        value: any;
    }
    export interface Context {
        auth: Auth;
        contentLanguage: string;
        displayLanguage: string;
        license: string;
        officeTheme: OfficeTheme;
        touchEnabled: boolean;
        ui: UI;
        host: HostType;
        platform: PlatformType;
        diagnostics: {
            host: HostType;
            platform: PlatformType;
            version: string;
        };
        requirements: {
            /**
             * Check if the specified requirement set is supported by the host Office application.
             * @param name - Set name. e.g.: "MatrixBindings".
             * @param minVersion - The minimum required version.
             */
            isSetSupported(name: string, minVersion?: number): boolean;
        }
    }
    /**
     * Provides specific information about an error that occurred during an asynchronous data operation.
     * 
     * Remarks: The Error object is accessed from the AsyncResult object that is returned in the function passed as the callback argument of an asynchronous data operation, such as the setSelectedDataAsync method of the Document object.
     */
    export interface Error {
        /**
         * Gets the numeric code of the error.
         */
        code: number;
        /**
         * Gets the name of the error.
         */
        message: string;
        /**
         * Gets a detailed description of the error.
         */
        name: string;
    }
    export interface UI {
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        */
        displayDialogAsync(startAddress: string): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param options - Optional. Accepts a DialogOptions object to define dialog behaviors.
        */
        displayDialogAsync(startAddress: string, options: DialogOptions): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt.
        */
        displayDialogAsync(startAddress: string, callback: (result: AsyncResult) => void): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param options - Optional. Accepts a DialogOptions object to define dialog behaviors.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt.
        */
        displayDialogAsync(startAddress: string, options: DialogOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Synchronously delivers a message from the dialog to its parent add-in.
         * @param messageObject - Accepts a message from the dialog to deliver to the add-in.
         */
        messageParent(messageObject: any): void;
        /**
         * Closes the UI container where the JavaScript is executing.
         *
         * Supported hosts: Outlook - Minimum requirement set: Mailbox 1.5
         *
         * The behavior of this method is specified by the following:
         *
         * Called from a UI-less command button: No effect. Any dialog opened by displayDialogAsync will remain open.
         *
         * Called from a taskpane: The taskpane will close. Any dialog opened by displayDialogAsync will also close. If the taskpane supports pinning and was pinned by the user, it will be un-pinned.
         *
         * Called from a module extension: No effect.
         */
        closeContainer(): void;
    }
    export interface DialogOptions {
        /**
         * Optional. Defines the width of the dialog as a percentage of the current display. Defaults to 99%. 250px minimum.
         */
        height?: number,
        /**
         * Optional. Defines the height of the dialog as a percentage of the current display. Defaults to 99%. 150px minimum.
         */
        width?: number,
        /**
         * Optional. Determines whether the dialog box should be displayed within an IFrame. This setting is only applicable in Office Online clients, and is ignored on other platforms.
         */
        displayInIframe?: boolean
    }
    export interface Auth {
        /**
        * Obtains an access token from AAD V 2.0 endpoint to grant the Office host application access to the add-in's web application.
        * @param callback - Optional. Accepts a callback method to handle the token acquisition attempt. If AsyncResult.status is "succeeded", then AsyncResult.value is the raw AAD v. 2.0-formatted access token.
        */
        getAccessTokenAsync(callback: (result: AsyncResult) => void): void;
        /**
        * Obtains an access token from AAD V 2.0 endpoint to grant the Office host application access to the add-in's web application.
        * @param options - Optional. Accepts an AuthOptions object to define sign-on behaviors.
        * @param callback - Optional. Accepts a callback method to handle the token acquisition attempt. If AsyncResult.status is "succeeded", then AsyncResult.value is the raw AAD v. 2.0-formatted access token.
        */
        getAccessTokenAsync(options: AuthOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface AuthOptions {
        /**
         * Optional. Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has been revoked.
         */
        forceConsent?: boolean,
        /**
         * Optional. Prompts the user to add (or to switch if already added) his or her Office account.
         */
        forceAddAccount?: boolean,
        /**
         * Optional. Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should be used with the authChallenge option.
         */
        authChallenge?: string
        /**
         * Optional. A user-defined item of any type that is returned in the AsyncResult object without being altered.
         */
        asyncContext?: any
    }
    export interface OfficeTheme {
        bodyBackgroundColor: string;
        bodyForegroundColor: string;
        controlBackgroundColor: string;
        controlForegroundColor: string;
    }
    /**
     * Dialog object returned as part of the displayDialogAsync callback. The object exposes methods for registering event handlers and closing the dialog
     */
    export interface DialogHandler {
        /**
         * When called from an active add-in dialog, asynchronously closes the dialog.
         */
        close(): void;
        /**
         * Adds an event handler for DialogMessageReceived or DialogEventReceived
         */
        addEventHandler(eventType: Office.EventType, handler: Function): void;

    }
}

export declare namespace Office {
    /**
     * Returns a promise of an object described in the expression. Callback is invoked only if method fails.
     * @param expression - The object to be retrieved. Example "bindings#BindingName", retrieves a binding promise for a binding named 'BindingName'
     * @param callback - The optional callback method
     */
    export function select(expression: string, callback?: (result: AsyncResult) => void): Binding;
    // Enumerations
    export enum ActiveView {
        Read,
        Edit
    }
    export enum BindingType {
        /**
         * Text based Binding
         */
        Text,
        /**
         * Matrix based Binding
         */
        Matrix,
        /**
         * Table based Binding
         */
        Table
    }
    export enum CoercionType {
        /**
         * Coerce as Text
         */
        Text,
        /**
         * Coerce as Matrix
         */
        Matrix,
        /**
         * Coerce as Table
         */
        Table,
        /**
         * Coerce as HTML
         */
        Html,
        /**
         * Coerce as Office Open XML
         */
        Ooxml,
        /**
         * Coerce as JSON object containing an array of the ids, titles, and indexes of the selected slides.
         */
        SlideRange,
        /**
        * Coerce as Image
        */
        Image
    }
    export enum DocumentMode {
        /**
         * Document in Read Only Mode
         */
        ReadOnly,
        /**
         * Document in Read/Write Mode
         */
        ReadWrite
    }
    export enum EventType {
        /**
         * Occurs when the user changes the current view of the document.
         */
        ActiveViewChanged,
        /**
         * Triggers when a binding level data change happens
         */
        BindingDataChanged,
        /**
         *  Triggers when a binding level selection happens
         */
        BindingSelectionChanged,
        /**
         * Triggers when Dialog sends a message via MessageParent.
         */
        DialogMessageReceived,
        /**
         * Triggers when Dialog has a event, such as dialog closed, dialog navigation failed.
         */
        DialogEventReceived,
        /**
         * Triggers when a document level selection happens
         */
        DocumentSelectionChanged,
        /**
         * Triggers when the active item changes
         */
        ItemChanged,
        /**
         * Triggers when a customXmlPart node was deleted
         */
        NodeDeleted,
        /**
         * Triggers when a customXmlPart node was inserted
         */
        NodeInserted,
        /**
         * Triggers when a customXmlPart node was replaced
         */
        NodeReplaced,
        /**
         * Triggers when settings change in a co-Auth session.
         */
        SettingsChanged,
        /**
         * Triggers when a Task selection happens in Project.
         */
        TaskSelectionChanged,
        /**
         *  Triggers when a Resource selection happens in Project.
         */
        ResourceSelectionChanged,
        /**
         * Triggers when a View selection happens in Project.
         */
        ViewSelectionChanged
    }
    export enum FileType {
        /**
         * Returns the file as plain text
         */
        Text,
        /**
         * Returns the file as a byte array
         */
        Compressed,
        /**
         * Returns the file in PDF format as a byte array
         */
        Pdf
    }
    export enum FilterType {
        /**
         * Returns all items
         */
        All,
        /**
         * Returns only visible items
         */
        OnlyVisible
    }
    export enum GoToType {
        /**
         * Goes to a binding object using the specified binding id.
         */
        Binding,
        /**
         * Goes to a named item using that item's name.
         * In Excel, you can use any structured reference for a named range or table: "Worksheet2!Table1"
         */
        NamedItem,
        /**
         * Goes to a slide using the specified id.
         */
        Slide,
        /**
         * Goes to the specified index by slide number or enum Office.Index
         */
        Index
    }
    export enum Index {
        First,
        Last,
        Next,
        Previous
    }
    export enum SelectionMode {
        Default,
        Selected,
        None
    }
    export enum ValueFormat {
        /**
         * Returns items without format
         */
        Unformatted,
        /**
         * Returns items with format
         */
        Formatted
    }
    // Objects
    export interface Binding {
        document: Document;
        /**
         * Id of the Binding
         */
        id: string;
        type: BindingType;
        /**
         * Adds an event handler to the object using the specified event type.
         * @param eventType - The event type. For binding it can be 'bindingDataChanged' and 'bindingSelectionChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current selection.
         * @param options - Syntax example: {coercionType: 'matrix,'valueFormat: 'formatted', filterType:'all'}
         *       coercionType: The expected shape of the selection. If not specified returns the bindingType shape. Use Office.CoercionType or text value.
         *       valueFormat: Get data with or without format. Use Office.ValueFormat or text value.
         *       startRow: Used in partial get for table/matrix. Indicates the start row.
         *       startColumn: Used in partial get for table/matrix. Indicates the start column.
         *       rowCount: Used in partial get for table/matrix. Indicates the number of rows from the start row.
         *       columnCount: Used in partial get for table/matrix. Indicates the number of columns from the start column.
         *       filterType: Get the visible or all the data. Useful when filtering data. Use Office.FilterType or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getDataAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler from the object using the specified event type.
         * @param eventType - The event type. For binding can be 'bindingDataChanged' and 'bindingSelectionChanged'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Writes the specified data into the current selection.
         * @param data - The data to be set. Either a string or value, 2d array or TableData object
         * @param options - Syntax example: {coercionType:Office.CoercionType.Matrix} or {coercionType: 'matrix'}
         *       coercionType: Explicitly sets the shape of the data object. Use Office.CoercionType or text value. If not supplied is inferred from the data type.
         *       startRow: Used in partial set for table/matrix. Indicates the start row.
         *       startColumn: Used in partial set for table/matrix. Indicates the start column.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setDataAsync(data: TableData | any, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Bindings {
        document: Document;
        /**
         * Creates a binding against a named object in the document
         * @param itemName - Name of the bindable object in the document. For Example 'MyExpenses' table in Excel."
         * @param bindingType - The Office BindingType for the data
         * @param options - Syntax example: {id: "BindingID"}
         *       id: Name of the binding, autogenerated if not supplied.
         *       asyncContext: Object keeping state for the callback
         *       columns: The string[] of the columns involved in the binding
         * @param callback - The optional callback method
         */
        addFromNamedItemAsync(itemName: string, bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Create a binding by prompting the user to make a selection on the document.
         * @param bindingType - The Office BindingType for the data
         * @param options - addFromPromptAsyncOptions- e.g. {promptText: "Please select data", id: "mySales"}
         *       promptText: Greet your users with a friendly word.
         *       asyncContext: Object keeping state for the callback
         *       id: Identifier.
         *       sampleData: A TableData that gives sample table in the Dialog.TableData.Headers is [][] of string.
         * @param callback - The optional callback method
         */
        addFromPromptAsync(bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Create a binding based on what the user's current selection.
         * @param bindingType - The Office BindingType for the data
         * @param options - addFromSelectionAsyncOptions- e.g. {id: "BindingID"}
         *       id: Identifier.
         *       asyncContext: Object keeping state for the callback
         *       columns: The string[] of the columns involved in the binding
         *       sampleData: A TableData that gives sample table in the Dialog.TableData.Headers is [][] of string.
         * @param callback - The optional callback method
         */
        addFromSelectionAsync(bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets an array with all the binding objects in the document.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getAllAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Retrieves a binding based on its Name
         * @param id - The binding id
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes the binding from the document
         * @param id - The binding id
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        releaseByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Context {
        document: Document;
    }
    export interface CustomXmlNode {
        baseName: string;
        namespaceUri: string;
        nodeType: string;
        /**
         * Gets the nodes associated with the xPath expression.
         * @param xPath - The xPath expression
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodesAsync(xPath: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the node value.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodeValueAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the text of an XML node in a custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTextAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the node's XML.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getXmlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the node value.
         * @param value - The value to be set on the node
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setNodeValueAsync(value: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously sets the text of an XML node in a custom XML part.
         * @param text - Required. The text value of the XML node.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setTextAsync(text: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the node XML.
         * @param xml - The XML to be set on the node
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setXmlAsync(xml: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlPart {
        builtIn: boolean;
        id: string;
        namespaceManager: CustomXmlPrefixMappings;
        /**
         * Adds an event handler to the object using the specified event type.
         * @param eventType - The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler?: (result: any) => void, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Deletes the Custom XML Part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        deleteAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the nodes associated with the xPath expression.
         * @param xPath - The xPath expression
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodesAsync(xPath: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the XML for the Custom XML Part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getXmlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler from the object using the specified event type.
         * @param eventType - The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlParts {
        /**
         * Asynchronously adds a new custom XML part to a file.
         * @param xml - The XML to add to the newly created custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        addAsync(xml: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the specified custom XML part by its id.
         * @param id - The id of the custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        getByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the specified custom XML part(s) by its namespace.
         * @param ns - The namespace to search.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        getByNamespaceAsync(ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlPrefixMappings {
        /**
         * Adds a namespace.
         * @param prefix - The namespace prefix
         * @param ns - The namespace URI
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addNamespaceAsync(prefix: string, ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets a namespace  with the specified prefix
         * @param prefix - The namespace prefix
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNamespaceAsync(prefix: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets a prefix  for  the specified URI
         * @param ns - The namespace URI
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getPrefixAsync(ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Document {
        bindings: Bindings;
        customXmlParts: CustomXmlParts;
        mode: DocumentMode;
        settings: Settings;
        url: string;
        /**
         * Adds an event handler for the specified event type.
         * @param eventType - The event type. For document can be 'DocumentSelectionChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current view of the presentation.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getActiveViewAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the entire file in slices of up to 4MB.
         * @param fileType - The format in which the file will be returned
         * @param options - Syntax example: {sliceSize:1024}
         *       sliceSize: Specifies the desired slice size (in bytes) up to 4MB. If not specified a default slice size of 4MB will be used.
         * @param callback - The optional callback method
         */
        getFileAsync(fileType: FileType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets file properties of the current document.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getFilePropertiesAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current selection.
         * @param coercionType - The expected shape of the selection.
         * @param options - Syntax example: {valueFormat: 'formatted', filterType:'all'}
         *       valueFormat: Get data with or without format. Use Office.ValueFormat or text value.
         *       filterType: Get the visible or all the data. Useful when filtering data. Use Office.FilterType or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedDataAsync(coercionType: CoercionType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Goes to the specified object or location in the document.
         * @param id - The identifier of the object or location to go to.
         * @param goToType - The type of the location to go to.
         * @param options - Syntax example: {asyncContext:context}
         *       selectionMode: Use Office.SelectionMode or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        goToByIdAsync(id: string | number, goToType: GoToType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler for the specified event type.
         * @param eventType - The event type. For document can be 'DocumentSelectionChanged'
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         *       handler: The name of the handler. If not specified all handlers are removed
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Writes the specified data into the current selection.
         * @param data - The data to be set. Either a string or value, 2d array or TableData object
         * @param options - Syntax example: {coercionType:Office.CoercionType.Matrix} or {coercionType: 'matrix'}
         *       coercionType: Explicitly sets the shape of the data object. Use Office.CoercionType or text value. If not supplied is inferred from the data type.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setSelectedDataAsync(data: string | TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
    }
    /**
     * Provides information about the document that raised the SelectionChanged event.
     */
    export interface DocumentSelectionChangedEventArgs {
        /**
         * Gets a Document object that represents the document that raised the SelectionChanged event.
         */
        document: Document;
        /**
         * Get an EventType enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }
    export interface File {
        size: number;
        sliceCount: number;
        /**
         * Closes the File.
         * @param callback - The optional callback method
         */
        closeAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the specified slice.
         * @param sliceIndex - The index of the slice to be retrieved
         * @param callback - The optional callback method
         */
        getSliceAsync(sliceIndex: number, callback?: (result: AsyncResult) => void): void;
    }
    export interface FileProperties {
        /**
         * File's URL
         */
        url: string
    }
    export interface MatrixBinding extends Binding {
        columnCount: number;
        rowCount: number;
    }
    export interface Settings {
        /**
         * Adds an event handler for the object using the specified event type.
         * @param eventType - The event type. For settings can be 'settingsChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Retrieves the setting with the specified name.
         * @param settingName - The name of the setting
         */
        get(name: string): any;
        /**
         * Gets the latest version of the settings object.
         * @param callback - The optional callback method
         */
        refreshAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Removes the setting with the specified name.
         * @param settingName - The name of the setting
         */
        remove(name: string): void;
        /**
         * Removes an event handler for the specified event type.
         * @param eventType - The event type. For settings can be 'settingsChanged'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Saves all settings.
         * @param options - Syntax example: {overwriteIfStale:false}
         *       overwriteIfStale: Indicates whether the setting will be replaced if stale.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        saveAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Sets a value for the setting with the specified name.
         * @param settingName - The name of the setting
         * @param value - The value for the setting
         */
        set(name: string, value: any): void;
    }
    export interface Slice {
        data: any;
        index: number;
        size: number;
    }
    export interface TableBinding extends Binding {
        columnCount: number;
        hasHeaders: boolean;
        rowCount: number;
        /**
         * Adds the specified columns to the table
         * @param tableData - A TableData object with the headers and rows
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addColumnsAsync(tableData: TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds the specified rows to the table
         * @param rows - A 2D array with the rows to add
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addRowsAsync(rows: TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Clears the table
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        deleteAllDataValuesAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Clears formatting on the bound table.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        clearFormatsAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the formatting on specified items in the table.
         * @param cellReference - An object literal containing name-value pairs that specify the range of cells to get formatting from.
         * @param formats - An array specifying the format properties to get.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getFormatsAsync(cellReference?: any, formats?: any[], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets formatting on specified items and data in the table.
         * @param formatsInfo - Array elements are themselves three-element arrays:[target, type, formats]
         *       target: The identifier of the item to format. String.
         *       type: The kind of item to format. String.
         *       formats: An object literal containing a list of property name-value pairs that define the formatting to apply.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setFormatsAsync(formatsInfo?: any[], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Updates table formatting options on the bound table.
         * @param tableOptions - An object literal containing a list of property name-value pairs that define the table options to apply.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setTableOptionsAsync(tableOptions: any, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export class TableData {
        constructor(rows: any[][], headers: any[]);
        constructor();
        headers: any[];
        rows: any[][];
    }
    export enum Table {
        All,
        Data,
        Headers
    }
    export interface TextBinding extends Binding { }
    export enum ProjectProjectFields {
        CurrencyDigits,
        CurrencySymbol,
        CurrencySymbolPosition,
        DurationUnits,
        GUID,
        Finish,
        Start,
        ReadOnly,
        VERSION,
        WorkUnits,
        ProjectServerUrl,
        WSSUrl,
        WSSList
    }
    export enum ProjectResourceFields {
        Accrual,
        ActualCost,
        ActualOvertimeCost,
        ActualOvertimeWork,
        ActualOvertimeWorkProtected,
        ActualWork,
        ActualWorkProtected,
        BaseCalendar,
        Baseline10BudgetCost,
        Baseline10BudgetWork,
        Baseline10Cost,
        Baseline10Work,
        Baseline1BudgetCost,
        Baseline1BudgetWork,
        Baseline1Cost,
        Baseline1Work,
        Baseline2BudgetCost,
        Baseline2BudgetWork,
        Baseline2Cost,
        Baseline2Work,
        Baseline3BudgetCost,
        Baseline3BudgetWork,
        Baseline3Cost,
        Baseline3Work,
        Baseline4BudgetCost,
        Baseline4BudgetWork,
        Baseline4Cost,
        Baseline4Work,
        Baseline5BudgetCost,
        Baseline5BudgetWork,
        Baseline5Cost,
        Baseline5Work,
        Baseline6BudgetCost,
        Baseline6BudgetWork,
        Baseline6Cost,
        Baseline6Work,
        Baseline7BudgetCost,
        Baseline7BudgetWork,
        Baseline7Cost,
        Baseline7Work,
        Baseline8BudgetCost,
        Baseline8BudgetWork,
        Baseline8Cost,
        Baseline8Work,
        Baseline9BudgetCost,
        Baseline9BudgetWork,
        Baseline9Cost,
        Baseline9Work,
        BaselineBudgetCost,
        BaselineBudgetWork,
        BaselineCost,
        BaselineWork,
        BudgetCost,
        BudgetWork,
        ResourceCalendarGUID,
        Code,
        Cost1,
        Cost10,
        Cost2,
        Cost3,
        Cost4,
        Cost5,
        Cost6,
        Cost7,
        Cost8,
        Cost9,
        ResourceCreationDate,
        Date1,
        Date10,
        Date2,
        Date3,
        Date4,
        Date5,
        Date6,
        Date7,
        Date8,
        Date9,
        Duration1,
        Duration10,
        Duration2,
        Duration3,
        Duration4,
        Duration5,
        Duration6,
        Duration7,
        Duration8,
        Duration9,
        Email,
        End,
        Finish1,
        Finish10,
        Finish2,
        Finish3,
        Finish4,
        Finish5,
        Finish6,
        Finish7,
        Finish8,
        Finish9,
        Flag10,
        Flag1,
        Flag11,
        Flag12,
        Flag13,
        Flag14,
        Flag15,
        Flag16,
        Flag17,
        Flag18,
        Flag19,
        Flag2,
        Flag20,
        Flag3,
        Flag4,
        Flag5,
        Flag6,
        Flag7,
        Flag8,
        Flag9,
        Group,
        Units,
        Name,
        Notes,
        Number1,
        Number10,
        Number11,
        Number12,
        Number13,
        Number14,
        Number15,
        Number16,
        Number17,
        Number18,
        Number19,
        Number2,
        Number20,
        Number3,
        Number4,
        Number5,
        Number6,
        Number7,
        Number8,
        Number9,
        OvertimeCost,
        OvertimeRate,
        OvertimeWork,
        PercentWorkComplete,
        CostPerUse,
        Generic,
        OverAllocated,
        RegularWork,
        RemainingCost,
        RemainingOvertimeCost,
        RemainingOvertimeWork,
        RemainingWork,
        ResourceGUID,
        Cost,
        Work,
        Start,
        Start1,
        Start10,
        Start2,
        Start3,
        Start4,
        Start5,
        Start6,
        Start7,
        Start8,
        Start9,
        StandardRate,
        Text1,
        Text10,
        Text11,
        Text12,
        Text13,
        Text14,
        Text15,
        Text16,
        Text17,
        Text18,
        Text19,
        Text2,
        Text20,
        Text21,
        Text22,
        Text23,
        Text24,
        Text25,
        Text26,
        Text27,
        Text28,
        Text29,
        Text3,
        Text30,
        Text4,
        Text5,
        Text6,
        Text7,
        Text8,
        Text9
    }
    export enum ProjectTaskFields {
        ActualCost,
        ActualDuration,
        ActualFinish,
        ActualOvertimeCost,
        ActualOvertimeWork,
        ActualStart,
        ActualWork,
        Text1,
        Text10,
        Finish10,
        Start10,
        Text11,
        Text12,
        Text13,
        Text14,
        Text15,
        Text16,
        Text17,
        Text18,
        Text19,
        Finish1,
        Start1,
        Text2,
        Text20,
        Text21,
        Text22,
        Text23,
        Text24,
        Text25,
        Text26,
        Text27,
        Text28,
        Text29,
        Finish2,
        Start2,
        Text3,
        Text30,
        Finish3,
        Start3,
        Text4,
        Finish4,
        Start4,
        Text5,
        Finish5,
        Start5,
        Text6,
        Finish6,
        Start6,
        Text7,
        Finish7,
        Start7,
        Text8,
        Finish8,
        Start8,
        Text9,
        Finish9,
        Start9,
        Baseline10BudgetCost,
        Baseline10BudgetWork,
        Baseline10Cost,
        Baseline10Duration,
        Baseline10Finish,
        Baseline10FixedCost,
        Baseline10FixedCostAccrual,
        Baseline10Start,
        Baseline10Work,
        Baseline1BudgetCost,
        Baseline1BudgetWork,
        Baseline1Cost,
        Baseline1Duration,
        Baseline1Finish,
        Baseline1FixedCost,
        Baseline1FixedCostAccrual,
        Baseline1Start,
        Baseline1Work,
        Baseline2BudgetCost,
        Baseline2BudgetWork,
        Baseline2Cost,
        Baseline2Duration,
        Baseline2Finish,
        Baseline2FixedCost,
        Baseline2FixedCostAccrual,
        Baseline2Start,
        Baseline2Work,
        Baseline3BudgetCost,
        Baseline3BudgetWork,
        Baseline3Cost,
        Baseline3Duration,
        Baseline3Finish,
        Baseline3FixedCost,
        Baseline3FixedCostAccrual,
        Basline3Start,
        Baseline3Work,
        Baseline4BudgetCost,
        Baseline4BudgetWork,
        Baseline4Cost,
        Baseline4Duration,
        Baseline4Finish,
        Baseline4FixedCost,
        Baseline4FixedCostAccrual,
        Baseline4Start,
        Baseline4Work,
        Baseline5BudgetCost,
        Baseline5BudgetWork,
        Baseline5Cost,
        Baseline5Duration,
        Baseline5Finish,
        Baseline5FixedCost,
        Baseline5FixedCostAccrual,
        Baseline5Start,
        Baseline5Work,
        Baseline6BudgetCost,
        Baseline6BudgetWork,
        Baseline6Cost,
        Baseline6Duration,
        Baseline6Finish,
        Baseline6FixedCost,
        Baseline6FixedCostAccrual,
        Baseline6Start,
        Baseline6Work,
        Baseline7BudgetCost,
        Baseline7BudgetWork,
        Baseline7Cost,
        Baseline7Duration,
        Baseline7Finish,
        Baseline7FixedCost,
        Baseline7FixedCostAccrual,
        Baseline7Start,
        Baseline7Work,
        Baseline8BudgetCost,
        Baseline8BudgetWork,
        Baseline8Cost,
        Baseline8Duration,
        Baseline8Finish,
        Baseline8FixedCost,
        Baseline8FixedCostAccrual,
        Baseline8Start,
        Baseline8Work,
        Baseline9BudgetCost,
        Baseline9BudgetWork,
        Baseline9Cost,
        Baseline9Duration,
        Baseline9Finish,
        Baseline9FixedCost,
        Baseline9FixedCostAccrual,
        Baseline9Start,
        Baseline9Work,
        BaselineBudgetCost,
        BaselineBudgetWork,
        BaselineCost,
        BaselineDuration,
        BaselineFinish,
        BaselineFixedCost,
        BaselineFixedCostAccrual,
        BaselineStart,
        BaselineWork,
        BudgetCost,
        BudgetFixedCost,
        BudgetFixedWork,
        BudgetWork,
        TaskCalendarGUID,
        ConstraintDate,
        ConstraintType,
        Cost1,
        Cost10,
        Cost2,
        Cost3,
        Cost4,
        Cost5,
        Cost6,
        Cost7,
        Cost8,
        Cost9,
        Date1,
        Date10,
        Date2,
        Date3,
        Date4,
        Date5,
        Date6,
        Date7,
        Date8,
        Date9,
        Deadline,
        Duration1,
        Duration10,
        Duration2,
        Duration3,
        Duration4,
        Duration5,
        Duration6,
        Duration7,
        Duration8,
        Duration9,
        Duration,
        EarnedValueMethod,
        FinishSlack,
        FixedCost,
        FixedCostAccrual,
        Flag10,
        Flag1,
        Flag11,
        Flag12,
        Flag13,
        Flag14,
        Flag15,
        Flag16,
        Flag17,
        Flag18,
        Flag19,
        Flag2,
        Flag20,
        Flag3,
        Flag4,
        Flag5,
        Flag6,
        Flag7,
        Flag8,
        Flag9,
        FreeSlack,
        HasRollupSubTasks,
        ID,
        Name,
        Notes,
        Number1,
        Number10,
        Number11,
        Number12,
        Number13,
        Number14,
        Number15,
        Number16,
        Number17,
        Number18,
        Number19,
        Number2,
        Number20,
        Number3,
        Number4,
        Number5,
        Number6,
        Number7,
        Number8,
        Number9,
        ScheduledDuration,
        ScheduledFinish,
        ScheduledStart,
        OutlineLevel,
        OvertimeCost,
        OvertimeWork,
        PercentComplete,
        PercentWorkComplete,
        Predecessors,
        PreleveledFinish,
        PreleveledStart,
        Priority,
        Active,
        Critical,
        Milestone,
        Overallocated,
        IsRollup,
        Summary,
        RegularWork,
        RemainingCost,
        RemainingDuration,
        RemainingOvertimeCost,
        RemainingWork,
        ResourceNames,
        Cost,
        Finish,
        Start,
        Work,
        StartSlack,
        Status,
        Successors,
        StatusManager,
        TotalSlack,
        TaskGUID,
        Type,
        WBS,
        WBSPREDECESSORS,
        WBSSUCCESSORS,
        WSSID
    }
    export enum ProjectViewTypes {
        Gantt,
        NetworkDiagram,
        TaskDiagram,
        TaskForm,
        TaskSheet,
        ResourceForm,
        ResourceSheet,
        ResourceGraph,
        TeamPlanner,
        TaskDetails,
        TaskNameForm,
        ResourceNames,
        Calendar,
        TaskUsage,
        ResourceUsage,
        Timeline
    }
    // Objects
    export interface Document {
        /**
         * Get Project field (Ex. ProjectWebAccessURL).
         * @param fieldId - Project level fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getProjectFieldAsync(fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get resource field for provided resource Id. (Ex.ResourceName)
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getResourceFieldAsync(resourceId: string, fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected Resource's Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedResourceAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected Task's Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedTaskAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected View Type (Ex. Gantt) and View Name.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedViewAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the Task Name, WSS Task Id, and ResourceNames for given taskId.
         * @param taskId - Either a string or value of the Task Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTaskAsync(taskId: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get task field for provided task Id. (Ex. StartDate).
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTaskFieldAsync(taskId: string, fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the WSS Url and list name for the Tasks List, the MPP is synced too.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getWSSUrlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
    }
}


////////////////////////////////////////////////////////////////
///////////////////// End Office namespace /////////////////////
////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////
//////////////// Begin OfficeExtension runtime /////////////////
////////////////////////////////////////////////////////////////

export declare namespace OfficeExtension {
    /** An abstract proxy object that represents an object in an Office document. You create proxy objects from the context (or from other proxy objects), add commands to a queue to act on the object, and then synchronize the proxy object state with the document by calling "context.sync()". */
    export class ClientObject {
        /** The request context associated with the object */
        context: ClientRequestContext;
        /** Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property. */
        isNullObject: boolean;
    }
}
export declare namespace OfficeExtension {
    export interface LoadOption {
        select?: string | string[];
        expand?: string | string[];
        top?: number;
        skip?: number;
    }
    export interface UpdateOptions {
        /**
         * Throw an error if the passed-in property list includes read-only properties (default = true).
         */
        throwOnReadOnly?: boolean
    }

    /** Contains debug information about the request context. */
    export interface RequestContextDebugInfo {
        /**
         * The statements to be executed in the host.
         *
         * These statements may not match the code exactly as written, but will be a close approximation.
         */
        pendingStatements: string[];
    }

    /** An abstract RequestContext object that facilitates requests to the host Office application. The "Excel.run" and "Word.run" methods provide a request context. */
    export class ClientRequestContext {
        constructor(url?: string);

        /** Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. */
        trackedObjects: TrackedObjects;

        /** Request headers */
        requestHeaders: { [name: string]: string };

        /** Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties. */
        load(object: ClientObject, option?: string | string[] | LoadOption): void;

        /**
        * Queues up a command to recursively load the specified properties of the object and its navigation properties.
        * You must call "context.sync()" before reading the properties.
        *
        * @param object - The object to be loaded.
        * @param options - The key-value pairing of load options for the types, such as { "Workbook": "worksheets,tables",  "Worksheet": "tables",  "Tables": "name" }
        * @param maxDepth - The maximum recursive depth.
        */
        loadRecursive(object: ClientObject, options: { [typeName: string]: string | string[] | LoadOption }, maxDepth?: number): void;

        /** Adds a trace message to the queue. If the promise returned by "context.sync()" is rejected due to an error, this adds a ".traceMessages" array to the OfficeExtension.Error object, containing all trace messages that were executed. These messages can help you monitor the program execution sequence and detect the cause of the error. */
        trace(message: string): void;

        /** Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context and retrieving properties of loaded Office objects for use in your code.�This method returns a promise, which is resolved when the synchronization is complete. */
        sync<T>(passThroughValue?: T): Promise<T>;

        /** Debug information */
        readonly debugInfo: RequestContextDebugInfo;
    }

    export interface EmbeddedOptions {
        sessionKey?: string,
        container?: HTMLElement,
        id?: string;
        timeoutInMilliseconds?: number;
        height?: string;
        width?: string;
    }

    export class EmbeddedSession {
        constructor(url: string, options?: EmbeddedOptions);
        public init(): Promise<any>;
    }
}
export declare namespace OfficeExtension {
    /** Contains the result for methods that return primitive types. The object's value property is retrieved from the document after "context.sync()" is invoked. */
    export class ClientResult<T> {
        /** The value of the result that is retrieved from the document after "context.sync()" is invoked. */
        value: T;
    }
}

export declare namespace OfficeExtension {
    /** Configuration */
    export var config: {
        /**
         * Determines whether to log additional error information upon failure.
         *
         * When this property is set to true, the error object will include a "debugInfo.fullStatements" property that lists all statements in the batch request, including all statements that precede and follow the point of failure.
         *
         * Setting this property to true will negatively impact performance and will log all statements in the batch request, including any statements that may contain potentially-sensitive data.
         * It is recommended that you only set this property to true during debugging and that you never log the value of error.debugInfo.fullStatements to an external database or analytics service.
         */
        extendedErrorLogging: boolean;
    };

    export interface DebugInfo {
        /** Error code string, such as "InvalidArgument". */
        code: string;
        /** The error message passed through from the host Office application. */
        message: string;
        /** Inner error, if applicable. */
        innerError?: DebugInfo | string;

        /** The object type and property or method name (or similar information), if available. */
        errorLocation?: string;

        /**
         * The statement that caused the error, if available.
         *
         * This statement will never contain any potentially-sensitive data and may not match the code exactly as written, but will be a close approximation.
         */
        statements?: string;

        /**
         * The statements that closely precede and follow the statement that caused the error, if available.
         *
         * These statements will never contain any potentially-sensitive data and may not match the code exactly as written, but will be a close approximation.
         */
        surroundingStatements?: string[];

        /**
         * All statements in the batch request (including any potentially-sensitive information that was specified in the request), if available.
         *
         * These statements may not match the code exactly as written, but will be a close approximation.
         */
        fullStatements?: string[];
    }

    /** The error object returned by "context.sync()", if a promise is rejected due to an error while processing the request. */
    export class Error {
        /** Error name: "OfficeExtension.Error".*/
        name: string;
        /** The error message passed through from the host Office application. */
        message: string;
        /** Stack trace, if applicable. */
        stack: string;
        /** Error code string, such as "InvalidArgument". */
        code: string;
        /** Trace messages (if any) that were added via a "context.trace()" invocation before calling "context.sync()". If there was an error, this contains all trace messages that were executed before the error occurred. These messages can help you monitor the program execution sequence and detect the case of the error. */
        traceMessages: Array<string>;
        /** Debug info (useful for detailed logging of the error, i.e., via JSON.stringify(...)). */
        debugInfo: DebugInfo;
        /** Inner error, if applicable. */
        innerError: Error;
    }
}
export declare namespace OfficeExtension {
    export class ErrorCodes {
        public static accessDenied: string;
        public static generalException: string;
        public static activityLimitReached: string;
        public static invalidObjectPath: string;
        public static propertyNotLoaded: string;
        public static valueNotLoaded: string;
        public static invalidRequestContext: string;
        public static invalidArgument: string;
        public static runMustReturnPromise: string;
        public static cannotRegisterEvent: string;
        public static apiNotFound: string;
        public static connectionFailure: string;
    }
}
export declare namespace OfficeExtension {
    /** An Promise object that represents a deferred interaction with the host Office application. The publically-consumable OfficeExtension.Promise is available starting in ExcelApi 1.2 and WordApi 1.2. Promises can be chained via ".then", and errors can be caught via ".catch". Remember to always use a ".catch" on the outer promise, and to return intermediary promises so as not to break the promise chain. When a "native" Promise implementation is available, OfficeExtension.Promise will switch to use the native Promise instead. */
    export const Promise: PromiseConstructor;

    export type IPromise<T> = Promise<T>;
}

export declare namespace OfficeExtension {
    /** Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information. */
    export class TrackedObjects {
        /** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
        add(object: ClientObject): void;
        /** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
        add(objects: ClientObject[]): void;
        /** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
        remove(object: ClientObject): void;
        /** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
        remove(objects: ClientObject[]): void;
    }
}

export declare namespace OfficeExtension {
    export class EventHandlers<T> {
        constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
        add(handler: (args: T) => Promise<any>): EventHandlerResult<T>;
        remove(handler: (args: T) => Promise<any>): void;
    }

    export class EventHandlerResult<T> {
        constructor(context: ClientRequestContext, handlers: EventHandlers<T>, handler: (args: T) => Promise<any>);
        /** The request context associated with the object */
        context: ClientRequestContext;
        remove(): void;
    }

    export interface EventInfo<T> {
        registerFunc: (callback: (args: any) => void) => Promise<any>;
        unregisterFunc: (callback: (args: any) => void) => Promise<any>;
        eventArgsTransformFunc: (args: any) => Promise<T>;
    }
}
export declare namespace OfficeExtension {
    /**
    * Request URL and headers
    */
    export interface RequestUrlAndHeaderInfo {
        /** Request URL */
        url: string;
        /** Request headers */
        headers?: {
            [name: string]: string;
        };
    }
}


////////////////////////////////////////////////////////////////
///////////////// End OfficeExtension runtime //////////////////
////////////////////////////////////////////////////////////////