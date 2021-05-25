import { Office as Outlook} from "../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin Office namespace ////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    /** A Promise object. Promises can be chained via ".then", and errors can be caught via ".catch".
     * When a browser-provided native Promise implementation is available, Office.Promise will switch to use the native Promise instead.
     */
    var Promise: IPromiseConstructor;

    // **Note**: this is a copy of the PromiseConstructor object from
    //     https://github.com/Microsoft/TypeScript/blob/master/lib/lib.es2015.promise.d.ts
    // It is necessary so that even with targeting "ES5" and not specifying any libs,
    //     developers will still get IntelliSense for "Office.Promise" just as they would with a regular Promise.
    // (because even though Promise is part of standard lib.d.ts, PromiseConstructor is not)
    export interface IPromiseConstructor {
        /**
         * A reference to the prototype.
         */
        readonly prototype: Promise<any>;

        /**
         * Creates a new Promise.
         * @param executor - A callback used to initialize the promise. This callback is passed two arguments:
         * a resolve callback used resolve the promise with a value or the result of another promise,
         * and a reject callback used to reject the promise with a provided reason or error.
         */
        new <T>(executor: (resolve: (value?: T | PromiseLike<T>) => void, reject: (reason?: any) => void) => void): Promise<T>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>, T10 | PromiseLike<T10>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8, T9, T10]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8, T9>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8, T9]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7, T8>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>]): Promise<[T1, T2, T3, T4, T5, T6, T7, T8]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6, T7>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>]): Promise<[T1, T2, T3, T4, T5, T6, T7]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5, T6>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>]): Promise<[T1, T2, T3, T4, T5, T6]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4, T5>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>]): Promise<[T1, T2, T3, T4, T5]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3, T4>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>]): Promise<[T1, T2, T3, T4]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2, T3>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>]): Promise<[T1, T2, T3]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T1, T2>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>]): Promise<[T1, T2]>;

        /**
         * Creates a Promise that is resolved with an array of results when all of the provided Promises
         * resolve, or rejected when any Promise is rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        all<T>(values: (T | PromiseLike<T>)[]): Promise<T[]>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>, T10 | PromiseLike<T10>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 | T9 | T10>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8, T9>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>, T9 | PromiseLike<T9>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 | T9>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7, T8>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>, T8 | PromiseLike<T8>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6, T7>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>, T7 | PromiseLike<T7>]): Promise<T1 | T2 | T3 | T4 | T5 | T6 | T7>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5, T6>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>, T6 | PromiseLike<T6>]): Promise<T1 | T2 | T3 | T4 | T5 | T6>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4, T5>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>, T5 | PromiseLike<T5>]): Promise<T1 | T2 | T3 | T4 | T5>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3, T4>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>, T4 | PromiseLike<T4>]): Promise<T1 | T2 | T3 | T4>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2, T3>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>, T3 | PromiseLike<T3>]): Promise<T1 | T2 | T3>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T1, T2>(values: [T1 | PromiseLike<T1>, T2 | PromiseLike<T2>]): Promise<T1 | T2>;

        /**
         * Creates a Promise that is resolved or rejected when any of the provided Promises are resolved
         * or rejected.
         * @param values - An array of Promises.
         * @returns A new Promise.
         */
        race<T>(values: (T | PromiseLike<T>)[]): Promise<T>;

        /**
         * Creates a new rejected promise for the provided reason.
         * @param reason - The reason the promise was rejected.
         * @returns A new rejected Promise.
         */
        reject(reason: any): Promise<never>;

        /**
         * Creates a new rejected promise for the provided reason.
         * @param reason - The reason the promise was rejected.
         * @returns A new rejected Promise.
         */
        reject<T>(reason: any): Promise<T>;

        /**
         * Creates a new resolved promise for the provided value.
         * @param value - A promise.
         * @returns A promise whose internal state matches the provided promise.
         */
        resolve<T>(value: T | PromiseLike<T>): Promise<T>;

        /**
         * Creates a new resolved promise.
         * @returns A resolved promise.
         */
        resolve(): Promise<void>;
    }

    /**
     * Gets the Context object that represents the runtime environment of the add-in and provides access to the top-level objects of the API.
     *
     * @remarks
     *
     * **Support details**
     *
     * A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application.
     * An empty cell indicates that the Office host application doesn't support this enumeration.
     *
     * For more information about Office host application and server requirements, see
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     *
     * *Supported hosts, by platform*
     *  <table>
     *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
     *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
     *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *  </table>
     */
    var context: Context;
    /**
     * Occurs when the runtime environment is loaded and the add-in is ready to start interacting with the application and hosted document.
     *
     * The reason parameter of the initialize event listener function returns an `InitializationReason` enumeration value that specifies how
     * initialization occurred. A task pane or content add-in can be initialized in two ways:
     *
     *  - The user just inserted it from Recently Used Add-ins section of the Add-in drop-down list on the Insert tab of the ribbon in the Office
     * host application, or from Insert add-in dialog box.
     *
     *  - The user opened a document that already contains the add-in.
     *
     * *Note*: The reason parameter of the initialize event listener function only returns an `InitializationReason` enumeration value for task pane
     * and content add-ins. It does not return a value for Outlook add-ins.
     *
     * @remarks
     *
     * **Support details**
     *
     * A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application.
     * An empty cell indicates that the Office host application doesn't support this method.
     *
     * For more information about Office host application and server requirements, see
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     *
     * *Supported hosts, by platform*
     *  <table>
     *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
     *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
     *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td> Y             </td></tr>
     *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
     *  </table>
     *
     * @param reason - Indicates how the app was initialized.
     */
    export function initialize(reason: InitializationReason): void;
    /**
    * Ensures that the Office JavaScript APIs are ready to be called by the add-in. If the framework hasn't initialized yet, the callback or promise
    * will wait until the Office host is ready to accept API calls. Note that though this API is intended to be used inside an Office add-in, it can
    * also be used outside the add-in. In that case, once Office.js determines that it is running outside of an Office host application, it will call
    * the callback and resolve the promise with "null" for both the host and platform.
    *
    * @param callback - An optional callback method, that will receive the host and platform info.
    *                   Alternatively, rather than use a callback, an add-in may simply wait for the Promise returned by the function to resolve.
    * @returns A Promise that contains the host and platform info, once initialization is completed.
    */
    export function onReady(callback?: (info: { host: HostType, platform: PlatformType }) => any): Promise<{ host: HostType, platform: PlatformType }>;
    /**
     * Toggles on and off the `Office` alias for the full `Microsoft.Office.WebExtension` namespace.
     *
     * @remarks
     *
     * **Support details**
     *
     * A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application.
     * An empty cell indicates that the Office host application doesn't support this method.
     *
     * For more information about Office host application and server requirements, see
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     *
     * *Supported hosts, by platform*
     *  <table>
     *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
     *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
     *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
     *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
     *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td>               </td></tr>
     *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
     *  </table>
     *
     * @param useShortNamespace - True to use the shortcut alias; otherwise false to disable it. The default is true.
     */
    export function useShortNamespace(useShortNamespace: boolean): void;
    /**
     * Provides a method for associating action names with functions that carry out an action.
     */
    const actions: Actions;
    /**
     * Represents the add-in.
     */
    const addin: Addin;
    /**
     * Represents the ribbon associated with the Office application.
     */
    const ribbon: Ribbon;
    /**
     * Checks if the specified requirement set is supported by the host Office application.
     * @param name - Set name; e.g., "MatrixBindings".
     * @param minVersion - The minimum required version; e.g., "1.4".
     */
    export function isSetSupported(name: string, minVersion?: string): boolean;
    // Enumerations
    /**
     * Provides options to determine the startup behavior of the add-in upon next start-up.
     */
    enum StartupBehavior {
        /**
         * The add-in does not load until opened by the user.
         */
        none = 'None',
        /**
         * Load the add-in but do not show UI.
         */
        load = 'Load',
    }
    /**
     * Visibility mode of the add-in.
     */
    enum VisibilityMode {
        /**
         * UI is Hidden
         */
        hidden = 'Hidden',
        /**
         * Displayed as taskpane
         */
        taskpane = 'Taskpane',
    }
    /**
     * Specifies the result of an asynchronous call.
     *
     * @remarks
     *
     * Returned by the `status` property of the {@link Office.AsyncResult | AsyncResult} object.
     */
    enum AsyncResultStatus {
        /**
         * The call succeeded.
         */
        Succeeded,
        /**
         * The call failed, check the error object.
         */
        Failed
    }
    /**
     * Specifies whether the add-in was just inserted or was already contained in the document.
     */
    enum InitializationReason {
        /**
         * The add-in was just inserted into the document.
         */
        Inserted,
        /**
         * The add-in is already part of the document that was opened.
         */
        DocumentOpened
    }
    /**
     * Specifies the host Office application in which the add-in is running.
     */
    enum HostType {
        /**
         * The Office host is Microsoft Word.
         */
        Word,
        /**
         * The Office host is Microsoft Excel.
         */
        Excel,
        /**
         * The Office host is Microsoft PowerPoint.
         */
        PowerPoint,
        /**
         * The Office host is Microsoft Outlook.
         */
        Outlook,
        /**
         * The Office host is Microsoft OneNote.
         */
        OneNote,
        /**
         * The Office host is Microsoft Project.
         */
        Project,
        /**
         * The Office host is Microsoft Access.
         *
         * **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint.
         * As an alternative, we recommend that you use {@link https://powerapps.microsoft.com/ | Microsoft PowerApps}
         * to build no-code business solutions for web and mobile devices.
         */
        Access
    }
    /**
     * Specifies the OS or other platform on which the Office host application is running.
     */
    enum PlatformType {
        /**
         * The platform is PC (Windows).
         */
        PC,
        /**
         * The platform is Office on the web (in a browser).
         */
        OfficeOnline,
        /**
         * The platform is Mac.
         */
        Mac,
        /**
         * The platform an iOS device.
         */
        iOS,
        /**
         * The platform is an Android device.
         */
        Android,
        /**
         * The platform is WinRT.
         */
        Universal
    }
    // Objects
        /**
        * An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.
        *
        * When the function you pass to the `callback` parameter of an "Async" method executes, it receives an AsyncResult object that you can access
        * from the `callback` function's only parameter.
        */
        export interface AsyncResult<T> {
        /**
        * Gets the user-defined item passed to the optional `asyncContext` parameter of the invoked method in the same state as it was passed in.
        * This returns the user-defined item (which can be of any JavaScript type: String, Number, Boolean, Object, Array, Null, or Undefined) passed
        * to the optional `asyncContext` parameter of the invoked method. Returns Undefined, if you didn't pass anything to the asyncContext parameter.
        */
        asyncContext: any;
        /**
        * Gets an object that may provide additional information if an {@link Office.Error | error} occurred.
        *
        * @remarks
        *
        * This property returns additional information if the following errors occur with these supported APIs.
        *
        * *Supported APIs*
        *
        * `Office.context.mailbox.item.getCallbackTokenAsync`, `Office.context.mailbox.item.getUserIdentityTokenAsync`
        *
        * *Supported errors*
        * <table>
        *   <tr><th>`AsyncResult.error.name`            </th><th>`AsyncResult.error.message`                                                                       </th><th>Description of diagnostics object returned                                                                                          </th></tr>
        *   <tr><td><strong>HTTPRequestFailure</strong> </td><td>The request has failed. Please look at the diagnostics object for the HTTP error code.            </td><td>The HTTP error code in a JSON object e.g., `{"HTTPCode":"401"}`                                                                     </td></tr>
        *   <tr><td><strong>InternalServerError</strong></td><td>The Exchange server returned an error. Please look at the diagnostics object for more information.</td><td>The error message from the Exchange server in a JSON object e.g., `{"ErrorText": "The mailbox database is temporarily unavailable"}`</td></tr>
        * </table>
        */
        diagnostics: any;
        /**
        * Gets an {@link Office.Error} object that provides a description of the error, if any error occurred.
        */
        error: Office.Error;
        /**
        * Gets the {@link Office.AsyncResultStatus} of the asynchronous operation.
        */
        status: AsyncResultStatus;
        /**
        * Gets the payload or content of this asynchronous operation, if any.
        *
        * @remarks
        *
        * You access the AsyncResult object in the function passed as the argument to the callback parameter of an "Async" method, such as the
        * `getSelectedDataAsync` and `setSelectedDataAsync` methods of the {@link Office.Document | Document} object.
        *
        * **Note**: What the value property returns for a particular "Async" method varies depending on the purpose and context of that method.
        * To determine what is returned by the value property for an "Async" method, refer to the "Callback value" section of the method's topic.
        */
        value: T;
    }
    /**
     * Used to associate an action name to a function.
     */
    export interface Actions {
        /**
         * Function to associate a name with the action function.
         */
        associate: (actionName: string, action: (arg?: any) => void) => void;
    }
    /**
     * Message used in the `onVisibilityModeChanged` invocation.
     */
    export interface VisibilityModeChangedMessage {
        /**
         * Visibility changed state.
         */
        visibilityMode: Office.VisibilityMode;
    }
    /**
     * Represents add-in level functionality for operating or configuring various aspects of the add-in.
     */
    export interface Addin {
        /**
         * Sets the startup behavior for the add-in for when the document is opened next time.
         * @param behavior - Specifies startup behavior of the add-in.
         */
        setStartupBehavior(behavior: Office.StartupBehavior): Promise<void>;
        /**
         * Gets the current startup behavior for the add-in.
         */
        getStartupBehavior(): Promise<Office.StartupBehavior>;
        /**
         * Shows the task pane associated with the add-in.
         * @returns A promise that is resolved when the UI is shown.
         */
        showAsTaskpane(): Promise<void>;
        /**
         * Hides the task pane.
         * @returns A promise that is resolved when the UI is hidden.
         */
        hide(): Promise<void>;
        /**
         * Adds a listener for the `onVisibilityModeChanged` event.
         * @param listener - The listener function that is called when the event is emitted. This function takes in a message for the receiving component.
         * @returns A promise that resolves to a function when the listener is added. Calling it removes the listener.
         */
        onVisibilityModeChanged(
            listener: (message: VisibilityModeChangedMessage) => void,
        ): Promise<() => Promise<void>>;
    }
    /**
     * An interface that contains all the functionality provided to manage the state of the Office ribbon.
     *
     * @remarks
     *
     * **Requirement set**: Ribbon 1.1
     */
    export interface Ribbon {
        /**
         * Registers a custom contextual tab with Office and defines the tab's controls.
         *
         * @remarks
         *
         * **Requirement set**: Ribbon 1.2
         * 
         * This method only requests that the tab be registered. The actual registration is controlled by the Office application and may not be complete when the returned `Promise` object is resolved.
         * For more information and code examples, see {@link https://docs.microsoft.com/office/dev/add-ins/design/contextual-tabs | Create custom contextual tabs}.
         * 
         * @param tabDefinition - Specifies the tab's properties and child controls and their properties. Pass a JSON string that conforms to the Office dynamic-ribbon JSON schema to `JSON.parse`, and then pass the returned object to this method.
         */
        requestCreateControls(tabDefinition: Object): Promise<void>;
        /**
         * Sends a request to Office to update the ribbon.
         *
         * @remarks
         *
         * **Requirement set**: Ribbon 1.1
         *
         * Note that this API is only to request an update. The actual UI update to the ribbon is controlled by the Office application and hence the exact timing of the ribbon update (or refresh) cannot be determined by the completion of this API.
         * 
         * For code examples, see  {@link https://docs.microsoft.com/office/dev/add-ins/design/disable-add-in-commands | Enable and Disable Add-in Commands} and {@link https://docs.microsoft.com/office/dev/add-ins/design/contextual-tabs | Create custom contextual tabs}.
         *
         * @param input - Represents the updates to be made to the ribbon. Note that only the changes specified in the input parameter are made.
         */
        requestUpdate(input: RibbonUpdaterData): Promise<void>;
    }
    /**
     * Specifies changes to the ribbon, such as the enabled or disabled status of a button.
     *
     * @remarks
     *
     * **Requirement set**: Ribbon 1.1
     */
    export interface RibbonUpdaterData {
        /**
         * Collection of tabs whose state is set with the call of `requestUpdate`.
         */
        tabs: Tab[];
    }
    /**
     * Represents an individual tab and the state it should have. For code examples, see  {@link https://docs.microsoft.com/office/dev/add-ins/design/disable-add-in-commands | Enable and Disable Add-in Commands} and {@link https://docs.microsoft.com/office/dev/add-ins/design/contextual-tabs | Create custom contextual tabs}.
     *
     * @remarks
     *
     * **Requirement set**: Ribbon 1.1
     */
    export interface Tab {
        /**
         * Identifier of the tab as specified in the manifest.
         */
        id: string;
        /**
         * Specifies one or more of the controls in the tab, such as menu items, buttons, etc.
         * 
         * @remarks
         * 
         * When the `Tab` object is part of a {@link Office.RibbonUpdaterData} object passed to the `requestUpdate` method of {@link Office.Ribbon}, this property specifies the IDs of the controls whose enabled status is to be changed. However, if there is a `groups` property on the tab, then this property is ignored and the `controls` properties of the specified groups must be used to change enabled status.
         */
        controls?: Control[];
        /**
         * Specifies whether the tab is visible on the ribbon. Used only with contextual tabs.
         * 
         * @remarks
         *
         * **Requirement set**: Ribbon 1.2
         */
        visible?: boolean;
        /**
         * Specifies one or more of the control groups on the tab.
         * 
         * @remarks
         * 
         * When the `Tab` object is part of an {@link Office.RibbonUpdaterData} object passed to the `requestUpdate` method of {@link Office.Ribbon}, the `controls` properties of the various {@link Office.Group} objects specify which controls have their enabled status changed; the `controls` property of the `Tab` object is ignored. 
         *
         * **Requirement set**: Ribbon 1.1
         */
        groups?: Group[];
    }
    /**
     * Represents a group of controls on a ribbon tab.
     *
     * **Requirement set**: Ribbon 1.1
     */
    export interface Group {
        /**
         * Identifier of the group as specified in the manifest.
         * 
         */
         id: string;
         /**
          * Specifies one or more of the controls in the group, such as menu items, buttons, etc.
          * 
          * @remarks
          * 
          * When the `Group` object is part of an {@link Office.RibbonUpdaterData} object passed to the `requestUpdate` method of {@link Office.Ribbon}, the `controls` properties of the various {@link Office.Group} objects specify which controls have their enabled status changed; the `controls` property of the `Group` object's parent `Tab` object is ignored. 
          */
         controls?: Control[];
    }
    /**
     * Represents an individual control or command and the state it should have.
     *
     * @remarks
     *
     * **Requirement set**: Ribbon 1.1
     */
    export interface Control {
        /**
         * Identifier of the control as specified in the manifest.
         */
        id: string;
        /**
         * Indicates whether the control should be visible or hidden. The default is true.
         */
        visible?: boolean;
        /**
         * Indicates whether the control should be enabled or disabled. The default is true.
         */
        enabled?: boolean;
    }
    /**
     * Represents the runtime environment of the add-in and provides access to key objects of the API.
     * The current context exists as a property of Office. It is accessed using `Office.context`.
     *
     * @remarks
     *
     * **Hosts**: Excel, Outlook, PowerPoint, Project, Word
     */
    export interface Context {
        /**
         * Provides information and access to the signed-in user.
         *
         * @beta
         */
        auth: Auth;
        /**
        * True, if the current platform allows the add-in to display a UI for selling or upgrading; otherwise returns False.
        *
        * @remarks
        *
        * **Hosts**: Excel, Word
        *
        * `commerceAllowed` is only supported in Office on iPad.
        *
        * The iOS App Store doesn't support apps with add-ins that provide links to additional payment systems. However, Office Add-ins running in
        * Office on the Windows desktop or in the browser do allow such links. If you want the UI of your add-in to provide a link to an
        * external payment system on platforms other than iOS, you can use the commerceAllowed property to control when that link is displayed.
        */
        commerceAllowed: boolean;
        /**
        * Gets the locale (language) specified by the user for editing the document or item.
        *
        * @remarks
        *
        * The `contentLanguage` value reflects the **Editing Language** setting specified with **File \> Options \> Language** in the Office host
        * application.
        *
        * **Support details**
        *
        * A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application.
        * An empty cell indicates that the Office host application doesn't support this enumeration.
        *
        * For more information about Office host application and server requirements, see
        * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
        *
        * *Supported hosts, by platform*
        *  <table>
        *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
        *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
        *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
        *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
        *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td>               </td></tr>
        *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td>               </td></tr>
        *  </table>
        */
        contentLanguage: string;
        /**
        * Gets information about the environment in which the add-in is running.
        */
        diagnostics: ContextInformation;
        /**
        * Gets the locale (language) specified by the user for the UI of the Office host application.
        *
        * @remarks
        *
        * The returned value is a string in the RFC 1766 Language tag format, such as en-US.
        *
        * The `displayLanguage` value reflects the current **Display Language** setting specified with **File \> Options \> Language** in the Office
        * host application.
        *
        * When using in Outlook, the applicable modes are Compose or Read.
        *
        * **Support details**
        *
        * A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application.
        * An empty cell indicates that the Office host application doesn't support this enumeration.
        *
        * For more information about Office host application and server requirements, see
        * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
        *
        * *Supported hosts, by platform*
        *  <table>
        *   <tr><th>                             </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th><th> Outlook on mobile devices </th><th> Office on Mac </th></tr>
        *   <tr><td><strong> Excel      </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
        *   <tr><td><strong> Outlook    </strong></td><td> Y                 </td><td> Y                     </td><td>                </td><td> Y                         </td><td> Y             </td></tr>
        *   <tr><td><strong> PowerPoint </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
        *   <tr><td><strong> Project    </strong></td><td> Y                 </td><td>                       </td><td>                </td><td>                           </td><td> Y             </td></tr>
        *   <tr><td><strong> Word       </strong></td><td> Y                 </td><td>                       </td><td> Y              </td><td>                           </td><td> Y             </td></tr>
        *  </table>
        */
        displayLanguage: string;
        /**
        * Gets an object that represents the document the content or task pane add-in is interacting with.
        */
        document: Office.Document;
        /**
        * Contains the Office application host in which the add-in is running.
        *
        * **Important**: In Outlook, this property is available from requirement set 1.5.
        * For all Mailbox requirement sets, you can use the `Office.context.diagnostics` property to get the host.
        */
        host: HostType;
        /**
        * Gets the license information for the user's Office installation.
        */
        license: string;
        /**
         * Provides access to the Microsoft Outlook add-in object model.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Key properties**:
         *
         * - `diagnostics`: Provides diagnostic information to an Outlook add-in.
         *
         * - `item`: Provides methods and properties for accessing a message or appointment in an Outlook add-in.
         *
         * - `userProfile`: Provides information about the user in an Outlook add-in.
         */
        mailbox: Outlook.Mailbox;
        /**
        * Provides access to the properties for Office theme colors.
        */
        officeTheme: OfficeTheme;
        /**
        * Provides the platform on which the add-in is running.
        *
        * **Important**: In Outlook, this property is available from requirement set 1.5.
        * For all Mailbox requirement sets, you can use the `Office.context.diagnostics` property to get the platform.
        */
        platform: PlatformType;
        /**
        * Provides a method for determining what requirement sets are supported on the current host and platform.
        */
        requirements: RequirementSetSupport;
        /**
         * Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.
         *
         * The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so it's available to
         * that add-in when it is running from any host client application used to access that mailbox.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        roamingSettings: Outlook.RoamingSettings;
        /**
        * Specifies whether the platform and device allows touch interaction.
        * True if the add-in is running on a touch device, such as an iPad; false otherwise.
        *
        * @remarks
        *
        * **Hosts**: Excel, PowerPoint, Word
        *
        * `touchEnabled` is only supported in Office on iPad.
        *
        * Use the touchEnabled property to determine when your add-in is running on a touch device and if necessary, adjust the kind of controls, and
        * size and spacing of elements in your add-in's UI to accommodate touch interactions.
        */
        touchEnabled: boolean;
        /**
        * Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes.
        */
        ui: UI;
    }
    /**
     * Provides specific information about an error that occurred during an asynchronous data operation.
     *
     * @remarks
     *
     * The Error object is accessed from the AsyncResult object that is returned in the function passed as the callback argument of an asynchronous
     * data operation, such as the `setSelectedDataAsync` method of the Document object.
     */
    export interface Error {
        /**
         * Gets the numeric code of the error. For a list of error codes, see {@link https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office-error-codes | JavaScript API for Office error codes}.
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
    export namespace AddinCommands {
        /**
         * The `Event` object is passed as a parameter to add-in functions invoked by UI-less command buttons. The object allows the add-in to identify
         * which button was clicked and to signal the host that it has completed its processing.
         *
         * @remarks
         *
         * See {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets | Add-in commands requirement sets} for more support information.
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level (Outlook)}**: Restricted
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        export interface Event {
            /**
             * Information about the control that triggered calling this function.
             *
             * @remarks
             *
             * This property is supported in Outlook only in {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets | requirement set} Mailbox 1.3 and later.
             */
            source:Source;

            /**
             * Indicates that the add-in has completed processing and will automatically be closed.
             *
             * This method must be called at the end of a function which was invoked by the following.
             *
             * - A UI-less button (i.e., an add-in command defined with an `Action` element where the `xsi:type` attribute is set to `ExecuteFunction`)
             *
             * - An {@link https://docs.microsoft.com/office/dev/add-ins/reference/manifest/event | event} defined in the
             * {@link https://docs.microsoft.com/office/dev/add-ins/reference/manifest/extensionpoint#events | Events extension point},
             * e.g., an `ItemSend` event
             *
             * [Api set: Mailbox 1.3]
             *
             * @remarks
             *
             * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
             *
             * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
             *
             * **Note**: The `options` parameter was introduced in Mailbox 1.8.
             *
             * @param options - Optional. An object that specifies behavior options for when the event is completed.
             */
            completed(options?: EventCompletedOptions): void;
        }

        /**
         * Specifies the behavior for when the event is completed.
         */
        export interface EventCompletedOptions {
            /**
             * A boolean value. When the completed method is used to signal completion of an event handler,
             * this value indicates if the handled event should continue execution or be canceled.
             * For example, an add-in that handles the `ItemSend` event can set `allowEvent` to `false` to cancel sending of the message.
             */
            allowEvent: boolean;
        }

        /**
         * Encapsulates source data for add-in events.
         */
        export interface Source {

            /**
             * The ID of the control that triggered calling this function. The ID comes from the manifest.
             */
            id: string;
        }
    }
    /**
     * Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.
     *
     * Visit "{@link https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins | Use the Dialog API in your Office Add-ins}"
     * for more information.
     */
    export interface UI {
        /**
         * Adds an event handler to the object using the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi 1.2}
         *
         * You can add multiple event handlers for the specified event type as long as the name of each event handler function is unique.
         *
         * @param eventType - Specifies the type of event to add. This must be `Office.EventType.DialogParentMessageReceived`.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.DialogParentMessageReceivedEventArgs}.
         * @param options - Optional. Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the handler registration returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: (result: DialogParentMessageReceivedEventArgs) => void, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        *
        * @remarks
        *
        * **Hosts**: Excel, Outlook, PowerPoint, Word
        *
        * **Requirement sets**:
        *
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi}
        *
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets | Mailbox 1.4}
        *
        * This method is available in the DialogApi requirement set for Excel, PowerPoint, or Word add-ins, and in the Mailbox requirement set 1.4
        * for Outlook. For more on how to specify a requirement set in your manifest, see
        * {@link https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements | Specify Office hosts and API requirements}.
        *
        * The initial page must be on the same domain as the parent page (the startAddress parameter). After the initial page loads, you can go to
        * other domains.
        *
        * Any page calling `Office.context.ui.messageParent` must also be on the same domain as the parent page.
        *
        * **Design considerations**:
        *
        * The following design considerations apply to dialog boxes:
        *
        * - An Office Add-in task pane can have only one dialog box open at any time. Multiple dialogs can be open at the same time from Add-in
        * Commands (custom ribbon buttons or menu items).
        *
        * - Every dialog box can be moved and resized by the user.
        *
        * - Every dialog box is centered on the screen when opened.
        *
        * - Dialog boxes appear on top of the host application and in the order in which they were created.
        *
        * Use a dialog box to:
        *
        * - Display authentication pages to collect user credentials.
        *
        * - Display an error/progress/input screen from a ShowTaskpane or ExecuteAction command.
        *
        * - Temporarily increase the surface area that a user has available to complete a task.
        *
        * Do not use a dialog box to interact with a document. Use a task pane instead.
        *
        * **displayDialogAsync Errors**:
        *
        * <table>
        *   <tr>
        *     <th>Code number</th>
        *     <th>Meaning</th>
        *   </tr>
        *   <tr>
        *     <td>12004</td>
        *     <td>The domain of the URL passed to displayDialogAsync is not trusted. The domain must be either the same domain as the host page (including protocol and port number), or it must be registered in the `AppDomains` section of the add-in manifest.</td>
        *   </tr>
        *   <tr>
        *     <td>12005</td>
        *     <td>The URL passed to displayDialogAsync uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</td>
        *   </tr>
        *   <tr>
        *     <td>12007</td>
        *     <td>A dialog box is already opened from the task pane. A task pane add-in can only have one dialog box open at a time.</td>
        *   </tr>
        *   <tr>
        *     <td>12009</td>
        *     <td>The user chose to ignore the dialog box. This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</td>
        *   </tr>
        * </table>
        *
        * In the callback function passed to the displayDialogAsync method, you can use the properties of the AsyncResult object to return the
        * following information.
        *
        * <table>
        *   <tr>
        *     <th>Property</th>
        *     <th>Use to</th>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.value</td>
        *     <td>Access the Dialog object.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.status</td>
        *     <td>Determine the success or failure of the operation.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.error</td>
        *     <td>Access an Error object that provides error information if the operation failed.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.asyncContext</td>
        *     <td>Access your user-defined object or value, if you passed one as the asyncContext parameter.</td>
        *   </tr>
        * </table>
        *
        * @param startAddress - Accepts the initial full HTTPS URL that opens in the dialog. Relative URLs must not be used.
        * @param options - Optional. Accepts an {@link Office.DialogOptions} object to define dialog display.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt. If successful, the AsyncResult.value is a Dialog object.
        */
        displayDialogAsync(startAddress: string, options?: DialogOptions, callback?: (result: AsyncResult<Dialog>) => void): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        *
        * @remarks
        *
        * **Hosts**: Excel, Outlook, PowerPoint, Word
        *
        * **Requirement sets**:
        *
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi}
        *
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets | Mailbox 1.4}
        *
        * This method is available in the DialogApi requirement set for Excel, PowerPoint, or Word add-ins, and in the Mailbox requirement set 1.4
        * for Outlook. For more on how to specify a requirement set in your manifest, see
        * {@link https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements | Specify Office hosts and API requirements}.
        *
        * The initial page must be on the same domain as the parent page (the startAddress parameter). After the initial page loads, you can go to
        * other domains.
        *
        * Any page calling `Office.context.ui.messageParent` must also be on the same domain as the parent page.
        *
        * **Design considerations**:
        *
        * The following design considerations apply to dialog boxes:
        *
        * - An Office Add-in task pane can have only one dialog box open at any time. Multiple dialogs can be open at the same time from Add-in
        * Commands (custom ribbon buttons or menu items).
        *
        * - Every dialog box can be moved and resized by the user.
        *
        * - Every dialog box is centered on the screen when opened.
        *
        * - Dialog boxes appear on top of the host application and in the order in which they were created.
        *
        * Use a dialog box to:
        *
        * - Display authentication pages to collect user credentials.
        *
        * - Display an error/progress/input screen from a ShowTaskpane or ExecuteAction command.
        *
        * - Temporarily increase the surface area that a user has available to complete a task.
        *
        * Do not use a dialog box to interact with a document. Use a task pane instead.
        *
        * **displayDialogAsync Errors**:
        *
        * <table>
        *   <tr>
        *     <th>Code number</th>
        *     <th>Meaning</th>
        *   </tr>
        *   <tr>
        *     <td>12004</td>
        *     <td>The domain of the URL passed to displayDialogAsync is not trusted. The domain must be either the same domain as the host page (including protocol and port number), or it must be registered in the `AppDomains` section of the add-in manifest.</td>
        *   </tr>
        *   <tr>
        *     <td>12005</td>
        *     <td>The URL passed to displayDialogAsync uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)</td>
        *   </tr>
        *   <tr>
        *     <td>12007</td>
        *     <td>A dialog box is already opened from the task pane. A task pane add-in can only have one dialog box open at a time.</td>
        *   </tr>
        *   <tr>
        *     <td>12009</td>
        *     <td>The user chose to ignore the dialog box. This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.</td>
        *   </tr>
        * </table>
        *
        * In the callback function passed to the displayDialogAsync method, you can use the properties of the AsyncResult object to return the
        * following information.
        *
        * <table>
        *   <tr>
        *     <th>Property</th>
        *     <th>Use to</th>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.value</td>
        *     <td>Access the Dialog object.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.status</td>
        *     <td>Determine the success or failure of the operation.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.error</td>
        *     <td>Access an Error object that provides error information if the operation failed.</td>
        *   </tr>
        *   <tr>
        *     <td>AsyncResult.asyncContext</td>
        *     <td>Access your user-defined object or value, if you passed one as the asyncContext parameter.</td>
        *   </tr>
        * </table>
        *
        * @param startAddress - Accepts the initial full HTTPS URL that opens in the dialog. Relative URLs must not be used.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt. If successful, the AsyncResult.value is a Dialog object.
        */
        displayDialogAsync(startAddress: string, callback?: (result: AsyncResult<Dialog>) => void): void;
        /**
         * Delivers a message from the dialog box to its parent/opener page. The page calling this API must be on the same domain as the parent.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi}
         *
         * @param message - Accepts a message from the dialog to deliver to the add-in. In addition to a boolean, anything that can serialized to a string including JSON and XML can be sent.
         */
        messageParent(message: boolean | string): void;
        /**
         * Closes the UI container where the JavaScript is executing.
         *
         * @remarks
         *
         * **Hosts**: Excel, Outlook (Minimum requirement set: Mailbox 1.5), PowerPoint, Word
         *
        * **Requirement sets**: 
        * 
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi}
        * 
        * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets | Mailbox 1.5}
         *
         * The behavior of this method is specified by the following:
         *
         * - Called from a UI-less command button: No effect. Any dialog opened by displayDialogAsync will remain open.
         *
         * - Called from a task pane: The task pane will close. Any dialog opened by displayDialogAsync will also close.
         * If the task pane supports pinning and was pinned by the user, it will be un-pinned.
         *
         * - Called from a module extension: No effect.
         */
        closeContainer(): void;
        /**
         * Opens a browser window and loads the specified URL. 
         * 
         * @remarks
         * 
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/open-browser-window-api-requirement-sets | OpenBrowserWindowApi 1.1}
         *  
         * @param url - The full URL to be opened including protocol (e.g., https), and port number, if any.
         */
        openBrowserWindow(url: string): void;
    }

    /**
     * Provides information about which Requirement Sets are supported in the current environment.
     */
    export interface RequirementSetSupport {
       /**
        * Check if the specified requirement set is supported by the host Office application.
        * @param name - The requirement set name (e.g., "ExcelApi").
        * @param minVersion - The minimum required version (e.g., "1.4").
        */
       isSetSupported(name: string, minVersion?: string): boolean;

       /**
        * Check if the specified requirement set is supported by the host Office application.
        *
        * **Warning**: This overload of `isSetSupported` (where `minVersionNumber` is a number) has been deprecated. Use the string overload of `isSetSupported` instead.
        *
        * @deprecated Use the string overload of `isSetSupported` instead.
        * @param name - The requirement set name (e.g., "ExcelApi").
        * @param minVersionNumber - The minimum required version (e.g., 1.4).
        */
       isSetSupported(name: string, minVersionNumber?: number): boolean;
    }

    /**
     * Provides options for how a dialog is displayed.
     */
    export interface DialogOptions {
        /**
         * Defines the height of the dialog as a percentage of the current display. Defaults to 80%. 250px minimum.
         */
        height?: number,
        /**
         * Defines the width of the dialog as a percentage of the current display. Defaults to 80%. 150px minimum.
         */
        width?: number,
        /**
         * Determines whether the dialog box should be displayed within an IFrame. This setting is only applicable in Office on the web, and is
         * ignored by other platforms. If false (default), the dialog will be displayed as a new browser window (pop-up). Recommended for
         * authentication pages that cannot be displayed in an IFrame. If true, the dialog will be displayed as a floating overlay with an IFrame.
         * This is best for user experience and performance.
         */
        displayInIframe?: boolean
        /**
         * Determines if the pop-up blocker dialog will be shown to the user. Defaults to true.
         *
         * `true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker.
         * `false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation).
         */
        promptBeforeOpen?: boolean;
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }

    /**
     * The Office Auth namespace, `Office.context.auth`, provides a method that allows the Office client application to obtain an access token to the add-in's web application.
     * Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.
     */
    export interface Auth {
        /**
         * Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. Enables add-ins to identify users.
         * Server-side code can use this token to access Microsoft Graph for the add-in's web application by using the
         * {@link https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of | "on behalf of" OAuth flow}.
         *
         * **Important**: In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.
         *
         * **Warning**: `getAccessTokenAsync` has been deprecated. Use `Office.auth.getAccessToken` instead.
         *
         * @deprecated Use `Office.auth.getAccessToken` instead.
         *
         * @remarks
         *
         * **Hosts**: Excel, OneNote, Outlook, PowerPoint, Word
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets | IdentityAPI}
         *
         * This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign in with Organizational
         * Accounts and Microsoft Accounts. Microsoft Azure returns tokens intended for both user account types to access resources in the Microsoft Graph.
         *
         * @param options - Optional. Accepts an `AuthOptions` object to define sign-on behaviors.
         * @param callback - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.
         *                   If `AsyncResult.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.
         */
        getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. Enables add-ins to identify users.
         * Server-side code can use this token to access Microsoft Graph for the add-in's web application by using the
         * {@link https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of | "on behalf of" OAuth flow}.
         *
         * **Important**: In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.
         *
         * **Warning**: `getAccessTokenAsync` has been deprecated. Use `Office.auth.getAccessToken` instead.
         *
         * @deprecated Use `Office.auth.getAccessToken` instead.
         *
         * @remarks
         *
         * **Hosts**: Excel, OneNote, Outlook, PowerPoint, Word
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets | IdentityAPI}
         *
         * This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign in with Organizational
         * Accounts and Microsoft Accounts. Microsoft Azure returns tokens intended for both user account types to access resources in the Microsoft Graph.
         *
         * @param callback - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.
         *                   If `AsyncResult.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.
         */
        getAccessTokenAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. Enables add-ins to identify users.
         * Server-side code can use this token to access Microsoft Graph for the add-in's web application by using the
         * {@link https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of | "on behalf of" OAuth flow}. 
         * This API requires a single sign-on configuration that bridges the add-in to an Azure application. Office users sign in with Organizational
         * Accounts and Microsoft Accounts. Microsoft Azure returns tokens intended for both user account types to access resources in the Microsoft Graph.
         *
         * @remarks
         *
         * **Hosts**: Excel, OneNote, Outlook, PowerPoint, Word
         * 
         * **Important**: In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox. 
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets | IdentityAPI}
         *
         * @param options - Optional. Accepts an `AuthOptions` object to define sign-on behaviors.
         * @returns Promise to the access token.
         */
        getAccessToken(options?: AuthOptions): Promise<string>;
    }
    /**
     * Provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessToken` method.
     */
    export interface AuthOptions {
        /**
         * Allows Office to get an access token silently or through interactive consent, if one is required. Default value is `false`.
         * If set to `false`, Office will silently try to get an access token. If it fails to do so, Office will return a descriptive error.
         * If set to `true`, Office will show an interactive consent UI after it fails to silently get an access token.
         * The prompt will only allow consent to the AAD profile scope, not to any Microsoft Graph scopes.
         */
        allowConsentPrompt?: boolean;
        /**
         * Allows Office to get an access token silently provided consent is present or show interactive UI to sign in the user. Default value is `false`.
         * If set to `false`, Office will silently try to get an access token. If it fails to do so, Office will return a descriptive error.
         * If set to `true`, Office will show an interactive sign-in UI after it fails to silently get an access token.
         */
        allowSignInPrompt?: boolean;
        /**
         * Prompts the user to add their Office account (or to switch to it, if it is already added). Default value is `false`.
         *
         * **Warning**: `forceAddAccount` has been deprecated. Use `allowSignInPrompt` instead.
         *
         * @deprecated Use `allowSignInPrompt` instead.
         */
        forceAddAccount?: boolean;
        /**
         * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has
         * been revoked. Default value is `false`.
         *
         * **Warning**: `forceConsent` has been deprecated. Use `allowConsentPrompt` instead.
         *
         * @deprecated Use `allowConsentPrompt` instead.
         */
        forceConsent?: boolean;
        /**
         * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor
         * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development
         * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try"
         * call of `getAccessToken` after Microsoft Graph has sent an error requesting the additional factor and containing the string that should
         * be used with the `authChallenge` option.
         */
        authChallenge?: string;
        /**
         * A user-defined item of any type that is returned, unchanged, in the `asyncContext` property of the `AsyncResult` object that is passed to a callback.
         */
        asyncContext?: any;
        /**
         * Causes Office to return a descriptive error when the add-in wants to access Microsoft Graph and the user/admin has not granted consent to Graph scopes. Default value is `false`.
         * Office only supports consent to Graph scopes when the add-in has been deployed by a tenant admin. This information will not be available during development.
         * Setting this option to `true` will cause Office to inform your add-in beforehand (by returning a descriptive error) if Graph access will fail.
         */
        forMSGraphAccess?: boolean;
    }
    /**
     * Provides an option for preserving context data of any type, unchanged, for use in a callback.
     */
    export interface AsyncContextOptions {
        /**
         * A user-defined item of any type that is returned, unchanged, in the `asyncContext` property of the `AsyncResult` object
         * that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides information about the environment in which the add-in is running.
     */
    export interface ContextInformation {
        /**
        * Gets the Office application host in which the add-in is running.
        */
        host: Office.HostType;
        /**
        * Gets the platform on which the add-in is running.
        */
        platform: Office.PlatformType;
        /**
        * Gets the version of Office on which the add-in is running.
        */
        version: string;
    }
    /**
     * Provides options for how to get the data in a binding.
     *
     * @remarks
     *
     * If the rows option is used, the value must be "thisRow".
     */
    export interface GetBindingDataOptions {
        /**
         * The expected shape of the selection. Use {@link Office.CoercionType} or text value. Default: The original, uncoerced type of the binding.
         */
        coercionType?: Office.CoercionType | string
        /**
         * Specifies whether values, such as numbers and dates, are returned with their formatting applied. Use Office.ValueFormat or text value.
         * Default: Unformatted data.
         */
        valueFormat?: Office.ValueFormat | string
        /**
        * For table or matrix bindings, specifies the zero-based starting row for a subset of the data in the binding. Default is first row.
        */
        startRow?: number
        /**
        * For table or matrix bindings, specifies the zero-based starting column for a subset of the data in the binding. Default is first column.
        */
        startColumn?: number
        /**
        * For table or matrix bindings, specifies the number of rows offset from the startRow. Default is all subsequent rows.
        */
        rowCount?: number
        /**
        * For table or matrix bindings, specifies the number of columns offset from the startColumn. Default is all subsequent columns.
        */
        columnCount?: number
        /**
        * Specify whether to get only the visible (filtered in) data or all the data (default is all). Useful when filtering data.
        * Use Office.FilterType or text value.
        */
        filterType?: Office.FilterType | string
        /**
        * Only for table bindings in content add-ins for Access. Specifies the pre-defined string "thisRow" to get data in the currently selected row.
        *
        * **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint.
        * As an alternative, we recommend that you use {@link https://powerapps.microsoft.com/ | Microsoft PowerApps}
        * to build no-code business solutions for web and mobile devices.
        */
        rows?: string
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for how to set the data in a binding.
     *
     * @remarks
     *
     * If the rows option is used, the value must be "thisRow".
     */
    export interface SetBindingDataOptions {
        /**
         * Use only with binding type table and when a TableData object is passed for the data parameter. An array of objects that specify a range of
         * columns, rows, or cells and specify, as key-value pairs, the cell formatting to apply to that range.
         *
         * Example: `[{cells: Office.Table.Data, format: {fontColor: "yellow"}}, {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]`
         */
        cellFormat?: RangeFormatConfiguration[]
        /**
         * Explicitly sets the shape of the data object. If not supplied is inferred from the data type.
         */
        coercionType?: Office.CoercionType | string
        /**
        * Only for table bindings in content add-ins for Access. Array of strings. Specifies the column names.
        *
        * **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint.
        * As an alternative, we recommend that you use {@link https://powerapps.microsoft.com/ | Microsoft PowerApps}
        * to build no-code business solutions for web and mobile devices.
        */
        columns?: string[]
        /**
        * Only for table bindings in content add-ins for Access. Specifies the pre-defined string "thisRow" to get data in the currently selected row.
        *
        * **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint.
        * As an alternative, we recommend that you use {@link https://powerapps.microsoft.com/ | Microsoft PowerApps}
        * to build no-code business solutions for web and mobile devices.
        */
        rows?: string
        /**
        * Specifies the zero-based starting row for a subset of the data in the binding. Only for table or matrix bindings. If omitted, data is set
        * starting in the first row.
        */
        startRow?: number
        /**
        * Specifies the zero-based starting column for a subset of the data. Only for table or matrix bindings. If omitted, data is set starting in
        * the first column.
        */
        startColumn?: number
        /**
        * For an inserted table, a list of key-value pairs that specify table formatting options, such as header row, total row, and banded rows.
        * Example: `{bandedRows: true,  filterButton: false}`
        */
        tableOptions?: object
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Specifies a range and its formatting.
     */
    export interface RangeFormatConfiguration {
        /**
         * Specifies the range. Example of using Office.Table enum: Office.Table.All. Example of using RangeCoordinates: `{row: 3, column: 4}` specifies
         * the cell in the 3rd (zero-based) row in the 4th (zero-based) column.
         */
         cells: Office.Table | RangeCoordinates
        /**
         * Specifies the formatting as key-value pairs. Example: `{borderColor: "white", fontStyle: "bold"}`
         */
         format: object
    }
    /**
     * Specifies a cell, or row, or column, by its zero-based row and/or column number. Example: `{row: 3, column: 4}` specifies the cell in the 3rd
     * (zero-based) row in the 4th (zero-based) column.
     */
    export interface RangeCoordinates {
        /**
         * The zero-based row of the range. If not specified, all cells, in the column specified by `column` are included.
         */
         row?: number
        /**
         * The zero-based column of the range. If not specified, all cells, in the row specified by `row` are included.
         */
         column?: number
    }
    /**
     * Provides options to determine which event handler or handlers are removed.
     */
    export interface RemoveHandlerOptions {
        /**
         * The handler to be removed. If a particular handler is not specified, then all handlers for the specified event type are removed.
         */
        handler?: (eventArgs?: Office.BindingDataChangedEventArgs | Office.BindingSelectionChangedEventArgs) => any;
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for configuring the binding that is created.
     */
    export interface AddBindingFromNamedItemOptions {
        /**
         * The unique ID of the binding. Autogenerated if not supplied.
         */
        id?: string
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for configuring the prompt and identifying the binding that is created.
     */
    export interface AddBindingFromPromptOptions {
        /**
         * The unique ID of the binding. Autogenerated if not supplied.
         */
        id?: string
        /**
         * Specifies the string to display in the prompt UI that tells the user what to select. Limited to 200 characters.
         * If no promptText argument is passed, "Please make a selection" is displayed.
         */
        promptText?: string
        /**
         * Specifies a table of sample data displayed in the prompt UI as an example of the kinds of fields (columns) that can be bound by your add-in.
         * The headers provided in the TableData object specify the labels used in the field selection UI.
         *
         * **Note**: This parameter is used only in add-ins for Access. It is ignored if provided when calling the method in an add-in for Excel.
         *
         * **Important**: We no longer recommend that you create and use Access web apps and databases in SharePoint.
         * As an alternative, we recommend that you use {@link https://powerapps.microsoft.com/ | Microsoft PowerApps}
         * to build no-code business solutions for web and mobile devices.
         */
        sampleData?: Office.TableData
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for identifying the binding that is created.
     */
    export interface AddBindingFromSelectionOptions {
        /**
         * The unique ID of the binding. Autogenerated if not supplied.
         */
        id?: string
        /**
         * The names of the columns involved in the binding.
         */
        columns?: string[]
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for setting the size of slices that the document will be divided into.
     */
    export interface GetFileOptions {
        /**
         * The the size of the slices in bytes. The maximum (and the default) is 4194304 (4MB).
         */
        sliceSize?: number
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for customizing what data is returned and how it is formatted.
     */
    export interface GetSelectedDataOptions {
        /**
         * Specify whether the data is formatted. Use Office.ValueFormat or string equivalent.
         */
        valueFormat?: Office.ValueFormat | string
        /**
         * Specify whether to get only the visible (that is, filtered-in) data or all the data. Useful when filtering data.
         * Use {@link Office.FilterType} or string equivalent. This parameter is ignored in Word documents.
         */
        filterType?: Office.FilterType | string
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for whether to select the location that is navigated to.
     *
     * @remarks
     *
     * The behavior caused by the {@link Office.SelectionMode | options.selectionMode} option varies by host:
     *
     * In Excel: `Office.SelectionMode.Selected` selects all content in the binding, or named item. `Office.SelectionMode.None` for text bindings,
     * selects the cell; for matrix bindings, table bindings, and named items, selects the first data cell (not first cell in header row for tables).
     *
     * In PowerPoint: `Office.SelectionMode.Selected` selects the slide title or first textbox on the slide.
     * `Office.SelectionMode.None` doesn't select anything.
     *
     * In Word: `Office.SelectionMode.Selected` selects all content in the binding. `Office.SelectionMode.None` for text bindings, moves the cursor to
     * the beginning of the text; for matrix bindings and table bindings, selects the first data cell (not first cell in header row for tables).
     */
    export interface GoToByIdOptions {
        /**
         * Specifies whether the location specified by the id parameter is selected (highlighted).
         * Use {@link Office.SelectionMode} or string equivalent. See the Remarks for more information.
         */
        selectionMode?: Office.SelectionMode | string
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for how to insert data to the selection.
     */
    export interface SetSelectedDataOptions {
        /**
         * Use only with binding type table and when a TableData object is passed for the data parameter. An array of objects that specify a range of
         * columns, rows, or cells and specify, as key-value pairs, the cell formatting to apply to that range.
         *
         * Example: `[{cells: Office.Table.Data, format: {fontColor: "yellow"}}, {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]`
         */
        cellFormat?: RangeFormatConfiguration[]
        /**
         * Explicitly sets the shape of the data object. If not supplied is inferred from the data type.
         */
        coercionType?: Office.CoercionType | string
        /**
         * For an inserted table, a list of key-value pairs that specify table formatting options, such as header row, total row, and banded rows.
         * Example: `{bandedRows: true,  filterButton: false}`
         */
        tableOptions?: object
        /**
        * This option is applicable for inserting images. Indicates the insert location in relation to the top of the slide for PowerPoint, and its
        * relation to the currently selected cell in Excel. This value is ignored for Word. This value is in points.
        */
        imageTop?: number
        /**
        * This option is applicable for inserting images. Indicates the image width. If this option is provided without the imageHeight, the image
        * will scale to match the value of the image width. If both image width and image height are provided, the image will be resized accordingly.
        * If neither the image height or width is provided, the default image size and aspect ratio will be used. This value is in points.
        */
        imageWidth?: number
        /**
        * This option is applicable for inserting images. Indicates the insert location in relation to the left side of the slide for PowerPoint, and
        * its relation to the currently selected cell in Excel. This value is ignored for Word. This value is in points.
        */
        imageLeft?: number
        /**
        * This option is applicable for inserting images. Indicates the image height. If this option is provided without the imageWidth, the image
        * will scale to match the value of the image height. If both image width and image height are provided, the image will be resized accordingly.
        * If neither the image height or width is provided, the default image size and aspect ratio will be used. This value is in points.
        */
        imageHeight?: number
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides options for saving settings.
     */
    export interface SaveSettingsOptions {
        /**
         * **Warning**: This setting has been deprecated and should not be used. It has no effect on most platforms and will cause errors if set to `false` in Excel on the web.
         * 
         * @deprecated `overwriteIfStale` is no longer supported.
         */
        overwriteIfStale?: boolean
        /**
         * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
         */
        asyncContext?: any
    }
    /**
     * Provides access to the properties for Office theme colors.
     *
     * Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with File \>
     * Office Account \> Office Theme UI, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and
     * task pane add-ins.
     *
     * @remarks
     *
     * **Hosts**: Excel, Outlook (in preview), PowerPoint, Word
     *
     * `OfficeTheme` is only supported in Office on Windows.
     */
    export interface OfficeTheme {
        /**
         * Gets the Office theme body background color as a hexadecimal color triplet (e.g., "FFA500").
         */
        bodyBackgroundColor: string;
        /**
         * Gets the Office theme body foreground color as a hexadecimal color triplet (e.g., "FFA500").
         */
        bodyForegroundColor: string;
        /**
         * Gets the Office theme control background color as a hexadecimal color triplet (e.g., "FFA500").
         */
        controlBackgroundColor: string;
        /**
         * Gets the Office theme control foreground color as a hexadecimal color triplet (e.g., "FFA500").
         */
        controlForegroundColor: string;
    }
    /**
     * The object that is returned when `UI.displayDialogAsync` is called. It exposes methods for registering event handlers and closing the dialog.
     *
     * @remarks
     *
     * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi}
     */
    export interface Dialog {
        /**
         * Called from a parent page to close the corresponding dialog box. 
         * 
         * This method is asynchronous. It does not take a callback parameter and it does not return a Promise object, so it cannot be awaited with either the `await` keyword or the `then` function. See this best practice for more information: {@link https://docs.microsoft.com/office/dev/add-ins/develop/dialog-best-practices#opening-another-dialog-immediately-after-closing-one | Opening another dialog immediately after closing one}
         */
        close(): void;
        /**
         * Registers an event handler. The two supported events are:
         *
         * - DialogMessageReceived. Triggered when the dialog box sends a message to its parent.
         *
         * - DialogEventReceived. Triggered when the dialog box has been closed or otherwise unloaded.
         * 
         * @param eventType - Must be either DialogMessageReceived or DialogEventReceived.
         * @param handler - A function which accepts either an object with a `message` property, if eventType is DialogMessageReceived, or an object with an `error` property, if eventType is DialogEventReceived.
         */
        addEventHandler(eventType: Office.EventType, handler: (args: {message: string | boolean} | {error: number}) => void): void;
        /**
         * Delivers a message from the host page, such as a task pane or a UI-less function file, to a dialog that was opened from the page.
         *
         * @remarks
         *
         * **Hosts**: Excel, Outlook (Minimum requirement set: Mailbox 1.9), PowerPoint, Word
         *
         * **Requirement sets**: 
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets | DialogApi 1.2}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets | Mailbox 1.9}
         *
         * @param message - Accepts a message from the host page to deliver to the dialog. Anything that can be serialized to a string, including JSON and XML, can be sent.
         */
        messageChild(message: string): void;
        /**
         * FOR INTERNAL USE ONLY. DO NOT CALL IN YOUR CODE.
         */
        sendMessage(name: string): void;
    }

    /**
     * Returns a promise of an object described in the expression. Callback is invoked only if method fails.
     *
     * @param expression - The object to be retrieved. Example "bindings#BindingName", retrieves a binding promise for a binding named 'BindingName'
     * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
     *
     * @remarks
     *
     * **Support details**
     *
     * A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application.
     * An empty cell indicates that the Office host application doesn't support this method.
     *
     * For more information about Office host application and server requirements, see
     * {@link https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins | Requirements for running Office Add-ins}.
     *
     * *Supported hosts, by platform*
     *  <table>
     *   <tr><th>                         </th><th> Office on Windows </th><th> Office in web browser </th><th> Office on iPad </th></tr>
     *   <tr><td><strong> Excel  </strong></td><td> Y                 </td><td> Y                     </td><td> Y              </td></tr>
     *   <tr><td><strong> Word   </strong></td><td> Y                 </td><td>                       </td><td> Y              </td></tr>
     *  </table>
     */
    export function select(expression: string, callback?: (result: AsyncResult<any>) => void): Binding;

    // Enumerations

    /**
     * Specifies the state of the active view of the document, for example, whether the user can edit the document.
     */
    enum ActiveView {
        /**
         * The active view of the host application only lets the user read the content in the document.
         */
        Read,
        /**
         * The active view of the host application lets the user edit the content in the document.
         */
        Edit
    }
    /**
     * Specifies the type of the binding object that should be returned.
     */
    enum BindingType {
        /**
         * Plain text. Data is returned as a run of characters.
         */
        Text,
        /**
         * Tabular data without a header row. Data is returned as an array of arrays, for example in this form:
         * [[row1column1, row1column2],[row2column1, row2column2]]
         */
        Matrix,
        /**
         * Tabular data with a header row. Data is returned as a {@link Office.TableData | TableData} object.
         */
        Table
    }
    /**
     * Specifies how to coerce data returned or set by the invoked method.
     *
     * @remarks
     * 
     * Application and platform support for each `CoercionType` is specified in the following requirement set descriptions.
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion}, (when using `Office.CoercionType.Html`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets | ImageCoercion 1.1} (when using `Office.CoercionType.Image`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection | Selection}
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
     * 
     * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12 | ImageCoercion 1.2} (when using `Office.CoercionType.XmlSvg`)
     * 
     */
    enum CoercionType {
        /**
         * Return or set data as text (string). Data is returned or set as a one-dimensional run of characters.
         */
        Text,
        /**
         * Return or set data as tabular data with no headers. Data is returned or set as an array of arrays containing one-dimensional runs of
         * characters. For example, three rows of  string values in two columns would be: [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]].
         *
         * **Note**: Only applies to data in Excel and Word.
         */
        Matrix,
        /**
         * Return or set data as tabular data with optional headers. Data is returned or set as an array of arrays with optional headers.
         *
         * **Note**: Only applies to data in Excel and Word.
         */
        Table,
        /**
         * Return or set data as HTML.
         *
         * **Note**: Only applies to data in add-ins for Word and Outlook add-ins for Outlook (compose mode).
         */
        Html,
        /**
         * Return or set data as Office Open XML.
         *
         * **Note**: Only applies to data in Word.
         */
        Ooxml,
        /**
         * Return a JSON object that contains an array of the ids, titles, and indexes of the selected slides. For example,
         * `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` for a selection of two slides.
         *
         * **Note**: Only applies to data in PowerPoint when calling the {@link Office.Document | Document}.getSelectedData method to get the current
         * slide or selected range of slides.
         */
        SlideRange,
        /**
        * Data is returned or set as an image stream.
        * **Note**: Only applies to data in Excel, Word, and PowerPoint.
        */
        Image,
        /**
         * Data is returned or set as XML data containing an SVG image.
         * **Note**: Only applies to data in Excel, Word, and PowerPoint.
         */
        XmlSvg
    }
    /**
     * Specifies whether the document in the associated application is read-only or read-write.
     *
     * @remarks
     *
     * Returned by the mode property of the {@link Office.Document | Document} object.
     */
    enum DocumentMode {
        /**
         * The document is read-only.
         */
        ReadOnly,
        /**
         * The document can be read and written to.
         */
        ReadWrite
    }
    /**
     * Specifies the type of the XML node.
     */
    enum CustomXMLNodeType {
        /**
         * The node is an attribute.
         */
        Attribute,
        /**
         * The node is CData.
         */
        CData,
        /**
         * The node is a comment.
         */
        NodeComment,
        /**
         * The node is an element.
         */
        Element,
        /**
         * The node is a Document element.
         */
        NodeDocument,
        /**
         * The node is a processing instruction.
         */
        ProcessingInstruction,
        /**
         * The node is text.
         */
        Text,
    }
    /**
     * Specifies the kind of event that was raised. Returned by the `type` property of an `*EventArgs` object.
     *
     * Add-ins for Project support the `Office.EventType.ResourceSelectionChanged`, `Office.EventType.TaskSelectionChanged`, and
     * `Office.EventType.ViewSelectionChanged` event types.
     *
     * Only task pane add-ins for Outlook support Mailbox API set event types.
     *
     * @remarks
     *
     * **`BindingDataChanged` and `BindingSelectionChanged` hosts**: Excel, Word.
     *
     */
    enum EventType {
        /**
         * A Document.ActiveViewChanged event was raised.
         *
         * @remarks
         *
         * **Hosts**: PowerPoint
         */
        ActiveViewChanged,
        /**
         * Triggers when any date or time of the selected appointment or series is changed in Outlook. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.7 | Office.AppointmentTimeChangedEventArgs}.
         *
         * [Api set: Mailbox 1.7]
         */
        AppointmentTimeChanged,
        /**
         * Triggers when an attachment is added to or removed from an item. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.8 | Office.AttachmentsChangedEventArgs}.
         *
         * [Api set: Mailbox 1.8]
         */
        AttachmentsChanged,
        /**
         * Occurs when data within the binding is changed.
         * To add an event handler for the BindingDataChanged event of a binding, use the addHandlerAsync method of the Binding object.
         * The event handler receives an argument of type {@link Office.BindingDataChangedEventArgs}.
         *
         * @remarks
         *
         * **Hosts**: Excel, Word
         */
        BindingDataChanged,
        /**
         * Occurs when the selection is changed within the binding. To add an event handler for the BindingSelectionChanged event of a binding, use
         * the addHandlerAsync method of the Binding object. The event handler receives an argument of type {@link Office.BindingSelectionChangedEventArgs}.
         *
         * @remarks
         *
         * **Hosts**: Excel, Word
         */
        BindingSelectionChanged,
        /**
         * Triggers when Dialog has an event, such as dialog closed or dialog navigation failed.
         */
        DialogEventReceived,
        /**
         * Triggers when a dialog sends a message via `messageParent`.
         */
        DialogMessageReceived,
        /**
         * Triggers when a host page sends a message to a child dialog box with `messageChild`.
         */
        DialogParentMessageReceived,
        /**
         * Triggers when a document-level selection happens.
         *
         * @remarks
         *
         * **Hosts**: Excel, Word
         */
        DocumentSelectionChanged,
        /**
         * Triggers when the appointment location is changed in Outlook. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.enhancedlocationschangedeventargs?view=outlook-js-1.8 | Office.EnhancedLocationsChangedEventArgs}.
         *
         * [Api set: Mailbox 1.8]
         */
        EnhancedLocationsChanged,
        /**
         * Triggers when a different Outlook item is selected for viewing while the task pane is pinned. Supported with task pane only.
         *
         * [Api set: Mailbox 1.5]
         */
        ItemChanged,
        /**
         * Triggers when a customXmlPart node is deleted.
         */
        NodeDeleted,
        /**
         * Triggers when a customXmlPart node is inserted.
         */
        NodeInserted,
        /**
         * Triggers when a customXmlPart node is replaced.
         */
        NodeReplaced,
        /**
         * Triggers when the OfficeTheme is changed in Outlook. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.officethemechangedeventargs?view=outlook-js-preview | Office.OfficeThemeChangedEventArgs}.
         *
         * [Api set: Mailbox preview]
         *
         * @beta
         */
        OfficeThemeChanged,
        /**
         * Triggers when the recipient list of the selected item or the appointment location is changed in Outlook. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.7 | Office.RecipientsChangedEventArgs}.
         *
         * [Api set: Mailbox 1.7]
         */
        RecipientsChanged,
        /**
         * Triggers when the recurrence pattern of the selected series is changed in Outlook. Supported with task pane only.
         *
         * The event handler receives an argument of type
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.7 | Office.RecurrenceChangedEventArgs}.
         *
         * [Api set: Mailbox 1.7]
         */
        RecurrenceChanged,
        /**
         * Triggers when a Resource selection happens in Project.
         */
        ResourceSelectionChanged,
        /**
         * A Settings.settingsChanged event was raised.
         *
         * @remarks
         *
         * **Hosts**: Excel, PowerPoint, Word
         */
        SettingsChanged,
        /**
         * Triggers when a Task selection happens in Project.
         */
        TaskSelectionChanged,
        /**
         * Triggers when a View selection happens in Project.
         */
        ViewSelectionChanged
    }
    /**
     * Specifies the format in which to return the document.
     */
    enum FileType {
        /**
         * Returns only the text of the document as a string.
         */
        Text,
        /**
         * Returns the entire document (.pptx, .docx, or .xlsx) in Office Open XML (OOXML) format as a byte array.
         */
        Compressed,
        /**
         * Returns the entire document in PDF format as a byte array.
         */
        Pdf
    }
    /**
     * Specifies whether filtering from the host application is applied when the data is retrieved.
     */
    enum FilterType {
        /**
         * Return all data (not filtered by the host application).
         */
        All,
        /**
         * Return only the visible data (as filtered by the host application).
         */
        OnlyVisible
    }
    /**
     * Specifies the type of place or object to navigate to.
     */
    enum GoToType {
        /**
         * Goes to a binding object using the specified binding id.
         *
         * Supported hosts: Excel, Word
         */
        Binding,
        /**
         * Goes to a named item using that item's name.
         * In Excel, you can use any structured reference for a named range or table: "Worksheet2!Table1"
         *
         * Supported hosts: Excel
         */
        NamedItem,
        /**
         * Goes to a slide using the specified id.
         *
         * Supported hosts: PowerPoint
         */
        Slide,
        /**
         * Goes to the specified index by slide number or {@link Office.Index}.
         *
         * Supported hosts: PowerPoint
         */
        Index
    }
    /**
     * Specifies the relative PowerPoint slide.
     */
    enum Index {
        /**
         * Represents the first PowerPoint slide
         */
        First,
        /**
         * Represents the last PowerPoint slide
         */
        Last,
        /**
         * Represents the next PowerPoint slide
         */
        Next,
        /**
         * Represents the previous PowerPoint slide
         */
        Previous
    }
    /**
     * Specifies whether to select (highlight) the location to navigate to (when using the {@link Office.Document | Document}.goToByIdAsync method).
     */
    enum SelectionMode {
        Default,
        /**
         * The location will be selected (highlighted).
         */
        Selected,
        /**
         * The cursor is moved to the beginning of the location.
         */
        None
    }
    /**
     * Specifies whether values, such as numbers and dates, returned by the invoked method are returned with their formatting applied.
     *
     * @remarks
     *
     * For example, if the valueFormat parameter is specified as "formatted", a number formatted as currency, or a date formatted as mm/dd/yy in the
     * host application will have its formatting preserved. If the valueFormat parameter is specified as "unformatted", a date will be returned in its
     * underlying sequential serial number form.
     */
    enum ValueFormat {
        /**
         * Return unformatted data.
         */
        Unformatted,
        /**
         * Return formatted data.
         */
        Formatted
    }
    // Objects
    /**
    * Represents a binding to a section of the document.
    *
    * The Binding object exposes the functionality possessed by all bindings regardless of type.
    *
    * The Binding object is never called directly. It is the abstract parent class of the objects that represent each type of binding:
    * {@link Office.MatrixBinding}, {@link Office.TableBinding}, or {@link Office.TextBinding}. All three of these objects inherit the getDataAsync
    * and setDataAsync methods from the Binding object that enable to you interact with the data in the binding. They also inherit the id and type
    * properties for querying those property values. Additionally, the MatrixBinding and TableBinding objects expose additional methods for matrix-
    * and table-specific features, such as counting the number of rows and columns.
    *
    * @remarks
    *
    * **Hosts**: Excel, Word
    *
    * **Requirement sets**:
    *
    * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
    *
    * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
    *
    * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
    */
    export interface Binding {
        /**
        * Get the Document object associated with the binding.
        */
        document: Office.Document;
        /**
         * A string that uniquely identifies this binding among the bindings in the same {@link Office.Document} object.
         */
        id: string;
        /**
        * Gets the type of the binding.
        */
        type: Office.BindingType;
        /**
         * Adds an event handler to the object for the specified {@link Office.EventType}. Supported EventTypes are
         * `Office.EventType.BindingDataChanged` and `Office.EventType.BindingSelectionChanged`.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents | BindingEvents}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - The event type. For bindings, it can be `Office.EventType.BindingDataChanged` or `Office.EventType.BindingSelectionChanged`.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.BindingDataChangedEventArgs} or {@link Office.BindingSelectionChangedEventArgs}.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, options?: Office.AsyncContextOptions, callback?: (result: Office.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler to the object for the specified {@link Office.EventType}. Supported EventTypes are
         * `Office.EventType.BindingDataChanged` and `Office.EventType.BindingSelectionChanged`.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents | BindingEvents}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - The event type. For bindings, it can be `Office.EventType.BindingDataChanged` or `Office.EventType.BindingSelectionChanged`.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.BindingDataChangedEventArgs} or {@link Office.BindingSelectionChangedEventArgs}.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, callback?: (result: Office.AsyncResult<void>) => void): void;
        /**
         * Returns the data contained within the binding.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * When called from a MatrixBinding or TableBinding, the getDataAsync method will return a subset of the bound values if the optional startRow,
         * startColumn, rowCount, and columnCount parameters are specified (and they specify a contiguous and valid range).
         *
         * @param options - Provides options for how to get the data in a binding.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the values in the specified binding.
         *                  If the `coercionType` parameter is specified (and the call is successful), the data is returned in the format described in the CoercionType enumeration topic.
         */
        getDataAsync<T>(options?: GetBindingDataOptions, callback?: (result: AsyncResult<T>) => void): void;
        /**
         * Returns the data contained within the binding.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * When called from a MatrixBinding or TableBinding, the getDataAsync method will return a subset of the bound values if the optional startRow,
         * startColumn, rowCount, and columnCount parameters are specified (and they specify a contiguous and valid range).
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the values in the specified binding.
         *                  If the `coercionType` parameter is specified (and the call is successful), the data is returned in the format described in the CoercionType enumeration topic.
         */
        getDataAsync<T>(callback?: (result: AsyncResult<T>) => void): void;
        /**
         * Removes the specified handler from the binding for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents | BindingEvents}
         *
         * @param eventType - The event type. For bindings, it can be `Office.EventType.BindingDataChanged` or `Office.EventType.BindingSelectionChanged`.
         * @param options - Provides options to determine which event handler or handlers are removed.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, options?: RemoveHandlerOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Removes the specified handler from the binding for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents | BindingEvents}
         *
         * @param eventType - The event type. For bindings, it can be `Office.EventType.BindingDataChanged` or `Office.EventType.BindingSelectionChanged`.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Writes data to the bound section of the document represented by the specified binding object.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * The value passed for data contains the data to be written in the binding. The kind of value passed determines what will be written as
         * described in the following table.
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>Plain text or anything that can be coerced to a string will be written.</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix")</td>
         *     <td>Tabular data without headers will be written. For example, to write data to three rows in two columns, you can pass an array like this: `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. To write a single column of three rows, pass an array like this: `[["R1C1"], ["R2C1"], ["R3C1"]]`.</td>
         *   </tr>
         *    <tr>
         *     <td>A `TableData` object</td>
         *     <td>A table with headers will be written.</td>
         *   </tr>
         * </table>
         *
         * Additionally, these application-specific actions apply when writing data to a binding. For Word, the specified data is written to the
         * binding as follows:
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>The specified text is written.</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix") or a `TableData` object</td>
         *     <td>A Word table is written.</td>
         *   </tr>
         *   <tr>
         *     <td>HTML</td>
         *     <td>The specified HTML is written. If any of the HTML you write is invalid, Word will not raise an error. Word will write as much of the HTML as it can and will omit any invalid data.</td>
         *   </tr>
         *   <tr>
         *     <td>Office Open XML ("Open XML")</td>
         *     <td>The specified the XML is written.</td>
         *   </tr>
         * </table>
         *
         * For Excel, the specified data is written to the binding as follows:
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>The specified text is inserted as the value of the first bound cell.You can also specify a valid formula to add that formula to the bound cell. For example, setting  data to `"=SUM(A1:A5)"` will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added formula (or any pre-existing formula) from the bound cell. If you call the Binding.getDataAsync method on the bound cell to read its data, the method can return only the data displayed in the cell (the formula's result).</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix"), and the shape exactly matches the shape of the binding specified</td>
         *     <td>The set of rows and columns are written.You can also specify an array of arrays that contain valid formulas to add them to the bound cells. For example, setting  data to `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` will add those two formulas to a binding that contains two cells. Just as when setting a formula on a single bound cell, you can't read the added formulas (or any pre-existing formulas) from the binding with the `Binding.getDataAsync` method - it returns only the data displayed in the bound cells.</td>
         *   </tr>
         *   <tr>
         *     <td>A `TableData` object, and the shape of the table matches the bound table.</td>
         *     <td>The specified set of rows and/or headers are written, if no other data in surrounding cells will be overwritten. **Note**: If you specify formulas in the TableData object you pass for the *data* parameter, you might not get the results you expect due to the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to write *data* that contains formulas to a bound table, try specifying the data as an array of arrays (instead of a TableData object), and specify the *coercionType* as Microsoft.Office.Matrix or "matrix".</td>
         *   </tr>
         * </table>
         *
         * For Excel on the web:
         *
         *  - The total number of cells in the value passed to the data parameter can't exceed 20,000 in a single call to this method.
         *
         *  - The number of formatting groups passed to the cellFormat parameter can't exceed 100.
         * A single formatting group consists of a set of formatting applied to a specified range of cells.
         *
         * In all other cases, an error is returned.
         *
         * The setDataAsync method will write data in a subset of a table or matrix binding if the optional startRow and startColumn parameters are
         * specified, and they specify a valid range.
         *
         * In the callback function passed to the setDataAsync method, you can use the properties of the AsyncResult object to return the following
         * information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param data - The data to be set in the current selection. Possible data types by host:
         *
         *        string: Excel on the web and Windows, and Word on the web and Windows only
         *
         *        array of arrays: Excel and Word only
         *
         *        {@link Office.TableData}: Excel and Word only
         *
         *        HTML: Word on the web and Windows only
         *
         *        Office Open XML: Word only
         *
         * @param options - Provides options for how to set the data in a binding.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setDataAsync(data: TableData | any, options?: SetBindingDataOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Writes data to the bound section of the document represented by the specified binding object.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * The value passed for data contains the data to be written in the binding. The kind of value passed determines what will be written as
         * described in the following table.
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>Plain text or anything that can be coerced to a string will be written.</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix")</td>
         *     <td>Tabular data without headers will be written. For example, to write data to three rows in two columns, you can pass an array like this: `[["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`. To write a single column of three rows, pass an array like this: `[["R1C1"], ["R2C1"], ["R3C1"]]`.</td>
         *   </tr>
         *    <tr>
         *     <td>A `TableData` object</td>
         *     <td>A table with headers will be written.</td>
         *   </tr>
         * </table>
         *
         * Additionally, these application-specific actions apply when writing data to a binding. For Word, the specified data is written to the
         * binding as follows:
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>The specified text is written.</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix") or a `TableData` object</td>
         *     <td>A Word table is written.</td>
         *   </tr>
         *   <tr>
         *     <td>HTML</td>
         *     <td>The specified HTML is written. If any of the HTML you write is invalid, Word will not raise an error. Word will write as much of the HTML as it can and will omit any invalid data.</td>
         *   </tr>
         *   <tr>
         *     <td>Office Open XML ("Open XML")</td>
         *     <td>The specified the XML is written.</td>
         *   </tr>
         * </table>
         *
         * For Excel, the specified data is written to the binding as follows:
         *
         * <table>
         *   <tr>
         *     <th>`data` value</th>
         *     <th>Data written</th>
         *   </tr>
         *   <tr>
         *     <td>A string</td>
         *     <td>The specified text is inserted as the value of the first bound cell.You can also specify a valid formula to add that formula to the bound cell. For example, setting  data to `"=SUM(A1:A5)"` will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added formula (or any pre-existing formula) from the bound cell. If you call the Binding.getDataAsync method on the bound cell to read its data, the method can return only the data displayed in the cell (the formula's result).</td>
         *   </tr>
         *   <tr>
         *     <td>An array of arrays ("matrix"), and the shape exactly matches the shape of the binding specified</td>
         *     <td>The set of rows and columns are written.You can also specify an array of arrays that contain valid formulas to add them to the bound cells. For example, setting  data to `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` will add those two formulas to a binding that contains two cells. Just as when setting a formula on a single bound cell, you can't read the added formulas (or any pre-existing formulas) from the binding with the `Binding.getDataAsync` method - it returns only the data displayed in the bound cells.</td>
         *   </tr>
         *   <tr>
         *     <td>A `TableData` object, and the shape of the table matches the bound table.</td>
         *     <td>The specified set of rows and/or headers are written, if no other data in surrounding cells will be overwritten. **Note**: If you specify formulas in the TableData object you pass for the *data* parameter, you might not get the results you expect due to the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to write *data* that contains formulas to a bound table, try specifying the data as an array of arrays (instead of a TableData object), and specify the *coercionType* as Microsoft.Office.Matrix or "matrix".</td>
         *   </tr>
         * </table>
         *
         * For Excel on the web:
         *
         *  - The total number of cells in the value passed to the data parameter can't exceed 20,000 in a single call to this method.
         *
         *  - The number of formatting groups passed to the cellFormat parameter can't exceed 100.
         * A single formatting group consists of a set of formatting applied to a specified range of cells.
         *
         * In all other cases, an error is returned.
         *
         * The setDataAsync method will write data in a subset of a table or matrix binding if the optional startRow and startColumn parameters are
         * specified, and they specify a valid range.
         *
         * In the callback function passed to the setDataAsync method, you can use the properties of the AsyncResult object to return the following
         * information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param data - The data to be set in the current selection. Possible data types by host:
         *
         *        string: Excel on the web and Windows, and Word on the web and Windows only
         *
         *        array of arrays: Excel and Word only
         *
         *        `TableData`: Excel and Word only
         *
         *        HTML: Word on the web and Windows only
         *
         *        Office Open XML: Word only
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setDataAsync(data: TableData | any, callback?: (result: AsyncResult<void>) => void): void;
    }

    /**
     * Provides information about the binding that raised the DataChanged event.
     */
    export interface BindingDataChangedEventArgs {
        /**
         * Gets an {@link Office.Binding} object that represents the binding that raised the DataChanged event.
         */
        binding: Binding;

        /**
         * Gets an {@link Office.EventType} enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }

    /**
     * Provides information about the binding that raised the SelectionChanged event.
     */
    export interface BindingSelectionChangedEventArgs {
        /**
         * Gets an {@link Office.Binding} object that represents the binding that raised the SelectionChanged event.
         */
        binding: Binding;
        /**
         * Gets the number of columns selected. If a single cell is selected returns 1.
         *
         * If the user makes a non-contiguous selection, the count for the last contiguous selection within the binding is returned.
         *
         * For Word, this property will work only for bindings of {@link Office.BindingType} "table". If the binding is of type "matrix", null is
         * returned. Also, the call will fail if the table contains merged cells, because the structure of the table must be uniform for this property
         * to work correctly.
         */
        columnCount: number;
        /**
         * Gets the number of rows selected. If a single cell is selected returns 1.
         *
         * If the user makes a non-contiguous selection, the count for the last contiguous selection within the binding is returned.
         *
         * For Word, this property will work only for bindings of {@link Office.BindingType} "table". If the binding is of type "matrix", null is
         * returned. Also, the call will fail if the table contains merged cells, because the structure of the table must be uniform for this property
         * to work correctly.
         */
        rowCount: number;
        /**
         * The zero-based index of the first column of the selection counting from the leftmost column in the binding.
         *
         * If the user makes a non-contiguous selection, the coordinates for the last contiguous selection within the binding are returned.
         *
         * For Word, this property will work only for bindings of {@link Office.BindingType} "table". If the binding is of type "matrix", null is
         * returned. Also, the call will fail if the table contains merged cells, because the structure of the table must be uniform for this property
         * to work correctly.
         */
        startColumn: number;
        /**
         * The zero-based index of the first row of the selection counting from the first row in the binding.
         *
         * If the user makes a non-contiguous selection, the coordinates for the last contiguous selection within the binding are returned.
         *
         * For Word, this property will work only for bindings of {@link Office.BindingType} "table". If the binding is of type "matrix", null is
         * returned. Also, the call will fail if the table contains merged cells, because the structure of the table must be uniform for this property
         * to work correctly.
         */
        startRow: number;
        /**
         * Gets an {@link Office.EventType} enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }

    /**
    * Represents the bindings the add-in has within the document.
    */
    export interface Bindings {
        /**
         * Gets an {@link Office.Document} object that represents the document associated with this set of bindings.
         */
        document: Document;
        /**
         * Creates a binding against a named object in the document.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * For Excel, the itemName parameter can refer to a named range or a table.
         *
         * By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on.
         * To assign a meaningful name for a table in the Excel UI, use the Table Name property on the Table Tools | Design tab of the ribbon.
         *
         *     **Note**: In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of
         * the table in this format: "Sheet1!Table1"
         *
         * For Word, the itemName parameter refers to the Title property of a Rich Text content control. (You can't bind to content controls other
         * than the Rich Text content control).
         *
         * By default, a content control has no Title value assigned. To assign a meaningful name in the Word UI, after inserting a Rich Text content
         * control from the Controls group on the Developer tab of the ribbon, use the Properties command in the Controls group to display the Content
         * Control Properties dialog box. Then set the Title property of the content control to the name you want to reference from your code.
         *
         *     **Note**: In Word, if there are multiple Rich Text content controls with the same Title property value (name), and you try to bind to one
         * these content controls with this method (by specifying its name as the itemName parameter), the operation will fail.
         *
         * @param itemName - Name of the bindable object in the document. For Example 'MyExpenses' table in Excel."
         * @param bindingType - The {@link Office.BindingType} for the data. The method returns null if the selected object cannot be coerced into the specified type.
         * @param options - Provides options for configuring the binding that is created.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the specified named item.
         */
        addFromNamedItemAsync(itemName: string, bindingType: BindingType, options?: AddBindingFromNamedItemOptions, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Creates a binding against a named object in the document.
         *
         * @remarks
         *
         * {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings},
         * {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings},
         * {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * For Excel, the itemName parameter can refer to a named range or a table.
         *
         * By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on.
         * To assign a meaningful name for a table in the Excel UI, use the Table Name property on the Table Tools | Design tab of the ribbon.
         *
         *     **Note**: In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of
         * the table in this format: "Sheet1!Table1"
         *
         * For Word, the itemName parameter refers to the Title property of a Rich Text content control. (You can't bind to content controls other
         * than the Rich Text content control).
         *
         * By default, a content control has no Title value assigned. To assign a meaningful name in the Word UI, after inserting a Rich Text content
         * control from the Controls group on the Developer tab of the ribbon, use the Properties command in the Controls group to display the Content
         * Control Properties dialog box. Then set the Title property of the content control to the name you want to reference from your code.
         *
         *     **Note**: In Word, if there are multiple Rich Text content controls with the same Title property value (name), and you try to bind to one
         * these content controls with this method (by specifying its name as the itemName parameter), the operation will fail.
         *
         * @param itemName - Name of the bindable object in the document. For Example 'MyExpenses' table in Excel."
         * @param bindingType - The {@link Office.BindingType} for the data. The method returns null if the selected object cannot be coerced into the specified type.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the specified named item.
         */
        addFromNamedItemAsync(itemName: string, bindingType: BindingType, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Create a binding by prompting the user to make a selection on the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * Adds a binding object of the specified type to the Bindings collection, which will be identified with the supplied id.
         * The method fails if the specified selection cannot be bound.
         *
         * @param bindingType - Specifies the type of the binding object to create. Required.
         *                    Returns null if the selected object cannot be coerced into the specified type.
         * @param options - Provides options for configuring the prompt and identifying the binding that is created.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the selection specified by the user.
         */
        addFromPromptAsync(bindingType: BindingType, options?: AddBindingFromPromptOptions, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Create a binding by prompting the user to make a selection on the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * Adds a binding object of the specified type to the Bindings collection, which will be identified with the supplied id.
         * The method fails if the specified selection cannot be bound.
         *
         * @param bindingType - Specifies the type of the binding object to create. Required.
         *                    Returns null if the selected object cannot be coerced into the specified type.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the selection specified by the user.
         */
        addFromPromptAsync(bindingType: BindingType, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Create a binding based on the user's current selection.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Adds the specified type of binding object to the Bindings collection, which will be identified with the supplied id.
         *
         * Note In Excel, if you call the addFromSelectionAsync method passing in the Binding.id of an existing binding, the Binding.type of that
         * binding is used, and its type cannot be changed by specifying a different value for the bindingType parameter.
         * If you need to use an existing id and change the bindingType, call the Bindings.releaseByIdAsync method first to release the binding, and
         * then call the addFromSelectionAsync method to reestablish the binding with a new type.
         *
         * @param bindingType - Specifies the type of the binding object to create. Required.
         *                    Returns null if the selected object cannot be coerced into the specified type.
         * @param options - Provides options for identifying the binding that is created.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the selection specified by the user.
         */
        addFromSelectionAsync(bindingType: BindingType, options?: AddBindingFromSelectionOptions, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Create a binding based on the user's current selection.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Adds the specified type of binding object to the Bindings collection, which will be identified with the supplied id.
         *
         * Note In Excel, if you call the addFromSelectionAsync method passing in the Binding.id of an existing binding, the Binding.type of that
         * binding is used, and its type cannot be changed by specifying a different value for the bindingType parameter.
         * If you need to use an existing id and change the bindingType, call the Bindings.releaseByIdAsync method first to release the binding, and
         * then call the addFromSelectionAsync method to reestablish the binding with a new type.
         *
         * @param bindingType - Specifies the type of the binding object to create. Required.
         *                    Returns null if the selected object cannot be coerced into the specified type.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object that represents the selection specified by the user.
         */
        addFromSelectionAsync(bindingType: BindingType, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Gets all bindings that were previously created.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array that contains each binding created for the referenced Bindings object.
         */
        getAllAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<Binding[]>) => void): void;
        /**
         * Gets all bindings that were previously created.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array that contains each binding created for the referenced Bindings object.
         */
        getAllAsync(callback?: (result: AsyncResult<Binding[]>) => void): void;
        /**
         * Retrieves a binding based on its Name
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Fails if the specified id does not exist.
         *
         * @param id - Specifies the unique name of the binding object. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object specified by the id in the call.
         */
        getByIdAsync(id: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Retrieves a binding based on its Name
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Fails if the specified id does not exist.
         *
         * @param id - Specifies the unique name of the binding object. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the Binding object specified by the id in the call.
         */
        getByIdAsync(id: string, callback?: (result: AsyncResult<Binding>) => void): void;
        /**
         * Removes the binding from the document
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Fails if the specified id does not exist.
         *
         * @param id - Specifies the unique name to be used to identify the binding object. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        releaseByIdAsync(id: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Removes the binding from the document
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings | MatrixBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings | TableBindings}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings | TextBindings}
         *
         * Fails if the specified id does not exist.
         *
         * @param id - Specifies the unique name to be used to identify the binding object. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        releaseByIdAsync(id: string, callback?: (result: AsyncResult<void>) => void): void;
    }
    /**
     * Represents an XML node in a tree in a document.
     *
     * @remarks
     *
     * **Hosts**: Word
     */
    export interface CustomXmlNode {
        /**
         * Gets the base name of the node without the namespace prefix, if one exists.
         */
        baseName: string;
        /**
         * Retrieves the string GUID of the CustomXMLPart.
         */
        namespaceUri: string;
        /**
         * Gets the type of the CustomXMLNode.
         */
        nodeType: string;
        /**
         * Gets the nodes associated with the XPath expression.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xPath - The XPath expression that specifies the nodes to get. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlNode objects that represent the nodes specified by the XPath expression passed to the `xPath` parameter.
         */
        getNodesAsync(xPath: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<CustomXmlNode[]>) => void): void;
        /**
         * Gets the nodes associated with the XPath expression.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xPath - The XPath expression that specifies the nodes to get. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlNode objects that represent the nodes specified by the XPath expression passed to the `xPath` parameter.
         */
        getNodesAsync(xPath: string, callback?: (result: AsyncResult<CustomXmlNode[]>) => void): void;
        /**
         * Gets the node value.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the value of the referenced node.
         */
        getNodeValueAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Gets the node value.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the value of the referenced node.
         */
        getNodeValueAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Gets the text of an XML node in a custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the inner text of the referenced nodes.
         */
        getTextAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Gets the text of an XML node in a custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the inner text of the referenced nodes.
         */
        getTextAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Gets the node's XML.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the XML of the referenced node.
         */
        getXmlAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Gets the node's XML.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the XML of the referenced node.
         */
        getXmlAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Sets the node value.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param value - The value to be set on the node
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setNodeValueAsync(value: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Sets the node value.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param value - The value to be set on the node
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setNodeValueAsync(value: string, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Asynchronously sets the text of an XML node in a custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param text - Required. The text value of the XML node.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setTextAsync(text: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Asynchronously sets the text of an XML node in a custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param text - Required. The text value of the XML node.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setTextAsync(text: string, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Sets the node XML.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xml - The XML to be set on the node
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setXmlAsync(xml: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Sets the node XML.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xml - The XML to be set on the node
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setXmlAsync(xml: string, callback?: (result: AsyncResult<void>) => void): void;
    }
    /**
     * Represents a single CustomXMLPart in an {@link Office.CustomXmlParts} collection.
     *
     * @remarks
     *
     * **Hosts**: Word
     */
    export interface CustomXmlPart {
        /**
         * True, if the custom XML part is built in; otherwise false.
         */
        builtIn: boolean;
        /**
         * Gets the GUID of the CustomXMLPart.
         */
        id: string;
        /**
         * Gets the set of namespace prefix mappings ({@link Office.CustomXmlPrefixMappings}) used against the current CustomXmlPart.
         */
        namespaceManager: CustomXmlPrefixMappings;
        /**
         * Adds an event handler to the object using the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - Specifies the type of event to add. For a CustomXmlPart object, the eventType parameter can be specified as
         *                  `Office.EventType.NodeDeleted`, `Office.EventType.NodeInserted`, and `Office.EventType.NodeReplaced`.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.NodeDeletedEventArgs},
         *                {@link Office.NodeInsertedEventArgs}, or {@link Office.NodeReplacedEventArgs}
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: (result: any) => void, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds an event handler to the object using the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - Specifies the type of event to add. For a CustomXmlPart object, the eventType parameter can be specified as
         *                  `Office.EventType.NodeDeleted`, `Office.EventType.NodeInserted`, and `Office.EventType.NodeReplaced`.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.NodeDeletedEventArgs},
         *                {@link Office.NodeInsertedEventArgs}, or {@link Office.NodeReplacedEventArgs}
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: (result: any) => void, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Deletes the Custom XML Part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        deleteAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Deletes the Custom XML Part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        deleteAsync(callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Asynchronously gets any CustomXmlNodes in this custom XML part which match the specified XPath.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xPath - An XPath expression that specifies the nodes you want returned. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlNode objects that represent the nodes specified by the XPath expression passed to the xPath parameter.
         */
        getNodesAsync(xPath: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<CustomXmlNode[]>) => void): void;
        /**
         * Asynchronously gets any CustomXmlNodes in this custom XML part which match the specified XPath.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xPath - An XPath expression that specifies the nodes you want returned. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlNode objects that represent the nodes specified by the XPath expression passed to the xPath parameter.
         */
        getNodesAsync(xPath: string, callback?: (result: AsyncResult<CustomXmlNode[]>) => void): void;
        /**
         * Asynchronously gets the XML inside this custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the XML of the referenced CustomXmlPart object.
         */
        getXmlAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the XML inside this custom XML part.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the XML of the referenced CustomXmlPart object.
         */
        getXmlAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Removes an event handler for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param eventType - Specifies the type of event to remove. For a CustomXmlPart object, the eventType parameter can be specified as
         *                  `Office.EventType.NodeDeleted`, `Office.EventType.NodeInserted`, and `Office.EventType.NodeReplaced`.
         * @param handler - The name of the handler to remove.
         * @param options - Provides options to determine which event handler or handlers are removed.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, handler?: (result: any) => void, options?: RemoveHandlerOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Removes an event handler for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param eventType - Specifies the type of event to remove. For a CustomXmlPart object, the eventType parameter can be specified as
         *                  `Office.EventType.NodeDeleted`, `Office.EventType.NodeInserted`, and `Office.EventType.NodeReplaced`.
         * @param handler - The name of the handler to remove.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, handler?: (result: any) => void, callback?: (result: AsyncResult<void>) => void): void;
    }

    /**
     * Provides information about the deleted node that raised the nodeDeleted event.
     */
    export interface NodeDeletedEventArgs {
        /**
         * Gets whether the node was deleted as part of an Undo/Redo action by the user.
         */
        isUndoRedo: boolean;
        /**
         * Gets the former next sibling of the node that was just deleted from the {@link Office.CustomXmlPart} object.
         */
        oldNextSibling: CustomXmlNode;
        /**
         * Gets the node which was just deleted from the {@link Office.CustomXmlPart} object.
         *
         * Note that this node may have children, if a subtree is being removed from the document. Also, this node will be a "disconnected" node in
         * that you can query down from the node, but you cannot query up the tree - the node appears to exist alone.
         */
        oldNode: CustomXmlNode;
    }

    /**
     * Provides information about the inserted node that raised the nodeInserted event.
     */
    export interface NodeInsertedEventArgs  {
        /**
         * Gets whether the node was inserted as part of an Undo/Redo action by the user.
         */
        isUndoRedo: boolean;
        /**
         * Gets the node that was just added to the CustomXMLPart object.
         *
         * Note that this node may have children, if a subtree was just added to the document.
         */
        newNode: CustomXmlNode;
    }

    /**
     * Provides information about the replaced node that raised the nodeReplaced event.
     */
    export interface NodeReplacedEventArgs  {
        /**
         * Gets whether the replaced node was inserted as part of an undo or redo operation by the user.
         */
        isUndoRedo: boolean;
        /**
         * Gets the node that was just added to the CustomXMLPart object.
         *
         * Note that this node may have children, if a subtree was just added to the document.
         */
        newNode: CustomXmlNode;
        /**
         * Gets the node which was just deleted (replaced) from the CustomXmlPart object.
         *
         * Note that this node may have children, if a subtree is being removed from the document. Also, this node will be a "disconnected" node in
         * that you can query down from the node, but you cannot query up the tree - the node appears to exist alone.
         */
        oldNode: CustomXmlNode;
    }

    /**
     * Represents a collection of CustomXmlPart objects.
     *
     * @remarks
     *
     * **Hosts**: Word
     */
    export interface CustomXmlParts {
        /**
         * Asynchronously adds a new custom XML part to a file.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xml - The XML to add to the newly created custom XML part.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the newly created CustomXmlPart object.
         */
        addAsync(xml: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<CustomXmlPart>) => void): void;
        /**
         * Asynchronously adds a new custom XML part to a file.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param xml - The XML to add to the newly created custom XML part.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the newly created CustomXmlPart object.
         */
        addAsync(xml: string, callback?: (result: AsyncResult<CustomXmlPart>) => void): void;
        /**
         * Asynchronously gets the specified custom XML part by its id.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param id - The GUID of the custom XML part, including opening and closing braces.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a CustomXmlPart object that represents the specified custom XML part.
         *                  If there is no custom XML part with the specified id, the method returns null.
         */
        getByIdAsync(id: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<CustomXmlPart>) => void): void;
        /**
         * Asynchronously gets the specified custom XML part by its id.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param id - The GUID of the custom XML part, including opening and closing braces.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a CustomXmlPart object that represents the specified custom XML part.
         *                  If there is no custom XML part with the specified id, the method returns null.
         */
        getByIdAsync(id: string, callback?: (result: AsyncResult<CustomXmlPart>) => void): void;
        /**
         * Asynchronously gets the specified custom XML part(s) by its namespace.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param ns - The namespace URI.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlPart objects that match the specified namespace.
        */
        getByNamespaceAsync(ns: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<CustomXmlPart[]>) => void): void;
        /**
         * Asynchronously gets the specified custom XML part(s) by its namespace.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * @param ns - The namespace URI.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array of CustomXmlPart objects that match the specified namespace.
        */
        getByNamespaceAsync(ns: string, callback?: (result: AsyncResult<CustomXmlPart[]>) => void): void;
    }
    /**
     * Represents a collection of CustomXmlPart objects.
     *
     * @remarks
     *
     * **Hosts**: Word
     */
    export interface CustomXmlPrefixMappings {
        /**
         * Asynchronously adds a prefix to namespace mapping to use when querying an item.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If no namespace is assigned to the requested prefix, the method returns an empty string ("").
         *
         * @param prefix - Specifies the prefix to add to the prefix mapping list. Required.
         * @param ns - Specifies the namespace URI to assign to the newly added prefix. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addNamespaceAsync(prefix: string, ns: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Asynchronously adds a prefix to namespace mapping to use when querying an item.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If no namespace is assigned to the requested prefix, the method returns an empty string ("").
         *
         * @param prefix - Specifies the prefix to add to the prefix mapping list. Required.
         * @param ns - Specifies the namespace URI to assign to the newly added prefix. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addNamespaceAsync(prefix: string, ns: string, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Asynchronously gets the namespace mapped to the specified prefix.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If the prefix already exists in the namespace manager, this method will overwrite the mapping of that prefix except when the prefix is one
         * added or used by the data store internally, in which case it will return an error.
         *
         * @param prefix - TSpecifies the prefix to get the namespace for. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the namespace mapped to the specified prefix.
         */
        getNamespaceAsync(prefix: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the namespace mapped to the specified prefix.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If the prefix already exists in the namespace manager, this method will overwrite the mapping of that prefix except when the prefix is one
         * added or used by the data store internally, in which case it will return an error.
         *
         * @param prefix - TSpecifies the prefix to get the namespace for. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the namespace mapped to the specified prefix.
         */
        getNamespaceAsync(prefix: string, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the prefix for the specified namespace.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If no prefix is assigned to the requested namespace, the method returns an empty string (""). If there are multiple prefixes specified in
         * the namespace manager, the method returns the first prefix that matches the supplied namespace.
         *
         * @param ns - Specifies the namespace to get the prefix for. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the prefix of the specified namespace.
         */
        getPrefixAsync(ns: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the prefix for the specified namespace.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts | CustomXmlParts}
         *
         * If no prefix is assigned to the requested namespace, the method returns an empty string (""). If there are multiple prefixes specified in
         * the namespace manager, the method returns the first prefix that matches the supplied namespace.
         *
         * @param ns - Specifies the namespace to get the prefix for. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is a string that contains the prefix of the specified namespace.
         */
        getPrefixAsync(ns: string, callback?: (result: AsyncResult<string>) => void): void;
    }
    /**
     * An abstract class that represents the document the add-in is interacting with.
     *
     * @remarks
     *
     * **Hosts**: Excel, PowerPoint, Project, Word
     */
    export interface Document {
        /**
         * Gets an object that provides access to the bindings defined in the document.
         *
         * @remarks
         *
         * You don't instantiate the Document object directly in your script. To call members of the Document object to interact with the current
         * document or worksheet, use `Office.context.document` in your script.
         */
        bindings: Bindings;
        /**
         * Gets an object that represents the custom XML parts in the document.
         */
        customXmlParts: CustomXmlParts;
        /**
         * Gets the mode the document is in.
         */
        mode: DocumentMode;
        /**
         * Gets an object that represents the saved custom settings of the content or task pane add-in for the current document.
         */
        settings: Settings;
        /**
         * Gets the URL of the document that the host application currently has open. Returns null if the URL is unavailable.
         */
        url: string;
        /**
         * Adds an event handler for a Document object event.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents | DocumentEvents}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - For a Document object event, the eventType parameter can be specified as `Office.EventType.Document.SelectionChanged` or
         *                  `Office.EventType.Document.ActiveViewChanged`, or the corresponding text value of this enumeration.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.DocumentSelectionChangedEventArgs}. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a Document object event.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents | DocumentEvents}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - For a Document object event, the eventType parameter can be specified as `Office.EventType.Document.SelectionChanged` or
         *                  `Office.EventType.Document.ActiveViewChanged`, or the corresponding text value of this enumeration.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.DocumentSelectionChangedEventArgs}. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Returns the state of the current view of the presentation (edit or read).
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview | ActiveView}
         *
         * Can trigger an event when the view changes.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the state of the presentation's current view.
         *                  The value returned can be either "edit" or "read". "edit" corresponds to any of the views in which you can edit slides,
         *                  such as Normal or Outline View. "read" corresponds to either Slide Show or Reading View.
         */
        getActiveViewAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<"edit" | "read">) => void): void;
        /**
         * Returns the state of the current view of the presentation (edit or read).
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview | ActiveView}
         *
         * Can trigger an event when the view changes.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the state of the presentation's current view.
         *                  The value returned can be either "edit" or "read". "edit" corresponds to any of the views in which you can edit slides,
         *                  such as Normal or Outline View. "read" corresponds to either Slide Show or Reading View.
         */
        getActiveViewAsync(callback?: (result: AsyncResult<"edit" | "read">) => void): void;
        /**
         * Returns the entire document file in slices of up to 4194304 bytes (4 MB). For add-ins on iPad, file slice is supported up to 65536 (64 KB).
         * Note that specifying file slice size of above permitted limit will result in an "Internal Error" failure.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile | CompressedFile} (when using `Office.FileType.Compressed`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file | File}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile | TextFile} (when using `Office.FileType.Text`)
         *
         * For add-ins running in Office host applications other than Office on iPad, the `getFileAsync` method supports getting files in slices of up
         * to 4194304 bytes (4 MB). For add-ins running in Office apps on iPad, the `getFileAsync` method supports getting files in slices of up to
         * 65536 (64 KB).
         *
         * The `fileType` parameter can be specified by using the {@link Office.FileType} enumeration or text values. But the possible values vary with
         * the host:
         *
         * *Supported FileTypes, by platform*
         *  <table>
         *   <tr><th>                             </th><th> Office on Windows           </th><th> Office on the web           </th><th> Office on iPad      </th><th> Office on Mac               </th></tr>
         *   <tr><td><strong> Excel      </strong></td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`, `Pdf`         </td><td>                     </td><td> `Compressed`, `Pdf`, `Text` </td></tr>
         *   <tr><td><strong> PowerPoint </strong></td><td> `Compressed`, `Pdf`         </td><td> `Compressed`, `Pdf`         </td><td> `Compressed`, `Pdf` </td><td> `Compressed`, `Pdf`         </td></tr>
         *   <tr><td><strong> Word       </strong></td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`        </td><td> `Compressed`, `Pdf`, `Text` </td></tr>
         *  </table>
         *
         * @param fileType - The format in which the file will be returned
         * @param options - Provides options for setting the size of slices that the document will be divided into.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the File object.
         */
        getFileAsync(fileType: FileType, options?: GetFileOptions, callback?: (result: AsyncResult<Office.File>) => void): void;
        /**
         * Returns the entire document file in slices of up to 4194304 bytes (4 MB). For add-ins on iPad, file slice is supported up to 65536 (64 KB).
         * Note that specifying file slice size of above permitted limit will result in an "Internal Error" failure.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile | CompressedFile} (when using `Office.FileType.Compressed`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file | File}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile | TextFile} (when using `Office.FileType.Text`)
         *
         * For add-ins running in Office host applications other than Office on iPad, the `getFileAsync` method supports getting files in slices of up
         * to 4194304 bytes (4 MB). For add-ins running in Office on iPad apps, the `getFileAsync` method supports getting files in slices of up to
         * 65536 (64 KB).
         *
         * The `fileType` parameter can be specified by using the {@link Office.FileType} enumeration or text values. But the possible values vary with
         * the host:
         *
         * *Supported FileTypes, by platform*
         *  <table>
         *   <tr><th>                             </th><th> Office on Windows           </th><th> Office on the web           </th><th> Office on iPad      </th><th> Office on Mac               </th></tr>
         *   <tr><td><strong> Excel      </strong></td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`, `Pdf`         </td><td>                     </td><td> `Compressed`, `Pdf`, `Text` </td></tr>
         *   <tr><td><strong> PowerPoint </strong></td><td> `Compressed`, `Pdf`         </td><td> `Compressed`, `Pdf`         </td><td> `Compressed`, `Pdf` </td><td> `Compressed`, `Pdf`         </td></tr>
         *   <tr><td><strong> Word       </strong></td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`, `Pdf`, `Text` </td><td> `Compressed`        </td><td> `Compressed`, `Pdf`, `Text` </td></tr>
         *  </table>
         *
         * @param fileType - The format in which the file will be returned
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the File object.
         */
        getFileAsync(fileType: FileType, callback?: (result: AsyncResult<Office.File>) => void): void;
        /**
         * Gets file properties of the current document.
         *
         * @remarks
         *
         * **Requirement sets**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * You get the file's URL with the url property `asyncResult.value.url`.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the file's properties (with the URL found at `asyncResult.value.url`).
         */
        getFilePropertiesAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<Office.FileProperties>) => void): void;
        /**
         * Gets file properties of the current document.
         *
         * @remarks
         *
         * **Requirement sets**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * You get the file's URL with the url property `asyncResult.value.url`.
         *
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the file's properties (with the URL found at `asyncResult.value.url`).
         */
        getFilePropertiesAsync(callback?: (result: AsyncResult<Office.FileProperties>) => void): void;
        /**
         * Reads the data contained in the current selection in the document.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection | Selection}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * In the callback function that is passed to the getSelectedDataAsync method, you can use the properties of the AsyncResult object to return
         * the following information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * The possible values for the {@link Office.CoercionType} parameter vary by the host.
         *
         * <table>
         *   <tr>
         *     <th>Host</th>
         *     <th>Supported coercionType</th>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, Project, and Word</td>
         *     <td>`Office.CoercionType.Text` (string)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Matrix` (array of arrays)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Table` (TableData object)</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Html`</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Ooxml` (Office Open XML)</td>
         *   </tr>
         *   <tr>
         *     <td>PowerPoint on the web and Windows</td>
         *     <td>`Office.CoercionType.SlideRange`</td>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, and Word</td>
         *     <td>`Office.CoercionType.XmlSvg`</td>
         *   </tr>
         * </table>
         *
         * @param coercionType - The type of data structure to return. See the remarks section for each host's supported coercion types.
         *
         * @param options - Provides options for customizing what data is returned and how it is formatted.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the data in the current selection.
         *                  This is returned in the data structure or format you specified with the coercionType parameter.
         *                  (See Remarks for more information about data coercion.)
         */
        getSelectedDataAsync<T>(coercionType: Office.CoercionType, options?: GetSelectedDataOptions, callback?: (result: AsyncResult<T>) => void): void;
        /**
         * Reads the data contained in the current selection in the document.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion} (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection | Selection}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * In the callback function that is passed to the getSelectedDataAsync method, you can use the properties of the AsyncResult object to return
         * the following information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * The possible values for the {@link Office.CoercionType} parameter vary by the host.
         *
         * <table>
         *   <tr>
         *     <th>Host</th>
         *     <th>Supported coercionType</th>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, Project, and Word</td>
         *     <td>`Office.CoercionType.Text` (string)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Matrix` (array of arrays)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Table` (TableData object)</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Html`</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Ooxml` (Office Open XML)</td>
         *   </tr>
         *   <tr>
         *     <td>PowerPoint on the web and Windows</td>
         *     <td>`Office.CoercionType.SlideRange`</td>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, and Word</td>
         *     <td>`Office.CoercionType.XmlSvg`</td>
         *   </tr>
         * </table>
         *
         * @param coercionType - The type of data structure to return. See the remarks section for each host's supported coercion types.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the data in the current selection.
         *                  This is returned in the data structure or format you specified with the coercionType parameter.
         *                  (See Remarks for more information about data coercion.)
         */
        getSelectedDataAsync<T>(coercionType: Office.CoercionType, callback?: (result: AsyncResult<T>) => void): void;
        /**
         * Goes to the specified object or location in the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * PowerPoint doesn't support the goToByIdAsync method in Master Views.
         *
         * The behavior caused by the selectionMode option varies by host:
         *
         * In Excel: `Office.SelectionMode.Selected` selects all content in the binding, or named item. Office.SelectionMode.None for text bindings,
         * selects the cell; for matrix bindings, table bindings, and named items, selects the first data cell (not first cell in header row for tables).
         *
         * In PowerPoint: `Office.SelectionMode.Selected` selects the slide title or first textbox on the slide.
         * `Office.SelectionMode.None` doesn't select anything.
         *
         * In Word: `Office.SelectionMode.Selected` selects all content in the binding. Office.SelectionMode.None for text bindings, moves the cursor
         * to the beginning of the text; for matrix bindings and table bindings, selects the first data cell (not first cell in header row for tables).
         *
         * @param id - The identifier of the object or location to go to.
         * @param goToType - The type of the location to go to.
         * @param options - Provides options for whether to select the location that is navigated to.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the current view.
         */
        goToByIdAsync(id: string | number, goToType: GoToType, options?: GoToByIdOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Goes to the specified object or location in the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * PowerPoint doesn't support the goToByIdAsync method in Master Views.
         *
         * The behavior caused by the selectionMode option varies by host:
         *
         * In Excel: `Office.SelectionMode.Selected` selects all content in the binding, or named item. Office.SelectionMode.None for text bindings,
         * selects the cell; for matrix bindings, table bindings, and named items, selects the first data cell (not first cell in header row for tables).
         *
         * In PowerPoint: `Office.SelectionMode.Selected` selects the slide title or first textbox on the slide.
         * `Office.SelectionMode.None` doesn't select anything.
         *
         * In Word: `Office.SelectionMode.Selected` selects all content in the binding. Office.SelectionMode.None for text bindings, moves the cursor
         * to the beginning of the text; for matrix bindings and table bindings, selects the first data cell (not first cell in header row for tables).
         *
         * @param id - The identifier of the object or location to go to.
         * @param goToType - The type of the location to go to.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the current view.
         */
        goToByIdAsync(id: string | number, goToType: GoToType, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Removes an event handler for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents | DocumentEvents}
         *
         * @param eventType - The event type. For document can be 'Document.SelectionChanged' or 'Document.ActiveViewChanged'.
         * @param options - Provides options to determine which event handler or handlers are removed.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, options?: RemoveHandlerOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Removes an event handler for the specified event type.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents | DocumentEvents}
         *
         * @param eventType - The event type. For document can be 'Document.SelectionChanged' or 'Document.ActiveViewChanged'.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Writes the specified data into the current selection.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion}, (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets | ImageCoercion 1.1} (when using `Office.CoercionType.Image`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection | Selection}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12 | ImageCoercion 1.2} (when using `Office.CoercionType.XmlSvg`)
         *
         * **Application-specific behaviors**
         *
         * The following application-specific actions apply when writing data to a selection.
         *
         * <table>
         *   <tr>
         *     <td>Word</td>
         *     <td>If there is no selection and the insertion point is at a valid location, the specified `data` is inserted at the insertion point</td>
         *     <td>If `data` is a string, the specified text is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is an array of arrays ("matrix") or a TableData object, a new Word table is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is HTML, the specified HTML is inserted. (**Important**: If any of the HTML you insert is invalid, Word won't raise an error. Word will insert as much of the HTML as it can and omits any invalid data).</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is Office Open XML, the specified XML is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is a base64 encoded image stream, the specified image is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>If there is a selection</td>
         *     <td>It will be replaced with the specified `data` following the same rules as above.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>Insert images</td>
         *     <td>Inserted images are placed inline. The imageLeft and imageTop parameters are ignored. The image aspect ratio is always locked. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         * </tr>
         *
         *   <tr>
         *     <td>Excel</td>
         *     <td>If a single cell is selected</td>
         *     <td>If `data` is a string, the specified text is inserted as the value of the current cell.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is an array of arrays ("matrix"), the specified set of rows and columns are inserted, if no other data in surrounding cells will be overwritten.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is a TableData object, a new Excel table with the specified set of rows and headers is inserted, if no other data in surrounding cells will be overwritten.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>If multiple cells are selected</td>
         *     <td>If the shape does not match the shape of `data`, an error is returned.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If the shape of the selection exactly matches the shape of `data`, the values of the selected cells are updated based on the values in `data`.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>Insert images</td>
         *     <td>Inserted images are floating. The position imageLeft and imageTop parameters are relative to currently selected cell(s). Negative imageLeft and imageTop values are allowed and possibly readjusted by Excel to position the image inside a worksheet. Image aspect ratio is locked unless both imageWidth and imageHeight parameters are provided. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>All other cases</td>
         *     <td>An error is returned.</td>
         *   </tr>
         *
         *   <tr>
         *     <td>Excel on the web</td>
         *     <td>In addition to the behaviors described for Excel above, these limits apply when writing data in Excel on the web</td>
         *     <td>The total number of cells you can write to a worksheet with the `data` parameter can't exceed 20,000 in a single call to this method.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>The number of formatting groups passed to the `cellFormat` parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells.</td>
         *   </tr>
         *
         *   <tr>
         *     <td>PowerPoint</td>
         *     <td>Insert image</td>
         *     <td>Inserted images are floating. The position imageLeft and imageTop parameters are optional but if provided, both should be present. If a single value is provided, it will be ignored. Negative imageLeft and imageTop values are allowed and can position an image outside of a slide. If no optional parameter is given and slide has a placeholder, the image will replace the placeholder in the slide. Image aspect ratio will be locked unless both imageWidth and imageHeight parameters are provided. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         *   </tr>
         * </table>
         *
         * **Hosts**
         *
         * The possible values for the {@link Office.CoercionType} parameter vary by the host.
         *
         * <table>
         *   <tr>
         *     <th>Host</th>
         *     <th>Supported coercionType</th>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, Project, and Word</td>
         *     <td>`Office.CoercionType.Text` (string)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Matrix` (array of arrays)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Table` (TableData object)</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Html`</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Ooxml` (Office Open XML)</td>
         *   </tr>
         *   <tr>
         *     <td>PowerPoint on the web and Windows</td>
         *     <td>`Office.CoercionType.SlideRange`</td>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, and Word</td>
         *     <td>`Office.CoercionType.XmlSvg`</td>
         *   </tr>
         * </table>
         *
         *
         * @param data - The data to be set. Either a string or  {@link Office.CoercionType} value, 2d array or TableData object.
         *
         * If the value passed for `data` is:
         *
         * - A string: Plain text or anything that can be coerced to a string will be inserted.
         * In Excel, you can also specify data as a valid formula to add that formula to the selected cell. For example, setting data to "=SUM(A1:A5)"
         * will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added
         * formula (or any pre-existing formula) from the bound cell. If you call the Document.getSelectedDataAsync method on the selected cell to
         * read its data, the method can return only the data displayed in the cell (the formula's result).
         *
         * - An array of arrays ("matrix"): Tabular data without headers will be inserted. For example, to write data to three rows in two columns,
         * you can pass an array like this: [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]. To write a single column of three rows, pass an
         * array like this: [["R1C1"], ["R2C1"], ["R3C1"]]
         *
         * In Excel, you can also specify data as an array of arrays that contains valid formulas to add them to the selected cells. For example if no
         * other data will be overwritten, setting data to [["=SUM(A1:A5)","=AVERAGE(A1:A5)"]] will add those two formulas to the selection. Just as
         * when setting a formula on a single cell as "text", you can't read the added formulas (or any pre-existing formulas) after they have been
         * set - you can only read the formulas' results.
         *
         * - A TableData object: A table with headers will be inserted.
         * In Excel, if you specify formulas in the TableData object you pass for the data parameter, you might not get the results you expect due to
         * the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to
         * write `data` that contains formulas to a selected table, try specifying the data as an array of arrays (instead of a TableData object), and
         * specify the coercionType as Microsoft.Office.Matrix or "matrix". However, this technique will block the "calculated columns" feature only
         * when one of the following conditions is met: (1) you are writing to all the cells of the column, or (2) there are already at least two
         * different formulas in the column.
         *
         * @param options - Provides options for how to insert data to the selection.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The AsyncResult.value property always returns undefined because there is no object or data to retrieve.
         */
        setSelectedDataAsync(data: string | TableData | any[][], options?: SetSelectedDataOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Writes the specified data into the current selection.
         *
         * @remarks
         *
         * **Requirement sets**:
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion | HtmlCoercion}, (when using `Office.CoercionType.Html`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets | ImageCoercion} (when using `Office.CoercionType.Image`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion | MatrixCoercion} (when using `Office.CoercionType.Matrix`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion | OoxmlCoercion} (when using `Office.CoercionType.Ooxml`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection | Selection}
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion | TableCoercion} (when using `Office.CoercionType.Table`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion | TextCoercion} (when using `Office.CoercionType.Text`)
         *
         * - {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12 | ImageCoercion 1.2} (when using `Office.CoercionType.XmlSvg`)
         *
         * **Application-specific behaviors**
         *
         * The following application-specific actions apply when writing data to a selection.
         *
         * <table>
         *   <tr>
         *     <td>Word</td>
         *     <td>If there is no selection and the insertion point is at a valid location, the specified `data` is inserted at the insertion point</td>
         *     <td>If `data` is a string, the specified text is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is an array of arrays ("matrix") or a TableData object, a new Word table is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is HTML, the specified HTML is inserted. (**Important**: If any of the HTML you insert is invalid, Word won't raise an error. Word will insert as much of the HTML as it can and omits any invalid data).</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is Office Open XML, the specified XML is inserted.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is a base64 encoded image stream, the specified image is inserted.</td></tr></td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>If there is a selection</td>
         *     <td>It will be replaced with the specified `data` following the same rules as above.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>Insert images</td>
         *     <td>Inserted images are placed inline. The imageLeft and imageTop parameters are ignored. The image aspect ratio is always locked. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         *   </tr>
         *
         *   <tr>
         *     <td>Excel</td>
         *     <td>If a single cell is selected</td>
         *     <td>If `data` is a string, the specified text is inserted as the value of the current cell.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is an array of arrays ("matrix"), the specified set of rows and columns are inserted, if no other data in surrounding cells will be overwritten.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If `data` is a TableData object, a new Excel table with the specified set of rows and headers is inserted, if no other data in surrounding cells will be overwritten.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>If multiple cells are selected</td><td>If the shape does not match the shape of `data`, an error is returned.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>If the shape of the selection exactly matches the shape of `data`, the values of the selected cells are updated based on the values in `data`.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>Insert images</td>
         *     <td>Inserted images are floating. The position imageLeft and imageTop parameters are relative to currently selected cell(s). Negative imageLeft and imageTop values are allowed and possibly readjusted by Excel to position the image inside a worksheet. Image aspect ratio is locked unless both imageWidth and imageHeight parameters are provided. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td>All other cases</td>
         *     <td>An error is returned.</td>
         *   </tr>
         *
         *   <tr>
         *     <td>Excel on the web</td>
         *     <td>In addition to the behaviors described for Excel above, these limits apply when writing data in Excel on the web</td>
         *     <td>The total number of cells you can write to a worksheet with the `data` parameter can't exceed 20,000 in a single call to this method.</td>
         *   </tr>
         *   <tr>
         *     <td></td>
         *     <td></td>
         *     <td>The number of formatting groups passed to the `cellFormat` parameter can't exceed 100. A single formatting group consists of a set of formatting applied to a specified range of cells.</td>
         *   </tr>
         *
         *   <tr>
         *     <td>PowerPoint</td>
         *     <td>Insert image</td>
         *     <td>Inserted images are floating. The position imageLeft and imageTop parameters are optional but if provided, both should be present. If a single value is provided, it will be ignored. Negative imageLeft and imageTop values are allowed and can position an image outside of a slide. If no optional parameter is given and slide has a placeholder, the image will replace the placeholder in the slide. Image aspect ratio will be locked unless both imageWidth and imageHeight parameters are provided. If only one of the imageWidth and imageHeight parameter is given, the other value will be automatically scaled to keep the original aspect ratio.</td>
         *   </tr>
         * </table>
         *
         * **Hosts**
         *
         * The possible values for the {@link Office.CoercionType} parameter vary by the host.
         *
         * <table>
         *   <tr>
         *     <th>Host</th>
         *     <th>Supported coercionType</th>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, Project, and Word</td>
         *     <td>`Office.CoercionType.Text` (string)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Matrix` (array of arrays)</td>
         *   </tr>
         *   <tr>
         *     <td>Excel and Word</td>
         *     <td>`Office.CoercionType.Table` (TableData object)</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Html`</td>
         *   </tr>
         *   <tr>
         *     <td>Word</td>
         *     <td>`Office.CoercionType.Ooxml` (Office Open XML)</td>
         *   </tr>
         *   <tr>
         *     <td>PowerPoint on the web and Windows</td>
         *     <td>`Office.CoercionType.SlideRange`</td>
         *   </tr>
         *   <tr>
         *     <td>Excel, PowerPoint, and Word</td>
         *     <td>`Office.CoercionType.XmlSvg`</td>
         *   </tr>
         * </table>
         *
         *
         * @param data - The data to be set. Either a string or  {@link Office.CoercionType} value, 2d array or TableData object.
         *
         * If the value passed for `data` is:
         *
         * - A string: Plain text or anything that can be coerced to a string will be inserted.
         * In Excel, you can also specify data as a valid formula to add that formula to the selected cell. For example, setting data to "=SUM(A1:A5)"
         * will total the values in the specified range. However, when you set a formula on the bound cell, after doing so, you can't read the added
         * formula (or any pre-existing formula) from the bound cell. If you call the Document.getSelectedDataAsync method on the selected cell to
         * read its data, the method can return only the data displayed in the cell (the formula's result).
         *
         * - An array of arrays ("matrix"): Tabular data without headers will be inserted. For example, to write data to three rows in two columns,
         * you can pass an array like this: [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]. To write a single column of three rows, pass an
         * array like this: [["R1C1"], ["R2C1"], ["R3C1"]]
         *
         * In Excel, you can also specify data as an array of arrays that contains valid formulas to add them to the selected cells. For example if no
         * other data will be overwritten, setting data to [["=SUM(A1:A5)","=AVERAGE(A1:A5)"]] will add those two formulas to the selection. Just as
         * when setting a formula on a single cell as "text", you can't read the added formulas (or any pre-existing formulas) after they have been
         * set - you can only read the formulas' results.
         *
         * - A TableData object: A table with headers will be inserted.
         * In Excel, if you specify formulas in the TableData object you pass for the data parameter, you might not get the results you expect due to
         * the "calculated columns" feature of Excel, which automatically duplicates formulas within a column. To work around this when you want to
         * write `data` that contains formulas to a selected table, try specifying the data as an array of arrays (instead of a TableData object), and
         * specify the coercionType as Microsoft.Office.Matrix or "matrix". However, this technique will block the "calculated columns" feature only
         * when one of the following conditions is met: (1) you are writing to all the cells of the column, or (2) there are already at least two
         * different formulas in the column.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The AsyncResult.value property always returns undefined because there is no object or data to retrieve.
         */
        setSelectedDataAsync(data: string | TableData | any[][], callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Project documents only. Get Project field (Ex. ProjectWebAccessURL).
         * @param fieldId - Project level fields.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the `fieldValue` property, which represents the value of the specified field.
         */
        getProjectFieldAsync(fieldId: number, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get Project field (Ex. ProjectWebAccessURL).
         * @param fieldId - Project level fields.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the `fieldValue` property, which represents the value of the specified field.
         *
         */
        getProjectFieldAsync(fieldId: number, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get resource field for provided resource Id. (Ex.ResourceName)
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getResourceFieldAsync(resourceId: string, fieldId: number, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get resource field for provided resource Id. (Ex.ResourceName)
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getResourceFieldAsync(resourceId: string, fieldId: number, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the current selected Resource's Id.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getSelectedResourceAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the current selected Resource's Id.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getSelectedResourceAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the current selected Task's Id.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getSelectedTaskAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the current selected Task's Id.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getSelectedTaskAsync(callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the current selected View Type (Ex. Gantt) and View Name.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `viewName` - The name of the view, as a ProjectViewTypes constant.
         *                  `viewType` - The type of view, as the integer value of a ProjectViewTypes constant.
         *
         */
        getSelectedViewAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the current selected View Type (Ex. Gantt) and View Name.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `viewName` - The name of the view, as a ProjectViewTypes constant.
         *                  `viewType` - The type of view, as the integer value of a ProjectViewTypes constant.
         *
         */
        getSelectedViewAsync(callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the Task Name, WSS Task Id, and ResourceNames for given taskId.
         * @param taskId - Either a string or value of the Task Id.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `taskName` - The name of the task.
         *                  `wssTaskId` - The ID of the task in the synchronized SharePoint task list. If the project is not synchronized with a SharePoint task list, the value is 0.
         *                  `resourceNames` - The comma-separated list of the names of resources that are assigned to the task.
         *
         */
        getTaskAsync(taskId: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the Task Name, WSS Task Id, and ResourceNames for given taskId.
         * @param taskId - Either a string or value of the Task Id.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `taskName` - The name of the task.
         *                  `wssTaskId` - The ID of the task in the synchronized SharePoint task list. If the project is not synchronized with a SharePoint task list, the value is 0.
         *                  `resourceNames` - The comma-separated list of the names of resources that are assigned to the task.
         *
         */
        getTaskAsync(taskId: string, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get task field for provided task Id. (Ex. StartDate).
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the `fieldValue` property, which represents the value of the specified field.
         *
         */
        getTaskFieldAsync(taskId: string, fieldId: number, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get task field for provided task Id. (Ex. StartDate).
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the `fieldValue` property, which represents the value of the specified field.
         *
         */
        getTaskFieldAsync(taskId: string, fieldId: number, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the WSS Url and list name for the Tasks List, the MPP is synced too.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `listName` - the name of the synchronized SharePoint task list.
         *                  `serverUrl` - the URL of the synchronized SharePoint task list.
         *
         */
        getWSSUrlAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the WSS Url and list name for the Tasks List, the MPP is synced too.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result contains the following properties:
         *                  `listName` - the name of the synchronized SharePoint task list.
         *                  `serverUrl` - the URL of the synchronized SharePoint task list.
         *
         */
        getWSSUrlAsync(callback?: (result: AsyncResult<any>) => void): void;
        /**
         * Project documents only. Get the maximum index of the collection of resources in the current project.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the highest index number in the current project's resource collection.
         *
         */
        getMaxResourceIndexAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<number>) => void): void;
        /**
         * Project documents only. Get the maximum index of the collection of resources in the current project.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the highest index number in the current project's resource collection.
         *
         */
        getMaxResourceIndexAsync(callback?: (result: AsyncResult<number>) => void): void;
        /**
         * Project documents only. Get the maximum index of the collection of tasks in the current project.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the highest index number in the current project's task collection.
         *
         */
        getMaxTaskIndexAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<number>) => void): void;
        /**
         * Project documents only. Get the maximum index of the collection of tasks in the current project.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the highest index number in the current project's task collection.
         *
         */
        getMaxTaskIndexAsync(callback?: (result: AsyncResult<number>) => void): void;
        /**
         * Project documents only. Get the GUID of the resource that has the specified index in the resource collection.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param resourceIndex - The index of the resource in the collection of resources for the project.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getResourceByIndexAsync(resourceIndex: number, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the GUID of the resource that has the specified index in the resource collection.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param resourceIndex - The index of the resource in the collection of resources for the project.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the resource as a string.
         *
         */
        getResourceByIndexAsync(resourceIndex: number, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the GUID of the task that has the specified index in the task collection.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param taskIndex - The index of the task in the collection of tasks for the project.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the task as a string.
         *
         */
        getTaskByIndexAsync(taskIndex: number, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Get the GUID of the task that has the specified index in the task collection.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param taskIndex - The index of the task in the collection of tasks for the project.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the GUID of the task as a string.
         *
         */
        getTaskByIndexAsync(taskIndex: number, callback?: (result: AsyncResult<string>) => void): void;
        /**
         * Project documents only. Set resource field for specified resource Id.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param fieldValue - Value of the target field.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setResourceFieldAsync(resourceId: string, fieldId: number, fieldValue: string | number | boolean | object, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Project documents only. Set resource field for specified resource Id.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param fieldValue - Value of the target field.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setResourceFieldAsync(resourceId: string, fieldId: number, fieldValue: string | number | boolean | object, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Project documents only. Set task field for specified task Id.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param fieldValue - Value of the target field.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setTaskFieldAsync(taskId: string, fieldId: number, fieldValue: string | number | boolean | object, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Project documents only. Set task field for specified task Id.
         *
         * **Important**: This API works only in Project 2016 on Windows desktop.
         *
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param fieldValue - Value of the target field.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setTaskFieldAsync(taskId: string, fieldId: number, fieldValue: string | number | boolean | object, callback?: (result: AsyncResult<void>) => void): void;
    }
    /**
     * Provides information about the document that raised the SelectionChanged event.
     */
    export interface DocumentSelectionChangedEventArgs {
        /**
         * Gets an {@link Office.Document} object that represents the document that raised the SelectionChanged event.
         */
        document: Document;
        /**
         * Get an {@link Office.EventType} enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }
    /**
     * Represents the document file associated with an Office Add-in.
     *
     * @remarks
     *
     * Access the File object with the AsyncResult.value property in the callback function passed to the Document.getFileAsync method.
     *
     */
    export interface File {
        /**
         * Gets the document file size in bytes.
         */
        size: number;
        /**
         * Gets the number of slices into which the file is divided.
         */
        sliceCount: number;
        /**
         * Closes the document file.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file | File}
         *
         * No more than two documents are allowed to be in memory; otherwise the Document.getFileAsync operation will fail. Use the File.closeAsync
         * method to close the file when you are finished working with it.
         *
         * In the callback function passed to the closeAsync method, you can use the properties of the AsyncResult object to return the following
         * information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        closeAsync(callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Returns the specified slice.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file | File}
         *
         * In the callback function passed to the getSliceAsync method, you can use the properties of the AsyncResult object to return the following
         * information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Access the Slice object.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param sliceIndex - Specifies the zero-based index of the slice to be retrieved. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is the {@link Office.Slice} object.
         */
        getSliceAsync(sliceIndex: number, callback?: (result: AsyncResult<Office.Slice>) => void): void;
    }
    export interface FileProperties {
        /**
         * File's URL
         */
        url: string
    }
    /**
     * Represents a binding in two dimensions of rows and columns.
     *
     * @remarks
     *
     * The MatrixBinding object inherits the id property, type property, getDataAsync method, and setDataAsync method from the Binding object.
     */
    export interface MatrixBinding extends Binding {
        /**
        * Gets the number of columns in the matrix data structure, as an integer value.
        */
        columnCount: number;
        /**
        * Gets the number of rows in the matrix data structure, as an integer value.
        */
        rowCount: number;
    }
    /**
     * Represents custom settings for a task pane or content add-in that are stored in the host document as name/value pairs.
     *
     * @remarks
     *
     * **Hosts**: Excel, PowerPoint, Word
     *
     * The settings created by using the methods of the Settings object are saved per add-in and per document.
     * That is, they are available only to the add-in that created them, and only from the document in which they are saved.
     *
     * The name of a setting is a string, while the value can be a string, number, boolean, null, object, or array.
     *
     * The Settings object is automatically loaded as part of the Document object, and is available by calling the settings property of that object
     * when the add-in is activated.
     *
     * The developer is responsible for calling the saveAsync method after adding or deleting settings to save the settings in the document.
     */
    export interface Settings {
        /**
         * Adds an event handler for the settingsChanged event.
         *
         * **Important**: Your add-in's code can register a handler for the settingsChanged event when the add-in is running with any Excel client, but
         * the event will fire only when the add-in is loaded with a spreadsheet that is opened in Excel on the web, and more than one user is editing the
         * spreadsheet (coauthoring). Therefore, effectively the settingsChanged event is supported only in Excel on the web in coauthoring scenarios.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - Specifies the type of event to add. Required.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.SettingsChangedEventArgs}. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no data or object to retrieve when adding an event handler.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for the settingsChanged event.
         *
         * **Important**: Your add-in's code can register a handler for the settingsChanged event when the add-in is running with any Excel client, but
         * the event will fire only when the add-in is loaded with a spreadsheet that is opened in Excel on the web, and more than one user is editing the
         * spreadsheet (coauthoring). Therefore, effectively the settingsChanged event is supported only in Excel on the web in coauthoring scenarios.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * You can add multiple event handlers for the specified eventType as long as the name of each event handler function is unique.
         *
         * @param eventType - Specifies the type of event to add. Required.
         * @param handler - The event handler function to add, whose only parameter is of type {@link Office.SettingsChangedEventArgs}. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no data or object to retrieve when adding an event handler.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         */
        addHandlerAsync(eventType: Office.EventType, handler: any, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Retrieves the specified setting.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings | Settings}
         *
         * @param settingName - The case-sensitive name of the setting to retrieve.
         * @returns An object that has property names mapped to JSON serialized values.
         */
        get(name: string): any;
        /**
         * Reads all settings persisted in the document and refreshes the content or task pane add-in's copy of those settings held in memory.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * This method is useful in Excel, Word, and PowerPoint coauthoring scenarios when multiple instances of the same add-in are working against
         * the same document. Because each add-in is working against an in-memory copy of the settings loaded from the document at the time the user
         * opened it, the settings values used by each user can get out of sync. This can happen whenever an instance of the add-in calls the
         * Settings.saveAsync method to persist all of that user's settings to the document. Calling the refreshAsync method from the event handler
         * for the settingsChanged event of the add-in will refresh the settings values for all users.
         *
         * In the callback function passed to the refreshAsync method, you can use the properties of the AsyncResult object to return the following
         * information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Access a Settings object with the refreshed values.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an {@link Office.Settings} object with the refreshed values.
         */
        refreshAsync(callback?: (result: AsyncResult<Office.Settings>) => void): void;
        /**
         * Removes the specified setting.
         *
         * **Important**: Be aware that the Settings.remove method affects only the in-memory copy of the settings property bag. To persist the removal of
         * the specified setting in the document, at some point after calling the Settings.remove method and before the add-in is closed, you must
         * call the Settings.saveAsync method.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings | Settings}
         *
         * null is a valid value for a setting. Therefore, assigning null to the setting will not remove it from the settings property bag.
         *
         * @param settingName - The case-sensitive name of the setting to remove.
         */
        remove(name: string): void;
        /**
         * Removes an event handler for the settingsChanged event.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * If the optional handler parameter is omitted when calling the removeHandlerAsync method, all event handlers for the specified eventType
         * will be removed.
         *
         * When the function you passed to the callback parameter executes, it receives an AsyncResult object that you can access from the callback
         * function's only parameter.
         *
         * In the callback function passed to the removeHandlerAsync method, you can use the properties of the AsyncResult object to return the
         * following information.
         *
         * @param eventType - Specifies the type of event to remove. Required.
         * @param options - Provides options to determine which event handler or handlers are removed.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, options?: RemoveHandlerOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Removes an event handler for the settingsChanged event.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * If the optional handler parameter is omitted when calling the removeHandlerAsync method, all event handlers for the specified eventType
         * will be removed.
         *
         * When the function you passed to the callback parameter executes, it receives an AsyncResult object that you can access from the callback
         * function's only parameter.
         *
         * In the callback function passed to the removeHandlerAsync method, you can use the properties of the AsyncResult object to return the
         * following information.
         *
         * @param eventType - Specifies the type of event to remove. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        removeHandlerAsync(eventType: Office.EventType, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Persists the in-memory copy of the settings property bag in the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings | Settings}
         *
         * Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the
         * set and get methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are
         * available the next time the add-in is used, use the saveAsync method.
         *
         * **Note**: The saveAsync method persists the in-memory settings property bag into the document file. However, the changes to the document file
         * itself are saved only when the user (or AutoRecover setting) saves the document to the file system. The refreshAsync method is only useful
         * in coauthoring scenarios when other instances of the same add-in might change the settings and those changes should be made available to
         * all instances.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param options - Provides options for saving settings.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        saveAsync(options?: SaveSettingsOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Persists the in-memory copy of the settings property bag in the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings | Settings}
         *
         * Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the
         * set and get methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are
         * available the next time the add-in is used, use the saveAsync method.
         *
         * **Note**: The saveAsync method persists the in-memory settings property bag into the document file. However, the changes to the document file
         * itself are saved only when the user (or AutoRecover setting) saves the document to the file system. The refreshAsync method is only useful
         * in coauthoring scenarios when other instances of the same add-in might change the settings and those changes should be made available to
         * all instances.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no object or data to retrieve.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        saveAsync(callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Sets or creates the specified setting.
         *
         * **Important**: Be aware that the Settings.set method affects only the in-memory copy of the settings property bag.
         * To make sure that additions or changes to settings will be available to your add-in the next time the document is opened, at some point
         * after calling the Settings.set method and before the add-in is closed, you must call the Settings.saveAsync method to persist settings in
         * the document.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings | Settings}
         *
         * The set method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name
         * in the in-memory copy of the settings property bag. After you call the Settings.saveAsync method, the value is stored in the document as
         * the serialized JSON representation of its data type.
         *
         * @param settingName - The case-sensitive name of the setting to set or create.
         * @param value - Specifies the value to be stored.
         */
        set(name: string, value: any): void;
    }
    /**
     * Provides information about the settings that raised the settingsChanged event.
     *
     * To add an event handler for the settingsChanged event, use the addHandlerAsync method of the
     * {@link Office.Settings} object.
     *
     * The settingsChanged event fires only when your add-in's script calls the Settings.saveAsync method to persist
     * the in-memory copy of the settings into the document file. The settingsChanged event is not triggered when the
     * Settings.set or Settings.remove methods are called.
     *
     * The settingsChanged event was designed to let you to handle potential conflicts when two or more users are
     * attempting to save settings at the same time when your add-in is used in a shared (coauthored) document.
     *
     * **Important**: Your add-in's code can register a handler for the settingsChanged event when the add-in
     * is running with any Excel client, but the event will fire only when the add-in is loaded with a spreadsheet
     * that is opened in Excel on the web, and more than one user is editing the spreadsheet (coauthoring).
     * Therefore, effectively the settingsChanged event is supported only in Excel on the web in coauthoring scenarios.
     */
    export interface SettingsChangedEventArgs {
        /**
         * Gets an {@link Office.Settings} object that represents the settings that raised the settingsChanged event.
         */
        settings: Settings;
        /**
         * Get an {@link Office.EventType} enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }
    /**
     * Provides information about the message from the parent page that raised the `DialogParentMessageReceived` event.
     *
     * To add an event handler for the `DialogParentMessageReceived` event, use the `addHandlerAsync` method of the
     * {@link Office.UI} object.
     *
     */
    export interface DialogParentMessageReceivedEventArgs {
        /**
         * Gets the content of the message sent from the parent page, which can be any string or stringified data.
         */
        message: string;
        /**
         * Get an {@link Office.EventType} enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }
    /**
     * Represents a slice of a document file. The Slice object is accessed with the `File.getSliceAsync` method.
     */
    export interface Slice {
        /**
         * Gets the raw data of the file slice in `Office.FileType.Text` or `Office.FileType.Compressed` format as specified
         * by the `fileType` parameter of the call to the `Document.getFileAsync` method.
         *
         * @remarks
         *
         * Files in the "compressed" format will return a byte array that can be transformed to a base64-encoded string if required.
         */
        data: any;
        /**
         * Gets the zero-based index of the file slice.
         */
        index: number;
        /**
         * Gets the size of the slice in bytes.
         */
        size: number;
    }
    /**
     * Represents a binding in two dimensions of rows and columns, optionally with headers.
     *
     * @remarks
     *
     * The TableBinding object inherits the `id` property, `type` property, `getDataAsync` method, and `setDataAsync` method from the
     * {@link Office.Binding} object.
     *
     * For Excel, note that after you establish a table binding, each new row a user adds to the table is automatically included in the binding and
     * rowCount increases.
     */
    export interface TableBinding extends Binding {
        /**
        * Gets the number of columns in the TableBinding, as an integer value.
        */
        columnCount: number;
        /**
        * True, if the table has headers; otherwise false.
        */
        hasHeaders: boolean;
         /**
        * Gets the number of rows in the TableBinding, as an integer value.
        *
        * @remarks
        *
        * When you insert an empty table by selecting a single row in Excel 2013 and Excel on the web (using Table on the Insert tab), both Office host
        * applications create a single row of headers followed by a single blank row. However, if your add-in's script creates a binding for this
        * newly inserted table (for example, by using the {@link Office.Bindings}.addFromSelectionAsync method), and then checks the value of the
        * rowCount property, the value returned will differ depending whether the spreadsheet is open in Excel 2013 or Excel on the web.
        *
        * - In Excel on the desktop, rowCount will return 0 (the blank row following the headers is not counted).
        *
        * - In Excel on the web, rowCount will return 1 (the blank row following the headers is counted).
        *
        * You can work around this difference in your script by checking if rowCount == 1, and if so, then checking if the row contains all empty
        * strings.
        */
        rowCount: number;
        /**
         * Adds the specified data to the table as additional columns.
         *
         * @remarks
         *
         * To add one or more columns specifying the values of the data and headers, pass a TableData object as the data parameter. To add one or more
         * columns specifying only the data, pass an array of arrays ("matrix") as the data parameter.
         *
         * The success or failure of an addColumnsAsync operation is atomic. That is, the entire add columns operation must succeed, or it will be
         * completely rolled back (and the AsyncResult.status property returned to the callback will report failure):
         *
         *  - Each row in the array you pass as the data argument must have the same number of rows as the table being updated. If not, the entire
         * operation will fail.
         *
         *  - Each row and cell in the array must successfully add that row or cell to the table in the newly added column(s). If any row or cell
         * fails to be set for any reason, the entire operation will fail.
         *
         *  - If you pass a TableData object as the data argument, the number of header rows must match that of the table being updated.
         *
         * Additional remark for Excel on the web: The total number of cells in the TableData object passed to the data parameter can't exceed 20,000 in
         * a single call to this method.
         *
         * @param tableData - An array of arrays ("matrix") or a TableData object that contains one or more columns of data to add to the table. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addColumnsAsync(tableData: TableData | any[][], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds the specified data to the table as additional columns.
         *
         * @remarks
         *
         * To add one or more columns specifying the values of the data and headers, pass a TableData object as the data parameter. To add one or more
         * columns specifying only the data, pass an array of arrays ("matrix") as the data parameter.
         *
         * The success or failure of an addColumnsAsync operation is atomic. That is, the entire add columns operation must succeed, or it will be
         * completely rolled back (and the AsyncResult.status property returned to the callback will report failure):
         *
         *  - Each row in the array you pass as the data argument must have the same number of rows as the table being updated. If not, the entire
         * operation will fail.
         *
         *  - Each row and cell in the array must successfully add that row or cell to the table in the newly added column(s). If any row or cell
         * fails to be set for any reason, the entire operation will fail.
         *
         *  - If you pass a TableData object as the data argument, the number of header rows must match that of the table being updated.
         *
         * Additional remark for Excel on the web: The total number of cells in the TableData object passed to the data parameter can't exceed 20,000 in
         * a single call to this method.
         *
         * @param tableData - An array of arrays ("matrix") or a TableData object that contains one or more columns of data to add to the table. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addColumnsAsync(tableData: TableData | any[][], callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds the specified data to the table as additional rows.
         *
         * @remarks
         *
         * The success or failure of an addRowsAsync operation is atomic. That is, the entire add columns operation must succeed, or it will be
         * completely rolled back (and the AsyncResult.status property returned to the callback will report failure):
         *
         *  - Each row in the array you pass as the data argument must have the same number of columns as the table being updated. If not, the entire
         * operation will fail.
         *
         *  - Each column and cell in the array must successfully add that column or cell to the table in the newly added rows(s). If any column or
         * cell fails to be set for any reason, the entire operation will fail.
         *
         *  - If you pass a TableData object as the data argument, the number of header rows must match that of the table being updated.
         *
         * Additional remark for Excel on the web: The total number of cells in the TableData object passed to the data parameter can't exceed 20,000 in
         * a single call to this method.
         *
         * @param rows - An array of arrays ("matrix") or a TableData object that contains one or more rows of data to add to the table. Required.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addRowsAsync(rows: TableData | any[][], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Adds the specified data to the table as additional rows.
         *
         * @remarks
         *
         * The success or failure of an addRowsAsync operation is atomic. That is, the entire add columns operation must succeed, or it will be
         * completely rolled back (and the AsyncResult.status property returned to the callback will report failure):
         *
         *  - Each row in the array you pass as the data argument must have the same number of columns as the table being updated. If not, the entire
         * operation will fail.
         *
         *  - Each column and cell in the array must successfully add that column or cell to the table in the newly added rows(s). If any column or
         * cell fails to be set for any reason, the entire operation will fail.
         *
         *  - If you pass a TableData object as the data argument, the number of header rows must match that of the table being updated.
         *
         * Additional remark for Excel on the web: The total number of cells in the TableData object passed to the data parameter can't exceed 20,000 in
         * a single call to this method.
         *
         * @param rows - An array of arrays ("matrix") or a TableData object that contains one or more rows of data to add to the table. Required.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        addRowsAsync(rows: TableData | any[][], callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Deletes all non-header rows and their values in the table, shifting appropriately for the host application.
         *
         * @remarks
         *
         * In Excel, if the table has no header row, this method will delete the table itself.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        deleteAllDataValuesAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Deletes all non-header rows and their values in the table, shifting appropriately for the host application.
         *
         * @remarks
         *
         * In Excel, if the table has no header row, this method will delete the table itself.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        deleteAllDataValuesAsync(callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Clears formatting on the bound table.
         *
         * @remarks
         *
         * See {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables#format-a-table | Format tables in add-ins for Excel} for more information.
         *
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        clearFormatsAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Clears formatting on the bound table.
         *
         * @remarks
         *
         * See {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables#format-a-table | Format tables in add-ins for Excel} for more information.
         *
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        clearFormatsAsync(callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Gets the formatting on specified items in the table.
         *
         * @remarks
         *
         * **Returned format structure**
         *
         * Each JavaScript object in the return value array has this form: `{cells:{ cell_range }, format:{ format_definition }}`
         *
         * The `cells:` property specifies the range you want format using one of the following values:
         *
         * **Supported ranges in cells property**
         *
         * <table>
         *   <tr>
         *     <th>cells range settings</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>`{row: n}`</td>
         *     <td>Specifies the range that is the zero-based nth row of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{column: n}`</td>
         *     <td>Specifies the range that is the zero-based nth column of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{row: i, column: j}`</td>
         *     <td>Specifies the single cell that is the ith row and jth column of the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.All`</td>
         *     <td>Specifies the entire table, including column headers, data, and totals (if any).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Data`</td>
         *     <td>Specifies only the data in the table (no headers and totals).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Headers`</td>
         *     <td>Specifies only the header row.</td>
         *   </tr>
         * </table>
         *
         * The `format:` property specifies values that correspond to a subset of the settings available in the Format Cells dialog box in Excel
         * (Right-click \> Format Cells or Home \> Format \> Format Cells).
         *
         * @param cellReference - An object literal containing name-value pairs that specify the range of cells to get formatting from.
         * @param formats - An array specifying the format properties to get.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array containing one or more JavaScript objects specifying the formatting of their corresponding cells.
         */
        getFormatsAsync(cellReference?: any, formats?: any[], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult< ({ cells: any, format: any})[]>) => void): void;
        /**
         * Gets the formatting on specified items in the table.
         *
         * @remarks
         *
         * **Returned format structure**
         *
         * Each JavaScript object in the return value array has this form: `{cells:{ cell_range }, format:{ format_definition }}`
         *
         * The `cells:` property specifies the range you want format using one of the following values:
         *
         * **Supported ranges in cells property**
         *
         * <table>
         *   <tr>
         *     <th>cells range settings</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>`{row: n}`</td>
         *     <td>Specifies the range that is the zero-based nth row of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{column: n}`</td>
         *     <td>Specifies the range that is the zero-based nth column of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{row: i, column: j}`</td>
         *     <td>Specifies the single cell that is the ith row and jth column of the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.All`</td>
         *     <td>Specifies the entire table, including column headers, data, and totals (if any).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Data`</td>
         *     <td>Specifies only the data in the table (no headers and totals).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Headers`</td>
         *     <td>Specifies only the header row.</td>
         *   </tr>
         * </table>
         *
         * The `format:` property specifies values that correspond to a subset of the settings available in the Format Cells dialog box in Excel
         * (Right-click \> Format Cells or Home \> Format \> Format Cells).
         *
         * @param cellReference - An object literal containing name-value pairs that specify the range of cells to get formatting from.
         * @param formats - An array specifying the format properties to get.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *                  The `value` property of the result is an array containing one or more JavaScript objects specifying the formatting of their corresponding cells.
         */
        getFormatsAsync(cellReference?: any, formats?: any[], callback?: (result: AsyncResult< ({ cells: any, format: any})[]>) => void): void;
        /**
         * Sets formatting on specified items and data in the table.
         *
         * @remarks
         *
         * **Specifying the cellFormat parameter**
         *
         * Use the cellFormat parameter to set or change cell formatting values, such as width, height, font, background, alignment, and so on.
         * The value you pass as the cellFormat parameter is an array that contains a list of one or more JavaScript objects that specify which cells
         * to target (`cells:`) and the formats (`format:`) to apply to them.
         *
         * Each JavaScript object in the cellFormat array has this form: `{cells:{ cell_range }, format:{ format_definition }}`
         *
         * The `cells:` property specifies the range you want format using one of the following values:
         *
         * **Supported ranges in cells property**
         *
         * <table>
         *   <tr>
         *     <th>cells range settings</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>`{row: n}`</td>
         *     <td>Specifies the range that is the zero-based nth row of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{column: n}`</td>
         *     <td>Specifies the range that is the zero-based nth column of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{row: i, column: j}`</td>
         *     <td>Specifies the single cell that is the ith row and jth column of the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.All`</td>
         *     <td>Specifies the entire table, including column headers, data, and totals (if any).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Data`</td>
         *     <td>Specifies only the data in the table (no headers and totals).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Headers`</td>
         *     <td>Specifies only the header row.</td>
         *   </tr>
         * </table>
         *
         * The `format:` property specifies values that correspond to a subset of the settings available in the Format Cells dialog box in Excel
         * (Right-click \> Format Cells or Home \> Format \> Format Cells).
         *
         * You specify the value of the `format:` property as a list of one or more property name - value pairs in a JavaScript object literal. The
         * property name specifies the name of the formatting property to set, and value specifies the property value.
         * You can specify multiple values for a given format, such as both a font's color and size.
         *
         * Here's three `format:` property value examples:
         *
         * `//Set cells: font color to green and size to 15 points.`
         *
         * `format: {fontColor : "green", fontSize : 15}`
         *
         * `//Set cells: border to dotted blue.`
         *
         * `format: {borderStyle: "dotted", borderColor: "blue"}`
         *
         * `//Set cells: background to red and alignment to centered.`
         *
         * `format: {backgroundColor: "red", alignHorizontal: "center"}`
         *
         *
         * You can specify number formats by specifying the number formatting "code" string in the `numberFormat:` property.
         * The number format strings you can specify correspond to those you can set in Excel using the Custom category on the Number tab of the Format Cells dialog box.
         * This example shows how to format a number as a percentage with two decimal places:
         *
         * `format: {numberFormat:"0.00%"}`
         *
         * For more detail, see how to {@link https://support.office.com/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 | Create a custom number format}.
         *
         * To set formatting on tables when writing data, use the tableOptions and cellFormat optional parameters of the
         * `Document.setSelectedDataAsync` or `TableBinding.setDataAsync` methods.
         *
         * Setting formatting with the optional parameters of the `Document.setSelectedDataAsync` and `TableBinding.setDataAsync` methods only works
         * to set formatting when writing data the first time.
         * To make formatting changes after writing data, use the following methods:
         *
         *  - To update cell formatting, such as font color and style, use the `TableBinding.setFormatsAsync` method (this method).
         *
         *  - To update table options, such as banded rows and filter buttons, use the `TableBinding.setTableOptions` method.
         *
         *  - To clear formatting, use the `TableBinding.clearFormats` method.
         *
         * For more details and examples, see
         * {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables#format-a-table | How to format tables in add-ins for Excel}.
         *
         * @param cellFormat - An array that contains one or more JavaScript objects that specify which cells to target and the formatting to apply to them.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setFormatsAsync(cellFormat: any[], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Sets formatting on specified items and data in the table.
         *
         * @remarks
         *
         * **Specifying the cellFormat parameter**
         *
         * Use the cellFormat parameter to set or change cell formatting values, such as width, height, font, background, alignment, and so on.
         * The value you pass as the cellFormat parameter is an array that contains a list of one or more JavaScript objects that specify which cells
         * to target (`cells:`) and the formats (`format:`) to apply to them.
         *
         * Each JavaScript object in the cellFormat array has this form: `{cells:{ cell_range }, format:{ format_definition }}`
         *
         * The `cells:` property specifies the range you want format using one of the following values:
         *
         * **Supported ranges in cells property**
         *
         * <table>
         *   <tr>
         *     <th>cells range settings</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>`{row: n}`</td>
         *     <td>Specifies the range that is the zero-based nth row of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{column: n}`</td>
         *     <td>Specifies the range that is the zero-based nth column of data in the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`{row: i, column: j}`</td>
         *     <td>Specifies the single cell that is the ith row and jth column of the table.</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.All`</td>
         *     <td>Specifies the entire table, including column headers, data, and totals (if any).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Data`</td>
         *     <td>Specifies only the data in the table (no headers and totals).</td>
         *   </tr>
         *   <tr>
         *     <td>`Office.Table.Headers`</td>
         *     <td>Specifies only the header row.</td>
         *   </tr>
         * </table>
         *
         * The `format:` property specifies values that correspond to a subset of the settings available in the Format Cells dialog box in Excel
         * (Right-click \> Format Cells or Home \> Format \> Format Cells).
         *
         * You specify the value of the `format:` property as a list of one or more property name - value pairs in a JavaScript object literal. The
         * property name specifies the name of the formatting property to set, and value specifies the property value.
         * You can specify multiple values for a given format, such as both a font's color and size.
         *
         * Here's three `format:` property value examples:
         *
         * `//Set cells: font color to green and size to 15 points.`
         *
         * `format: {fontColor : "green", fontSize : 15}`
         *
         * `//Set cells: border to dotted blue.`
         *
         * `format: {borderStyle: "dotted", borderColor: "blue"}`
         *
         * `//Set cells: background to red and alignment to centered.`
         *
         * `format: {backgroundColor: "red", alignHorizontal: "center"}`
         *
         *
         * You can specify number formats by specifying the number formatting "code" string in the `numberFormat:` property.
         * The number format strings you can specify correspond to those you can set in Excel using the Custom category on the Number tab of the Format Cells dialog box.
         * This example shows how to format a number as a percentage with two decimal places:
         *
         * `format: {numberFormat:"0.00%"}`
         *
         * For more detail, see how to {@link https://support.office.com/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 | Create a custom number format}.
         *
         * To set formatting on tables when writing data, use the tableOptions and cellFormat optional parameters of the
         * `Document.setSelectedDataAsync` or `TableBinding.setDataAsync` methods.
         *
         * Setting formatting with the optional parameters of the `Document.setSelectedDataAsync` and `TableBinding.setDataAsync` methods only works
         * to set formatting when writing data the first time.
         * To make formatting changes after writing data, use the following methods:
         *
         *  - To update cell formatting, such as font color and style, use the `TableBinding.setFormatsAsync` method (this method).
         *
         *  - To update table options, such as banded rows and filter buttons, use the `TableBinding.setTableOptions` method.
         *
         *  - To clear formatting, use the `TableBinding.clearFormats` method.
         *
         * For more details and examples, see
         * {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables#format-a-table | How to format tables in add-ins for Excel}.
         *
         * @param cellFormat - An array that contains one or more JavaScript objects that specify which cells to target and the formatting to apply to them.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         */
        setFormatsAsync(cellFormat: any[], callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Updates table formatting options on the bound table.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * In the callback function passed to the goToByIdAsync method, you can use the properties of the AsyncResult object to return the following information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no data or object to retrieve when setting formats.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param tableOptions - An object literal containing a list of property name-value pairs that define the table options to apply.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setTableOptionsAsync(tableOptions: any, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult<void>) => void): void;
        /**
         * Updates table formatting options on the bound table.
         *
         * @remarks
         *
         * **Requirement set**: {@link https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set | Not in a set}
         *
         * In the callback function passed to the goToByIdAsync method, you can use the properties of the AsyncResult object to return the following information.
         *
         * <table>
         *   <tr>
         *     <th>Property</th>
         *     <th>Use to...</th>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.value</td>
         *     <td>Always returns undefined because there is no data or object to retrieve when setting formats.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.status</td>
         *     <td>Determine the success or failure of the operation.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.error</td>
         *     <td>Access an Error object that provides error information if the operation failed.</td>
         *   </tr>
         *   <tr>
         *     <td>AsyncResult.asyncContext</td>
         *     <td>A user-defined item of any type that is returned in the AsyncResult object without being altered.</td>
         *   </tr>
         * </table>
         *
         * @param tableOptions - An object literal containing a list of property name-value pairs that define the table options to apply.
         * @param callback - Optional. A function that is invoked when the callback returns, whose only parameter is of type {@link Office.AsyncResult}.
         *
         */
        setTableOptionsAsync(tableOptions: any, callback?: (result: AsyncResult<void>) => void): void;
    }
    /**
     * Represents the data in a table or an {@link Office.TableBinding}.
     */
    export class TableData {
        constructor(rows: any[][], headers: any[]);
        constructor();
        /**
         * Gets or sets the headers of the table.
         *
         * @remarks
         *
         * To specify headers, you must specify an array of arrays that corresponds to the structure of the table. For example, to specify headers
         * for a two-column table you would set the header property to [['header1', 'header2']].
         *
         * If you specify null for the headers property (or leaving the property empty when you construct a TableData object), the following results
         * occur when your code executes:
         *
         * - If you insert a new table, the default column headers for the table are created.
         *
         * - If you overwrite or update an existing table, the existing headers are not altered.
         */
        headers: any[];
        /**
         * Gets or sets the rows in the table. Returns an array of arrays that contains the data in the table.
         * Returns an empty array if there are no rows.
         *
         * @remarks
         *
         * To specify rows, you must specify an array of arrays that corresponds to the structure of the table. For example, to specify two rows of
         * string values in a two-column table you would set the rows property to [['a', 'b'], ['c', 'd']].
         *
         * If you specify null for the rows property (or leave the property empty when you construct a TableData object), the following results occur
         * when your code executes:
         *
         * - If you insert a new table, a blank row will be inserted.
         *
         * - If you overwrite or update an existing table, the existing rows are not altered.
         */
        rows: any[][];
    }
    /**
     * Specifies enumerated values for the `cells` property in the cellFormat parameter of
     * {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables#format-a-table | table formatting methods}.
     */
    enum Table {
        /**
         * The entire table, including column headers, data, and totals (if any).
         */
        All,
        /**
         * Only the data (no headers and totals).
         */
        Data,
        /**
         * Only the header row.
         */
        Headers
    }
    /**
     * Represents a bound text selection in the document.
     *
     * The TextBinding object inherits the id property, type property, getDataAsync method, and setDataAsync method from the {@link Office.Binding}
     * object. It does not implement any additional properties or methods of its own.
     */
    export interface TextBinding extends Binding { }
    /**
     * Specifies the project fields that are available as a parameter for the {@link Office.Document | Document}.getProjectFieldAsync method.
     *
     * @remarks
     *
     * A ProjectProjectFields constant can be used as a parameter of the {@link Office.Document | Document}.getProjectFieldAsync method.
     */
    enum ProjectProjectFields {
        /**
         * The number of digits after the decimal for the currency.
         */
        CurrencyDigits,
        /**
         * The currency symbol.
         */
        CurrencySymbol,
        /**
         * The placement of the currency symbol: Not specified = -1; Before the value with no space ($0) = 0; After the value with no space (0$) = 1;
         * Before the value with a space ($ 0) = 2; After the value with a space (0 $) = 3.
         */
        CurrencySymbolPosition,
        DurationUnits,
        /**
         * The GUID of the project.
         */
        GUID,
        /**
         * The project finish date.
         */
        Finish,
        /**
         * The project start date.
         */
        Start,
        /**
         * Specifies whether the project is read-only.
         */
        ReadOnly,
        /**
         * The project version.
         */
        VERSION,
        /**
         * The work units of the project, such as days or hours.
         */
        WorkUnits,
        /**
         * The Project Web App URL, for projects that are stored in Project Server.
         */
        ProjectServerUrl,
        /**
         * The SharePoint URL, for projects that are synchronized with a SharePoint list.
         */
        WSSUrl,
        /**
         * The name of the SharePoint list, for projects that are synchronized with a tasks list.
         */
        WSSList
    }
    /**
     * Specifies the resource fields that are available as a parameter for the {@link Office.Document | Document}.getResourceFieldAsync method.
     *
     * @remarks
     *
     * A ProjectResourceFields constant can be used as a parameter of the {@link Office.Document | Document}.getResourceFieldAsync method.
     *
     * For more information about working with fields in Project, see
     * {@link https://support.office.com/article/Available-fields-reference-615a4563-1cc3-40f4-b66f-1b17e793a460 | Available fields} reference. In
     * Project Help, search for Available fields.
     */
    enum ProjectResourceFields {
        /**
         * The accrual method that defines how a task accrues the cost of the resource: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Accrual,
        /**
         * The calculated actual cost of the resource for assignments in the project.
         */
        ActualCost,
        /**
         * The actual overtime cost for a resource.
         */
        ActualOvertimeCost,
        /**
         * The actual overtime work for a resource, in minutes.
         */
        ActualOvertimeWork,
        /**
         * The actual overtime work for the resource that has been protected (made read-only).
         */
        ActualOvertimeWorkProtected,
        /**
         * The actual work that the resource has done on assignments in the project.
         */
        ActualWork,
        /**
         * The actual work for the resource that has been protected (made read-only).
         */
        ActualWorkProtected,
        /**
         * The name of the base calendar for the resource.
         */
        BaseCalendar,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline10BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline10BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline10Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline10Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline1BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline1BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline1Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline1Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline2BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline2BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline2Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline2Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline3BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline3BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline3Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline3Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline4BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline4BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline4Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline4Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline5BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline5BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline5Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline5Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline6BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline6BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline6Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline6Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline7BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline7BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline7Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline7Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline8BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline8BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline8Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline8Work,
        /**
         * The budget cost for the baseline resource.
         */
        Baseline9BudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        Baseline9BudgetWork,
        /**
         * The cost for the baseline resource.
         */
        Baseline9Cost,
        /**
         * The work for the baseline resource, in minutes.
         */
        Baseline9Work,
        /**
         * The budget cost for the baseline resource.
         */
        BaselineBudgetCost,
        /**
         * The budget work for the baseline resource, in hours.
         */
        BaselineBudgetWork,
        /**
         * The baseline cost for the resource for assignments in the project.
         */
        BaselineCost,
        /**
         * The baseline work for the resource for assignments in the project, in minutes.
         */
        BaselineWork,
        /**
         * The budget cost for the resource.
         */
        BudgetCost,
        /**
         * The budget work for the resource.
         */
        BudgetWork,
        /**
         * The GUID of the resource calendar.
         */
        ResourceCalendarGUID,
        /**
         * The code value of the resource.
         */
        Code,
        /**
         * A cost field for the resource.
         */
        Cost1,
        /**
         * A cost field for the resource.
         */
        Cost10,
        /**
         * A cost field for the resource.
         */
        Cost2,
        /**
         * A cost field for the resource.
         */
        Cost3,
        /**
         * A cost field for the resource.
         */
        Cost4,
        /**
         * A cost field for the resource.
         */
        Cost5,
        /**
         * A cost field for the resource.
         */
        Cost6,
        /**
         * A cost field for the resource.
         */
        Cost7,
        /**
         * A cost field for the resource.
         */
        Cost8,
        /**
         * A cost field for the resource.
         */
        Cost9,
        /**
         * The date the resource was created.
         */
        ResourceCreationDate,
        /**
         * A date field for the resource.
         */
        Date1,
        /**
         * A date field for the resource.
         */
        Date10,
        /**
         * A date field for the resource.
         */
        Date2,
        /**
         * A date field for the resource.
         */
        Date3,
        /**
         * A date field for the resource.
         */
        Date4,
        /**
         * A date field for the resource.
         */
        Date5,
        /**
         * A date field for the resource.
         */
        Date6,
        /**
         * A date field for the resource.
         */
        Date7,
        /**
         * A date field for the resource.
         */
        Date8,
        /**
         * A date field for the resource.
         */
        Date9,
        /**
         * A duration field for the resource.
         */
        Duration1,
        /**
         * A duration field for the resource.
         */
        Duration10,
        /**
         * A duration field for the resource.
         */
        Duration2,
        /**
         * A duration field for the resource.
         */
        Duration3,
        /**
         * A duration field for the resource.
         */
        Duration4,
        /**
         * A duration field for the resource.
         */
        Duration5,
        /**
         * A duration field for the resource.
         */
        Duration6,
        /**
         * A duration field for the resource.
         */
        Duration7,
        /**
         * A duration field for the resource.
         */
        Duration8,
        /**
         * A duration field for the resource.
         */
        Duration9,
        /**
         * The email address of the resource.
         */
        Email,
        /**
         * The end date of the resource availability.
         */
        End,
        /**
         * A finish field for the task.
         */
        Finish1,
        /**
         * A finish field for the task.
         */
        Finish10,
        /**
         * A finish field for the task.
         */
        Finish2,
        /**
         * A finish field for the task.
         */
        Finish3,
        /**
         * A finish field for the task.
         */
        Finish4,
        /**
         * A finish field for the task.
         */
        Finish5,
        /**
         * A finish field for the task.
         */
        Finish6,
        /**
         * A finish field for the task.
         */
        Finish7,
        /**
         * A finish field for the task.
         */
        Finish8,
        /**
         * A finish field for the task.
         */
        Finish9,
        /**
         * A Boolean flag field for the resource.
         */
        Flag10,
        /**
         * A Boolean flag field for the resource.
         */
        Flag1,
        /**
         * A Boolean flag field for the resource.
         */
        Flag11,
        /**
         * A Boolean flag field for the resource.
         */
        Flag12,
        /**
         * A Boolean flag field for the resource.
         */
        Flag13,
        /**
         * A Boolean flag field for the resource.
         */
        Flag14,
        /**
         * A Boolean flag field for the resource.
         */
        Flag15,
        /**
         * A Boolean flag field for the resource.
         */
        Flag16,
        /**
         * A Boolean flag field for the resource.
         */
        Flag17,
        /**
         * A Boolean flag field for the resource.
         */
        Flag18,
        /**
         * A Boolean flag field for the resource.
         */
        Flag19,
        /**
         * A Boolean flag field for the resource.
         */
        Flag2,
        /**
         * A Boolean flag field for the resource.
         */
        Flag20,
        /**
         * A Boolean flag field for the resource.
         */
        Flag3,
        /**
         * A Boolean flag field for the resource.
         */
        Flag4,
        /**
         * A Boolean flag field for the resource.
         */
        Flag5,
        /**
         * A Boolean flag field for the resource.
         */
        Flag6,
        /**
         * A Boolean flag field for the resource.
         */
        Flag7,
        /**
         * A Boolean flag field for the resource.
         */
        Flag8,
        /**
         * A Boolean flag field for the resource.
         */
        Flag9,
        /**
         * The group the resource belongs to.
         */
        Group,
        /**
         * The percentage of work units that the resource has assigned in the project. If the resource is working full-time on the project, Units = 100.
         */
        Units,
        /**
         * The name of the resource.
         */
        Name,
        /**
         * The text value of the notes regarding the resource.
         */
        Notes,
        /**
         * A number field for the resource.
         */
        Number1,
        /**
         * A number field for the resource.
         */
        Number10,
        /**
         * A number field for the resource.
         */
        Number11,
        /**
         * A number field for the resource.
         */
        Number12,
        /**
         * A number field for the resource.
         */
        Number13,
        /**
         * A number field for the resource.
         */
        Number14,
        /**
         * A number field for the resource.
         */
        Number15,
        /**
         * A number field for the resource.
         */
        Number16,
        /**
         * A number field for the resource.
         */
        Number17,
        /**
         * A number field for the resource.
         */
        Number18,
        /**
         * A number field for the resource.
         */
        Number19,
        /**
         * A number field for the resource.
         */
        Number2,
        /**
         * A number field for the resource.
         */
        Number20,
        /**
         * A number field for the resource.
         */
        Number3,
        /**
         * A number field for the resource.
         */
        Number4,
        /**
         * A number field for the resource.
         */
        Number5,
        /**
         * A number field for the resource.
         */
        Number6,
        /**
         * A number field for the resource.
         */
        Number7,
        /**
         * A number field for the resource.
         */
        Number8,
        /**
         * A number field for the resource.
         */
        Number9,
        /**
         * The overtime cost for a resource.
         */
        OvertimeCost,
        /**
         * The overtime rate for a resource.
         */
        OvertimeRate,
        /**
         * The overtime work for a resource.
         */
        OvertimeWork,
        /**
         * The percentage of work complete for a resource.
         */
        PercentWorkComplete,
        /**
         * The cost per use of the resource.
         */
        CostPerUse,
        /**
         * Indicates whether the resource is a generic resource (identified by skill rather than by name).
         */
        Generic,
        /**
         * Indicates whether the resource is overallocated.
         */
        OverAllocated,
        /**
         * The amount of regular work for the resource.
         */
        RegularWork,
        /**
         * The remaining cost for the resource.
         */
        RemainingCost,
        /**
         * The remaining overtime cost for the resource.
         */
        RemainingOvertimeCost,
        /**
         * The remaining overtime work for the resource, in minutes.
         */
        RemainingOvertimeWork,
        /**
         * The remaining work for the resource, in minutes.
         */
        RemainingWork,
        /**
         * The ID of the resource.
         */
        ResourceGUID,
        /**
         * The total cost of the resource.
         */
        Cost,
        /**
         * The total work for the resource, in minutes.
         */
        Work,
        /**
         * The start date for the resource.
         */
        Start,
        /**
         * A start field for the resource.
         */
        Start1,
        /**
         * A start field for the resource.
         */
        Start10,
        /**
         * A start field for the resource.
         */
        Start2,
        /**
         * A start field for the resource.
         */
        Start3,
        /**
         * A start field for the resource.
         */
        Start4,
        /**
         * A start field for the resource.
         */
        Start5,
        /**
         * A start field for the resource.
         */
        Start6,
        /**
         * A start field for the resource.
         */
        Start7,
        /**
         * A start field for the resource.
         */
        Start8,
        /**
         * A start field for the resource.
         */
        Start9,
        /**
         * The standard rate of pay for the resource, in cost per hour.
         */
        StandardRate,
        /**
         * A text field for the resource.
         */
        Text1,
        /**
         * A text field for the resource.
         */
        Text10,
        /**
         * A text field for the resource.
         */
        Text11,
        /**
         * A text field for the resource.
         */
        Text12,
        /**
         * A text field for the resource.
         */
        Text13,
        /**
         * A text field for the resource.
         */
        Text14,
        /**
         * A text field for the resource.
         */
        Text15,
        /**
         * A text field for the resource.
         */
        Text16,
        /**
         * A text field for the resource.
         */
        Text17,
        /**
         * A text field for the resource.
         */
        Text18,
        /**
         * A text field for the resource.
         */
        Text19,
        /**
         * A text field for the resource.
         */
        Text2,
        /**
         * A text field for the resource.
         */
        Text20,
        /**
         * A text field for the resource.
         */
        Text21,
        /**
         * A text field for the resource.
         */
        Text22,
        /**
         * A text field for the resource.
         */
        Text23,
        /**
         * A text field for the resource.
         */
        Text24,
        /**
         * A text field for the resource.
         */
        Text25,
        /**
         * A text field for the resource.
         */
        Text26,
        /**
         * A text field for the resource.
         */
        Text27,
        /**
         * A text field for the resource.
         */
        Text28,
        /**
         * A text field for the resource.
         */
        Text29,
        /**
         * A text field for the resource.
         */
        Text3,
        /**
         * A text field for the resource.
         */
        Text30,
        /**
         * A text field for the resource.
         */
        Text4,
        /**
         * A text field for the resource.
         */
        Text5,
        /**
         * A text field for the resource.
         */
        Text6,
        /**
         * A text field for the resource.
         */
        Text7,
        /**
         * A text field for the resource.
         */
        Text8,
        /**
         * A text field for the resource.
         */
        Text9
    }
    /**
     * Specifies the task fields that are available as a parameter for the {@link Office.Document | Document}.getTaskFieldAsync method.
     *
     * @remarks
     *
     * A ProjectTaskFields constant can be used as a parameter of the {@link Office.Document | Document}.getTaskFieldAsync method.
     *
     * For more information about working with fields in Project, see the
     * {@link https://support.office.com/article/Available-fields-reference-615a4563-1cc3-40f4-b66f-1b17e793a460 | Available fields} reference.
     * In Project Help, search for Available fields.
     */
    enum ProjectTaskFields {
        /**
         * The current actual cost for the task.
         */
        ActualCost,
        /**
         * The actual duration of the task, in minutes.
         */
        ActualDuration,
        /**
         * The actual finish date of the task.
         */
        ActualFinish,
        /**
         * The actual overtime cost for the task.
         */
        ActualOvertimeCost,
        /**
         * The actual overtime work for the task, in minutes.
         */
        ActualOvertimeWork,
        /**
         * The actual start date of the task.
         */
        ActualStart,
        /**
         * The actual work for the task, in minutes.
         */
        ActualWork,
        /**
         * A text field for the task.
         */
        Text1,
        /**
         * A text field for the task.
         */
        Text10,
        /**
         * A finish field for the task.
         */
        Finish10,
        /**
         * A start field for the task.
         */
        Start10,
        /**
         * A text field for the task.
         */
        Text11,
        /**
         * A text field for the task.
         */
        Text12,
        /**
         * A text field for the task.
         */
        Text13,
        /**
         * A text field for the task.
         */
        Text14,
        /**
         * A text field for the task.
         */
        Text15,
        /**
         * A text field for the task.
         */
        Text16,
        /**
         * A text field for the task.
         */
        Text17,
        /**
         * A text field for the task.
         */
        Text18,
        /**
         * A text field for the task.
         */
        Text19,
        /**
         * A finish field for the task.
         */
        Finish1,
        /**
         * A start field for the task.
         */
        Start1,
        /**
         * A text field for the task.
         */
        Text2,
        /**
         * A text field for the task.
         */
        Text20,
        /**
         * A text field for the task.
         */
        Text21,
        /**
         * A text field for the task.
         */
        Text22,
        /**
         * A text field for the task.
         */
        Text23,
        /**
         * A text field for the task.
         */
        Text24,
        /**
         * A text field for the task.
         */
        Text25,
        /**
         * A text field for the task.
         */
        Text26,
        /**
         * A text field for the task.
         */
        Text27,
        /**
         * A text field for the task.
         */
        Text28,
        /**
         * A text field for the task.
         */
        Text29,
        /**
         * A finish field for the task.
         */
        Finish2,
        /**
         * A start field for the task.
         */
        Start2,
        /**
         * A text field for the task.
         */
        Text3,
        /**
         * A text field for the task.
         */
        Text30,
        /**
         * A finish field for the task.
         */
        Finish3,
        /**
         * A start field for the task.
         */
        Start3,
        /**
         * A text field for the task.
         */
        Text4,
        /**
         * A finish field for the task.
         */
        Finish4,
        /**
         * A start field for the task.
         */
        Start4,
        /**
         * A text field for the task.
         */
        Text5,
        /**
         * A finish field for the task.
         */
        Finish5,
        /**
         * A start field for the task.
         */
        Start5,
        /**
         * A text field for the task.
         */
        Text6,
        /**
         * A finish field for the task.
         */
        Finish6,
        /**
         * A start field for the task.
         */
        Start6,
        /**
         * A text field for the task.
         */
        Text7,
        /**
         * A finish field for the task.
         */
        Finish7,
        /**
         * A start field for the task.
         */
        Start7,
        /**
         * A text field for the task.
         */
        Text8,
        /**
         * A finish field for the task.
         */
        Finish8,
        /**
         * A start field for the task.
         */
        Start8,
        /**
         * A text field for the task.
         */
        Text9,
        /**
         * A finish field for the task.
         */
        Finish9,
        /**
         * A start field for the task.
         */
        Start9,
        /**
         * The budget cost for the baseline task.
         */
        Baseline10BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline10BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline10Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline10Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline10Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline10FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline10FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline10Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline10Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline1BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline1BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline1Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline1Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline1Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline1FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline1FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline1Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline1Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline2BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline2BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline2Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline2Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline2Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline2FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline2FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline2Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline2Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline3BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline3BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline3Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline3Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline3Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline3FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline3FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline3Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline3Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline4BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline4BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline4Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline4Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline4Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline4FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline4FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline4Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline4Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline5BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline5BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline5Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline5Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline5Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline5FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline5FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline5Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline5Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline6BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline6BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline6Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline6Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline6Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline6FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline6FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline6Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline6Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline7BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline7BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline7Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline7Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline7Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline7FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline7FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline7Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline7Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline8BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline8BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline8Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline8Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline8Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline8FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline8FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline8Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline8Work,
        /**
         * The budget cost for the baseline task.
         */
        Baseline9BudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        Baseline9BudgetWork,
        /**
         * The cost for the baseline task.
         */
        Baseline9Cost,
        /**
         * The duration for the baseline task, in minutes.
         */
        Baseline9Duration,
        /**
         * The finish date for the baseline task.
         */
        Baseline9Finish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        Baseline9FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        Baseline9FixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        Baseline9Start,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        Baseline9Work,
        /**
         * The budget cost for the baseline task.
         */
        BaselineBudgetCost,
        /**
         * The budget work for the baseline task, in hours.
         */
        BaselineBudgetWork,
        /**
         * The cost for the baseline task.
         */
        BaselineCost,
        /**
         * The duration for the baseline task, in minutes.
         */
        BaselineDuration,
        /**
         * The finish date for the baseline task.
         */
        BaselineFinish,
        /**
         * The fixed cost of any non-resource expense for the baseline task.
         */
        BaselineFixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        BaselineFixedCostAccrual,
        /**
         * The start date for the baseline task.
         */
        BaselineStart,
        /**
         * The total person-hours scheduled for the baseline task, in minutes.
         */
        BaselineWork,
        /**
         * The budget cost for the task.
         */
        BudgetCost,
        BudgetFixedCost,
        BudgetFixedWork,
        /**
        * The budget work for the task, in hours.
        */
        BudgetWork,
        /**
         * The GUID of the task calendar.
         */
        TaskCalendarGUID,
        /**
         * A constraint date for the task.
         */
        ConstraintDate,
        /**
         * A constraint type for the task: As Soon As Possible = 0, As Late As Possible = 1, Must Start On = 2, Must Finish On = 3,
         * Start No Earlier Than = 4, Start No Later Than = 5, Finish No Earlier Than = 6, Finish No Later Than = 7.
         */
        ConstraintType,
        /**
         * A cost field of the task.
         */
        Cost1,
        /**
         * A cost field of the task.
         */
        Cost10,
        /**
         * A cost field of the task.
         */
        Cost2,
        /**
         * A cost field of the task.
         */
        Cost3,
        /**
         * A cost field of the task.
         */
        Cost4,
        /**
         * A cost field of the task.
         */
        Cost5,
        /**
         * A cost field of the task.
         */
        Cost6,
        /**
         * A cost field of the task.
         */
        Cost7,
        /**
         * A cost field of the task.
         */
        Cost8,
        /**
         * A cost field of the task.
         */
        Cost9,
        /**
         * A date field of the task.
         */
        Date1,
        /**
         * A date field of the task.
         */
        Date10,
        /**
         * A date field of the task.
         */
        Date2,
        /**
         * A date field of the task.
         */
        Date3,
        /**
         * A date field of the task.
         */
        Date4,
        /**
         * A date field of the task.
         */
        Date5,
        /**
         * A date field of the task.
         */
        Date6,
        /**
         * A date field of the task.
         */
        Date7,
        /**
         * A date field of the task.
         */
        Date8,
        /**
         * A date field of the task.
         */
        Date9,
        /**
         * The deadline for a task.
         */
        Deadline,
        /**
         * A duration field of the task.
         */
        Duration1,
        /**
         * A duration field of the task.
         */
        Duration10,
        /**
         * A duration field of the task.
         */
        Duration2,
        /**
         * A duration field of the task.
         */
        Duration3,
        /**
         * A duration field of the task.
         */
        Duration4,
        /**
         * A duration field of the task.
         */
        Duration5,
        /**
         * A duration field of the task.
         */
        Duration6,
        /**
         * A duration field of the task.
         */
        Duration7,
        /**
         * A duration field of the task.
         */
        Duration8,
        /**
         * A duration field of the task.
         */
        Duration9,
        /**
         * A duration field of the task.
         */
        Duration,
        /**
         * The method for calculating earned value for the task.
         */
        EarnedValueMethod,
        /**
         * The duration between the Early Finish and Late Finish dates for the task, in minutes.
         */
        FinishSlack,
        /**
         * The fixed cost for the task.
         */
        FixedCost,
        /**
         * The accrual method that defines how the baseline task accrues fixed costs: Accrues when the task starts = 1,
         * accrues when the task ends = 2, accrues as the task progresses (prorated) = 3.
         */
        FixedCostAccrual,
        /**
         * A Boolean flag field for the task.
         */
        Flag10,
        /**
         * A Boolean flag field for the task.
         */
        Flag1,
        /**
         * A Boolean flag field for the task.
         */
        Flag11,
        /**
         * A Boolean flag field for the task.
         */
        Flag12,
        /**
         * A Boolean flag field for the task.
         */
        Flag13,
        /**
         * A Boolean flag field for the task.
         */
        Flag14,
        /**
         * A Boolean flag field for the task.
         */
        Flag15,
        /**
         * A Boolean flag field for the task.
         */
        Flag16,
        /**
         * A Boolean flag field for the task.
         */
        Flag17,
        /**
         * A Boolean flag field for the task.
         */
        Flag18,
        /**
         * A Boolean flag field for the task.
         */
        Flag19,
        /**
         * A Boolean flag field for the task.
         */
        Flag2,
        /**
         * A Boolean flag field for the task.
         */
        Flag20,
        /**
         * A Boolean flag field for the task.
         */
        Flag3,
        /**
         * A Boolean flag field for the task.
         */
        Flag4,
        /**
         * A Boolean flag field for the task.
         */
        Flag5,
        /**
         * A Boolean flag field for the task.
         */
        Flag6,
        /**
         * A Boolean flag field for the task.
         */
        Flag7,
        /**
         * A Boolean flag field for the task.
         */
        Flag8,
        /**
         * A Boolean flag field for the task.
         */
        Flag9,
        /**
         * The amount of time that the task can be delayed without delaying its successor tasks.
         */
        FreeSlack,
        /**
         * Indicates whether the task has rollup subtasks.
         */
        HasRollupSubTasks,
        /**
         * The index of the selected task. After the project summary task, the index of the first task in a project is 1.
         */
        ID,
        /**
         * The name of the task.
         */
        Name,
        /**
         * The text value of the notes regarding the task.
         */
        Notes,
        /**
         * A number field for the task.
         */
        Number1,
        /**
         * A number field for the task.
         */
        Number10,
        /**
         * A number field for the task.
         */
        Number11,
        /**
         * A number field for the task.
         */
        Number12,
        /**
         * A number field for the task.
         */
        Number13,
        /**
         * A number field for the task.
         */
        Number14,
        /**
         * A number field for the task.
         */
        Number15,
        /**
         * A number field for the task.
         */
        Number16,
        /**
         * A number field for the task.
         */
        Number17,
        /**
         * A number field for the task.
         */
        Number18,
        /**
         * A number field for the task.
         */
        Number19,
        /**
         * A number field for the task.
         */
        Number2,
        /**
         * A number field for the task.
         */
        Number20,
        /**
         * A number field for the task.
         */
        Number3,
        /**
         * A number field for the task.
         */
        Number4,
        /**
         * A number field for the task.
         */
        Number5,
        /**
         * A number field for the task.
         */
        Number6,
        /**
         * A number field for the task.
         */
        Number7,
        /**
         * A number field for the task.
         */
        Number8,
        /**
         * A number field for the task.
         */
        Number9,
        /**
         * The scheduled (as opposed to actual) duration of the task.
         */
        ScheduledDuration,
        /**
         * The scheduled (as opposed to actual) finish date of the task.
         */
        ScheduledFinish,
        /**
         * The scheduled (as opposed to actual) start date of the task.
         */
        ScheduledStart,
        /**
         * The level of the task in the outline hierarchy.
         */
        OutlineLevel,
        /**
         * The overtime cost for the task.
         */
        OvertimeCost,
        /**
         * The overtime work for the task.
         */
        OvertimeWork,
        /**
         * The percent complete status of the task.
         */
        PercentComplete,
        /**
         * The percentage of work completed for the task.
         */
        PercentWorkComplete,
        /**
         * The IDs of the task's predecessors.
         */
        Predecessors,
        /**
         * The finish date of a task before leveling occurred.
         */
        PreleveledFinish,
        /**
         * The start date of a task before leveling occurred.
         */
        PreleveledStart,
        /**
         * The priority of the task, with values from 0 (low) to 1000 (high). The default priority value is 500.
         */
        Priority,
        /**
         * Indicates whether the task is active.
         */
        Active,
        /**
         * Indicates whether the task is on the critical path.
         */
        Critical,
        /**
         * Indicates whether the task is a milestone.
         */
        Milestone,
        /**
         * Indicates whether any assignments for a task are overallocated.
         */
        Overallocated,
        /**
         * Indicates whether subtask information is rolled up to the summary task bar.
         */
        IsRollup,
        /**
         * Indicates whether the task is a summary task.
         */
        Summary,
        /**
         * The amount of regular work for the task.
         */
        RegularWork,
        /**
         * The remaining cost for the task.
         */
        RemainingCost,
        /**
         * The remaining duration for the task, in minutes.
         */
        RemainingDuration,
        /**
         * The remaining overtime cost for the task.
         */
        RemainingOvertimeCost,
        /**
         * The remaining work for the task, in minutes.
         */
        RemainingWork,
        /**
         * The names of the resources assigned to a task.
         */
        ResourceNames,
        /**
         * The total cost of the task.
         */
        Cost,
        /**
         * The finish date of the task.
         */
        Finish,
        /**
         * The start date of the task.
         */
        Start,
        /**
         * The total person-hours scheduled for the task, in minutes.
         */
        Work,
        /**
         * The duration between the Early Start and Late Start dates for the task.
         */
        StartSlack,
        /**
         * The status of the task: Complete = 0, on schedule = 1, late = 2, future task = 3, status not available = 4.
         */
        Status,
        /**
         * The IDs of the task's successors.
         */
        Successors,
        /**
         * The enterprise resource responsible for accepting or rejecting assignment progress updates for the task.
         */
        StatusManager,
        /**
         * The total slack time for the task, in minutes.
         */
        TotalSlack,
        /**
         * The GUID of the task.
         */
        TaskGUID,
        /**
         * The way the task is calculated: Fixed units = 0, fixed duration = 1, fixed work = 2.
         */
        Type,
        /**
         * The work breakdown structure code of the task.
         */
        WBS,
        /**
         * The work breakdown structure codes of the task predecessors, separated by the list separator.
         */
        WBSPREDECESSORS,
        /**
         * The work breakdown structure codes of the task successors, separated by the list separator.
         */
        WBSSUCCESSORS,
        /**
         * The ID of the task in a SharePoint list, for a project that is synchronized with a SharePoint tasks list.
         */
        WSSID
    }
    /**
     * Specifies the types of views that the {@link Office.Document | Document}.getSelectedViewAsync method can recognize.
     *
     * @remarks
     *
     * The {@link Office.Document | Document}.getSelectedViewAsync method returns the ProjectViewTypes constant value and name that corresponds to the
     * active view.
     */
    enum ProjectViewTypes {
        /**
         * The Gantt chart view.
         */
        Gantt,
        /**
         * The Network Diagram view.
         */
        NetworkDiagram,
        /**
         * The Task Diagram view.
         */
        TaskDiagram,
        /**
         * The Task form view.
         */
        TaskForm,
        /**
         * The Task Sheet view.
         */
        TaskSheet,
        /**
         * The Resource Form view.
         */
        ResourceForm,
        /**
         * The Resource Sheet view.
         */
        ResourceSheet,
        /**
         * The Resource Graph view.
         */
        ResourceGraph,
        /**
         * The Team Planner view.
         */
        TeamPlanner,
        /**
         * The Task Details view.
         */
        TaskDetails,
        /**
         * The Task Name Form view.
         */
        TaskNameForm,
        /**
         * The Resource Names view.
         */
        ResourceNames,
        /**
         * The Calendar view.
         */
        Calendar,
        /**
         * The Task Usage view.
         */
        TaskUsage,
        /**
         * The Resource Usage view.
         */
        ResourceUsage,
        /**
         * The Timeline view.
         */
        Timeline
    }
}


////////////////////////////////////////////////////////////////
///////////////////// End Office namespace /////////////////////
////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////
//////////////// Begin OfficeExtension runtime /////////////////
////////////////////////////////////////////////////////////////

export declare namespace OfficeExtension {
    /**
     * An abstract proxy object that represents an object in an Office document. 
     * You create proxy objects from the context (or from other proxy objects), add commands to a queue to act on the object, and then synchronize the 
     * proxy object state with the document by calling `context.sync()`.
     */
    export class ClientObject {
        /** The request context associated with the object */
        context: ClientRequestContext;
        /**
         *  Returns a boolean value for whether the corresponding object is a null object. You must call `context.sync()` before reading the 
         * isNullObject property.
         */
        isNullObject: boolean;
    }

    /**
     * Specifies which properties of an object should be loaded. This load happens when the sync() method is executed.
     * This synchronizes the states between Office objects and corresponding JavaScript proxy objects.
     * 
     * @remarks
     * 
     * For Word, the preferred method for specifying the properties and paging information is by using a string literal. 
     * The first two examples show the preferred way to request the text and font size properties for paragraphs in a paragraph collection:
     * 
     * `context.load(paragraphs, 'text, font/size');`
     * 
     * `paragraphs.load('text, font/size');`
     * 
     * Here is a similar example using object notation (includes paging):
     * 
     * `context.load(paragraphs, {select: 'text, font/size', expand: 'font', top: 50, skip: 0});`
     * 
     * `paragraphs.load({select: 'text, font/size', expand: 'font', top: 50, skip: 0});`
     * 
     * Note that if we don't specify the specific properties on the font object in the select statement, the expand statement by itself would 
     * indicate that all of the font properties are loaded.
     */
    export interface LoadOption {
        /**
         * A comma-delimited string, or array of strings, that specifies the properties to load.
         */
        select?: string | string[];
        /**
         * A comma-delimited string, or array of strings, that specifies the navigation properties to load.
         */
        expand?: string | string[];
        /**
         * Only usable on collection types. Specifies the maximum number of collection items that can be included in the result.
         */
        top?: number;
        /**
         * Only usable on collection types. Specifies the number of items in the collection that are to be skipped and not included in the result. 
         * If top is specified, the result set will start after skipping the specified number of items.
         */
        skip?: number;
    }
    /**
     * Provides an option for suppressing an error when the object that is used to set multiple properties tries to set read-only properties.
     */
    export interface UpdateOptions {
        /**
         * Throw an error if the passed-in property list includes read-only properties (default = true).
         */
        throwOnReadOnly?: boolean
    }

    /**
     * Additional options passed into `{Host}.run(...)`.
     */
    export interface RunOptions<T> {
        /**
         * The URL of the remote workbook and the request headers to be sent.
         */
        session?: RequestUrlAndHeaderInfo | T;
    
        /**
         *  A previously-created context, or API object, or array of objects. 
         * The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up 
         * by `context.sync()`.
         */
        previousObjects?: ClientObject | ClientObject[] | ClientRequestContext;
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

    /**
     * An abstract RequestContext object that facilitates requests to the host Office application. 
     * The `Excel.run` and `Word.run` methods provide a request context.
     */
    export class ClientRequestContext {
        constructor(url?: string);

        /** Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. */
        trackedObjects: TrackedObjects;

        /** Request headers */
        requestHeaders: { [name: string]: string };

        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. 
         * 
         * @param object - The object whose properties are loaded.
         * @param option - A comma-delimited string, or array of strings, that specifies the properties to load, or an 
         * {@link OfficeExtension.LoadOption} object.
         */
        load(object: ClientObject, option?: string | string[] | LoadOption): void;

        /**
        * Queues up a command to recursively load the specified properties of the object and its navigation properties.
        * 
        * You must call `context.sync()` before reading the properties.
        *
        * @param object - The object to be loaded.
        * @param options - The key-value pairing of load options for the types, such as 
        *                `{ "Workbook": "worksheets,tables",  "Worksheet": "tables",  "Tables": "name" }`
        * @param maxDepth - The maximum recursive depth.
        */
        loadRecursive(object: ClientObject, options: { [typeName: string]: string | string[] | LoadOption }, maxDepth?: number): void;

        /**
         * Adds a trace message to the queue. If the promise returned by `context.sync()` is rejected due to an error, this adds a ".traceMessages" 
         * array to the OfficeExtension.Error object, containing all trace messages that were executed. 
         * These messages can help you monitor the program execution sequence and detect the cause of the error. 
         */
        trace(message: string): void;

        /**
         * Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context 
         * and retrieving properties of loaded Office objects for use in your code. 
         * This method returns a promise, which is resolved when the synchronization is complete.
         */
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

    /** Contains the result for methods that return primitive types. The object's value property is retrieved from the document after `context.sync()` is invoked. */
    export class ClientResult<T> {
        /** The value of the result that is retrieved from the document after `context.sync()` is invoked. */
        value: T;
    }

    /** Configuration */
    var config: {
        /**
         * Determines whether to log additional error information upon failure.
         *
         * When this property is set to true, the error object will include a "debugInfo.fullStatements" property that lists all statements in the 
         * batch request, including all statements that precede and follow the point of failure.
         *
         * Setting this property to true will negatively impact performance and will log all statements in the batch request, including any statements 
         * that may contain potentially-sensitive data.
         * It is recommended that you only set this property to true during debugging and that you never log the value of 
         * error.debugInfo.fullStatements to an external database or analytics service.
         */
        extendedErrorLogging: boolean;
    };
    /**
     * Provides information about an error.
     */
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
         * This statement will never contain any potentially-sensitive data and may not match the code exactly as written, 
         * but will be a close approximation.
         */
        statements?: string;
        /**
         * The statements that closely precede and follow the statement that caused the error, if available.
         *
         * These statements will never contain any potentially-sensitive data and may not match the code exactly as written, 
         * but will be a close approximation.
         */
        surroundingStatements?: string[];
        /**
         * All statements in the batch request (including any potentially-sensitive information that was specified in the request), if available.
         *
         * These statements may not match the code exactly as written, but will be a close approximation.
         */
        fullStatements?: string[];
    }

    /** The error object returned by `context.sync()`, if a promise is rejected due to an error while processing the request. */
    export class Error {
        /** Error name: "OfficeExtension.Error".*/
        name: string;
        /** The error message passed through from the host Office application. */
        message: string;
        /** Stack trace, if applicable. */
        stack: string;
        /** Error code string, such as "InvalidArgument". */
        code: string;
        /**
         * Trace messages (if any) that were added via a `context.trace()` invocation before calling `context.sync()`. 
         * If there was an error, this contains all trace messages that were executed before the error occurred. 
         * These messages can help you monitor the program execution sequence and detect the case of the error.
         */
        traceMessages: string[];
        /** Debug info (useful for detailed logging of the error, i.e., via `JSON.stringify(...)`). */
        debugInfo: DebugInfo;
        /** Inner error, if applicable. */
        innerError: Error;
    }

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

    /**
     * A Promise object that represents a deferred interaction with the host Office application. 
     * The publicly-consumable OfficeExtension.Promise is available starting in ExcelApi 1.2 and WordApi 1.2. 
     * Promises can be chained via ".then", and errors can be caught via ".catch". 
     * Remember to always use a ".catch" on the outer promise, and to return intermediary promises so as not to break the promise chain. 
     * When a browser-provided native Promise implementation is available, OfficeExtension.Promise will switch to use the native Promise instead.
     */
    const Promise: Office.IPromiseConstructor;
    export type IPromise<T> = Promise<T>;

    /** Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information. */
    export class TrackedObjects {
        /** 
         * Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. 
         * If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, 
         * and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object 
         * to the tracked object collection when the object was first created. 
         * 
         * This method also has the following signature: 
         * 
         * `add(objects: ClientObject[]): void;` Where objects is an array of objects to be tracked.
         */
        add(object: ClientObject): void;
        /**
         * Track a set of objects  for automatic adjustment based on surrounding changes in the document. Only some object types require this. 
         * If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, 
         * and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object 
         * to the tracked object collection when the object was first created.
         */
        add(objects: ClientObject[]): void;
        /** 
         * Release the memory associated with an object that was previously added to this collection. 
         * Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. 
         * You will need to call `context.sync()` before the memory release takes effect.
         * 
         * This method also has the following signature: 
         * 
         * `remove(objects: ClientObject[]): void;` Where objects is an array of objects to be removed.
         */
        remove(object: ClientObject): void;
        /**
         * Release the memory associated with an object that was previously added to this collection. 
         * Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. 
         * You will need to call `context.sync()` before the memory release takes effect.
         */
        remove(objects: ClientObject[]): void;
    }

    export class EventHandlers<T> {
        constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
        /**
         * Adds a function to be called when the event is triggered.
         * @param handler - A promise-based function that takes in any relevant event arguments.
         */
        add(handler: (args: T) => Promise<any>): EventHandlerResult<T>;
        /**
         * Removes the specified function from the event handler list so that it will not be called on subsequent events. 
         * 
         * **Note**: The same {@link OfficeExtension.ClientRequestContext | RequestContext} object that the handler was added in must be used when removing the handler. 
         * More information can be found in {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events#remove-an-event-handler | Remove an event handler}. 
         * 
         * @param handler - A reference to a function previously provided to the `add` method as an event handler. 
         */
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