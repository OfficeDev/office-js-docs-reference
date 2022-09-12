import {Office as CommonAPI} from "../../api-extractor-inputs-office/office"
////////////////////////////////////////////////////////////////
////////////////////// Begin Exchange APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export namespace MailboxEnums {
        
        
        
        /**
         * Specifies an attachment's type.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum AttachmentType {
            /**
             * The attachment is a file.
             */
            File = "file",
            /**
             * The attachment is an Exchange item.
             */
            Item = "item",
            /**
             * The attachment is stored in a cloud location, such as OneDrive.
             *
             * **Important**: In Read mode, the `id` property of the attachment's {@link Office.AttachmentDetails | details} object
             * contains a URL to the file.
             * From requirement set 1.8, the `url` property included in the attachment's
             * {@link https://docs.microsoft.com/javascript/api/outlook/office.attachmentdetailscompose?view=outlook-js-1.8 | details} object
             * contains a URL to the file in Compose mode.
             */
            Cloud = "cloud"
        }
        
        
        
        
        /**
         * Specifies an entity's type.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum EntityType {
            /**
             * Specifies that the entity is a meeting suggestion.
             */
            MeetingSuggestion = "meetingSuggestion",
            /**
             * Specifies that the entity is a task suggestion.
             */
            TaskSuggestion = "taskSuggestion",
            /**
             * Specifies that the entity is a postal address.
             */
            Address = "address",
            /**
             * Specifies that the entity is an SMTP email address.
             */
            EmailAddress = "emailAddress",
            /**
             * Specifies that the entity is an Internet URL.
             */
            Url = "url",
            /**
             * Specifies that the entity is a US phone number.
             */
            PhoneNumber = "phoneNumber",
            /**
             * Specifies that the entity is a contact.
             */
            Contact = "contact"
        }
        
        
        /**
         * Specifies the notification message type for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum ItemNotificationMessageType {
            /**
             * The notification message is a progress indicator.
             */
            ProgressIndicator = "progressIndicator",
            /**
             * The notification message is an informational message.
             */
            InformationalMessage = "informationalMessage",
            /**
             * The notification message is an error message.
             */
            ErrorMessage = "errorMessage",
                    }
        /**
         * Specifies an item's type.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum ItemType {
            /**
             * An email, meeting request, meeting response, or meeting cancellation.
             */
            Message = "message",
            /**
             * An appointment item.
             */
            Appointment = "appointment"
        }
        
        
        /**
         * Represents the current view of Outlook on the web.
         */
        enum OWAView {
            /**
             * One-column view. Displayed when the screen is narrow. Outlook on the web uses this single-column layout on the entire screen of a smartphone.
             */
            OneColumn = "OneColumn",
            /**
             * Two-column view. Displayed when the screen is wider. Outlook on the web uses this view on most tablets.
             */
            TwoColumns = "TwoColumns",
            /**
             Three-column view. Displayed when the screen is wide. For example, Outlook on the web uses this view in a full screen window on a desktop 
             computer.
             */
            ThreeColumns = "ThreeColumns"
        }
        /**
         * Specifies the type of recipient for an appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * **Important**: A `recipientType` property value isn't returned by the 
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-1.7#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync} 
         * and {@link https://docs.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-1.7#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         */
        enum RecipientType {
            /**
             * Specifies that the recipient is a distribution list containing a list of email addresses.
             */
            DistributionList = "distributionList",
            /**
             * Specifies that the recipient is an SMTP email address that is on the Exchange server.
             */
            User = "user",
            /**
             * Specifies that the recipient is an SMTP email address that is not on the Exchange server.
             */
            ExternalUser = "externalUser",
            /**
             * Specifies that the recipient is not one of the other recipient types.
             */
            Other = "other"
        }
        
        
        /**  
         * Specifies the type of response to a meeting invitation.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum ResponseType {
            /**
             * There has been no response from the attendee.
             */
            None = "none",
            /**
             * The attendee is the meeting organizer.
             */
            Organizer = "organizer",
            /**
             * The meeting request was tentatively accepted by the attendee.
             */
            Tentative = "tentative",
            /**
             * The meeting request was accepted by the attendee.
             */
            Accepted = "accepted",
            /**
             * The meeting request was declined by the attendee.
             */
            Declined = "declined"
        }
        /**
         * Specifies the version of the REST API that corresponds to a REST-formatted item ID.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum RestVersion {
            /**
             * Version 1.0.
             */
            v1_0 = "v1.0",
            /**
             * Version 2.0.
             */
            v2_0 = "v2.0",
            /**
             * Beta.
             */
            Beta = "beta"
        }
        /**
         * Specifies the source of the selected data in an item (see `Office.mailbox.item.getSelectedDataAsync` for details).
         * 
         * @remarks
         * [Api set: Mailbox 1.2]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         */
        enum SourceProperty {
            /**
             * The source of the data is from the body of the item.
             */
            Body = "body",
            /**
             * The source of the data is from the subject of the item.
             */
            Subject = "subject"
        }
        
    }
    /**
     * Provides an option for the data format.
     */
    export interface CoercionTypeOptions {
        /**
         * The desired data format.
         */
        coercionType?: CommonAPI.CoercionType | string;
    }
    /**
     * The subclass of {@link Office.Item | Item} dealing with appointments.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Child interfaces:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     * 
     * - {@link Office.AppointmentRead | AppointmentRead}
     */
    export interface Appointment extends Item {
    }
    /**
     * The appointment organizer mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Parent interfaces:
     *
     * - {@link Office.ItemCompose | ItemCompose}
     * 
     * - {@link Office.Appointment | Appointment}
     */
    export interface AppointmentCompose extends Appointment, ItemCompose {
         /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        body: Body;
        
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The `end` property is a {@link Office.Time | Time} object expressed as a Coordinated Universal Time (UTC) date and time value. 
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on 
         * the client to UTC for the server.
         *
         * **Important**: In the Windows client, you can't use this property to update the end of a recurrence.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        end: Time;
        
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets or sets the location of an appointment. The `location` property returns a {@link Office.Location | Location} object that provides methods that are 
         * used to get and set the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        location: Location;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        notificationMessages: NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
         * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
         * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        optionalAttendees: Recipients;
        
        
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the
         * required attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
         * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        requiredAttendees: Recipients;
        
        
        /**
         * Gets or sets the date and time that the appointment is to begin.
         *
         * The `start` property is a {@link Office.Time | Time} object expressed as a Coordinated Universal Time (UTC) date and time value. 
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the start time, you should use the `convertToUtcClientTime` method to convert the local time on 
         * the client to UTC for the server.
         *
         * **Important**: In the Windows client, you can't use this property to update the start of a recurrence.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        start: Time;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        subject: Subject;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         * 
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body,
         *                    and should not be displayed in the attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: CommonAPI.AsyncContextOptions & { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         * 
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        
        
        
        
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. 
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or 
         * a code that indicates any error that occurred while attaching the item. 
         * You can use the `options` parameter to pass state information to the callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that 
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If adding the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. 
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or 
         * a code that indicates any error that occurred while attaching the item. 
         * You can use the `options` parameter to pass state information to the callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that 
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If adding the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the `close` method depends on the current state of the item being composed. 
         * If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.
         *
         * **Note**: In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, 
         * discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        close(): void;
        
        
        
        
        
        
        
        
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. 
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`. 
         * To access the `source` property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present. 
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. 
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`. 
         * To access the `source` property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present. 
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        
        
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. 
         * This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the 
         * current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property. 
         * This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to 
         * the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. 
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. 
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment 
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session. 
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to 
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type `Office.AsyncResult`. 
         *                 If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. 
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment 
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session. 
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to 
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type `Office.AsyncResult`. 
         *                 If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        
        
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function. 
         * In Outlook on the web or Outlook in online mode, the item is saved to the server. 
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal 
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. 
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. 
         * Until the item is synced, using the item ID will return an error.
         *
         * **Note**: In Outlook on Mac, only build 16.35.308 or later supports saving a meeting.
         * Otherwise, the `saveAsync` method fails when called from a meeting in compose mode.
         * For a workaround, see {@link https://docs.microsoft.com/outlook/troubleshoot/calendars/cannot-save-meeting-as-draft-in-outlook-for-mac | Cannot save a meeting as a draft in Outlook for Mac by using Office JS API}.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`. 
         */
        saveAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function. 
         * In Outlook on the web or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal 
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. 
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware that 
         * when Outlook is in cached mode, it may take some time before the item is actually synced to the server. 
         * Until the item is synced, using the item ID will return an error.
         *
         * **Note**: In Outlook on Mac, only build 16.35.308 or later supports saving a meeting.
         * Otherwise, the `saveAsync` method fails when called from a meeting in compose mode.
         * For a workaround, see {@link https://docs.microsoft.com/outlook/troubleshoot/calendars/cannot-save-meeting-as-draft-in-outlook-for-mac | Cannot save a meeting as a draft in Outlook for Mac by using Office JS API}.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is 
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. 
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. 
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web and Windows. 
         *                      If the field is an HTML editor, only the text data is inserted, even if the data is HTML. 
         *                      If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the 
         *                      default style is applied in Outlook on desktop clients. 
         *                      If the field is a text field, an `InvalidDataFormat` error is returned. 
         *                      If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; 
         *                      if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is 
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. 
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. 
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The `AppointmentForm` object is used to access the currently selected appointment.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface AppointmentForm {
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        body: Body | string;
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the `convertToLocalClientTime` method to 
         * convert the `end` property value to the client's local date and time.
         *
         * *Read mode*
         *
         * The `end` property returns a `Date` object.
         *
         * *Compose mode*
         *
         * The `end` property returns a `Time` object.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on 
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        end: Time | Date;
        /**
        * Gets or sets the location of an appointment.
        *
        * *Read mode*
        *
        * The `location` property returns a string that contains the location of the appointment.
        *
        * *Compose mode*
        *
        * The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        * 
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       location: Location | string;
       /**
        * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
        *
        * *Read mode*
        *
        * The `optionalAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
        * each optional attendee to the meeting. Collection size limits:
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * - New Mac UI, web browser, Android: No limit
        *
        * *Compose mode*
        *
        * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
        * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        * 
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       optionalAttendees: Recipients[] | EmailAddressDetails[];
       /**
        * Provides access to the resources of an event. Returns an array of strings containing the resources required for the appointment.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        * 
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       resources: string[];

       /**
        * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
        *
        * *Read mode*
        *
        * The `requiredAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
        * each required attendee to the meeting. Collection size limits:
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * - New Mac UI, web browser, Android: No limit
        *
        * *Compose mode*
        *
        * The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the
        * required attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       requiredAttendees: Recipients[] | EmailAddressDetails[];
       /**
        * Gets or sets the date and time that the appointment is to begin.
        *
        * The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the `convertToLocalClientTime` method
        * to convert the value to the client's local date and time.
        *
        * *Read mode*
        *
        * The `start` property returns a `Date` object.
        *
        * *Compose mode*
        *
        * The `start` property returns a `Time` object.
        *
        * When you use the `Time.setAsync` method to set the start time, you should use the `convertToUtcClientTime` method to convert the local time on
        * the client to UTC for the server.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       start: Time | Date;
       /**
        * Gets or sets the description that appears in the subject field of an item.
        *
        * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
        *
        * *Read mode*
        *
        * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
        *
        * *Compose mode*
        *
        * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
        *
        * @remarks
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
        *
        * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       subject: Subject | string;
    }
    /**
     * The appointment attendee mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Parent interfaces:
     *
     * - {@link Office.ItemRead | ItemRead}
     * 
     * - {@link Office.Appointment | Appointment}
     */
    export interface AppointmentRead extends Appointment, ItemRead {
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Note**: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, refer to 
         * {@link https://support.microsoft.com/office/434752e1-02d3-4e90-9124-8b81e49a8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        body: Body;
        
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        dateTimeModified: Date;
        /**
         * Gets the date and time that the appointment is to end.
         *
         * The `end` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value. 
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on 
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        end: Date;
        
        /**
         * Gets the Exchange Web Services item class of the selected item.
         *
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         * 
         * The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.
         * 
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurrence.</td>
         *     <td>IPM.Appointment,IPM.Appointment.Occurrence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note,IPM.Schedule.Meeting.Request,IPM.Schedule.Meeting.Neg,IPM.Schedule.Meeting.Pos,IPM.Schedule.Meeting.Tent,IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         * 
         */
        itemClass: string;
        /**
         * Gets the {@link https://docs.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}
         * for the current item.
         *
         * The `itemId` property is not available in compose mode. 
         * If an item identifier is required, the `saveAsync` method can be used to save the item to the store, which will return the item identifier 
         * in the `asyncResult.value` parameter in the callback function.
         *
         * **Note**: The identifier returned by the `itemId` property is the same as the
         * {@link https://docs.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}. 
         * The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. 
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`. 
         * For more details, see {@link https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api#get-the-item-id | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of an appointment.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        location: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The `normalizedSubject` property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by email programs. 
         * To get the subject of the item with the prefixes intact, use the `subject` property.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        notificationMessages: NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each optional attendee to the meeting. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - New Mac UI, web browser, Android: No limit
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        optionalAttendees: EmailAddressDetails[];
        /**
         * Gets the meeting organizer's email properties.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        organizer: EmailAddressDetails;
        
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `requiredAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each required attendee to the meeting. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - New Mac UI, web browser, Android: No limit
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        requiredAttendees: EmailAddressDetails[];
        /**
         * Gets the date and time that the appointment is to begin.
         *
         * The `start` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value. 
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        start: Date;
        
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        subject: string;

        
        
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        
        
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        
        
        
        
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null. 
         * If no entities of the specified type are present in the item's body, the method returns an empty array. 
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param entityType - One of the `EntityType` enumeration values.
         * 
         * While the minimum permission level to use this method is `Restricted`, some entity types require `ReadItem` to access, as specified in the following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Level</th>
         *   </tr>
         *   <tr>
         *     <td>Address</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>Contact</td>
         *     <td>Contact</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>EmailAddress</td>
         *     <td>String</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>MeetingSuggestion</td>
         *     <td>MeetingSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>PhoneNumber</td>
         *     <td>PhoneNumber</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>TaskSuggestion</td>
         *     <td>TaskSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>URL</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         * </table>
         */
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.
         *
         * The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the `ItemHasKnownEntity` rule element 
         * in the manifest XML file with the specified `FilterName` element value.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         * 
         * @returns If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, 
         * the method returns `null`. 
         * If the `name` parameter matches an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match,
         * the method return an empty array.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns string values in the selected item that match the regular expressions defined in the manifest XML file.
         *
         * The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file. 
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item. 
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results. 
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. 
         * The name of each array is equal to the corresponding value of the RegExName attribute of the matching `ItemHasRegularExpressionMatch` rule 
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in the manifest XML file.
         *
         * The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule
         * element in the manifest XML file with the specified `RegExName` element value.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item. 
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        /**
         * Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getSelectedEntities(): Entities;
        /**
         * Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. 
         * Highlighted matches apply to contextual add-ins.
         *
         * The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file.
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getSelectedRegExMatches(): any;
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis.
         * This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the asyncResult.value property.
         * This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to
         * the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        
        
    }
    
    
    
    /**
     * Represents an attachment on an item from the server. Read mode only.
     *
     * An array of `AttachmentDetails` objects is returned as the attachments property of an appointment or message item.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     *
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface AttachmentDetails {
        /**
         * Gets a value that indicates the type of an attachment.
         */
        attachmentType: MailboxEnums.AttachmentType | string;
        /**
         * Gets the MIME content type of the attachment.
         *
         * **Warning**: While the `contentType` value is a direct lookup of the attachment's extension, the internal mapping isn't actively maintained
         * so this property has been deprecated. If you require specific types, grab the attachment's extension and process accordingly. For details,
         * refer to the {@link https://devblogs.microsoft.com/microsoft365dev/outlook-javascript-api-deprecation-for-attachmentdetails-contenttype-property/ | related blog post }.
         *
         * @deprecated If you require specific content types, grab the attachment's extension and process accordingly.
         */
        contentType: string;
        /**
         * Gets the Exchange attachment ID of the attachment.
         * However, if the attachment type is `MailboxEnums.AttachmentType.Cloud`, then a URL for the file is returned.
         */
        id: string;
        /**
         * Gets a value that indicates whether the attachment should be displayed in the body of the item.
         */
        isInline: boolean;
        /**
         * Gets the name of the attachment.
         *
         * **Important**: For message or appointment items that were attached by drag-and-drop or "Attach Item",
         * `name` includes a file extension in Outlook on Mac, but excludes the extension on the web or Windows.
         */
        name: string;
        /**
         * Gets the size of the attachment in bytes.
         */
        size: number;
    }
    
    /**
     * The body object provides methods for adding and updating the content of the message or appointment. 
     * It is returned in the body property of the selected item.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **Known issue with HTML table border colors**
     *
     * Outlook on Windows: If you're setting various cell borders to different colors in an HTML table in Compose mode, a cell's borders may not reflect
     * the expected color. For the known behavior, visit {@link https://github.com/OfficeDev/office-js/issues/1818 | OfficeDev/office-js issue #1818}.
     *
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     *
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Body {
        
        
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by `coercionType`.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously. 
         * The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param coercionType - The format for the returned body.
         * @param options - An object literal that contains one or more of the following properties:
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. The body is provided in the requested format in the `asyncResult.value` property.
         */
        getAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by `coercionType`.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously. 
         * The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param coercionType - The format for the returned body.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. The body is provided in the requested format in the `asyncResult.value` property.
         */
        getAsync(coercionType: CommonAPI.CoercionType | string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  The content type is returned as one of the `CoercionType` values in the `asyncResult.value` property.
         */
        getTypeAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<CommonAPI.CoercionType>) => void): void;
        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  The content type is returned as one of the `CoercionType` values in the `asyncResult.value` property.
         */
        getTypeAsync(callback?: (asyncResult: CommonAPI.AsyncResult<CommonAPI.CoercionType>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The `prependAsync` method inserts the specified string at the beginning of the item body.
         * After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` in order to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) will not necessarily exactly contain the value that was passed in the `prependAsync` method previously.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         * 
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        prependAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The `prependAsync` method inserts the specified string at the beginning of the item body.
         * After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` in order to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) will not necessarily exactly contain the value that was passed in the `prependAsync` method previously.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         * 
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        prependAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method
         * previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP" 
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**: In Outlook on Windows and on Mac, the add-in user won't be able to revert this action with the **Undo** command.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         * 
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method
         * previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP" 
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**: In Outlook on Windows and on Mac, the add-in user won't be able to revert this action with the **Undo** command.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         * 
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in
         * the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the
         * UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 1,000,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         * 
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in
         * the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the
         * UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 1,000,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         * 
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        
        
     }
    
    
    /**
     * Represents the details about a contact (similar to what's on a physical contact or business card) extracted from the item's body. Read mode only.
     *
     * The list of contacts extracted from the body of an email message or appointment is returned in the `contacts` property of the
     * {@link Office.Entities | Entities} object returned by the `getEntities` or `getEntitiesByType` method of the current item.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface Contact {
        /**
         * An array of strings containing the mailing and street addresses associated with the contact. Nullable.
         */
        addresses: string[];
        /**
         * A string containing the name of the business associated with the contact. Nullable.
         */
        businessName: string;
        /**
         * An array of strings containing the SMTP email addresses associated with the contact. Nullable.
         */
        emailAddresses: string[];
        /**
         * A string containing the name of the person associated with the contact. Nullable.
         */
        personName: string;
        /**
         * An array containing a `PhoneNumber` object for each phone number associated with the contact. Nullable.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * An array of strings containing the Internet URLs associated with the contact. Nullable.
         */
        urls: string[];
    }
    /**
     * The `CustomProperties` object represents custom properties that are specific to a particular item and specific to a mail add-in for Outlook.
     * For example, there might be a need for a mail add-in to save some data that is specific to the current email message that activated the add-in. 
     * If the user revisits the same message in the future and activates the mail add-in again, the add-in will be able to retrieve the data that had 
     * been saved as custom properties. **Important**: The maximum length of a `CustomProperties` JSON object is 2500 characters.
     *
     * Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins cannot access their custom properties.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface CustomProperties {
        /**
         * Returns the value of the specified custom property.
         * 
         * @returns The value of the specified custom property.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param name - The name of the custom property to be returned.
         */
        get(name: string): any;
        
        /**
         * Removes the specified property from the custom property collection.
         *
         * To make the removal of the property permanent, you must call the `saveAsync` method of the `CustomProperties` object.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param name - The `name` of the property to be removed.
         */
        remove(name: string): void;
        /**
         * Saves item-specific custom properties to the server.
         *
         * You must call the `saveAsync` method to persist any changes made with the `set` method or the `remove` method of the `CustomProperties` object.
         * The saving action is asynchronous.
         *
         * It's a good practice to have your callback function check for and handle errors from `saveAsync`.
         * In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes
         * disconnected. 
         * If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error.
         * Your callback function should handle this error accordingly.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<void>) => void, asyncContext?: any): void;
        /**
         * Saves item-specific custom properties to the server.
         *
         * You must call the `saveAsync` method to persist any changes made with the `set` method or the `remove` method of the `CustomProperties` object.
         * The saving action is asynchronous.
         *
         * It's a good practice to have your callback function check for and handle errors from `saveAsync`.
         * In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes
         * disconnected. 
         * If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error.
         * Your callback function should handle this error accordingly.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
         saveAsync(asyncContext?: any): void;
         /**
         * Sets the specified property to the specified value.
         *
         * The `set` method sets the specified property to the specified value. You must use the `saveAsync` method to save the property to the server.
         *
         * The `set` method creates a new property if the specified property does not already exist; 
         * otherwise, the existing value is replaced with the new value. 
         * The `value` parameter can be of any type; however, it is always passed to the server as a string.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The name of the property to be set.
         * @param value - The value of the property to be set.
         */
        set(name: string, value: string): void;
    }
    /**
     * Provides diagnostic information to an Outlook add-in.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     * 
     * Starting with Mailbox requirement set 1.5, you can also use the 
     * {@link https://docs.microsoft.com/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics} 
     * property to get similar information.
     */
    export interface Diagnostics {
        /**
         * Gets a string that represents the name of the host application.
         *
         * A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         *
         * **Note**: The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        hostName: string;
        /**
         * Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").
         *
         * If the mail add-in is running in Outlook on a desktop or mobile client, the `hostVersion` property returns the version of the host 
         * application, Outlook. In Outlook on the web, the property returns the version of the Exchange Server.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        hostVersion: string;
        /**
         * Gets a string that represents the current view of Outlook on the web.
         *
         * The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.
         *
         * If the host application is not Outlook on the web, then accessing this property results in undefined.
         *
         * Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:
         *
         * - `OneColumn`, which is displayed when the screen is narrow. Outlook on the web uses this single-column layout on the entire screen of a 
         * smartphone.
         *
         * - `TwoColumns`, which is displayed when the screen is wider. Outlook on the web uses this view on most tablets.
         *
         * - `ThreeColumns`, which is displayed when the screen is wide. For example, Outlook on the web uses this view in a full screen window on a 
         * desktop computer.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        OWAView: MailboxEnums.OWAView | "OneColumn" | "TwoColumns" | "ThreeColumns";
    }
    /**
     * Provides the email properties of the sender or specified recipients of an email message or appointment.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface EmailAddressDetails {
        /**
         * Gets the SMTP email address.
         */
        emailAddress: string;
        /**
         * Gets the display name associated with an email address.
         */
        displayName: string;
        /**
         * Gets the response that an attendee returned for an appointment. 
         * This property applies to only an attendee of an appointment, as represented by the `optionalAttendees` or `requiredAttendees` property. 
         * This property returns undefined in other scenarios.
         */
        appointmentResponse: MailboxEnums.ResponseType | string;
        /**
         * Gets the email address type of a recipient.
         * 
         * @remarks
         * **Important**: A `recipientType` property value isn't returned by the 
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-1.7#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync} 
         * and {@link https://docs.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-1.7#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         */
        recipientType: MailboxEnums.RecipientType | string;
    }
    /**
     * Represents an email account on an Exchange Server.
     *
     * `EmailUser` objects are primarily received in {@link Office.MeetingSuggestion | MeetingSuggestion} and
     * {@link Office.TaskSuggestion | TaskSuggestion} entities extracted from an Outlook item. To learn more about this scenario, refer to
     * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/extract-entity-strings-from-an-item | Extract entity strings from an Outlook item}.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface EmailUser {
        /**
         * Gets the display name associated with an email address.
         */
        displayName: string;
        /**
         * Gets the SMTP email address.
         */
        emailAddress: string;
    }
    
    
    /**
     * Represents a collection of entities found in an email message or appointment. Read mode only.
     *
     * The `Entities` object is a container for the entity arrays returned by the `getEntities` and `getEntitiesByType` methods when the item 
     * (either an email message or an appointment) contains one or more entities that have been found by the server. 
     * You can use these entities in your code to provide additional context information to the viewer, such as a map to an address found in the item, 
     * or to open a dialer for a phone number found in the item.
     *
     * If no entities of the type specified in the property are present in the item, the property associated with that entity is null. 
     * For example, if a message contains a street address and a phone number, the addresses property and phoneNumbers property would contain 
     * information, and the other properties would be null.
     *
     * To be recognized as an address, the string must contain a United States postal address that has at least a subset of the elements of a street 
     * number, street name, city, state, and zip code.
     *
     * To be recognized as a phone number, the string must contain a North American phone number format.
     *
     * Entity recognition relies on natural language recognition that is based on machine learning of large amounts of data. 
     * The recognition of an entity is non-deterministic and success sometimes relies on the particular context in the item.
     *
     * When the property arrays are returned by the `getEntitiesByType` method, only the property for the specified entity contains data; 
     * all other properties are null.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface Entities {
        /**
         * Gets the physical addresses (street or mailing addresses) found in an email message or appointment.
         */
        addresses: string[];
        /**
         * Gets the contacts found in an email address or appointment.
         */
        contacts: Contact[];
        /**
         * Gets the email addresses found in an email message or appointment.
         */
        emailAddresses: string[];
        /**
         * Gets the meeting suggestions found in an email message.
         */
        meetingSuggestions: MeetingSuggestion[];
        /**
         * Gets the phone numbers found in an email message or appointment.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * Gets the task suggestions found in an email message or appointment.
         */
        taskSuggestions: string[];
        /**
         * Gets the Internet URLs present in an email message or appointment.
         */
        urls: string[];
    }
    
    
    
    
    /**
     * The item namespace is used to access the currently selected message, meeting request, or appointment. 
     * You can determine the type of the item by using the `itemType` property.
     *
     * To see the full member list, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     *
     * If you want to see IntelliSense for only a specific type or mode, cast this item to one of the following:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     * 
     * - {@link Office.AppointmentRead | AppointmentRead}
     * 
     * - {@link Office.MessageCompose | MessageCompose}
     * 
     * - {@link Office.MessageRead | MessageRead}
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer, Appointment Attendee, Message Compose, Message Read
     */
    export interface Item {
    }
    /**
     * The compose mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Child interfaces:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     * 
     * - {@link Office.MessageCompose | MessageCompose}
     */
    export interface ItemCompose extends Item {
    }
    /**
     * The read mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Child interfaces:
     *
     * - {@link Office.AppointmentRead | AppointmentRead}
     * 
     * - {@link Office.MessageRead | MessageRead}
     */
    export interface ItemRead extends Item {
    }
    /**
     * Represents a date and time in the local client's time zone. Read mode only.
     *
     * @remarks
     *
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface LocalClientTime {
        /**
         * Integer value representing the month, beginning with 0 for January to 11 for December.
         */
        month: number;
        /**
         * Integer value representing the day of the month.
         */
        date: number;
        /**
         * Integer value representing the year.
         */
        year: number;
        /**
         * Integer value representing the hour on a 24-hour clock.
         */
        hours: number;
        /**
         * Integer value representing the minutes.
         */
        minutes: number;
        /**
         * Integer value representing the seconds.
         */
        seconds: number;
        /**
         * Integer value representing the milliseconds.
         */
        milliseconds: number;
        /**
         * Integer value representing the number of minutes difference between the local time zone and UTC.
         */
        timezoneOffset: number;
    }
    /**
     * Provides methods to get and set the location of a meeting in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Location {
        /**
         * Gets the location of an appointment.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the location of an appointment.
         * The location of the appointment is provided as a string in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the location of an appointment.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the location of an appointment.
         * The location of the appointment is provided as a string in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the location of an appointment. 
         * Setting the location of an appointment overwrites the current location.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - DataExceedsMaximumSize: The location parameter is longer than 255 characters.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. If setting the location fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(location: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the location of an appointment. 
         * Setting the location of an appointment overwrites the current location.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - DataExceedsMaximumSize: The location parameter is longer than 255 characters.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. If setting the location fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(location: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    
    
    /**
     * Provides access to the Microsoft Outlook add-in object model.
     *
     * Key properties:
     *
     * - `diagnostics`: Provides diagnostic information to an Outlook add-in.
     *
     * - `item`: Provides methods and properties for accessing a message or appointment in an Outlook add-in.
     *
     * - `userProfile`: Provides information about the user in an Outlook add-in.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Mailbox {
        /**
         * Provides diagnostic information to an Outlook add-in.
         * 
         * Contains the following members:
         * 
         *  - `hostName` (string): A string that represents the name of the host application. 
         * It should be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         * **Note**: The "Outlook" value is returned for Outlook on desktop clients (i.e., Windows and Mac).
         * 
         *  - `hostVersion` (string): A string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0"). 
         * If the mail add-in is running in Outlook on desktop or mobile clients, the `hostVersion` property returns the version of the 
         * host application, Outlook. In Outlook on the web, the property returns the version of the Exchange Server.
         * 
         *  - `OWAView` (`MailboxEnums.OWAView` or string): An enum (or string literal) that represents the current view of Outlook on the web. 
         * If the host application is not Outlook on the web, then accessing this property results in undefined. 
         * Outlook on the web has three views (`OneColumn` - displayed when the screen is narrow, `TwoColumns` - displayed when the screen is wider, 
         * and `ThreeColumns` - displayed when the screen is wide) that correspond to the width of the screen and the window, and the number of columns 
         * that can be displayed.
         *
         * More information is under {@link Office.Diagnostics}. 
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * Starting with Mailbox requirement set 1.5, you can also use the 
         * {@link https://docs.microsoft.com/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics} 
         * property to get similar information.
         */
        diagnostics: Diagnostics;
        /**
         * Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.
         *
         * Your app must have the `ReadItem` permission specified in its manifest to call the `ewsUrl` member in read mode.
         *
         * In compose mode you must call the `saveAsync` method before you can use the `ewsUrl` member. 
         * Your app must have `ReadWriteItem` permissions to call the `saveAsync` method.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. 
         * For example, you can create a remote service to {@link https://docs.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         */
        ewsUrl: string;
        /**
         * The mailbox item. Depending on the context in which the add-in opened, the item type may vary.
         * If you want to see IntelliSense for only a specific type or mode, cast this item to one of the following:
         *
         * {@link Office.MessageCompose | MessageCompose}, {@link Office.MessageRead | MessageRead},
         * {@link Office.AppointmentCompose | AppointmentCompose}, {@link Office.AppointmentRead | AppointmentRead}
         *
         * **Important**: `item` can be null if your add-in supports pinning the task pane. For details on how to handle, see
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane#implement-the-event-handler | Implement a pinnable task pane in Outlook}.
         */
        item?: Item & ItemCompose & ItemRead & Message & MessageCompose & MessageRead & Appointment & AppointmentCompose & AppointmentRead;
        
        /**
         * Gets the URL of the REST endpoint for this email account.
         *
         * Your app must have the `ReadItem` permission specified in its manifest to call the `restUrl` member in read mode.
         *
         * In compose mode you must call the `saveAsync` method before you can use the `restUrl` member. 
         * Your app must have `ReadWriteItem` permissions to call the `saveAsync` method.
         *
         * However, in delegate or shared scenarios, you should instead use the `targetRestUrl` property of the
         * {@link https://docs.microsoft.com/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8 | SharedProperties} object
         * (introduced in requirement set 1.8). For more information, see the
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/delegate-access | delegate access} article.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `restUrl` value can be used to make {@link https://docs.microsoft.com/outlook/rest/ | REST API} calls to the user's mailbox.
         */
        restUrl: string;
        /**
         * Information about the user associated with the mailbox. This includes their account type, display name, email address, and time zone.
         * 
         * More information is under {@link Office.UserProfile}
         */
        userProfile: UserProfile;

        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Converts an item ID formatted for REST into EWS format.
         *
         * Item IDs retrieved via a REST API (such as the Outlook Mail API or the Microsoft Graph) use a different format than the format used by
         * Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - An item ID formatted for the Outlook REST APIs.
         * @param restVersion - A value indicating the version of the Outlook REST API used to retrieve the item ID.
         */
        convertToEwsId(itemId: string, restVersion: MailboxEnums.RestVersion | string): string;
        /**
         * Gets a dictionary containing time information in local client time.
         *
         * The dates and times used by a mail app for Outlook on the web or desktop clients can use different time zones. 
         * Outlook uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). 
         * You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that 
         * the user expects.
         *
         * If the mail app is running in Outlook on desktop clients, the `convertToLocalClientTime` method will return a dictionary object
         * with the values set to the client computer time zone. 
         * If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object
         * with the values set to the time zone specified in the EAC.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param timeValue - A `Date` object.
         */
        convertToLocalClientTime(timeValue: Date): LocalClientTime;
        /**
         * Converts an item ID formatted for EWS into REST format.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the
         * {@link https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations | Outlook Mail API}
         * or the {@link https://graph.microsoft.io/ | Microsoft Graph}.
         * The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.
         *
         * @param itemId - An item ID formatted for Exchange Web Services (EWS)
         * @param restVersion - A value indicating the version of the Outlook REST API that the converted ID will be used with.
         */
        convertToRestId(itemId: string, restVersion: MailboxEnums.RestVersion | string): string;
        /**
         * Gets a `Date` object from a dictionary containing time information.
         *
         * The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a `Date` object with the correct values for the 
         * local date and time.
         * 
         * @returns A Date object with the time expressed in UTC.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param input - The local time value to convert.
         */
        convertToUtcClientTime(input: LocalClientTime): Date;
        /**
         * Displays an existing calendar appointment.
         *
         * The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on 
         * mobile devices.
         *
         * In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment
         * of a recurring series. However, you can't display an instance of the series because you can't access the properties
         * (including the item ID) of instances of a recurring series.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and 
         * no error message is returned.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         */
        displayAppointmentForm(itemId: string): void;
        
        
        /**
         * Displays an existing message.
         *
         * The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and
         * no error message is returned.
         *
         * Do not use the `displayMessageForm` with an itemId that represents an appointment. Use the `displayAppointmentForm` method to display
         * an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message.
         */
        displayMessageForm(itemId: string): void;
        
        
        /**
         * Displays a form for creating a new calendar appointment.
         *
         * The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting.
         * If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.
         *
         * In Outlook on the web, this method always displays a form with an attendees field.
         * If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.
         * If you have specified attendees, the form would include the attendees and a **Send** button.
         *
         * In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or
         * `resources` parameter, this method displays a meeting form with a **Send** button.
         * If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - An `AppointmentForm` describing the new appointment. All properties are optional.
         */
        displayNewAppointmentForm(parameters: AppointmentForm): void;
        
        
         /**
         * Displays a form for creating a new message.
         *
         * The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form 
         * fields are automatically populated with the contents of the parameters.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - A dictionary containing all values to be filled in for the user in the new form. All parameters are optional.
         * 
         *        `toRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object 
         *        for each of the recipients on the **To** line. The array is limited to a maximum of 100 entries.
         * 
         *        `ccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object 
         *        for each of the recipients on the **Cc** line. The array is limited to a maximum of 100 entries.
         * 
         *        `bccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object 
         *        for each of the recipients on the **Bcc** line. The array is limited to a maximum of 100 entries.
         * 
         *        `subject`: A string containing the subject of the message. The string is limited to a maximum of 255 characters.
         * 
         *        `htmlBody`: The HTML body of the message. The body content is limited to a maximum size of 32 KB.
         * 
         *        `attachments`: An array of JSON objects that are either file or item attachments.
         * 
         *        `attachments.type`: Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         * 
         *        `attachments.name`: A string that contains the name of the attachment, up to 255 characters in length.
         * 
         *        `attachments.url`: Only used if type is set to file. The URI of the location for the file. **Important**: This link must be
         *        publicly accessible, without need for authentication by Exchange Online servers. However, with on-premises Exchange, the link can
         *        be accessible on a private network as long as it doesn't need further authentication.
         * 
         *        `attachments.isInline`: Only used if type is set to file. If true, indicates that the attachment will be shown inline in the 
         *        message body, and should not be displayed in the attachment list.
         * 
         *        `attachments.itemId`: Only used if type is set to item. The EWS item id of the existing e-mail you want to attach to the new message. 
         *        This is a string up to 100 characters.
         */
        displayNewMessageForm(parameters: any): void;
        
        
        /**
         * Gets a string that contains a token used to call REST APIs or Exchange Web Services (EWS).
         *
         * The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. 
         * The lifetime of the callback token is 5 minutes.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of `ReadItem`.
         *
         * Calling the `getCallbackTokenAsync` method in compose mode requires you to have saved the item.
         * The `saveAsync` method requires a minimum permission level of `ReadWriteItem`.
         *
         * **Important**: For guidance on delegate or shared scenarios, see the
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/delegate-access | delegate access} article.
         *
         * *REST Tokens*
         *
         * When a REST token is requested (`options.isRest` = `true`), the resulting token will not work to authenticate EWS calls.
         * The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the
         * `ReadWriteMailbox` permission in its manifest.
         * If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts,
         * including the ability to send mail.
         *
         * The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.
         *
         * This API works for the following scopes:
         * 
         * - `Mail.ReadWrite`
         * 
         * - `Mail.Send`
         * 
         * - `Calendars.ReadWrite`
         * 
         * - `Contacts.ReadWrite`
         *
         * *EWS Tokens*
         *
         * When an EWS token is requested (`options.isRest` = `false`), the resulting token will not work to authenticate REST API calls.
         * The token will be limited in scope to accessing the current item.
         *
         * The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.
         *
         * You can pass both the token and either an attachment identifier or item identifier to an external system. That system uses
         * the token as a bearer authorization token to call the Exchange Web Services (EWS)
         * {@link https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation | GetAttachment} operation or
         * {@link https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation | GetItem} operation to return an
         * attachment or item. For example, you can create a remote service to
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * **Note**: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * **Errors**:
         * 
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         * 
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         * 
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `isRest`: Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.
         *        `asyncContext`: Any state data that is passed to the asynchronous method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         */
        getCallbackTokenAsync(options: CommonAPI.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets a string that contains a token used to get an attachment or item from an Exchange Server.
         *
         * The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. 
         * The lifetime of the callback token is 5 minutes.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * You can pass both the token and either an attachment identifier or item identifier to an external system. That system uses
         * the token as a bearer authorization token to call the Exchange Web Services (EWS)
         * {@link https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation | GetAttachment} operation or
         * {@link https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation | GetItem} operation to return an
         * attachment or item. For example, you can create a remote service to
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of `ReadItem`.
         *
         * Calling the `getCallbackTokenAsync` method in compose mode requires you to have saved the item.
         * The `saveAsync` method requires a minimum permission level of `ReadWriteItem`.
         *
         * **Important**: For guidance on delegate or shared scenarios, see the
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/delegate-access | delegate access} article.
         *
         * @remarks
         * [Api set: All support Read mode; Mailbox 1.3 introduced Compose mode support]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         * 
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         * 
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         * 
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         * 
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getCallbackTokenAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Gets a token identifying the user and the Office Add-in.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `getUserIdentityTokenAsync` method returns a token that you can use to identify and 
         * {@link https://docs.microsoft.com/office/dev/add-ins/outlook/authentication | authenticate the add-in and user with an external system}.
         *
         * **Errors**:
         * 
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         * 
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         * 
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         * 
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         *                 The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getUserIdentityTokenAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user's mailbox.
         *
         * In these cases, add-ins should use REST APIs to access the user's mailbox instead.
         *
         * The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.
         *
         * You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.
         *
         * The XML request must specify UTF-8 encoding: `\<?xml version="1.0" encoding="utf-8"?\>`.
         *
         * Your add-in must have the `ReadWriteMailbox` permission to use the `makeEwsRequestAsync` method.
         * For information about using the `ReadWriteMailbox` permission and the EWS operations that you can call with the `makeEwsRequestAsync` method,
         * see {@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Specify permissions for mail add-in access to the user's mailbox}.
         *
         * The XML result of the EWS call is provided as a string in the `asyncResult.value` property. 
         * If the result exceeds 1 MB in size, an error message is returned instead.
         *
         * **Note**: This method is not supported in the following scenarios:
         * 
         * - In Outlook on iOS or Android.
         * 
         * - When the add-in is loaded in a Gmail mailbox.
         *
         * **Note**: The server administrator must set `OAuthAuthentication` to `true` on the Client Access Server EWS directory to enable the 
         * `makeEwsRequestAsync` method to make EWS requests.
         *
         * *Version differences*
         *
         * When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set 
         * the encoding value to ISO-8859-1.
         *
         * `<?xml version="1.0" encoding="iso-8859-1"?>`
         *
         * You do not need to set the encoding value when your mail app is running in Outlook on the web.
         * You can determine whether your mail app is running in Outlook or Outlook on the web by using the `mailbox.diagnostics.hostName` property.
         * You can determine what version of Outlook is running by using the `mailbox.diagnostics.hostVersion` property.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteMailbox`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param data - The EWS request.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         *                 The `value` property of the result is the XML of the EWS request provided as a string. 
         *                 If the result exceeds 1 MB in size, an error message is returned instead.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        makeEwsRequestAsync(data: any, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    
    /**
     * Represents a suggested meeting found in an item. Read mode only.
     *
     * The list of meetings suggested in an email message is returned in the `meetingSuggestions` property of the `Entities` object that is returned when
     * the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * The start and end values are string representations of a `Date` object that contains the date and time at which the suggested meeting is to
     * begin and end. The values are in the default time zone specified for the current user.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface MeetingSuggestion {
        /**
         * Gets the attendees for a suggested meeting.
         */
        attendees: EmailUser[];
        /**
         * Gets the date and time that a suggested meeting is to end.
         */
        end: string;
        /**
         * Gets the location of a suggested meeting.
         */
        location: string;
        /**
         * Gets a string that was identified as a meeting suggestion.
         */
        meetingString: string;
        /**
         * Gets the date and time that a suggested meeting is to begin.
         */
        start: string;
        /**
         * Gets the subject of a suggested meeting.
         */
        subject: string;
    }
    /**
     * A subclass of {@link Office.Item | Item} for messages.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Child interfaces:
     *
     * - {@link Office.MessageCompose | MessageCompose}
     * 
     * - {@link Office.MessageRead | MessageRead}
     */
    export interface Message extends Item {
    }
     /**
     * The message compose mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Parent interfaces:
     *
     * - {@link Office.ItemCompose | ItemCompose}
     * 
     * - {@link Office.Message | Message}
     */
    export interface MessageCompose extends Message, ItemCompose {
        /**
         * Gets an object that provides methods to get or update the recipients on the **Bcc** (blind carbon copy) line of a message.
         *
         * Depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients you can get or update.
         * See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        bcc: Recipients;
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        body: Body;
        
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depend on the mode of the 
         * current item.
         *
         * The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the
         * **Cc** line of the message. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients
         * you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        cc: Recipients;
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. 
         * If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change 
         * and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form. 
         * If the user sets a subject and saves the item, the `conversationId` property will return a value.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        conversationId: string;
        
        
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or
         * an appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        notificationMessages: NotificationMessages;
        
        
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        subject: Subject;
        /**
         * Provides access to the recipients on the **To** line of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the
         * **To** line of the message. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients
         * you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        to: Recipients;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         * 
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the 
         *        attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of 
         *                 the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: CommonAPI.AsyncContextOptions & { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         * 
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of 
         *                 the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        
         
        
        
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. 
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or 
         * a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the 
         * callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that 
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of 
         *                 the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. 
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or 
         * a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the 
         * callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that 
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         * 
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property. 
         *                 If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of 
         *                 the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed. 
         * If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         *
         * **Note**: In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, 
         * discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        close(): void;
        
        
        
        
        
        
        
        
        
        
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. 
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present. 
         *                     If `Html`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. 
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present. 
         *                     If `Html`, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        
        
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. 
         * This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the 
         * current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the asyncResult.value property. 
         * This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to 
         * the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. 
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment 
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session. 
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to 
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing the attachment fails, the `asyncResult.error` property will contain an error code
         *                 with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. 
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment 
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session. 
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to 
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing the attachment fails, the `asyncResult.error` property will contain an error code
         *                 with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        
        
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server.
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the itemId will return an error.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. 
         */
        saveAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server.
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the `itemId` will return an error.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. 
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is 
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. 
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. 
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web and desktop clients. 
         *        If the field is an HTML editor, only the text data is inserted, even if the data is HTML. 
         *        If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is
         *        applied in Outlook on desktop clients. If the field is a text field, an `InvalidDataFormat` error is returned.
         *        If `coercionType` is not set, the result depends on the field:
         *        if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
         setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
         /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is 
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. 
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         * 
         * **Errors**:
         * 
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         * 
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. 
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The message read mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces. 
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://docs.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.6/office.context.mailbox.item | Object Model} page.
     * 
     * Parent interfaces:
     *
     * - {@link Office.ItemRead | ItemRead}
     * 
     * - {@link Office.Message | Message}
     */
    export interface MessageRead extends Message, ItemRead {
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         * 
         * **Note**: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. 
         * For more information, see 
         * {@link https://support.microsoft.com/office/434752e1-02d3-4e90-9124-8b81e49a8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        body: Body;
        
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depend on the mode of the 
         * current item.
         *
         * The `cc` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **Cc** line of the message. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - Web browser: 20 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        cc: EmailAddressDetails[];
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms.
         * If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change
         * and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form.
         * If the user sets a subject and saves the item, the `conversationId` property will return a value.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        conversationId: string;
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        dateTimeModified: Date;
        /**
         * Gets the date and time that the appointment is to end.
         *
         * The `end` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value. 
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on 
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        end: Date;
        /**
         * Gets the email address of the sender of a message.
         *
         * The `from` and `sender` properties represent the same person unless the message is sent by a delegate.
         * In that case, the `from` property represents the delegator, and the `sender` property represents the delegate.
         *
         * **Note**: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is undefined.
         * 
         * The `from` property returns an `EmailAddressDetails` object.
         * 
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        from: EmailAddressDetails;
        /**
         * Gets the internet message identifier for an email message.
         *
         * **Important**: In the **Sent Items** folder, the `internetMessageId` may not be available yet on recently sent items. In that case,
         * consider using {@link https://docs.microsoft.com/office/dev/add-ins/outlook/web-services | Exchange Web Services} to get this
         * {@link https://docs.microsoft.com/exchange/client-developer/web-service-reference/internetmessageid | property from the server}.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        internetMessageId: string;
        /**
         * Gets the Exchange Web Services item class of the selected item.
         * 
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class
         * `IPM.Appointment.Contoso`.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * The `itemClass` property specifies the message class of the selected item.
         * The following are the default message classes for the message or appointment item.
         * 
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurrence.</td>
         *     <td>IPM.Appointment,IPM.Appointment.Occurrence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note,IPM.Schedule.Meeting.Request,IPM.Schedule.Meeting.Neg,IPM.Schedule.Meeting.Pos,IPM.Schedule.Meeting.Tent,IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         * 
         */
        itemClass: string;
        /**
         * Gets the {@link https://docs.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}
         * for the current item.
         *
         * The `itemId` property is not available in compose mode.
         * If an item identifier is required, the `saveAsync` method can be used to save the item to the store, which will return the item identifier
         * in the `asyncResult.value` parameter in the callback function.
         *
         * **Note**: The identifier returned by the `itemId` property is the same as the
         * {@link https://docs.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}.
         * The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api#get-the-item-id | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or 
         * an appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of a meeting request.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        location: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The `normalizedSubject` property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by
         * email programs. To get the subject of the item with the prefixes intact, use the `subject` property.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        notificationMessages: NotificationMessages;
        
        
        /**
         * Gets the email address of the sender of an email message.
         *
         * The `from` and `sender` properties represent the same person unless the message is sent by a delegate.
         * In that case, the `from` property represents the delegator, and the `sender` property represents the delegate.
         *
         * **Note**: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is undefined.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        sender: EmailAddressDetails;
        /**
         * Gets the date and time that the appointment is to begin.
         *
         * The `start` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        start: Date;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        subject: string;
        /**
         * Provides access to the recipients on the **To** line of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `to` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **To** line of the message. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - Web browser: 20 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        to: EmailAddressDetails[];

        
         
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        
        
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        
        
        
        
        
         
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         * 
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns `null`. 
         * If no entities of the specified type are present in the item's body, the method returns an empty array. 
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param entityType - One of the `EntityType` enumeration values.
         * 
         * While the minimum permission level to use this method is `Restricted`, some entity types require `ReadItem` to access, as specified in the 
         * following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Level</th>
         *   </tr>
         *   <tr>
         *     <td>Address</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>Contact</td>
         *     <td>Contact</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>EmailAddress</td>
         *     <td>String</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>MeetingSuggestion</td>
         *     <td>MeetingSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>PhoneNumber</td>
         *     <td>PhoneNumber</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>TaskSuggestion</td>
         *     <td>TaskSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>URL</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         * </table>
         */
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.
         *
         * The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the `ItemHasKnownEntity` rule element 
         * in the manifest XML file with the specified `FilterName` element value.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         * 
         * @returns If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, 
         * the method returns `null`. 
         * If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, 
         * the method return an empty array.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns string values in the selected item that match the regular expressions defined in the manifest XML file.
         *
         * The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter
         * the body and should not attempt to return the entire body of the item. 
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. 
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in the manifest XML file.
         *
         * The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the
         * `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter
         * the body and should not attempt to return the entire body of the item. 
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        /**
         * Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getSelectedEntities(): Entities;
        /**
         * Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. 
         * Highlighted matches apply to contextual add-ins.
         *
         * The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in
         * each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. 
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or 
         * the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getSelectedRegExMatches(): any;
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. 
         * This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
         *
         * @remarks
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. 
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        
        
    }
    
    /**
     * An array of `NotificationMessageDetails` objects are returned by the `NotificationMessages.getAllAsync` method.
     *
     * @remarks
     * [Api set: Mailbox 1.3]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface NotificationMessageDetails {
        /**
         * The identifier for the notification message.
         */
        key?: string;
        /**
         * Specifies the `ItemNotificationMessageType` of message.
         *
         * If type is `ProgressIndicator` or `ErrorMessage`, an icon is automatically supplied
         * and the message is not persistent. Therefore the icon and persistent properties are not valid for these types of messages.
         * Including them will result in an `ArgumentException`.
         *
         * If type is `ProgressIndicator`, the developer should remove or replace the progress indicator when the action is complete.
         */
        type: MailboxEnums.ItemNotificationMessageType | string;
        /**
         * A reference to an icon that is defined in the manifest in the `Resources` section. It appears in the infobar area.
         * It is only applicable if the type is `InformationalMessage`. Specifying this parameter for an unsupported type results in an exception.
         * 
         * **Note**: At present, the custom icon is displayed in Outlook on Windows only and not on other clients (e.g., Mac, web browser).
         */
        icon?: string;
        /**
         * The text of the notification message. Maximum length is 150 characters.
         * If the developer passes in a longer string, an `ArgumentOutOfRange` exception is thrown.
         */
        message: string;
        /**
         * Specifies if the message should be persistent. Only applicable when type is `InformationalMessage`.
         * If true, the message remains until removed by this add-in or dismissed by the user.
         * If false, it is removed when the user navigates to a different item.
         * For error notifications, the message persists until the user sees it once.
         * Specifying this parameter for an unsupported type throws an exception.
         */
        persistent?: Boolean;
        
    }
    /**
     * The `NotificationMessages` object is returned as the `notificationMessages` property of an item.
     *
     * @remarks
     * [Api set: Mailbox 1.3]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface NotificationMessages {
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a `NumberOfNotificationMessagesExceeded` error.
         *
         * **Important**:
         *
         * - Only one notification of type {@link https://docs.microsoft.com/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype#fields | InsightMessage}
         * is allowed per add-in. Attempting to add more will throw an error.
         *
         * - In modern Outlook on the web, you can add an `InsightMessage` notification only in Compose mode.
         * 
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - A developer-specified key used to reference this notification message.
         *             Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item.
         *                    It contains a `NotificationMessageDetails` object.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a `NumberOfNotificationMessagesExceeded` error.
         *
         * **Important**:
         *
         * - Only one notification of type {@link https://docs.microsoft.com/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype#fields | InsightMessage}
         * is allowed per add-in. Attempting to add more will throw an error.
         *
         * - In modern Outlook on the web, you can add an `InsightMessage` notification only in Compose mode.
         * 
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - A developer-specified key used to reference this notification message.
         *             Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item.
         *                    It contains a `NotificationMessageDetails` object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `NotificationMessageDetails` objects.
         */
        getAllAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<NotificationMessageDetails[]>) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `NotificationMessageDetails` objects.
         */
        getAllAsync(callback?: (asyncResult: CommonAPI.AsyncResult<NotificationMessageDetails[]>) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param key - The key for the notification message to remove.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        removeAsync(key: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param key - The key for the notification message to remove.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        removeAsync(key: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, `replaceAsync` will add the notification.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. 
         *                    It contains a `NotificationMessageDetails` object.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, `replaceAsync` will add the notification.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. 
         *                    It contains a `NotificationMessageDetails` object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    
    /**
     * Represents a phone number identified in an item. Read mode only.
     *
     * An array of `PhoneNumber` objects containing the phone numbers found in an email message is returned in the `phoneNumbers` property of the
     * `Entities` object that is returned when you call the `getEntities` method on the selected item.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface PhoneNumber {
        /**
         * Gets a string containing a phone number. This string contains only the digits of the telephone number and excludes characters
         * like parentheses and hyphens, if they exist in the original item.
         */
        phoneString: string;
        /**
         * Gets the text that was identified in an item as a phone number.
         */
        originalPhoneString: string;
        /**
         * Gets a string that identifies the type of phone number: Home, Work, Mobile, Unspecified.
         */
        type: string;
    }
    /**
     * Represents recipients of an item. Compose mode only.
     * 
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Recipients {
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be added:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         * 
         * @param recipients - The recipients to add to the recipients list.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
         addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
         /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be added:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         * 
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the `asyncResult.value` property will contain an array of {@link Office.EmailAddressDetails | EmailAddressDetails}
         * objects. Collection size limits:
         *
         * - Windows, classic Mac UI, web browser: 500 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Important**: In Outlook on the web and on Windows, if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * `EmailAddressDetails` object instead of the contact's saved name. For more details, refer to the
         * {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         * 
         * **Important**: The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
         * 
         * - If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
         * 
         * - A Teams meeting status icon appears before the recipient's name or email address.
         * 
         * - A semicolon appears after the recipient's name or email address.
         * 
         * - The recipient's name or email address is underlined or enclosed in a box.
         * 
         * To resolve an email address once it's added to a mail item, the sender must use the **Tab** key or select a suggested contact or email address from 
         * the auto-complete list.
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `EmailAddressDetails` objects.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the `asyncResult.value` property will contain an array of {@link Office.EmailAddressDetails | EmailAddressDetails}
         * objects. Collection size limits:
         *
         * - Windows, classic Mac UI, web browser: 500 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on the web and on Windows, if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * `EmailAddressDetails` object instead of the contact's saved name. For more details, refer to the
         * {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         * 
         * **Important**: The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
         * 
         * - If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
         * 
         * - A Teams meeting status icon appears before the recipient's name or email address.
         * 
         * - A semicolon appears after the recipient's name or email address.
         * 
         * - The recipient's name or email address is underlined or enclosed in a box.
         * 
         * To resolve an email address once it's added to a mail item, the sender must use the **Tab** key or select a suggested contact or email address from 
         * the auto-complete list.
         * 
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `EmailAddressDetails` objects.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be set:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         * 
         * @param recipients - The recipients to add to the recipients list.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be set:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of 
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
     
    
    
    
    
    
    /**
     * A file or item attachment. Used when displaying a reply form.
     */
    export interface ReplyFormAttachment {
        /**
         * Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         */
        type: string;
        /**
         * A string that contains the name of the attachment, up to 255 characters in length.
         */
        name: string;
        /**
         * Only used if type is set to file. The URI of the location for the file.
         *
         * **Important**: This link must be publicly accessible, without need for authentication by Exchange Online servers. However, with
         * on-premises Exchange, the link can be accessible on a private network as long as it doesn't need further authentication.
         */
        url?: string;
        /**
         * Only used if type is set to file. If true, indicates that the attachment will be shown inline in the message body, and should not be 
         * displayed in the attachment list.
         */
        inLine?: boolean;
        /**
         * Only used if type is set to item. The EWS item id of the attachment. This is a string up to 100 characters.
         */
        itemId?: string;
    }
    /**
     * A ReplyFormData object that contains body or attachment data and a callback function. Used when displaying a reply form.
     */
    export interface ReplyFormData {
        /**
         * A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
         */
        htmlBody?: string;
        /**
         * An array of {@link Office.ReplyFormAttachment | ReplyFormAttachment} that are either file or item attachments.
         */
        attachments?: ReplyFormAttachment[];
        /**
         * When the reply display call completes, the function passed in the callback parameter is called with a single parameter, 
         * `asyncResult`, which is an `Office.AsyncResult` object.
         */
        callback?: (asyncResult: CommonAPI.AsyncResult<any>) => void;
        /**
         * An object literal that contains the following property.
         * `asyncContext`: Developers can provide any object they wish to access in the callback function.
         */
        options?: CommonAPI.AsyncContextOptions;
    }
    /**
     * The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user.
     * That is, they are available only to the add-in that created them, and only from the user's mailbox in which they are saved.
     *
     * While the Outlook add-in API limits access to these settings to only the add-in that created them, these settings should not be considered
     * secure storage. They can be accessed by Exchange Web Services or Extended MAPI.
     * They should not be used to store sensitive information such as user credentials or security tokens.
     *
     * The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.
     *
     * The `RoamingSettings` object is accessible via the `roamingSettings` property in the `Office.context` namespace.
     *
     * **Important**:
     *
     * - The `RoamingSettings` object is initialized from the persisted storage only when the add-in is first loaded.
     * For task panes, this means that it is only initialized when the task pane first opens.
     * If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values, even if
     * your add-in has persisted changes.
     * The persisted changes will not be available until the task pane (or item in the case of UI-less add-ins) is closed and reopened.
     *
     * - When set and saved through Outlook on Windows or Mac, these settings are reflected in Outlook on the web only after a browser refresh.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface RoamingSettings {
        /**
         * Retrieves the specified setting.
         *
         * @returns Type: String | Number | Boolean | Object | Array
         * 
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to retrieve.
         */
        get(name: string): any;
        /**
         * Removes the specified setting
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to remove.
         */
        remove(name: string): void;
        /**
         * Saves the settings.
         *
         * Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use 
         * the set and get methods to work with the in-memory copy of the settings property bag. 
         * When you want to persist the settings so that they are available the next time the add-in is used, use the saveAsync method.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        saveAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets or creates the specified setting.
         *
         * The `set` method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name. 
         * The value is stored in the document as the serialized JSON representation of its data type.
         *
         * A maximum of 32KB is available for the settings of each add-in.
         *
         * Any changes made to settings using the `set` method will not be saved to the server until the `saveAsync` method is called.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `Restricted`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to set or create.
         * @param value - Specifies the value to be stored.
         */
        set(name: string, value: any): void;
    }
    
    
    
    /**
     * Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Subject {
        /**
         * Gets the subject of an appointment or message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is the subject of the item.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the subject of an appointment or message.
         * 
         * The getAsync method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is the subject of the item.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the subject of an appointment or message.
         * Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The subject parameter is longer than 255 characters.
         * 
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If setting the subject fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(subject: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the subject of an appointment or message.
         * Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The subject parameter is longer than 255 characters.
         * 
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If setting the subject fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(subject: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents a suggested task identified in an item. Read mode only.
     *
     * The list of tasks suggested in an email message is returned in the `taskSuggestions` property of the {@link Office.Entities | Entities} object
     * that is returned when the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface TaskSuggestion {
        /**
         * Gets the users that should be assigned a suggested task.
         */
        assignees: EmailUser[];
        /**
         * Gets the text of an item that was identified as a task suggestion.
         */
        taskString: string;
    }
    /**
     * The `Time` object is returned as the start or end property of an appointment in compose mode.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Time {
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a `Date` object in the `asyncResult.value` property. The value is in Coordinated Universal Time (UTC).
         * You can convert the UTC time to the local client time by using the `convertToLocalClientTime` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                  of type `Office.AsyncResult`. The `value` property of the result is a `Date` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a `Date` object in the `asyncResult.value` property. The value is in Coordinated Universal Time (UTC).
         * You can convert the UTC time to the local client time by using the `convertToLocalClientTime` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter
         *                  of type `Office.AsyncResult`. The `value` property of the result is a `Date` object.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the `setAsync` method is called on the start property, the `end` property will be adjusted to maintain the duration of the appointment as
         * previously set. If the `setAsync` method is called on the `end` property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the `convertToUtcClientTime` method.
         *
         * **Important**: In the Windows client, you can't use this method to update the start or end of a recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `InvalidEndTime`: The appointment end time is before the appointment start time.
         * 
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param options - An object literal that contains one or more of the following properties.
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *               type `Office.AsyncResult`. If setting the date and time fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(dateTime: Date, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the `setAsync` method is called on the start property, the `end` property will be adjusted to maintain the duration of the appointment as
         * previously set. If the `setAsync` method is called on the `end` property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the `convertToUtcClientTime` method.
         *
         * **Important**: In the Windows client, you can't use this method to update the start or end of a recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadWriteItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Errors**:
         * 
         * - `InvalidEndTime`: The appointment end time is before the appointment start time.
         * 
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *               type `Office.AsyncResult`. If setting the date and time fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(dateTime: Date, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Information about the user associated with the mailbox. This includes their account type, display name, email address, and time zone.
     *
     * @remarks
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
     * 
     * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface UserProfile {
        /**
         * Gets the account type of the user associated with the mailbox. 
         *
         * **Note**: This member is currently only supported in Outlook 2016 or later on Mac, build 16.9.1212 and greater.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The possible account types are listed in the following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>enterprise</td>
         *     <td>The mailbox is on an on-premises Exchange server.</td>
         *   </tr>
         *   <tr>
         *     <td>gmail</td>
         *     <td>The mailbox is associated with a Gmail account.</td>
         *   </tr>
         *   <tr>
         *     <td>office365</td>
         *     <td>The mailbox is associated with a Microsoft 365 work or school account.</td>
         *   </tr>
         *   <tr>
         *     <td>outlookCom</td>
         *     <td>The mailbox is associated with a personal Outlook.com account.</td>
         *   </tr>
         * </table>
         * 
         * **Note**: For hybrid Exchange environments, the returned account type value depends on where the mailbox is hosted.
         * If the mailbox is on an on-premises server, the account type value is **enterprise**. However, if it's hosted on 
         * Exchange Online, the account type value is **office365**.
         */
        accountType: string;
        /**
         * Gets the user's display name.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        displayName: string;
        /**
         * Gets the user's SMTP email address.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        emailAddress: string;
        /**
         * Gets the user's time zone in Windows format.
         *
         * The system time zone is usually returned. However, in Outlook on the web, the default time zone in the calendar preferences is returned instead.
         *
         * @remarks
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: `ReadItem`
         * 
         * **{@link https://docs.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        timeZone: string;
    }
}


////////////////////////////////////////////////////////////////
/////////////////////// End Exchange APIs //////////////////////
////////////////////////////////////////////////////////////////