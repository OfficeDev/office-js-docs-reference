import {Office as CommonAPI} from "../../api-extractor-inputs-office/office"
////////////////////////////////////////////////////////////////
////////////////////// Begin Exchange APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export namespace MailboxEnums {
        
        
        
        
        /**
         * Specifies the attachment's type.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
             * {@link https://learn.microsoft.com/javascript/api/outlook/office.attachmentdetailscompose?view=outlook-js-1.8 | details} object
             * contains a URL to the file in Compose mode.
             */
            Cloud = "cloud"
        }
        
        
        
        
        /**
         * Specifies an entity's type.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Specifies an item's type.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Specifies the location from which an add-in wants to access data.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         *
         * **Important**: This enum is only supported in Outlook on Android and on iOS starting in Version 4.2443.0. To learn more about APIs supported in Outlook on mobile devices, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         */
        enum OpenLocation {
            /**
             * A location associated with an account within an add-in.
             */
            AccountDocument,
            /**
             * The device's camera.
             */
            Camera,
            /**
             * Local storage on a device.
             */
            Local,
            /**
             * OneDrive for Business.
             *
             * **Important**: For OneDrive Personal, use OTHER.
             */
            OnedriveForBusiness,
            /**
             * Other cloud storage providers, including OneDrive Personal.
             */
            Other,
            /**
             * The device's photo library.
             */
            PhotoLibrary,
            /**
             * SharePoint. Includes both SharePoint Online and SharePoint on-premises (if accessed with a Microsoft Entra ID account).
             */
            SharePoint
        }
        /**
         * Represents the current view of Outlook on the web.
         */
        enum OWAView {
            /**
             * Narrow one-column view. Displayed when the screen width is less than 436 pixels.
             * For example, Outlook on the web uses this view on the entire screen of older smartphones.
             */
            OneColumnNarrow = "OneColumnNarrow",
            /**
             * One-column view. Displayed when the screen width is greater than or equal to 436 pixels,
             * but less than 536 pixels. For example, Outlook on the web uses this view on the entire screen of newer smartphones.
             */
            OneColumn = "OneColumn",
            /**
             * Two-column view. Displayed when the screen width is greater than or equal to 536 pixels,
             * but less than 780 pixels. For example, Outlook on the web uses this view on most tablets.
             */
            TwoColumns = "TwoColumns",
            /**
             * Three-column view. Displayed when the screen width is greater than or equal to 780 pixels.
             * For example, Outlook on the web uses this view in a full screen window on a desktop computer.
             */
            ThreeColumns = "ThreeColumns"
        }
        /**
         * Specifies the type of recipient of a message or appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * **Important**: A `recipientType` property value isn't returned by the 
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-1.7#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync} 
         * and {@link https://learn.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-1.7#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         */
        enum RecipientType {
            /**
             * Specifies the recipient is a distribution list containing a list of email addresses.
             */
            DistributionList = "distributionList",
            /**
             * Specifies the recipient is an SMTP email address on the Exchange server.
             */
            User = "user",
            /**
             * Specifies the recipient is an SMTP email address that isn't on the Exchange server. It also refers to a recipient added from a personal Outlook address book.
             * 
             * **Important**: In Outlook on the web, on Windows ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic
             * (starting with Version 2210, Build 15813.20002)), and on Mac, Global Address Book (GAL) recipients saved to a personal address book return the `ExternalUser` value,
             * even if their SMTP email address appears on the Exchange server. Recipients return a `User` value only if they're directly added or resolved against the GAL.
             */
            ExternalUser = "externalUser",
            /**
             * Specifies the recipient isn't one of the other recipient types. It also refers to a recipient that isn't resolved against the Exchange address book,
             * and is therefore treated as an external SMTP address.
             *
             * **Important**: In Outlook on Android and on iOS, Global Address Book (GAL) recipients saved to a personal address book return
             * the `Other` value, even if their SMTP email address appears on the Exchange server. Recipients return a `User` value only if they're directly
             * added or resolved against the GAL.
             */
            Other = "other"
        }
        
        
        /**
         * Specifies the type of response to a meeting invitation.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Specifies the location in which an add-in wants to save data.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         *
         * **Important**: This enum is only supported in Outlook on Android and on iOS starting in Version 4.2443.0. To learn more about APIs supported in Outlook on mobile devices, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         */
        enum SaveLocation {
            /**
             * A location associated with an account within an add-in.
             */
            AccountDocument,
            /**
             * Box.
             */
            Box,
            /**
             * Dropbox.
             */
            Dropbox,
            /**
             * Google Drive.
             */
            GoogleDrive,
            /**
             * Local storage on a device.
             */
            Local,
            /**
             * OneDrive for Business.
             *
             * **Important**: For OneDrive Personal, use OTHER.
             */
            OnedriveForBusiness,
            /**
             * Other cloud storage providers, including OneDrive Personal.
             */
            Other,
            /**
             * The device's photo library.
             */
            PhotoLibrary,
            /**
             * SharePoint. Includes both SharePoint Online and SharePoint on-premises (if accessed with a Microsoft Entra ID account).
             */
            SharePoint
        }
        
        /**
         * Specifies the source of the selected data in an item (see `Office.mailbox.item.getSelectedDataAsync` for details).
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        end: Time;
        
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets or sets the location of an appointment. The `location` property returns a {@link Office.Location | Location} object that provides methods that are
         * used to get and set the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        location: Location;
        
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
         * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
         * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        subject: Subject;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * @remarks
         * [Api set: Mailbox 1.1 for Outlook on Windows (classic) and on Mac, Mailbox 1.8 for Outlook on the web and new Outlook on Windows]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Important**:
         *
         * - This method isn't supported in Outlook on iOS or Android. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - In recent builds of classic Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, use the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * - The URI of the file to be attached must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header that
         * specifies `no-cache`, `no-store`, or similar options in the HTTP response. However, when you're developing the add-in and making changes to files,
         * caching can prevent you from seeing your changes. We recommend using `Cache-Control` headers during development.
         *
         * - You can use the same URI with the `removeAttachmentAsync` method to remove the attachment in the same session.
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
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline as an image in the message body and won't be displayed in the attachment list.
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
         * @remarks
         * [Api set: Mailbox 1.1 for Outlook on Windows (classic) and on Mac, Mailbox 1.8 for Outlook on the web and new Outlook on Windows]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Important**:
         *
         * - This method isn't supported in Outlook on iOS or Android. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - In recent builds of classic Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, use the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * - The URI of the file to be attached must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header that
         * specifies `no-cache`, `no-store`, or similar options in the HTTP response. However, when you're developing the add-in and making changes to files,
         * caching can prevent you from seeing your changes. We recommend using `Cache-Control` headers during development.
         *
         * - You can use the same URI with the `removeAttachmentAsync` method to remove the attachment in the same session.
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
         * If your Office Add-in is running in Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the `addItemAttachmentAsync` method can attach items to items other than the item that you're editing. However, this isn't supported and isn't recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * If your Office Add-in is running in Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the `addItemAttachmentAsync` method can attach items to items other than the item that you're editing. However, this isn't supported and isn't recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        
        
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
         * in the same session. In Outlook on the web, on mobile devices, and in {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * *Important**: The `removeAttachmentAsync` method doesn't remove inline attachments from a mail item.
         * To remove an inline attachment, first get the item's body, then remove any references of the attachment from its contents.
         * Use the {@link https://learn.microsoft.com/javascript/api/outlook/office.body | Office.Body} APIs to get and set the body of an item.
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows (new and classic).
         * @param options - An object literal that contains one or more of the following properties:-
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
         * in the same session. In Outlook on the web, on mobile devices, and in {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * *Important**: The `removeAttachmentAsync` method doesn't remove inline attachments from a mail item.
         * To remove an inline attachment, first get the item's body, then remove any references of the attachment from its contents.
         * Use the {@link https://learn.microsoft.com/javascript/api/outlook/office.body | Office.Body} APIs to get and set the body of an item.
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows (new and classic).
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        
        
        
        
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *        If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
          * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web, on Windows
         *        ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), and on Mac.
         *        If the field is an HTML editor, only the text data is inserted, even if the data is HTML.
         *        If the data is HTML and the field supports HTML (the subject doesn't), the current style is applied in
         *        Outlook on the web and new Outlook on Windows. The default style is applied in Outlook on Windows (classic) and on Mac.
         *        If the field is a text field, an `InvalidDataFormat` error is returned.
         *        If `coercionType` is not set, the result depends on the field:
         *        if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *        type `Office.AsyncResult`.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface AppointmentForm {
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        * - Web browser, new Mac UI, Android: No limit
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * *Compose mode*
        *
        * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
        * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       optionalAttendees: Recipients[] | EmailAddressDetails[];
       /**
        * Provides access to the resources of an event. Returns an array of strings containing the resources required for the appointment.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        * - Web browser, new Mac UI, Android: No limit
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * *Compose mode*
        *
        * The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the
        * required attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       subject: Subject | string;
    }
    /**
     * The appointment attendee mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Note**: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        body: Body;
        
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**: This property isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        end: Date;
        
        /**
         * Gets the Exchange Web Services item class of the selected appointment.
         *
         * Returns `IPM.Appointment` for non-recurring appointments and `IPM.Appointment.Occurrence` for recurring appointments.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**: You can create custom classes that extend a default item class. For example, `IPM.Appointment.Contoso`.
         */
        itemClass: string;
        /**
         * Gets the {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services (EWS) item identifier}
         * of the current item.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - The `itemId` property isn't available in compose mode.
         * If an item identifier is required, the `Office.context.mailbox.item.saveAsync` method can be used to save the item to the store, which will return the item identifier
         * in the `asyncResult.value` parameter in the callback function. If the item is already saved, you can call the `Office.context.mailbox.item.getItemIdAsync` method instead.
         *
         * - The item ID returned isn't identical to the Outlook Entry ID or the ID used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of an appointment.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        normalizedSubject: string;
        
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each optional attendee to the meeting. The maximum number of attendees returned varies per Outlook client.
         *
         * - Windows: 500 attendees
         *
         * - Android, classic Mac UI, iOS: 100 attendees
         *
         * - New Mac UI, web browser: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        optionalAttendees: EmailAddressDetails[];
        /**
         * Gets the meeting organizer's email properties.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        organizer: EmailAddressDetails;
        
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `requiredAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each required attendee to the meeting. The maximum number of attendees returned varies per Outlook client.
         *
         * - Windows: 500 attendees
         *
         * - Android, classic Mac UI, iOS: 100 attendees
         *
         * - New Mac UI, web browser: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**: In Outlook on the web and on Windows (new and classic), the appointment organizer is included in the object returned by the `requiredAttendees` property.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        subject: string;
        
        
        
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * - If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * - When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        
        
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * - If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * - When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        
        
        
        
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.
         * If no entities of the specified type are present in the item's body, the method returns an empty array.
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param entityType - One of the `EntityType` enumeration values.
         */
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): Array<string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion>;
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in an add-in only manifest file.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @returns
         * The entities that match the regular expression defined in the `ItemHasKnownEntity` rule element in the
         * add-in manifest file with the specified `FilterName` element value. If there's no `ItemHasKnownEntity` element in the manifest with a
         * `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter matches an
         * `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method returns an empty array.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): Array<string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion>;
        
        
        /**
         * Returns string values in the selected item that match the regular expressions defined in an add-in only manifest file.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the add-in manifest file.
         * The name of each array is equal to the corresponding value of the RegExName attribute of the matching `ItemHasRegularExpressionMatch` rule.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property
         * of the item that's specified by that rule. The `PropertyName` simple type defines the supported properties.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * - This method is used with the {@link https://learn.microsoft.com/javascript/api/manifest/rule | activation rules feature for Outlook add-ins},
         * which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | unified manifest for Microsoft 365}.
         *
         * - If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item doesn't always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in an add-in only manifest file.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the add-in manifest file,
         * with the specified `RegExName` element value.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * - This method is used with the {@link https://learn.microsoft.com/javascript/api/manifest/rule | activation rules feature for Outlook add-ins},
         * which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | unified manifest for Microsoft 365}.
         *
         * - If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item doesn't always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        
        
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
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
     * An array of `AttachmentDetails` objects is returned as the `attachments` property of an appointment or message item.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface AttachmentDetails {
        /**
         * Gets a value that indicates the attachment's type.
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
         * Gets a value that indicates whether the attachment appears as an image in the body of the item instead of in the attachment list.
         */
        isInline: boolean;
        /**
         * Gets the name of the attachment.
         *
         * **Important**: For message or appointment items that were attached by drag-and-drop or "Attach Item",
         * `name` includes a file extension in Outlook on Mac, but excludes the extension on the web or on Windows.
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Body {
        
        
        
        
        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  The content type is returned as one of the `CoercionType` values in the `asyncResult.value` property.
         */
        getTypeAsync(callback?: (asyncResult: CommonAPI.AsyncResult<CommonAPI.CoercionType>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Recommended**: Call `getTypeAsync`, then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**:
         *
         * - After the content is prepended, the position of the cursor depends on which client the add-in is running. In Outlook on the web and on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), the cursor position remains the same in the preexisting content of the body.
         * For example, if the cursor was positioned at the beginning of the body prior to the `prependAsync` call, it will appear between the prepended content and the preexisting
         * content of the body after the call. In Outlook on Mac, the cursor position isn't preserved. The cursor disappears after the `prependAsync` call and only reappears when the
         * user selects something in the body of the mail item.
         *
         * - When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) won't necessarily contain the exact value that was passed in the previous `prependAsync` call.
         *
         * - When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * - In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - SVG files aren't supported. Use JPG or PNG files instead.
         *
         * - The `prependAsync` method doesn't support inline CSS. Use internal or external CSS instead.
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        prependAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Recommended**: Call `getTypeAsync`, then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**:
         *
         * - After the content is prepended, the position of the cursor depends on which client the add-in is running. In Outlook on the web and on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), the cursor position remains the same in the preexisting content of the body.
         * For example, if the cursor was positioned at the beginning of the body prior to the `prependAsync` call, it will appear between the prepended content and the preexisting
         * content of the body after the call. In Outlook on Mac, the cursor position isn't preserved. The cursor disappears after the `prependAsync` call and only reappears when the
         * user selects something in the body of the mail item.
         *
         * - When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) won't necessarily contain the exact value that was passed in the previous `prependAsync` call.
         *
         * - When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * - In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - SVG files aren't supported. Use JPG or PNG files instead.
         *
         * - The `prependAsync` method doesn't support inline CSS. Use internal or external CSS instead.
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
         * Replaces the selection in the body with the specified text.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in
         * the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the
         * UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important*:
         *
         * - When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * - SVG files aren't supported. Use JPG or PNG files instead.
         *
         * - The `setSelectedDataAsync` method doesn't support inline CSS. Use internal or external CSS instead.
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 1,000,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important*:
         *
         * - When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * - SVG files aren't supported. Use JPG or PNG files instead.
         *
         * - The `setSelectedDataAsync` method doesn't support inline CSS. Use internal or external CSS instead.
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
     * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
     * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
     * For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     *
     * **Important**: Entity-based contextual Outlook add-ins will be retired in Q2 of 2024. The work to retire this feature will start in May and continue
     * until the end of June. After June, contextual add-ins will no longer be able to detect entities in mail items to perform tasks on them.
     * The following APIs will also be retired.
     *
     * - `Office.context.mailbox.item.getEntities`
     * - `Office.context.mailbox.item.getEntitiesByType`
     * - `Office.context.mailbox.item.getFilteredEntitiesByName`
     * - `Office.context.mailbox.item.getSelectedEntities`
     *
     * To help minimize potential disruptions, the following will still be supported after entity-based contextual add-ins are retired.
     *
     * - An alternative implementation of the **Join Meeting** button, which is activated by online meeting add-ins, is being developed. Once support for
     * entity-based contextual add-ins ends, online meeting add-ins will automatically transition to the alternative implementation to activate the
     * **Join Meeting** button.
     *
     * - Regular expression rules will continue to be supported after entity-based contextual add-ins are retired. We recommend updating your contextual add-in
     * to use regular expression rules as an alternative solution. For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * For more information, see
     * {@link https://devblogs.microsoft.com/microsoft365dev/retirement-of-entity-based-contextual-outlook-add-ins | Retirement of entity-based contextual Outlook add-ins}.
     */
    export interface Contact {
        /**
         * An array of strings containing the mailing and street addresses associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        addresses: string[];
        /**
         * A string containing the name of the business associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        businessName: string;
        /**
         * An array of strings containing the SMTP email addresses associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        emailAddresses: string[];
        /**
         * A string containing the name of the person associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        personName: string;
        /**
         * An array containing a `PhoneNumber` object for each phone number associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * An array of strings containing the Internet URLs associated with the contact. Nullable.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        urls: string[];
    }
    /**
     * The `CustomProperties` object represents custom properties that are specific to a particular mail item and specific to an Outlook add-in.
     * For example, there might be a need for an add-in to save some data that's specific to the current message that activated the add-in.
     * If the user revisits the same message in the future and activates the add-in again, the add-in will be able to retrieve the data that had
     * been saved as custom properties. 
     *
     * To learn more about `CustomProperties`, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * When using custom properties in your add-in, keep in mind that:
     * 
     * - Custom properties saved while in compose mode aren't transmitted to recipients of the mail item. When a message or appointment with custom
     * properties is sent, its properties can be accessed from the item in the Sent Items folder.
     * If you want to make custom data accessible to recipients, consider using
     * {@link https://learn.microsoft.com/javascript/api/outlook/office.internetheaders | InternetHeaders} instead.
     * 
     * - The maximum length of a `CustomProperties` JSON object is 2500 characters.
     *
     * - Outlook on Mac doesn't cache custom properties. If the user's network goes down, mail add-ins can't access their custom properties.
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface CustomProperties {
        /**
         * Returns the value of the specified custom property.
         *
         * @returns The value of the specified custom property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The `name` of the property to be removed.
         */
        remove(name: string): void;
        /**
         * Saves custom properties to a message or appointment.
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
         * [Api set: Mailbox 1.1]
         * 
         * **Important**: In Outlook on Windows, custom properties saved while in compose mode only persist after the item being composed is closed or
         * after `Office.context.mailbox.item.saveAsync` is called.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<void>) => void, asyncContext?: any): void;
        /**
         * Saves custom properties to a message or appointment.
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
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
        saveAsync(asyncContext?: any): void;
        /**
         * Sets the specified property to the specified value.
         *
         * The `set` method sets the specified property to the specified value. To ensure that the set property and value persist on the mail item,
         * you must call the `saveAsync` method.
         *
         * The `set` method creates a new property if the specified property does not already exist;
         * otherwise, the existing value is replaced with the new value.
         * The `value` parameter can be of any type; however, it is always passed to the server as a string.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     * 
     * Starting with Mailbox requirement set 1.5, you can also use the 
     * {@link https://learn.microsoft.com/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics}
     * property to get similar information.
     */
    export interface Diagnostics {
        /**
         * Gets a string that represents the type of Outlook client.
         *
         * The string can be one of the following values: `Outlook`, `newOutlookWindows`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         *
         * @remarks
         *
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**: The `Outlook` value is returned for Outlook on Windows (classic) and on Mac. `newOutlookWindows` is returned for
         * {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows}.
         */
        hostName: string;
        /**
         * Gets a string that represents the version of either the Outlook client or the Exchange Server (for example, "15.0.468.0").
         *
         * If the mail add-in is running in Outlook on Windows (classic), on Mac, or on mobile devices, the `hostVersion` property returns the version of the
         * Outlook client. In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the property returns the version of the Exchange Server.
         *
         * @remarks
         *
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        hostVersion: string;
        /**
         * Gets a string that represents the current view of Outlook on the web.
         *
         * The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.
         *
         * If the application is not Outlook on the web, then accessing this property results in undefined.
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
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        OWAView: MailboxEnums.OWAView | "OneColumn" | "TwoColumns" | "ThreeColumns";
    }
    /**
     * Provides the email properties of the sender or specified recipients of an email message or appointment.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * **Important**:
         *
         * - A `recipientType` property value isn't returned by the 
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-1.7#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync}
         * and {@link https://learn.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-1.7#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         *
         * - While composing a mail item, when you switch to a sender account that's on a different domain than that of the previously selected sender account,
         * the value of the `recipientType` property for existing recipients isn't updated and will still be based on the domain of the previously selected account.
         * To get the correct recipient types after switching accounts, you must first remove the existing recipients, then add them back to the mail item.
         */
        recipientType: MailboxEnums.RecipientType | string;
    }
    /**
     * Represents an email account on an Exchange Server.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
     * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
     * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
     * For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface Entities {
        /**
         * Gets the physical addresses (street or mailing addresses) found in an email message or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        addresses: string[];
        /**
         * Gets the contacts found in an email address or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        contacts: Contact[];
        /**
         * Gets the email addresses found in an email message or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        emailAddresses: string[];
        /**
         * Gets the meeting suggestions found in an email message.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        meetingSuggestions: MeetingSuggestion[];
        /**
         * Gets the phone numbers found in an email message or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * Gets the task suggestions found in an email message or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        taskSuggestions: string[];
        /**
         * Gets the Internet URLs present in an email message or appointment.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        urls: string[];
    }
    
    
    
    
    /**
     * The item namespace is used to access the currently selected message, meeting request, or appointment.
     * You can determine the type of the item by using the `itemType` property.
     *
     * To see the full member list, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer, Appointment Attendee, Message Compose, Message Read
     */
    export interface Item {
    }
    /**
     * The compose mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - DataExceedsMaximumSize: The location parameter is longer than 255 characters.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Mailbox {
        /**
         * Provides diagnostic information to an Outlook add-in.
         *
         * Contains the following members.
         *
         *  - `hostName` (string): A string that represents the name of the Office application.
         * It should be one of the following values: `Outlook`, `newOutlookWindows`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         * **Note**: The "Outlook" value is returned for Outlook on Windows (classic) and on Mac.
         *
         *  - `hostVersion` (string): A string that represents the version of either the Office application or the Exchange Server (e.g., "15.0.468.0").
         * If the mail add-in is running in Outlook on Windows (classic), on Mac, or on mobile devices, the `hostVersion` property returns the version of the
         * Outlook client. In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the property returns the version of the Exchange Server.
         *
         *  - `OWAView` (`MailboxEnums.OWAView` or string): An enum (or string literal) that represents the current view of Outlook on the web.
         * If the application is not Outlook on the web, then accessing this property results in undefined.
         * Outlook on the web has three views (`OneColumn` - displayed when the screen is narrow, `TwoColumns` - displayed when the screen is wider,
         * and `ThreeColumns` - displayed when the screen is wide) that correspond to the width of the screen and the window, and the number of columns
         * that can be displayed.
         *
         * More information is under {@link Office.Diagnostics}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * Starting with Mailbox requirement set 1.5, you can also use the 
         * {@link https://learn.microsoft.com/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics}
         * property to get similar information.
         */
        diagnostics: Diagnostics;
        /**
         * Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**:
         *
         * - Your app must have the **read item** permission specified in its manifest to call the `ewsUrl` member in read mode.
         *
         * - In compose mode, you must call the `saveAsync` method before you can use the `ewsUrl` member.
         * Your app must have **read/write item** permissions to call the `saveAsync` method.
         *
         * - This property isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox.
         * For example, you can create a remote service to {@link https://learn.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         */
        ewsUrl: string;
        /**
         * The mailbox item. Depending on the context in which the add-in opened, the item type may vary.
         * If you want to see IntelliSense for only a specific type or mode, cast this item to one of the following:
         *
         * {@link Office.MessageCompose | MessageCompose}, {@link Office.MessageRead | MessageRead},
         * {@link Office.AppointmentCompose | AppointmentCompose}, {@link Office.AppointmentRead | AppointmentRead}
         *
         * **Important**:
         *
         * - When calling `Office.context.mailbox.item` on a message, note that the Reading Pane in the Outlook client must be turned on.
         * For guidance on how to configure the Reading Pane, see
         * {@link https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0 | Use and configure the Reading Pane to preview messages}.
         *
         * - `item` can be null if your add-in supports pinning the task pane. For details on how to handle, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane#implement-the-event-handler | Implement a pinnable task pane in Outlook}.
         */
        item?: Item & ItemCompose & ItemRead & Message & MessageCompose & MessageRead & Appointment & AppointmentCompose & AppointmentRead;
        
        
        /**
         * Information about the user associated with the mailbox. This includes their account type, display name, email address, and time zone.
         *
         * More information is under {@link Office.UserProfile}
         */
        userProfile: UserProfile;

        
        
        
        /**
         * Gets a dictionary containing time information in local client time.
         *
         * The time zone used by the Outlook client varies by platform.
         * Outlook on Windows (classic) and on Mac use the client computer time zone.
         * Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows} use the time zone set on the Exchange Admin Center (EAC).
         * You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.
         *
         * In Outlook on Windows (classic) and on Mac, the `convertToLocalClientTime` method returns a dictionary object with the values set to the client computer time zone.
         * In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows}, the `convertToLocalClientTime` method returns a dictionary object
         * with the values set to the time zone specified in the EAC.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param timeValue - A `Date` object.
         */
        convertToLocalClientTime(timeValue: Date): LocalClientTime;
        
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param input - The local time value to convert.
         */
        convertToUtcClientTime(input: LocalClientTime): Date;
        /**
         * Displays an existing calendar appointment.
         *
         * The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop.
         *
         * In Outlook on Mac, you can use this method to display a single appointment that isn't part of a recurring series, or the master appointment
         * of a recurring series. However, you can't display an instance of the series because you can't access the properties
         * (including the item ID) of instances of a recurring series.
         *
         * In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and
         * no error message is returned.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**: This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         */
        displayAppointmentForm(itemId: string): void;
        
        
        /**
         * Displays an existing message.
         *
         * The `displayMessageForm` method opens an existing message in a new window on the desktop.
         *
         * In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier doesn't identify an existing message, no message will be displayed on the client computer, and
         * no error message is returned.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**:
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - Don't use the `displayMessageForm` with an itemId that represents an appointment. Use the `displayAppointmentForm` method to display
         * an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.
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
         * In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * this method always displays a form with an attendees field.
         * If you don't specify any attendees as input arguments, the method displays a form with a **Save** button.
         * If you have specified attendees, the form would include the attendees and a **Send** button.
         *
         * In Outlook on Windows (classic) and on Mac, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or
         * `resources` parameter, this method displays a meeting form with a **Send** button.
         * If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * **Important**: This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param parameters - An `AppointmentForm` describing the new appointment. All properties are optional.
         */
        displayNewAppointmentForm(parameters: AppointmentForm): void;
        
        
         
        
        
        
        /**
         * Gets a string that contains a token used to get an attachment or item from an Exchange Server.
         *
         * The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox.
         * The lifetime of the callback token is 5 minutes.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: All support Read mode; Mailbox 1.3 introduced Compose mode support]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**:
         *
         * - In February 2025, legacy Exchange {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#exchange-user-identity-token | user identity} and
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#callback-tokens | callback} tokens will be turned off by default for all Exchange Online tenants.
         * This is part of {@link https://blogs.microsoft.com/on-the-issues/2023/11/02/secure-future-initiative-sfi-cybersecurity-cyberattacks/ | Microsoft's Secure Future Initiative},
         * which gives organizations the tools needed to respond to the current threat landscape. Exchange user identity tokens will still work for Exchange on-premises.
         * Nested app authentication (NAA) is the recommended approach for tokens going forward. For more information, see the {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - You can pass both the token and either an attachment identifier or item identifier to an external system. That system uses
         * the token as a bearer authorization token to call the Exchange Web Services (EWS)
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation | GetAttachment} or
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation | GetItem} operation to return an
         * attachment or item. For example, you can create a remote service to
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * - Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **read item**.
         *
         * - Calling the `getCallbackTokenAsync` method in compose mode requires you to have saved the item.
         * The `saveAsync` method requires a minimum permission level of **read/write item**.
         *
         * - This method isn't supported in Outlook on Android or on iOS. EWS operations aren't supported in add-ins running in Outlook on mobile clients.
         * For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - This method isn't supported if you load an add-in in an Outlook.com or Gmail mailbox.
         *
         * - For guidance on delegate or shared scenarios, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Errors**:
         *
         * If your call fails, use the {@link https://learn.microsoft.com/javascript/api/office/office.asyncresult#office-office-asyncresult-diagnostics-member | asyncResult.diagnostics}
         * property to view details about the error.
         *
         * - `GenericTokenError: An internal error has occurred.` - In Exchange Online environments, this error occurs when the token can't be retrieved because legacy Exchange tokens
         * for Outlook add-ins are turned off. We recommend using NAA as a single sign-on solution for your add-in. For guidance on how to implement NAA, see the
         * {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - `HTTPRequestFailure: The request has failed. Please look at the diagnostics object for the HTTP error code.`
         *
         * - `InternalServerError: The Exchange server returned an error. Please look at the diagnostics object for more information.` - In Exchange Online environments,
         * this error occurs when the token can't be retrieved because legacy Exchange tokens for Outlook add-ins are turned off. We recommend using NAA as a single sign-on solution for your add-in.
         * For guidance on how to implement NAA, see the {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - `NetworkError: The user is no longer connected to the network. Please check your network connection and try again.`
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getCallbackTokenAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Returns true if the current mailbox is managed by {@link https://learn.microsoft.com/mem/intune/fundamentals/what-is-intune | Microsoft Intune}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         *
         * **Important**: This method is only supported in Outlook on Android and on iOS starting in Version 4.2443.0.To learn more about APIs supported in Outlook on mobile devices, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * **Errors**:
         *
         * - `MAMServiceNotAvailable`: The client is unable to fetch the mobile application management (MAM) policy.
         *
         * @returns True if the current mailbox is managed by Microsoft Intune.
         */
        getIsIdentityManaged(): boolean;
        /**
         * Returns true if an organization's {@link https://learn.microsoft.com/mem/intune/apps/app-management | Intune mobile application management (MAM) policy}
         * allows an add-in to access data from the specified location.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         *
         * **Important**: This method is only supported in Outlook on Android and on iOS starting in Version 4.2443.0. To learn more about APIs supported in Outlook on mobile devices, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * **Errors**:
         *
         * - `InvalidOpenLocationInput`: The value of the specified location is invalid.
         *
         * - `MAMServiceNotAvailable`: The client is unable to fetch the MAM policy.
         *
         * @param openLocation - The location from which the add-in is attempting to access data.
         *
         * @returns True if an organization's Intune MAM policy allows an add-in to access data from the specified location.
         */
        getIsOpenFromLocationAllowed(openLocation: MailboxEnums.OpenLocation): boolean;
        /**
         * Returns true if an organization's {@link https://learn.microsoft.com/mem/intune/apps/app-management | Intune mobile application management (MAM) policy}
         * allows an add-in to save data to the specified location.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         *
         * **Important**: This method is only supported in Outlook on Android and on iOS starting in Version 4.2443.0. To learn more about APIs supported in Outlook on mobile devices, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * **Errors**:
         *
         * - `InvalidSaveLocationInput`: The value of the specified location is invalid.
         *
         * - `MAMServiceNotAvailable`: The client is unable to fetch the MAM policy.
         *
         * @param saveLocation - The location in which the add-in is attempting to save data.
         *
         * @returns True if an organization's Intune MAM policy allows an add-in to save data to the specified location.
         */
        getIsSaveToLocationAllowed(saveLocation: MailboxEnums.SaveLocation): boolean;
        
        
        /**
         * Gets a token identifying the user and the Office Add-in.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**:
         *
         * - In February 2025, legacy Exchange {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#exchange-user-identity-token | user identity} and
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#callback-tokens | callback} tokens will be turned off by default for all Exchange Online tenants.
         * This is part of {@link https://blogs.microsoft.com/on-the-issues/2023/11/02/secure-future-initiative-sfi-cybersecurity-cyberattacks/ | Microsoft's Secure Future Initiative},
         * which gives organizations the tools needed to respond to the current threat landscape. Exchange user identity tokens will still work for Exchange on-premises.
         * Nested app authentication (NAA) is the recommended approach for tokens going forward. For more information, see the {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - The `getUserIdentityTokenAsync` method returns a token that you can use to identify and
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication | authenticate the add-in and user with an external system}.
         *
         * - This method isn't supported if you load an add-in in an Outlook.com or Gmail mailbox.
         *
         * **Errors**:
         *
         * If your call fails, use the {@link https://learn.microsoft.com/javascript/api/office/office.asyncresult#office-office-asyncresult-diagnostics-member | asyncResult.diagnostics}
         * property to view details about the error.
         *
         * - `GenericTokenError: An internal error has occurred.` - In Exchange Online environments, this error occurs when the token can't be retrieved because legacy Exchange tokens
         * for Outlook add-ins are turned off. We recommend using NAA as a single sign-on solution for your add-in. For guidance on how to implement NAA, see the
         * {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - `HTTPRequestFailure: The request has failed. Please look at the diagnostics object for the HTTP error code.`
         *
         * - `InternalServerError: The Exchange server returned an error. Please look at the diagnostics object for more information.` - In Exchange Online environments,
         * this error occurs when the token can't be retrieved because legacy Exchange tokens for Outlook add-ins are turned off. We recommend using NAA as a single sign-on solution for your add-in.
         * For guidance on how to implement NAA, see the {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - `NetworkError: The user is no longer connected to the network. Please check your network connection and try again.`
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
         * The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Important**:
         *
         * - In February 2025, legacy Exchange {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#exchange-user-identity-token | user identity} and
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication#callback-tokens | callback} tokens will be turned off by default for all Exchange Online tenants.
         * This is part of {@link https://blogs.microsoft.com/on-the-issues/2023/11/02/secure-future-initiative-sfi-cybersecurity-cyberattacks/ | Microsoft's Secure Future Initiative},
         * which gives organizations the tools needed to respond to the current threat landscape. Exchange user identity tokens will still work for Exchange on-premises.
         * Nested app authentication (NAA) is the recommended approach for tokens going forward. For more information, see the {@link https://aka.ms/naafaq | FAQ page}.
         *
         * - To enable the `makeEwsRequestAsync` method to make EWS requests, the server administrator must set `OAuthAuthentication` to `true` on the
         * Client Access Server EWS directory .
         *
         * - Your add-in must have the **read/write mailbox** permission to use the `makeEwsRequestAsync` method.
         * For information about using the **read/write mailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Specify permissions for mail add-in access to the user's mailbox}.
         *
         * - If your add-in needs to access Folder Associated Items or its XML request must specify UTF-8 encoding (`\<?xml version="1.0" encoding="utf-8"?\>`),
         * it must use Microsoft Graph or REST APIs to access the user's mailbox instead.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - This method isn't supported when the add-in is loaded in a Gmail mailbox.
         *
         * - When you use the `makeEwsRequestAsync` method in add-ins that run in Outlook versions earlier than Version 15.0.4535.1004, you must set
         * the encoding value to ISO-8859-1 (`<?xml version="1.0" encoding="iso-8859-1"?>`). To determine the version of an Outlook client, use the
         * `mailbox.diagnostics.hostVersion` property. You don't need to set the encoding value when your add-in is running in Outlook on the web or
         * {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows}.
         * To determine the Outlook client in which your add-in is running, use the `mailbox.diagnostics.hostName` property.
         *
         * @param data - The EWS request.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                   `asyncResult`, which is an `Office.AsyncResult` object. The XML response of the EWS request is provided as a string
         *                   in the `asyncResult.value` property. In Outlook on the web, on Windows
         *                   ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic (starting in Version 2303, Build 16225.10000)),
         *                   and on Mac (starting in Version 16.73 (23042601)), if the response exceeds 5 MB in size, an error message is returned in the `asyncResult.error` property.
         *                   In earlier versions of Outlook on Windows (classic) and on Mac, an error message is returned if the response exceeds 1 MB in size.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        makeEwsRequestAsync(data: any, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        
        
    }
    
    
    /**
     * Represents a suggested meeting found in an item. Read mode only.
     *
     * The list of meetings suggested in an email message is returned in the `meetingSuggestions` property of the `Entities` object that's returned when
     * the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * The start and end values are string representations of a `Date` object that contains the date and time at which the suggested meeting is to
     * begin and end. The values are in the default time zone specified for the current user.
     *
     * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
     * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
     * For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface MeetingSuggestion {
        /**
         * Gets the attendees for a suggested meeting.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        attendees: EmailUser[];
        /**
         * Gets the date and time that a suggested meeting is to end.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        end: string;
        /**
         * Gets the location of a suggested meeting.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        location: string;
        /**
         * Gets a string that was identified as a meeting suggestion.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        meetingString: string;
        /**
         * Gets the date and time that a suggested meeting is to begin.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        start: string;
        /**
         * Gets the subject of a suggested meeting.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        subject: string;
    }
    /**
     * A subclass of {@link Office.Item | Item} for messages.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
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
     * **Important**:
     *
     * - This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
     *
     * - When calling `Office.context.mailbox.item` on a message, note that the Reading Pane in the Outlook client must be turned on.
     * For guidance on how to configure the Reading Pane, see
     * {@link https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0 | Use and configure the Reading Pane to preview messages}.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        bcc: Recipients;
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        itemType: MailboxEnums.ItemType | string;
        
        
        
        
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        to: Recipients;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * @remarks
         * [Api set: Mailbox 1.1 for Outlook on Windows (classic) and on Mac, Mailbox 1.8 for Outlook on the web and new Outlook on Windows]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Important**:
         *
         * - This method isn't supported in Outlook on iOS or Android. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - In recent builds of classic Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, use the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * - The URI of the file to be attached must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header that
         * specifies `no-cache`, `no-store`, or similar options in the HTTP response. However, when you're developing the add-in and making changes to files,
         * caching can prevent you from seeing your changes. We recommend using `Cache-Control` headers during development.
         *
         * - You can use the same URI with the `removeAttachmentAsync` method to remove the attachment in the same session.
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
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline as an image in the message body and won't be displayed in the attachment list.
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
         * @remarks
         * [Api set: Mailbox 1.1 for Outlook on Windows (classic) and on Mac, Mailbox 1.8 for Outlook on the web and new Outlook on Windows]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Important**:
         *
         * - This method isn't supported in Outlook on iOS or Android. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * - In recent builds of classic Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, use the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * - The URI of the file to be attached must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header that
         * specifies `no-cache`, `no-store`, or similar options in the HTTP response. However, when you're developing the add-in and making changes to files,
         * caching can prevent you from seeing your changes. We recommend using `Cache-Control` headers during development.
         *
         * - You can use the same URI with the `removeAttachmentAsync` method to remove the attachment in the same session.
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
         * If your Office Add-in is running in Outlook on the web or {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing. However, this isn't supported and isn't recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * If your Office Add-in is running in Outlook on the web or {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing. However, this isn't supported and isn't recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `Html`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
         * in the same session. In Outlook on the web, on mobile devices, and in {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * *Important**: The `removeAttachmentAsync` method doesn't remove inline attachments from a mail item.
         * To remove an inline attachment, first get the item's body, then remove any references of the attachment from its contents.
         * Use the {@link https://learn.microsoft.com/javascript/api/outlook/office.body | Office.Body} APIs to get and set the body of an item.
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows (new and classic).
         * @param options - An object literal that contains one or more of the following properties:-
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
         * in the same session. In Outlook on the web, on mobile devices, and in {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * *Important**: The `removeAttachmentAsync` method doesn't remove inline attachments from a mail item.
         * To remove an inline attachment, first get the item's body, then remove any references of the attachment from its contents.
         * Use the {@link https://learn.microsoft.com/javascript/api/outlook/office.body | Office.Body} APIs to get and set the body of an item.
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows (new and classic).
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing the attachment fails, the `asyncResult.error` property will contain an error code
         *                 with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        
        
        
        
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web, on Windows
         *        ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), and on Mac.
         *        If the field is an HTML editor, only the text data is inserted, even if the data is HTML.
         *        If the data is HTML and the field supports HTML (the subject doesn't), the current style is applied in
         *        Outlook on the web and new Outlook on Windows. The default style is applied in Outlook on Windows (classic) and on Mac.
         *        If the field is a text field, an `InvalidDataFormat` error is returned.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
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
     * **Important**:
     *
     * - This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/requirement-set-1.2/office.context.mailbox.item | Object Model} page.
     *
     * - When calling `Office.context.mailbox.item` on a message, note that the Reading Pane in the Outlook client must be turned on.
     * For guidance on how to configure the Reading Pane, see
     * {@link https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0 | Use and configure the Reading Pane to preview messages}.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        body: Body;
        
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `cc` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **Cc** line of the message. The maximum number of recipients returned varies per Outlook client.
         *
         * - classic Windows: 500 recipients
         *
         * - Android, classic Mac UI, iOS: 100 recipients
         *
         * - Web browser, new Outlook: 20 recipients (collapsed view), 500 recipients (expanded view)
         *
         * - New Mac UI: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        conversationId: string;
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**: This property isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        from: EmailAddressDetails;
        /**
         * Gets the internet message identifier for an email message.
         *
         * **Important**: In the **Sent Items** folder, the `internetMessageId` may not be available yet on recently sent items. In that case,
         * consider using {@link https://learn.microsoft.com/office/dev/add-ins/outlook/web-services | Exchange Web Services} to get this
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/internetmessageid | property from the server}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        internetMessageId: string;
        /**
         * Gets the Exchange Web Services item class of the selected message.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Important**:
         *
         * The following table lists the default item classes for messages.
         *
         * <table>
         *   <tr>
         *     <th>Item class</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>IPM.Note</td>
         *     <td>New messages and message replies</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Note.SMIME</td>
         *     <td>Encrypted messages that can also be signed</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Note.SMIME.MultipartSigned</td>
         *     <td>Clear-signed messages</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Schedule.Meeting.Request</td>
         *     <td>Meeting requests</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Schedule.Meeting.Canceled</td>
         *     <td>Meeting cancellations</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Schedule.Meeting.Resp.Neg</td>
         *     <td>Responses to decline meeting requests</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Schedule.Meeting.Resp.Pos</td>
         *     <td>Responses to accept meeting requests</td>
         *   </tr>
         *   <tr>
         *     <td>IPM.Schedule.Meeting.Resp.Tent</td>
         *     <td>Responses to tentatively accept meeting requests</td>
         *   </tr>
         * </table>
         *
         * You can create custom classes that extend a default item class. For example, `IPM.Note.Contoso`.
         */
        itemClass: string;
        /**
         * Gets the {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services (EWS) item identifier}
         * of the current item.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Important**:
         *
         * - The `itemId` property isn't available in compose mode.
         * If an item identifier is required, the `Office.context.mailbox.item.saveAsync` method can be used to save the item to the store, which will return the item identifier
         * in the `asyncResult.value` parameter in the callback function. If the item is already saved, you can call the `Office.context.mailbox.item.getItemIdAsync` method instead.
         *
         * - The item ID returned isn't identical to the Outlook Entry ID or the ID used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of a meeting request.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        normalizedSubject: string;
        
        
        
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        subject: string;
        /**
         * Provides access to the recipients on the **To** line of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `to` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **To** line of the message. The maximum number of recipients returned varies per Outlook client.
         *
         * - classic Windows: 500 recipients
         *
         * - Android, classic Mac UI, iOS: 100 recipients
         *
         * - Web browser, new Outlook: 20 recipients (collapsed view), 500 recipients (expanded view)
         *
         * - New Mac UI: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        to: EmailAddressDetails[];
        
        
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * - If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.
         *
         * - When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        
        
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Important**:
         *
         * - In Outlook on the web and {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows},
         * the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * - If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * - When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        
        
        
        
        
        
        
        
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns `null`.
         * If no entities of the specified type are present in the item's body, the method returns an empty array.
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param entityType - One of the `EntityType` enumeration values.
         */
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): Array<string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion>;
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in an add-in only manifest file.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         *
         * @returns
         * The entities that match the regular expression defined in the `ItemHasKnownEntity` rule element in the
         * add-in manifest file with the specified `FilterName` element value. If there's no `ItemHasKnownEntity` element in the manifest with a
         * `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter matches an
         * `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method returns an empty array.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): Array<string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion>;
        
        
        /**
         * Returns string values in the selected item that match the regular expressions defined in an add-in only manifest file.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the add-in manifest file.
         * The name of each array is equal to the corresponding value of the RegExName attribute of the matching `ItemHasRegularExpressionMatch` rule.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property
         * of the item that's specified by that rule. The `PropertyName` simple type defines the supported properties.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * - This method is used with the {@link https://learn.microsoft.com/javascript/api/manifest/rule | activation rules feature for Outlook add-ins},
         * which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | unified manifest for Microsoft 365}.
         *
         * - If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item doesn't always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in an add-in only manifest file.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the add-in manifest file,
         * with the specified `RegExName` element value.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Important**:
         *
         * - Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * - This method is used with the {@link https://learn.microsoft.com/javascript/api/manifest/rule | activation rules feature for Outlook add-ins},
         * which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | unified manifest for Microsoft 365}.
         *
         * - If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item doesn't always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * - This method isn't supported in Outlook on Android or on iOS. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        
        
        
        
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        
        
    }
    
    
    
    
    
    /**
     * Represents a phone number identified in an item. Read mode only.
     *
     * An array of `PhoneNumber` objects containing the phone numbers found in an email message is returned in the `phoneNumbers` property of the
     * `Entities` object that's returned when you call the `getEntities` method on the selected item.
     *
     * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
     * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
     * For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface PhoneNumber {
        /**
         * Gets a string containing a phone number. This string contains only the digits of the telephone number and excludes characters
         * like parentheses and hyphens, if they exist in the original item.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        phoneString: string;
        /**
         * Gets the text that was identified in an item as a phone number.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        originalPhoneString: string;
        /**
         * Gets a string that identifies the type of phone number: Home, Work, Mobile, Unspecified.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        type: string;
    }
    /**
     * Represents recipients of an item. Compose mode only.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Recipients {
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: With the `addAsync` method, you can add a maximum of 100 recipients to a mail item in Outlook on the web, on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), on Mac (classic UI), on Android, and on iOS.
         * However, take note of the following:
         *
         * - In Outlook on the web, on Windows (new and classic), and on Mac (classic UI), you can have a maximum of 500 recipients in a target field.
         * If you need to add more than 100 recipients to a mail item, you can call `addAsync` repeatedly, but be mindful of the recipient limit of the field.
         *
         * - In Outlook on Android and on iOS, the `addAsync` method isn't supported in Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * There's no recipient limit if you call `addAsync` in Outlook on Mac (new UI).
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list. The array of recipients can contain strings of SMTP email addresses,
         *        {@link Office.EmailUser | EmailUser} objects, or {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: With the `addAsync` method, you can add a maximum of 100 recipients to a mail item in Outlook on the web, on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), on Mac (classic UI), on Android, and on iOS.
         * However, take note of the following:
         *
         * - In Outlook on the web, on Windows (new and classic), and on Mac (classic UI), you can have a maximum of 500 recipients in a target field.
         * If you need to add more than 100 recipients to a mail item, you can call `addAsync` repeatedly, but be mindful of the recipient limit of the field.
         *
         * - In Outlook on Android and on iOS, the `addAsync` method isn't supported in Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * There's no recipient limit if you call `addAsync` in Outlook on Mac (new UI).
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list. The array of recipients can contain strings of SMTP email addresses,
         *        {@link Office.EmailUser | EmailUser} objects, or {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**:
         *
         * The maximum number of recipients returned by this method varies per Outlook client.
         *
         * - Windows ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), web browser, Mac (classic UI): 500 recipients
         *
         * - Android, iOS: 100 recipients
         *
         * - Mac (new UI): No limit
         *
         * In classic Outlook on Windows, the appointment organizer is included in the object returned by the `getAsync` method when you create a new appointment or edit an
         * existing one. In Outlook on the web and new Outlook on Windows, the organizer is only included in the returned object when you edit an existing appointment.
         *
         * The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
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
         * In Outlook on the web and on Windows (new and classic), if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * {@link Office.EmailAddressDetails | EmailAddressDetails} object instead of the contact's saved name.
         * For more details, see {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         *
         * While composing a mail item, when you switch to a sender account that's on a different domain than that of the previously selected sender account,
         * the value of the `recipientType` property for existing recipients isn't updated and will still be based on the domain of the previously selected account.
         * To get the correct recipient types after switching accounts, you must first remove the existing recipients, then add them back to the mail item.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 of type `Office.AsyncResult`. The `asyncResult.value` property of the result is an array of
         *                 {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**:
         *
         * The maximum number of recipients returned by this method varies per Outlook client.
         *
         * - Windows ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), web browser, Mac (classic UI): 500 recipients
         *
         * - Android, iOS: 100 recipients
         *
         * - Mac (new UI): No limit
         *
         * The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
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
         * In Outlook on the web and on Windows (new and classic), if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * {@link Office.EmailAddressDetails | EmailAddressDetails} object instead of the contact's saved name.
         * For more details, see {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         *
         * While composing a mail item, when you switch to a sender account that's on a different domain than that of the previously selected sender account,
         * the value of the `recipientType` property for existing recipients isn't updated and will still be based on the domain of the previously selected account.
         * To get the correct recipient types after switching accounts, you must first remove the existing recipients, then add them back to the mail item.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 of type `Office.AsyncResult`. The `asyncResult.value` property of the result is an array of
         *                 {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: With the `setAsync` method, you can set a maximum of 100 recipients in Outlook on the web, on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), on Mac (classic UI), on Android, and on iOS.
         * However, take note of the following:
         *
         * - In Outlook on the web, on Windows (new and classic), and on Mac (classic UI), you can have a maximum of 500 recipients in a target field.
         * If you need to set more than 100 recipients, you can call `setAsync` repeatedly, but be mindful of the recipient limit of the field.
         *
         * - In Outlook on Android and on iOS, the `setAsync` method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * There's no recipient limit if you call `setAsync` in Outlook on Mac (new UI).
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list. The array of recipients can contain strings of SMTP email addresses,
         *        {@link Office.EmailUser | EmailUser} objects, or {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: With the `setAsync` method, you can set a maximum of 100 recipients in Outlook on the web, on Windows
         * ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} and classic), on Mac (classic UI), on Android, and on iOS.
         * However, take note of the following:
         *
         * - In Outlook on the web, on Windows (new and classic), and on Mac (classic UI), you can have a maximum of 500 recipients in a target field.
         * If you need to set more than 100 recipients, you can call `setAsync` repeatedly, but be mindful of the recipient limit of the field.
         *
         * - In Outlook on Android and on iOS, the `setAsync` method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * There's no recipient limit if you call `setAsync` in Outlook on Mac (new UI).
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list. The array of recipients can contain strings of SMTP email addresses,
         *        {@link Office.EmailUser | EmailUser} objects, or {@link Office.EmailAddressDetails | EmailAddressDetails} objects.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
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
         * Only used if type is set to item. The EWS item ID of the attachment. This is a string up to 100 characters.
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
         * An object literal that contains the following property:-
         * `asyncContext`: Developers can provide any object they wish to access in the callback function.
         */
        options?: CommonAPI.AsyncContextOptions;
    }
    /**
     * The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user.
     * That is, they are available only to the add-in that created them, and only from the user's mailbox in which they are saved.
     *
     * While the Outlook add-in API limits access to these settings to only the add-in that created them, these settings shouldn't be considered
     * secure storage. They can be accessed by Exchange Web Services or Extended MAPI.
     * They shouldn't be used to store sensitive information, such as user credentials or security tokens.
     *
     * The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.
     *
     * The `RoamingSettings` object is accessible via the `roamingSettings` property in the `Office.context` namespace.
     * 
     * To learn more about `RoamingSettings`, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **Important**:
     *
     * - The `RoamingSettings` object is initialized from the persisted storage only when the add-in is first loaded.
     * For task panes, this means that it's only initialized when the task pane first opens.
     * If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values, even if
     * your add-in has persisted changes.
     * The persisted changes will not be available until the task pane (or item in the case of UI-less add-ins) is closed and reopened.
     *
     * - When set and saved through Outlook on Windows ({@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new} or classic) or on Mac,
     * these settings are reflected in Outlook on the web only after a browser refresh.
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface RoamingSettings {
        /**
         * Retrieves the specified setting.
         *
         * @returns Type: String | Number | Boolean | Object | Array
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to retrieve.
         */
        get(name: string): any;
        /**
         * Removes the specified setting.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to remove.
         */
        remove(name: string): void;
        /**
         * Saves the settings.
         *
         * Any settings previously saved by an add-in are loaded when it's initialized, so during the lifetime of the session you can just use
         * the set and get methods to work with the in-memory copy of the settings property bag.
         * When you want to persist the settings so that they're available the next time the add-in is used, use the `saveAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        saveAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets or creates the specified setting.
         *
         * The `set` method creates a new setting of the specified name if it doesn't already exist, or sets an existing setting of the specified name.
         * The value is stored in the document as the serialized JSON representation of its data type.
         *
         * A maximum of 32KB is available for the settings of each add-in. An error with code 9057 is thrown when that size limit is exceeded.
         *
         * Any changes made to settings using the `set` method will not be saved to the server until the `saveAsync` method is called.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is the subject of the item.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the subject of an appointment or message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The subject parameter is longer than 255 characters.
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on Android and on iOS, this method isn't supported in the Message Compose mode. Only the Appointment Organizer mode is
         * supported. For more information on supported APIs in Outlook mobile, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-mobile-apis | Outlook JavaScript APIs supported in Outlook on mobile devices}.
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
     * that's returned when the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
     * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
     * For guidance on how to implement these rules, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
     *
     * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface TaskSuggestion {
        /**
         * Gets the users that should be assigned a suggested task.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        assignees: EmailUser[];
        /**
         * Gets the text of an item that was identified as a task suggestion.
         *
         * **Warning**: Entity-based contextual Outlook add-ins are now retired. However, regular expression rules are still supported.
         * We recommend updating your contextual add-in to use regular expression rules as an alternative solution.
         * For guidance on how to implement these rules, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | Contextual Outlook add-ins}.
         *
         * @deprecated Use {@link https://learn.microsoft.com/office/dev/add-ins/outlook/contextual-outlook-add-ins | regular expression rules} instead.
         */
        taskString: string;
    }
    /**
     * The `Time` object is returned as the start or end property of an appointment in compose mode.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidEndTime`: The appointment end time is before the appointment start time.
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param options - An object literal that contains one or more of the following properties:-
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
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
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
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface UserProfile {
        
        /**
         * Gets the user's display name.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        displayName: string;
        /**
         * Gets the user's SMTP email address.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        emailAddress: string;
        /**
         * Gets the user's time zone in Windows format.
         *
         * The system time zone is usually returned. However, in Outlook on the web and in {@link https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627 | new Outlook on Windows}),
         * the default time zone in the calendar preferences is returned instead.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        timeZone: string;
    }
}


////////////////////////////////////////////////////////////////
/////////////////////// End Exchange APIs //////////////////////
////////////////////////////////////////////////////////////////