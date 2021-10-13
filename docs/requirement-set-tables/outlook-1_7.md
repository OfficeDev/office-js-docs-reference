| Class | Fields | Description |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/outlook.appointmentcompose)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentcompose#addHandlerAsync_eventType__handler__callback__asyncResult_)|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentcompose#addHandlerAsync_eventType__handler__options__callback__asyncResult_)|Adds an event handler for a supported event.|
||[organizer](/javascript/api/outlook/outlook.appointmentcompose#organizer)|Gets the organizer for the specified meeting.|
||[recurrence](/javascript/api/outlook/outlook.appointmentcompose#recurrence)|Gets or sets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentcompose#removeHandlerAsync_eventType__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentcompose#removeHandlerAsync_eventType__options__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/outlook.appointmentcompose#seriesId)|Gets the id of the series that an instance belongs to.|
|[AppointmentRead](/javascript/api/outlook/outlook.appointmentread)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentread#addHandlerAsync_eventType__handler__callback__asyncResult_)|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentread#addHandlerAsync_eventType__handler__options__callback__asyncResult_)|Adds an event handler for a supported event.|
||[recurrence](/javascript/api/outlook/outlook.appointmentread#recurrence)|Gets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentread#removeHandlerAsync_eventType__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.appointmentread#removeHandlerAsync_eventType__options__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/outlook.appointmentread#seriesId)|Gets the ID of the series that an instance belongs to.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/outlook.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#end)||
||[start](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#start)||
||[type](/javascript/api/outlook/outlook.appointmenttimechangedeventargs#type)||
|[From](/javascript/api/outlook/outlook.from)|[getAsync(callback?: (asyncResult: Office.AsyncResult<EmailAddressDetails>) => void)](/javascript/api/outlook/outlook.from#getAsync_callback__asyncResult_)|Gets the from value of a message.|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<EmailAddressDetails>) => void)](/javascript/api/outlook/outlook.from#getAsync_options__callback__asyncResult_)|Gets the from value of a message.|
|[MessageCompose](/javascript/api/outlook/outlook.messagecompose)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messagecompose#addHandlerAsync_eventType__handler__callback__asyncResult_)|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messagecompose#addHandlerAsync_eventType__handler__options__callback__asyncResult_)|Adds an event handler for a supported event.|
||[from](/javascript/api/outlook/outlook.messagecompose#from)|Gets the email address of the sender of a message.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messagecompose#removeHandlerAsync_eventType__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messagecompose#removeHandlerAsync_eventType__options__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/outlook.messagecompose#seriesId)|Gets the ID of the series that an instance belongs to.|
|[MessageRead](/javascript/api/outlook/outlook.messageread)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messageread#addHandlerAsync_eventType__handler__callback__asyncResult_)|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messageread#addHandlerAsync_eventType__handler__options__callback__asyncResult_)|Adds an event handler for a supported event.|
||[recurrence](/javascript/api/outlook/outlook.messageread#recurrence)|Gets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messageread#removeHandlerAsync_eventType__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.messageread#removeHandlerAsync_eventType__options__callback__asyncResult_)|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/outlook.messageread#seriesId)|Gets the id of the series that an instance belongs to.|
|[Organizer](/javascript/api/outlook/outlook.organizer)|[getAsync(callback?: (asyncResult: Office.AsyncResult<EmailAddressDetails>) => void)](/javascript/api/outlook/outlook.organizer#getAsync_callback__asyncResult_)|Gets the organizer value of an appointment as an {@link Office.EmailAddressDetails | EmailAddressDetails} object|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<EmailAddressDetails>) => void)](/javascript/api/outlook/outlook.organizer#getAsync_options__callback__asyncResult_)|Gets the organizer value of an appointment as an {@link Office.EmailAddressDetails | EmailAddressDetails} object|
|[RecipientsChangedEventArgs](/javascript/api/outlook/outlook.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/outlook.recipientschangedeventargs#changedRecipientFields)||
||[type](/javascript/api/outlook/outlook.recipientschangedeventargs#type)||
|[RecipientsChangedFields](/javascript/api/outlook/outlook.recipientschangedfields)|[bcc](/javascript/api/outlook/outlook.recipientschangedfields#bcc)|Gets if recipients in the **bcc** field were changed.|
||[cc](/javascript/api/outlook/outlook.recipientschangedfields#cc)|Gets if recipients in the **cc** field were changed.|
||[optionalAttendees](/javascript/api/outlook/outlook.recipientschangedfields#optionalAttendees)|Gets if optional attendees were changed.|
||[requiredAttendees](/javascript/api/outlook/outlook.recipientschangedfields#requiredAttendees)|Gets if required attendees were changed.|
||[resources](/javascript/api/outlook/outlook.recipientschangedfields#resources)|Gets if resources were changed.|
||[to](/javascript/api/outlook/outlook.recipientschangedfields#to)|Gets if recipients in the **to** field were changed.|
|[Recurrence](/javascript/api/outlook/outlook.recurrence)|[getAsync(callback?: (asyncResult: Office.AsyncResult<Recurrence>) => void)](/javascript/api/outlook/outlook.recurrence#getAsync_callback__asyncResult_)|Returns the current recurrence object of an appointment series.|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<Recurrence>) => void)](/javascript/api/outlook/outlook.recurrence#getAsync_options__callback__asyncResult_)|Returns the current recurrence object of an appointment series.|
||[recurrenceProperties](/javascript/api/outlook/outlook.recurrence#recurrenceProperties)|Gets or sets the properties of the recurring appointment series.|
||[recurrenceTimeZone](/javascript/api/outlook/outlook.recurrence#recurrenceTimeZone)|Gets or sets the properties of the recurring appointment series.|
||[recurrenceType](/javascript/api/outlook/outlook.recurrence#recurrenceType)|Gets or sets the type of the recurring appointment series.|
||[seriesTime](/javascript/api/outlook/outlook.recurrence#seriesTime)|The {@link Office.SeriesTime | SeriesTime} object enables you to manage the start and end dates of the recurring appointment series and|
||[setAsync(recurrencePattern: Recurrence, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.recurrence#setAsync_recurrencePattern__callback__asyncResult_)|Sets the recurrence pattern of an appointment series.|
||[setAsync(recurrencePattern: Recurrence, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult<void>) => void)](/javascript/api/outlook/outlook.recurrence#setAsync_recurrencePattern__options__callback__asyncResult_)|Sets the recurrence pattern of an appointment series.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/outlook.recurrencechangedeventargs)|[recurrence](/javascript/api/outlook/outlook.recurrencechangedeventargs#recurrence)||
||[type](/javascript/api/outlook/outlook.recurrencechangedeventargs#type)||
|[RecurrenceProperties](/javascript/api/outlook/outlook.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/outlook.recurrenceproperties#dayOfMonth)|Represents the day of the month.|
||[dayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#dayOfWeek)|Represents the day of the week or type of day, for example, weekend day vs weekday.|
||[days](/javascript/api/outlook/outlook.recurrenceproperties#days)|Represents the set of days for this recurrence.|
||[firstDayOfWeek](/javascript/api/outlook/outlook.recurrenceproperties#firstDayOfWeek)|Represents your chosen first day of the week otherwise the default is the value in the current user's settings.|
||[interval](/javascript/api/outlook/outlook.recurrenceproperties#interval)|Represents the period between instances of the same recurring series.|
||[month](/javascript/api/outlook/outlook.recurrenceproperties#month)|Represents the month.|
||[weekNumber](/javascript/api/outlook/outlook.recurrenceproperties#weekNumber)|Represents the number of the week in the selected month e.g., 'first' for first week of the month.|
|[RecurrenceTimeZone](/javascript/api/outlook/outlook.recurrencetimezone)|[name](/javascript/api/outlook/outlook.recurrencetimezone#name)|Represents the name of the recurrence time zone.|
||[offset](/javascript/api/outlook/outlook.recurrencetimezone#offset)|Integer value representing the difference in minutes between the local time zone and UTC at the date that the meeting series began.|
|[SeriesTime](/javascript/api/outlook/outlook.seriestime)|[getDuration()](/javascript/api/outlook/outlook.seriestime#getDuration__)|Gets the duration in minutes of a usual instance in a recurring appointment series.|
||[getEndDate()](/javascript/api/outlook/outlook.seriestime#getEndDate__)|Gets the end date of a recurrence pattern in the following|
||[getEndTime()](/javascript/api/outlook/outlook.seriestime#getEndTime__)|Gets the end time of a usual appointment or meeting request instance of a recurrence pattern in whichever time zone that the user or|
||[getStartDate()](/javascript/api/outlook/outlook.seriestime#getStartDate__)|Gets the start date of a recurrence pattern in the following|
||[getStartTime()](/javascript/api/outlook/outlook.seriestime#getStartTime__)|Gets the start time of a usual appointment instance of a recurrence pattern in whichever time zone that the user/add-in set the|
||[setDuration(minutes: number)](/javascript/api/outlook/outlook.seriestime#setDuration_minutes_)|Sets the duration of all appointments in a recurrence pattern.|
||[setEndDate(date: string)](/javascript/api/outlook/outlook.seriestime#setEndDate_date_)|Sets the end date of a recurring appointment series.|
||[setEndDate(year: number, month: number, day: number)](/javascript/api/outlook/outlook.seriestime#setEndDate_year__month__day_)|Sets the end date of a recurring appointment series.|
||[setStartDate(date:string)](/javascript/api/outlook/outlook.seriestime#setStartDate_date_)|Sets the start date of a recurring appointment series.|
||[setStartDate(year:number, month:number, day:number)](/javascript/api/outlook/outlook.seriestime#setStartDate_year__month__day_)|Sets the start date of a recurring appointment series.|
||[setStartTime(hours: number, minutes: number)](/javascript/api/outlook/outlook.seriestime#setStartTime_hours__minutes_)|Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set|
||[setStartTime(time: string)](/javascript/api/outlook/outlook.seriestime#setStartTime_time_)|Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set|
