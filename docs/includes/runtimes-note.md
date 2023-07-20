> [!IMPORTANT]
> For the shared JavaScript runtime, this element enables the ribbon, task pane, and other supported components to use the same runtime. However, the SharedRuntime requirement set is only available in some Office applications. For more information, see [Shared runtime requirement sets](/office/dev/add-ins/reference/requirement-sets/shared-runtime-requirement-sets).
>
> **Outlook**
>
> - For event-based activation, this element enables your add-in to run on composing a new item, for example. For supported clients and other information, see [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch).
> - For integrated spam reporting (preview), this element enables your add-in to process unsolicited messages. To learn more about how to implement the integrated spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting).
>
> Note that the event-based activation and integrated spam reporting features must use the same runtime. Multiple runtimes aren't currently supported in Outlook.
