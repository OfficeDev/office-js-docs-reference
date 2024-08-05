---
title: FunctionFile element in the manifest file
description: Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.
ms.date: 03/20/2023
ms.localizationpriority: medium
---

# FunctionFile element

Specifies the source code file for operations that an add-in exposes in one of the following ways.

- Add-in commands that execute a JavaScript function instead of displaying UI.
- Keyboard shortcuts that execute a JavaScript function.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

The **\<FunctionFile\>** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The `resid` attribute of the **\<FunctionFile\>** element can be no more than 32 characters and is set to the value of the `id` attribute of a **\<Url\>** element in the [Resources](resources.md) element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by [function command](/office/dev/add-ins/design/add-in-commands) buttons, as defined by the [Control element](control.md).

> [!NOTE]
> When the add-in is configured to use a [shared runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime), the functions in the code file run in the same JavaScript runtime (and share a common global namespace) as the JavaScript in the add-in's task pane (if any).
>
> The **\<FunctionFile\>** element and the associated code file also have a special role to play with [custom keyboard shortcuts](/office/dev/add-ins/design/keyboard-shortcuts), which require a shared runtime.

The following is an example of the **\<FunctionFile\>** element.

```XML
<DesktopFormFactor>
  <FunctionFile resid="Commands.Url" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- Information about this extension point. -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed. -->
</DesktopFormFactor>

...

<Resources>
    <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://www.contoso.com/commands.html" />
    </bt:Urls>

    <!-- Define other resources as needed. -->
</Resources>
```

The JavaScript in the HTML file indicated by the **\<FunctionFile\>** element must [initialize Office.js](/office/dev/add-ins/develop/initialize-add-in) and define named functions that take a single parameter: [event](/javascript/api/office/office.addincommands.event). It should also call [event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)) when it has finished execution. Functions in Outlook add-ins should use the [notification APIs](/javascript/api/outlook/office.notificationmessages) to indicate progress, success, or failure to the user. The name of the functions are used in the [FunctionName](action.md#functionname) element for function command buttons.

You can define and register the function specified by the **\<FunctionName\>** element in a separate JavaScript file that is loaded by the HTML file. The following is an example of such a file.

```js
// Initialize the Office Add-in.
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

// The command function.
async function highlightSelection(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
          await Excel.run(async (context) => {
              const range = context.workbook.getSelectedRange();
              range.format.fill.color = "yellow";
              await context.sync();
          });
      } catch (error) {
          // Note: In a production add-in, notify the user through your add-in's UI.
          console.error(error);
      }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

// You must register the function with the following line.
Office.actions.associate("highlightSelection", highlightSelection);
```

> [!IMPORTANT]
> The call to `event.completed` signals that you've successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls `event.completed`, the next queued call to that function runs. You must call `event.completed`; otherwise your function will not run.
