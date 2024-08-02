---
title: Runtimes in the manifest file 
description: The Runtimes element specifies your add-in's runtime.
ms.date: 07/14/2023
ms.localizationpriority: medium
---

# Runtimes element

Specifies the runtime of your add-in. Child of the [Host](host.md) element.

> [!NOTE]
> When running in Office on Windows, an add-in that has a **\<Runtimes\>** element in its manifest does not necessarily run in the same webview control as it otherwise would. For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a **\<Runtimes\>** element. However, when those conditions are not met, an add-in with a **\<Runtimes\>** element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [SharedRuntime 1.1](../requirement-sets/common/shared-runtime-requirement-sets.md) (Only when used in a task pane add-in.)
- [Mailbox 1.10 and later](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (When used in an Outlook add-in that implements [event-based activation](/office/dev/add-ins/outlook/autolaunch).)
- [Mailbox preview](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview) (When used in an Outlook add-in that implements the [integrated spam reporting (preview)](/office/dev/add-ins/outlook/spam-reporting) feature.)

[!include[Runtimes support](../includes/runtimes-note.md)]

## Syntax

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## Contained in

[Host](host.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Yes |  The runtime for your add-in. **Important**: At present, you can only define one **\<Runtime\>** element. |

## See also

- [Runtime](runtime.md)
- [Runtimes in Office Add-ins](/office/dev/add-ins/testing/runtimes)
- [Configure your Office Add-in to use a shared JavaScript runtime](/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
- [Configure your Outlook add-in for event-based activation](/office/dev/add-ins/outlook/autolaunch)
- [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting)
