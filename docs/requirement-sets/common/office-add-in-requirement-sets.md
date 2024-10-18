---
title: Office Common API requirement sets
description: Learn more about the Office Common API requirement sets.
ms.date: 10/17/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Office Common API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!TIP]
> Looking for the *application-specific* API requirement sets? See the following API requirement sets.
>
> - [Excel JavaScript API requirement sets](../excel/excel-api-requirement-sets.md) (ExcelApi)
> - [Word JavaScript API requirement sets](../word/word-api-requirement-sets.md) (WordApi)
> - [OneNote JavaScript API requirement sets](../onenote/onenote-api-requirement-sets.md) (OneNoteApi)
> - [PowerPoint JavaScript API requirement sets](../powerpoint/powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [Understanding Outlook API requirement sets](../outlook/outlook-api-requirement-sets.md) (Mailbox)

## Common API requirement sets

The following sections list the Common API requirement sets, the methods in each set, and the Office client applications that support that requirement set. All of these API requirement sets are version 1.1, unless otherwise specified.

> [!TIP]
> Need information about where add-ins and requirement sets are supported by Office application and version? See [Office client application and platform availability for Office Add-ins](/office/dev/add-ins/overview/office-add-in-availability).

### ActiveView

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li></ul> | <ul><li>Document.getActiveViewAsync</li></ul> |

---

### AddInCommands

See [Add-in command requirement sets](add-in-commands-requirement-sets.md).

---

### BindingEvents

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows<li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Binding.addHandlerAsync</li><li>Binding.removeHandlerAsync</li></ul> |

---

### CompressedFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports output to Office Open XML (OOXML) format as a byte array<br>(Office.FileType.Compressed) when using the Document.getFileAsync method. |

---

### CustomXmlParts

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>CustomXmlNode.getNodesAsync</li><li>CustomXmlNode.getNodeValueAsync</li><li>CustomXmlNode.getTextAsync</li><li>CustomXmlNode.getXmlAsync</li><li>CustomXmlNode.setNodeValueAsync</li><li>CustomXmlNode.setTextAsync</li><li>CustomXmlNode.setXmlAsync</li><li>CustomXmlPart.addHandlerAsync</li><li>CustomXmlPart.deleteAsync</li><li>CustomXmlPart.getNodesAsync</li><li>CustomXmlPart.getXmlAsync</li><li>CustomXmlPart.removeHandlerAsync</li><li>CustomXmlParts.addAsync</li><li>CustomXmlParts.getByIdAsync</li><li>CustomXmlParts.getByNamespaceAsync</li><li>CustomXmlPrefixMappings.addNamespaceAsync</li><li>CustomXmlPrefixMappings.getNamespaceAsync</li><li>CustomXmlPrefixMappings.getPrefixAsync</li></ul> |

---

### DevicePermissionService

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Device Permission Service requirement sets](device-permission-service-requirement-sets.md). | <ul><li>DevicePermission.requestPermissions</li><li>DevicePermission.requestPermissionsAsync</li></ul> |

---

### DialogApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Dialog API requirement sets](dialog-api-requirement-sets.md). | <ul><li>UI.messageParent</li><li>UI.displayDialogAsync</li><li>UI.closeContainer</li><li>UI.Dialog</li></ul> |

---

### DialogOrigin

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Dialog Origin requirement sets](dialog-origin-requirement-sets.md). | Cross-domain support for:<ul><li>UI.messageParent</li><li>UI.Dialog.messageChild</li></ul> |

---

### DocumentEvents

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>OneNote on the web</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Document.addHandlerAsync</li><li>Document.removeHandlerAsync</li></ul> |

---

### File

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Document.getFileAsync</li><li>File.closeAsync</li><li>File.getSliceAsync</li></ul> |

---

### HtmlCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>OneNote on the web</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports coercion to HTML (Office.CoercionType.Html) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### IdentityAPI

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Identity API requirement sets](identity-api-requirement-sets.md). | <ul><li>Auth.getAccessToken</li></ul> |

---

### ImageCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Image Coercion requirement sets](image-coercion-requirement-sets.md). | <ul><li>Document.setSelectedDataAsync</li></ul> |

---

### KeyboardShortcuts

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Keyboard Shortcuts requirement sets](keyboard-shortcuts-requirement-sets.md). | <ul><li>Office.actions.areShortcutsInUse</li><li>Office.actions.getShortcuts</li><li>Office.actions.replaceShortcuts</li></ul> |

---

### Mailbox

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Outlook on the web</li><li>[new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)</li><li>classic Outlook on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Outlook on Android</li><li>Outlook on Mac</li><li>Outlook on iOS</li></ul> | See [Understanding Outlook API requirement sets](../outlook/outlook-api-requirement-sets.md). |

---

### MatrixBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Bindings.addFromNamedItemAsync</li><li>Bindings.addFromSelectionAsync</li><li>Bindings.getAllAsync</li><li>Bindings.getByIdAsync</li><li>Bindings.releaseByIdAsync</li><li>Binding.getDataAsync</li><li>Binding.setDataAsync</li></ul> |

---

### MatrixCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports coercion to the "matrix" (array of arrays) data structure (Office.CoercionType.Matrix) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### NestedAppAuth

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Nested app auth requirement sets](nested-app-auth-requirement-sets.md). | <ul><li>Office.auth.getAuthContext</li></ul> |

---

### OoxmlCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports coercion to Open Office XML (OOXML) format (Office.CoercionType.Ooxml) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### OpenBrowserWindowApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Open Browser Window API requirement sets](open-browser-window-api-requirement-sets.md). | <ul><li>Office.context.ui.openBrowserWindow</li></ul> |

---

### PdfFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports output to PDF format (Office.FileType.Pdf) when using the Document.getFileAsync method. |

---

### RibbonApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Ribbon API requirement sets](ribbon-api-requirement-sets.md). | <ul><li>Office.ribbon.requestUpdate</li></ul> |

---

### Selection

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Project on Windows</li><ul><li>volume-licensed perpetual Office 2016</li></ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Document.getSelectedDataAsync</li><li>Document.setSelectedDataAsync</li></ul> |

---

### Settings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>OneNote on the web</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Settings.get</li><li>Settings.remove</li><li>Settings.saveAsync</li><li>Settings.set</li></ul> |

---

### SharedRuntime

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Shared runtime requirement sets](shared-runtime-requirement-sets.md). | <ul><li>Office.addin.getStartupBehavior</li><li>Office.addin.hide</li><li>Office.addin.onVisibilityModeChanged</li><li>Office.addin.setStartupBehavior</li><li>Office.addin.showAsTaskpane</li><li>Office.BeforeDocumentCloseNotification.disable</li><li>Office.BeforeDocumentCloseNotification.enable</li><li>Office.BeforeDocumentCloseNotification.onCloseActionCancelled</li></ul> |

---

### TableBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li> Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Bindings.addFromNamedItemAsync</li><li>Bindings.addFromSelectionAsync</li><li>Bindings.getAllAsync</li><li>Bindings.getByIdAsync</li><li>Bindings.releaseByIdAsync</li><li>Binding.addColumnsAsync</li><li>Binding.addRowsAsync</li><li>Binding.deleteAllDataValuesAsync</li><li>Binding.getDataAsync</li><li>Binding.setDataAsync</li></ul> |

---

### TableCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports coercion to the "table" data structure (Office.CoercionType.Table) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### TextBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | <ul><li>Bindings.addFromNamedItemAsync</li><li>Bindings.addFromSelectionAsync</li><li>Bindings.getAllAsync</li><li>Bindings.getByIdAsync</li><li>Bindings.releaseByIdAsync</li><li>Binding.getDataAsync</li><li>Binding.setDataAsync</li></ul> |

---

### TextCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on iPad</li><li>OneNote on the web</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Project on Windows</li><ul><li>volume-licensed perpetual Office 2016</li></ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports coercion to text format (Office.CoercionType.Text) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### TextFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| <ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> | Supports output to text format (Office.FileType.Text) when using the Document.getFileAsync method. |

---

## Methods that aren't part of a requirement set

The following methods in the Office JavaScript API aren't part of a requirement set. If your add-in requires any of these methods, use the **\<Methods\>** and **\<Method\>** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an `if` statement. For more information, see [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

| Method name | Minimum Office application support |
|:-----|:-----|
| Bindings.addFromPromptAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li></ul> |
| Document.getFilePropertiesAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> |
| Document.getProjectFieldAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getResourceFieldAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getSelectedResourceAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getSelectedTaskAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getSelectedViewAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getTaskAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.getTaskFieldAsync | <ul><li>Project Standard 2016</li><li>Project Professional 2016</li></ul> |
| Document.goToByIdAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on Mac</li><li>PowerPoint on iPad</li><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on Mac</li><li>Word on iPad</li></ul> |
| Settings.addHandlerAsync | <ul><li>Excel on the web</li></ul> |
| Settings.refreshAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>PowerPoint on the web</li><li>PowerPoint on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Word on the web</li><li>Word on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul> |
| Settings.removeHandlerAsync | <ul><li>Excel on the web</li></ul> |
| TableBinding.clearFormatsAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li></ul> |
| TableBinding.setFormatsAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li></ul> |
| TableBinding.setTableOptionsAsync | <ul><li>Excel on the web</li><li>Excel on Windows</li><ul><li>Microsoft 365 subscription</li><li>perpetual Office 2016</li></ul><li>Excel on Mac</li><li>Excel on iPad</li></ul> |

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
