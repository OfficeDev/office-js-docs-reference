---
title: Office Common API requirement sets
description: Learn more about the Office Common API requirement sets.
ms.date: 02/29/2024
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

> [!IMPORTANT]
> We no longer recommend that you create and use Access web apps and databases in SharePoint. As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.

## Common API requirement sets

The following sections list the Common API requirement sets, the methods in each set, and the Office client applications that support that requirement set. All of these API requirement sets are version 1.1, unless otherwise specified.

> [!TIP]
> Need information about where add-ins and requirement sets are supported by Office application and version? See [Office client application and platform availability for Office Add-ins](/office/dev/add-ins/overview/office-add-in-availability).

### ActiveView

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac | - Document.getActiveViewAsync |

---

### AddInCommands

See [Add-in command requirement sets](add-in-commands-requirement-sets.md).

---

### BindingEvents

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Access Web Apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Binding.addHandlerAsync<br>- Binding.removeHandlerAsync |

---

### CompressedFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2016<br>- Excel on the web<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports output to Office Open XML (OOXML) format as a byte array<br>(Office.FileType.Compressed) when using the Document.getFileAsync method. |

---

### CustomXmlParts

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - CustomXmlNode.getNodesAsync<br>- CustomXmlNode.getNodeValueAsync<br>- CustomXmlNode.getTextAsync<br>- CustomXmlNode.getXmlAsync<br>- CustomXmlNode.setNodeValueAsync<br>- CustomXmlNode.setTextAsync<br>- CustomXmlNode.setXmlAsync<br>- CustomXmlPart.addHandlerAsync<br>- CustomXmlPart.deleteAsync<br>- CustomXmlPart.getNodesAsync<br>- CustomXmlPart.getXmlAsync<br>- CustomXmlPart.removeHandlerAsync<br>- CustomXmlParts.addAsync<br>- CustomXmlParts.getByIdAsync<br>- CustomXmlParts.getByNamespaceAsync<br>- CustomXmlPrefixMappings.addNamespaceAsync<br>- CustomXmlPrefixMappings.getNamespaceAsync<br>- CustomXmlPrefixMappings.getPrefixAsync |

---

### DialogApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Dialog API requirement sets](dialog-api-requirement-sets.md). | - UI.messageParent<br>- UI.displayDialogAsync<br>- UI.closeContainer<br>- UI.Dialog |

---

### DialogOrigin

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Dialog Origin requirement sets](dialog-origin-requirement-sets.md). | Cross-domain support for:<br>- UI.messageParent<br>- UI.Dialog.messageChild |

---

### DocumentEvents

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- OneNote on the web<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Document.addHandlerAsync<br>- Document.removeHandlerAsync |

---

### File

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Document.getFileAsync<br>- File.closeAsync<br>- File.getSliceAsync |

---

### HtmlCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - OneNote on the web<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports coercion to HTML (Office.CoercionType.Html) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### IdentityAPI

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Identity API requirement sets](identity-api-requirement-sets.md). | - Auth.getAccessToken |

---

### ImageCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Image Coercion requirement sets](image-coercion-requirement-sets.md). | - Document.setSelectedDataAsync |

---

### KeyboardShortcuts

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Keyboard Shortcuts requirement sets](keyboard-shortcuts-requirement-sets.md). | - Office.actions.areShortcutsInUse<br>- Office.actions.getShortcuts<br>- Office.actions.replaceShortcuts |

---

### Mailbox

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Outlook on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- [new Outlook on Windows (preview)](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)<br>- Outlook on the web<br>- Outlook on Android<br>- Outlook on Mac<br>- Outlook on iOS | See [Understanding Outlook API requirement sets](../outlook/outlook-api-requirement-sets.md). |

---

### MatrixBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on the web<br>- Word on iPad<br>- Word on Mac | - Bindings.addFromNamedItemAsync<br>- Bindings.addFromSelectionAsync<br>- Bindings.getAllAsync<br>- Bindings.getByIdAsync<br>- Bindings.releaseByIdAsync<br>- Binding.getDataAsync<br>- Binding.setDataAsync |

---

### MatrixCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports coercion to the "matrix" (array of arrays) data structure (Office.CoercionType.Matrix) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### OoxmlCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports coercion to Open Office XML (OOXML) format (Office.CoercionType.Ooxml) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### OpenBrowserWindowApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Open Browser Window API requirement sets](open-browser-window-api-requirement-sets.md). | - Office.context.ui.openBrowserWindow |

---

### PartialTableBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| Access Web Apps ||

---

### PdfFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports output to PDF format (Office.FileType.Pdf)<br>when using the Document.getFileAsync method. |

---

### RibbonApi

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Ribbon API requirement sets](ribbon-api-requirement-sets.md). | - Office.ribbon.requestUpdate |

---

### Selection

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Project on Windows<br>-- volume-licensed perpetual Office 2013<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Document.getSelectedDataAsync<br>- Document.setSelectedDataAsync |

---

### Settings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Access Web Apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- OneNote on the web<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Settings.get<br>- Settings.remove<br>- Settings.saveAsync<br>- Settings.set |

---

### SharedRuntime

| Minimum Office application support | Methods in set |
|:-----|:-----|
| See [Shared runtime requirement sets](shared-runtime-requirement-sets.md). | - Office.addin.getStartupBehavior<br>- Office.addin.hide<br>- Office.addin.onVisibilityModeChanged<br>- Office.addin.setStartupBehavior<br>- Office.addin.showAsTaskpane<br>- Office.BeforeDocumentCloseNotification.disable<br>- Office.BeforeDocumentCloseNotification.enable<br>- Office.BeforeDocumentCloseNotification.onCloseActionCancelled |

---

### TableBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Access Web Apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Bindings.addFromNamedItemAsync<br>- Bindings.addFromSelectionAsync<br>- Bindings.getAllAsync<br>- Bindings.getByIdAsync<br>- Bindings.releaseByIdAsync<br>- Binding.addColumnsAsync<br>- Binding.addRowsAsync<br>- Binding.deleteAllDataValuesAsync<br>- Binding.getDataAsync<br>- Binding.setDataAsync |

---

### TableCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Access Web Apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports coercion to the "table" data structure (Office.CoercionType.Table) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### TextBindings

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | - Bindings.addFromNamedItemAsync<br>- Bindings.addFromSelectionAsync<br>- Bindings.getAllAsync<br>- Bindings.getByIdAsync<br>- Bindings.releaseByIdAsync<br>- Binding.getDataAsync<br>- Binding.setDataAsync |

---

### TextCoercion

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- OneNote on the web<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Project on Windows<br>-- volume-licensed perpetual Office 2013<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports coercion to text format (Office.CoercionType.Text) when reading and writing data using the Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync, or Binding.setDataAsync methods. |

---

### TextFile

| Minimum Office application support | Methods in set |
|:-----|:-----|
| - Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on Mac<br>- Word on the web<br>- Word on iPad | Supports output to text format (Office.FileType.Text) when using the Document.getFileAsync method. |

---

## Methods that aren't part of a requirement set

The following methods in the Office JavaScript API aren't part of a requirement set. If your add-in requires any of these methods, use the **\<Methods\>** and **\<Method\>** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an `if` statement. For more information, see [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).

| Method name | Minimum Office application support |
|:-----|:-----|
| Bindings.addFromPromptAsync | - Access web apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac |
| Document.getFilePropertiesAsync | - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on the web<br>- Word on iPad<br>- Word on Mac |
| Document.getProjectFieldAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getResourceFieldAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getSelectedResourceAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getSelectedTaskAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getSelectedViewAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getTaskAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.getTaskFieldAsync | - Project Standard 2013<br>- Project Professional 2013 |
| Document.goToByIdAsync | - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- PowerPoint on iPad<br>- PowerPoint on Mac<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on the web<br>- Word on iPad<br>- Word on Mac |
| Settings.addHandlerAsync | - Access web apps<br>- Excel on the web |
| Settings.refreshAsync | - Access web apps<br>- Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- PowerPoint on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- PowerPoint on the web<br>- Word on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Word on the web |
| Settings.removeHandlerAsync | - Access web apps<br>- Excel on the web |
| TableBinding.clearFormatsAsync | - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac |
| TableBinding.setFormatsAsync | - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac |
| TableBinding.setTableOptionsAsync | - Excel on Windows<br>-- Microsoft 365 subscription<br>-- perpetual Office 2013<br>- Excel on the web<br>- Excel on iPad<br>- Excel on Mac |

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests)
