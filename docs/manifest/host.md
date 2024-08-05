---
title: Host element in the manifest file
description: Specifies an individual Office application type where the add-in should activate.
ms.date: 04/15/2024
ms.localizationpriority: medium
---

# Host element

Specifies an individual Office application type where the add-in should activate.

> [!IMPORTANT]
> The **\<Host\>** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node. However, the functionality is the same.  

## Basic manifest

When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.

### Attributes

| Attribute     | Type   | Required | Description                                      |
|:--------------|:-------:|:---------:|:-------------------------------------------------|
| [Name](#name) | string | Yes | The name of the type of Office client application. |

### Name

Specifies the Host type targeted by this add-in. The value must be one of the following:

- `Document` (Word)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)
- `Database` (Access (obsolete))

### Example

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## VersionOverrides node

When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute. 

This element overrides the **\<Hosts\>** element in the basic manifest.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

### Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [xsi:type](#xsitype)  |  Yes  | Specifies the Office application where these settings apply.|

### Child elements

|  Element |  Required  |  Description  |
|:-----|:-----:|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Yes   |  Defines the settings for the desktop form factor. |
|  [MobileFormFactor](mobileformfactor.md)    |  No   |  Defines the settings for the mobile form factor. **Note:** This element is only supported in Outlook on Android and iOS. |
|  [AllFormFactors](allformfactors.md)    |  No   |  Defines the settings for all form factors. Only used by custom functions in Excel. |
|  [Runtimes](runtimes.md)    |  No   |  Specifies the runtimes of your add-in. |

### xsi:type

Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply. The value must be one of the following:

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## Host example

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
