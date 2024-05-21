---
title: MoreInfo element in the manifest file
description: The MoreInfo element specifies the custom text and URL that direct users to informational resources from the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 05/20/2024
ms.localizationpriority: medium
---

# MoreInfo element

Specifies the custom text and URL to provide informational resources to the users from the preprocessing dialog of a spam-reporting add-in in Outlook. The information provided in this element helps users identify and report unsolicited messages.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.14](../requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14.md)

## Contained in

- [PreProcessingDialog](preprocessingdialog.md)

## Attributes

None.

## Child elements

| Element | Required | Description |
| ------- | ------- | -------|
| **MoreInfoText** | Yes | Specifies additional information in the preprocessing dialog of a spam-reporting add-in to help users report unsolicited messages. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **MoreInfoUrl** | Yes | Specifies the URL of a site containing informational resources in the preprocessing dialog of a spam-reporting add-in. Its **resid** attribute must be set to the value of the **id** attribute of a [Url](url.md) in the [Urls](urls.md) element under the [Resources](resources.md) element. |

## Example

```xml
<PreProcessingDialog>
  ...
  <MoreInfo>
    <MoreInfoText resid="MoreInfo.Label"/>
    <MoreInfoUrl resid="MoreInfo.Url"/>
  </MoreInfo>
</PreProcessingDialog>
```
