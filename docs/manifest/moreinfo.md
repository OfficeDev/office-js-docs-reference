---
title: MoreInfo element in the manifest file (preview)
description: The MoreInfo element specifies the custom text and URL that direct users to informational resources from the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 04/11/2024
ms.localizationpriority: medium
---

# MoreInfo element (preview)

Specifies a link with custom text and URL to provide informational resources to the users from the preprocessing dialog of a spam-reporting add-in in Outlook. The information provided in this element helps users identify and report unsolicited messages.

In the preprocessing dialog, the link is automatically prefixed with "For more info go to: ". It appears after the text provided in the [Description](preprocessingdialog.md#child-elements) child element of the **\<PreProcessingDialog\>** element.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox preview](../requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview.md)

## Contained in

- [PreProcessingDialog (preview)](preprocessingdialog.md)

## Attributes

None.

## Child elements

| Element | Required | Description |
| ------- | ------- | -------|
| **MoreInfoText** | Yes | Specifies the custom text of a link in the preprocessing dialog of a spam-reporting add-in. This link is used to provide additional information to help users report unsolicited messages. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element.<br><br>**Important**: In supported versions prior to Outlook on Windows Version 2404 (Build 17526.15020), the custom text provided in the **\<MoreInfoText\>** element isn't used as link text. Instead, the text is prepended to the URL provided in the **\<MoreInfoUrl\>** element. In these versions, the static string "For more info go to: " isn't added to the preprocessing dialog. |
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
