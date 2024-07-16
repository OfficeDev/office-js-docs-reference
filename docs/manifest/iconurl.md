---
title: IconUrl element in the manifest file
description: The IconUrl element specifies the URL of the image that represents your Office Add-in in the insertion UX, AppSource, and the tab bar.
ms.date: 06/19/2024
ms.localizationpriority: medium
---

# IconUrl element

Specifies the full, absolute URL of the image that is used to represent your Office Add-in in the insertion UX, [AppSource](https://appsource.microsoft.com), and the vertical task pane tab bar.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<IconUrl DefaultValue="string" />
```

## Can contain

- [Override](override.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----:|:-----:|:-----|
|DefaultValue|string|Yes|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.|

## Remarks

The image must be in one of the following file formats. 

- BMP
- EXIF
- GIF
- JPG
- PNG
- TIFF

The image resolution requirements are as follows.

| Add-in type | Resolution (pixels) |
|-------------|---------------------|
| Content     | 32 x 32             |
| Mail        | 64 x 64             |
| Task pane   | 32 x 32             |

You should also specify an icon for use with Office client applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element. For more information, see the section _Create a consistent visual identity_ in [Create effective listings in AppSource and within Office](/partner-center/marketplace-offers/create-effective-office-store-listings#create-a-consistent-visual-identity).

For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. 

The image is also used on the vertical tab bar for the task pane when more than one task pane is open. The tab bar appears beside the task pane whenever a second task pane is opened, regardless of whether it is a task pane in the same add-in or a different add-in. The following image shows the tab bar when the [Script Lab](/office/dev/add-ins/overview/explore-with-script-lab) add-in and another add-in have both been started and both the **Code** and **Run** task panes of Script Lab have been opened.

:::image type="content" source="/javascript/api/images/built-in-vertical-tab-bar.png" alt-text="The top inch of a task pane, with three square tabs to the right of the upper right corner. Two of the tabs have the Script Lab icon. The third has an icon for a different add-in.":::

For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.

Changing the value of the `IconUrl` element at runtime is not currently supported.

## Examples

```XML
<IconUrl DefaultValue="https://localhost:3000/assets/images/icon-32.png" />
```

```XML
<IconUrl DefaultValue="https://script-lab.azureedge.net/assets/images/icon-32.png" />
```