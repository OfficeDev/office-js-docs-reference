---
title: Form element in the manifest file
description: UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).
ms.date: 04/15/2024
ms.localizationpriority: medium
---

# Form element

UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).

> [!IMPORTANT]
> The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server).

**Add-in type:** Mail

## Syntax

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## Contained in

- [FormSettings](formsettings.md)

## Can contain

|**Element**|
|:-----|
|[DesktopSettings](desktopsettings.md)|
|[TabletSettings](tabletsettings.md)|
|[PhoneSettings](phonesettings.md)|

## Attributes

|Attribute|Required|Description|
|:-----|:-----:|:-----|
|xsi:type|Yes|Specifies where the add-in appears in Outlook. If your add-in should appear when a user reads messages or appointments, set the attribute to `ItemRead`. However, if your add-in should appear when a user composes a reply, creates a new message or appointment, or edits an existing appointment, set the attribute to `ItemEdit`.|
