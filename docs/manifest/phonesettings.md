---
title: PhoneSettings element in the manifest file
description: The PhoneSettings element specifies the source location and control settings that apply when your mail add-in is used on a phone.
ms.date: 04/15/2024
ms.localizationpriority: medium
---

# PhoneSettings element

Specifies source location and control settings that apply when your mail add-in is used on a phone.

> [!IMPORTANT]
> The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server). To support Outlook on Android and iOS, see [Add-ins for Outlook on mobile devices](/office/dev/add-ins/outlook/outlook-mobile-addins).

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

- [Form](form.md)
