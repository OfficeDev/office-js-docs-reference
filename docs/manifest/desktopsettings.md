---
title: DesktopSettings element in the manifest file
description: Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.
ms.date: 04/15/2024
ms.localizationpriority: medium
---

# DesktopSettings element

Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.

> [!IMPORTANT]
> The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server).

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

## Child elements

| Element | Required | Description |
|:-----|:-----:|:-----|
| [SourceLocation element](sourcelocation.md) | Yes | The location of your add-in's source files. |
| [RequestedHeight element](requestedheight.md) | Yes | The initial height of your add-in.<br><br>**Important**: The **\<RequestedHeight\>** element is only required if the `xsi:type` of the parent **\<Form\>** element is `ItemRead`. |
