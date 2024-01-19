---
title: FormSettings element in the manifest file
description: Specifies source location and control settings for your mail add-in.
ms.date: 01/18/2024
ms.localizationpriority: medium
---

# FormSettings element

Specifies source location and control settings for your mail add-in in older Outlook clients that only support up to [Mailbox requirement set 1.2](../requirement-sets/outlook/requirement-set-1.2/outlook-requirement-set-1.2.md).

> [!NOTE]
> Because the **\<FormSettings\>** element is required for manifest validation, it must be defined in all mail add-ins, including those that support Mailbox requirement set 1.3 or later. **\<FormSettings\>** is ignored when your manifest contains a [VersionOverrides](versionoverrides.md) element.

**Add-in type:** Mail

## Syntax

```XML
<FormSettings>
    <Form xsi:type="ItemRead">
        ...
    </Form>
</FormSettings>
```

## Contained in

- [OfficeApp](officeapp.md)

## Can contain

- [Form](form.md)
