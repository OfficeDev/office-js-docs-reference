---
title: EquivalentAddin element in the manifest file
description: Specifies backwards compatibility for an equivalent COM or VSTO add-in or XLL.
ms.date: 07/12/2025
ms.localizationpriority: medium
---

# EquivalentAddin element

Specifies backwards compatibility for an equivalent COM add-in, VSTO add-in, or XLL.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Add-in type:** Task pane, Mail, Custom function

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Syntax

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## Contained in

- [EquivalentAddins](equivalentaddins.md)

## Must contain

- [Type](type.md)

## Can contain

- [ProgId](progid.md)
- [FileName](filename.md)

## Remarks

To specify a COM or VSTO add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements. 

> [!NOTE]
>
> - Although the term "ProgId" is usually associated with only COM add-ins, in the manifest it refers to the name of either a COM or VSTO add-in.
> - Use `COM` as the value of the `Type` element for both COM and VSTO add-ins.

To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.

## See also

- [Make your custom functions compatible with XLL user-defined functions](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Make your Office Add-in compatible with an existing COM VSTO add-in](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)