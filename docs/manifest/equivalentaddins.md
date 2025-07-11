---
title: EquivalentAddins element in the manifest file
description: Specifies backwards compatibility with one or more equivalent COM or VSTO add-ins or XLLs.
ms.date: 07/12/2025
ms.localizationpriority: medium
---

# EquivalentAddins element

Specifies compatibility with one or more equivalent COM add-ins, VSTO add-ins, or XLLs.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Add-in type:** Task pane, Mail, Custom function

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Syntax

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## Contained in

- [VersionOverrides](versionoverrides.md)

## Must contain

- [EquivalentAddin](equivalentaddin.md)

## See also

- [Make your custom functions compatible with XLL user-defined functions](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Make your Office Add-in compatible with an existing COM or VSTO add-in](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)