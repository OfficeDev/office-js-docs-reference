---
title: EquivalentAddins element in the manifest file
description: Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.
ms.date: 06/12/2025
ms.localizationpriority: medium
---

# EquivalentAddins element

Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

**Add-in type:** Task pane, Mail, Custom function

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

> [!NOTE]
> Some child elements are not valid in the Mail schemas. See [Can contain](#can-contain).

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

## Can contain

The **\<EquivalentAddins\>** element can contain the following child element.

|Element|Content|Mail|TaskPane|
|:-----|:-----:|:-----:|:-----:|
|[Effect](#effect)|No|No|Yes|

### Effect

Specifies either that the COM add-in is disabled and hidden (instead of the Office Web Add-in) when they conflict, or specifies that the user chooses which to disable and hide. There are two possible values.

- **DisableWithNotification**: All of the COM add-ins specified in the child **\<EquivalentAddin\>** elements will be disabled and hidden. A popup dialog notifies the user that this happening.
- **UserOptionToDisable**: The user is prompted to choose whether to disable and hide COM add-ins specified in the child **\<EquivalentAddin\>** elements or to disable and hide the Office Add-in.

> [!NOTE]
> If the **\<Effect\>** element is not present, the COM add-ins are enabled and the Office Add-in is disabled and hidden on the Windows computer. 

The following is an example. The **\<Effect\>** element must be after all the **\<EquivalentAddin\>** elements.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
    <Effect>UserOptionToDisable</Effect>
  </EquivalentAddins>
</VersionOverrides>
```
> [!IMPORTANT]
> The **\<Effect\>** element is not available in Outlook.

## See also

- [Make your custom functions compatible with XLL user-defined functions](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Make your Office Add-in compatible with an existing COM add-in](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)