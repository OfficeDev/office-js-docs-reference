---
title: OfficeApp element in the manifest file
description: The OfficeApp element is the root element of an Office Add-in manifest.
ms.date: 11/06/2020
ms.localizationpriority: medium
---

# OfficeApp element

The root element in the manifest of an Office Add-in.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## Contained in

None

## Must contain

The **OfficeApp** element must contain the following child elements depending on the add-in type.

|Element|Content|Mail|TaskPane|
|:-----|:-----:|:-----:|:-----:|
|[Id](id.md)|Yes|Yes|Yes|
|[Version](version.md)|Yes|Yes|Yes|
|[ProviderName](providername.md)|Yes|Yes|Yes|
|[DefaultLocale](defaultlocale.md)|Yes|Yes|Yes|
|[DefaultSettings](defaultsettings.md)|Yes|No|Yes|
|[DisplayName](displayname.md)|Yes|Yes|Yes|
|[Description](description.md)|Yes|Yes|Yes|
|[FormSettings](formsettings.md)|No|Yes|No|
|[Permissions](permissions.md)|Yes|No|Yes|
|[Rule](rule.md)|No|Yes|No|

## Can contain

The **OfficeApp** element can contain the following child elements depending on the add-in type.

|Element|Content|Mail|TaskPane|
|:-----|:-----:|:-----:|:-----:|
|[AlternateId](alternateid.md)|Yes|Yes|Yes|
|[IconUrl](iconurl.md)|Yes|Yes|Yes|
|[HighResolutionIconUrl](highresolutioniconurl.md)|Yes|Yes|Yes|
|[SupportUrl](supporturl.md)|Yes|Yes|Yes|
|[AppDomains](appdomains.md)|Yes|Yes|Yes|
|[Hosts](hosts.md)|Yes|Yes|Yes|
|[Requirements](requirements.md)|Yes|Yes|Yes|
|[AllowSnapshot](allowsnapshot.md)|Yes|No|No|
|[Permissions](permissions.md)|No|Yes|No|
|[DisableEntityHighlighting](disableentityhighlighting.md)|No|Yes|No|
|[Dictionary](dictionary.md)|No|No|Yes|
|[VersionOverrides](versionoverrides.md)|Yes|Yes|Yes|
|[ExtendedOverrides](extendedoverrides.md)|No|No|Yes|

## Attributes

|Attribute|Description|
|:-----|:-----|
|xmlns|Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`|
