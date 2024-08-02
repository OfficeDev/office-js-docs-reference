---
title: Image Coercion requirement sets
description: Support for Image Coercion requirement sets with Office Add-ins across Excel, OneNote, PowerPoint, and Word.
ms.date: 04/15/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Image Coercion requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## ImageCoercion 1.1

ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method. The following applications are supported.

- Excel on Windows
  - Microsoft 365 subscription
  - perpetual Office 2016 and later
- Excel on Mac
- Excel on iPad
- OneNote on the web
- PowerPoint on the web
- PowerPoint on Windows
  - Microsoft 365 subscription
  - perpetual Office 2016 and later
- PowerPoint on Mac
- PowerPoint on iPad
- Word on the web
- Word on Windows
  - Microsoft 365 subscription
  - perpetual Office 2016 and later
- Word on Mac
- Word on iPad

## ImageCoercion 1.2

ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method. The following applications are supported.

- Excel on Windows
  - Microsoft 365 subscription
  - retail perpetual Office 2016 and later
  - volume-licensed perpetual Office 2021 and later
- Excel on Mac
- PowerPoint on the web
- PowerPoint on Windows
  - Microsoft 365 subscription
  - retail perpetual Office 2016 and later
  - volume-licensed perpetual Office 2021 and later
- PowerPoint on Mac
- Word on Windows
  - Microsoft 365 subscription
  - retail perpetual Office 2016 and later
  - volume-licensed perpetual Office 2021 and later
- Word on Mac

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office applications and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins manifest](/office/dev/add-ins/develop/add-in-manifests)
