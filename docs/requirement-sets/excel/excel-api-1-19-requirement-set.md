---
title: Excel JavaScript API requirement set 1.19
description: Details about the ExcelApi 1.19 requirement set.
ms.date: 05/13/2025
ms.topic: whats-new
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.19

The ExcelApi 1.19 adds capabilities for charts and shapes, to help you better visualize your data in Excel. It also includes updates to the data types feature, such as support for [linked data types](/office/dev/add-ins/excel/excel-data-types-linked-entity-cell-values), [dot notation](/office/dev/add-ins/excel/excel-add-ins-dot-functions), and expanded options for [basic cell values](/office/dev/add-ins/excel/excel-data-types-add-properties-to-basic-cell-values).

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Chart labels | Control the appearance of chart labels. | [ChartDataLabel](/javascript/api/excel/excel.chartdatalabel), [ChartDataLabelAnchor](/javascript/api/excel/excel.chartdatalabelanchor), [ChartLeaderLines](/javascript/api/excel/excel.chartleaderlines), [ChartLeaderLinesFormat](/javascript/api/excel/excel.chartleaderlinesformat) |
| Linked data types | Adds support for data types connected to Excel from external sources. To learn more, see [Create linked entity cell values](/office/dev/add-ins/excel/excel-data-types-linked-entity-cell-values) | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.19. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.19 or earlier, see [Excel APIs in requirement set 1.19 or earlier](/javascript/api/excel?view=excel-js-1.19&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_19.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.19&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
