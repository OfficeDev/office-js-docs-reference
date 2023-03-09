---
title: Excel JavaScript API requirement set 1.9
description: Details about the ExcelApi 1.9 requirement set.
ms.date: 04/01/2021
ms.topic: whats-new
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.9

More than 500 new Excel APIs were introduced with the 1.9 requirement set. The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Shapes](/office/dev/add-ins/excel/excel-add-ins-shapes) | Insert, position, and format images, geometric shapes and text boxes. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Auto Filter](/office/dev/add-ins/excel/excel-add-ins-worksheets#filter-data) | Add filters to ranges. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges) | Support for discontinuous ranges. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Special Cells](/office/dev/add-ins/excel/excel-add-ins-multiple-ranges#get-special-cells-from-multiple-ranges) | Get cells containing dates, comments, or formulas within a range. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](/office/dev/add-ins/excel/excel-add-ins-ranges-string-match) | Find values or formulas within a range or worksheet. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copy and Paste](/office/dev/add-ins/excel/excel-add-ins-ranges-cut-copy-paste) | Copy values, formats, and formulas from one range to another. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](/office/dev/add-ins/excel/performance#suspend-calculation-temporarily) | Greater control over the Excel calculation engine. | [Application](/javascript/api/excel/excel.application) |
| New Charts | Explore our new supported chart types: maps, box and whisker, waterfall, sunburst, pareto. and funnel. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | New capabilities with range formats. | [Range](/javascript/api/excel/excel.rangeformat) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.9. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.9 or earlier, see [Excel APIs in requirement set 1.9 or earlier](/javascript/api/excel?view=excel-js-1.9&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_9.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
