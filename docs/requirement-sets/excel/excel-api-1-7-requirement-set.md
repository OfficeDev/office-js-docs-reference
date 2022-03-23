---
title: Excel JavaScript API requirement set 1.7
description: 'Details about the ExcelApi 1.7 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.7

The Excel JavaScript API requirement set 1.7 features include APIs for charts, events, worksheets, ranges, document properties, named items, protection options and styles.

## Customize charts

With the new chart APIs, you can create additional chart types, add a data series to a chart, set the chart title, add an axis title, add display unit, add a trendline with moving average, change a trendline to linear, and more. The following are some examples.

- Chart axis - get, set, format and remove axis unit, label and title in a chart.
- Chart series - add, set, and delete a series in a chart.  Change series markers, plot orders and sizing.
- Chart trendlines - add, get, and format trendlines in a chart.
- Chart legend - format the legend font in a chart.
- Chart point - set chart point color.
- Chart title substring -  get and set title substring for a chart.
- Chart type - option to create more chart types.

## Events

Excel events APIs provide a variety of event handlers that allow your add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. For a list of events that are currently available, see [Work with Events using the Excel JavaScript API](/office/dev/add-ins/excel/excel-add-ins-events.md).

## Customize the appearance of worksheets and ranges

Using the new APIs, you can customize the appearance of worksheets in multiple ways:

- Freeze panes to keep specific rows or columns visible when you scroll in the worksheet. For example, if the first row in your worksheet contains headers, you might freeze that row so that the column headers will remain visible as you scroll down the worksheet.
- Modify the worksheet tab color.
- Add worksheet headings.

You can customize the appearance of ranges in multiple ways:

- Set the cell style for a range to ensure sure that all cells in the range have consistent formatting. A cell style is a defined set of formatting characteristics, such as fonts and font sizes, number formats, cell borders, and cell shading. Use any of Excel's built-in cell styles or create your own custom cell style.
- Set the text orientation for a range.
- Add or modify a hyperlink on a range that links to another location in the workbook or to an external location.

## Manage document properties

Using the document properties APIs, you can access built-in document properties and also create and manage custom document properties to store state of the workbook and drive workflow and business logic.

## Copy worksheets

Using the worksheet copy APIs, you can copy the data and format from one worksheet to a new worksheet within the same workbook and reduce the amount of data transfer needed.

## Handle ranges with ease

Using the various range APIs, you can do things such as get the surrounding region, get a resized range, and more. These APIs should make tasks like range manipulation and addressing much more efficient.

In addition:

- Workbook and worksheet protection options - use these APIs to protect data in a worksheet and the workbook structure.
- Update a named item - use this API to update a named item.
- Get active cell  - use this API to get the active cell of a workbook.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.7. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.7 or earlier, see [Excel APIs in requirement set 1.7 or earlier](/javascript/api/excel?view=excel-js-1.7&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_7.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
