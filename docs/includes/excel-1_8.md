| Class | Fields | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|The operator to use for validating the data.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|Represents chart axes.|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|Represents the data labels on the chart.|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|Encapsulates the format properties for the chart area.|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|Represents the legend for the chart.|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|Occurs when the chart is activated.|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|Occurs when the chart is deactivated.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|Gets the ID of the chart that is activated.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is activated.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|Gets the ID of the chart that is added to the worksheet.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is added.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|Represents the border format of chart area, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|Represents the font attributes (font name, font size, color, etc.) for the current object.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|Specifies the alignment for the specified axis tick label.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|Specifies the group for the specified axis.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|Specifies the base unit for the specified category axis.|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|Specifies the category axis type.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|Specifies the custom axis display unit value.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|Represents the axis display unit.|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|Specifies the height, in points, of the chart axis.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|Specifies if the value axis crosses the category axis between categories.|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|Specifies the distance, in points, from the left edge of the axis to the left of chart area.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|Specifies chart fill formatting.|
||[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|Specifies the font attributes (font name, font size, color, etc.) for a chart axis element.|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|Specifies chart line formatting.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[border](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|Specifies the chart axis title's border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|Specifies the chart axis title's fill formatting.|
||[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|Clear the border format of a chart element.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|Occurs when a chart is activated.|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|Occurs when a chart is deactivated.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|Occurs when a chart is deleted.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|Specifies if the data label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|Represents the format of chart data label.|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|String value that represents the formula of chart data label using A1-style notation.|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|Returns the height, in points, of the chart data label.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|Represents the horizontal alignment for chart data label.|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[border](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|Represents the fill format of the current chart data label.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|Represents the font attributes (such as font name, font size, and color) for a chart data label.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|Specifies if data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|Specifies the horizontal alignment for chart data label.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|Gets the ID of the chart that is deactivated.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deactivated.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|Gets the ID of the chart that is deleted from the worksheet.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deleted.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|Specifies the height of the legend entry on the chart legend.|
||[index](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|Specifies the index of the legend entry in the chart legend.|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|Specifies the left value of a chart legend entry.|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|Specifies the top of a chart legend entry.|
||[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|Represents the visibility of a chart legend entry.|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|Represents the width of the legend entry on the chart Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[border](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|Represents the font attributes such as font name, font size, and color of a chart legend.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|Specifies the formatting of a chart plot area.|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|Specifies the height value of a plot area.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|Specifies the inside height value of a plot area.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|Specifies the inside left value of a plot area.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|Specifies the inside top value of a plot area.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|Specifies the inside width value of a plot area.|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|Specifies the left value of a plot area.|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|Specifies the position of a plot area.|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|Specifies the top value of a plot area.|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|Specifies the width value of a plot area.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[border](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|Specifies the border attributes of a chart plot area.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|Specifies the fill format of an object, which includes background formatting information.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|Represents the number of periods that the trendline extends backward.|
||[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|Delete the trendline object.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|Represents the number of periods that the trendline extends forward.|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|Represents the intercept value of the trendline.|
||[label](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|Represents the label of a chart trendline.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|Represents the period of a chart trendline.|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|Represents the name of the trendline.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|Represents the order of a chart trendline.|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|True if the r-squared value for the trendline is displayed on the chart.|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|Represents the type of a chart trendline.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|Specifies if the trendline label automatically generates appropriate text based on context.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|The format of the chart trendline label.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|String value that represents the formula of the chart trendline label using A1-style notation.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|Returns the height, in points, of the chart trendline label.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|Represents the horizontal alignment of the chart trendline label.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[border](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|Specifies the border format, which includes color, linestyle, and weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|Specifies the fill format of the current chart trendline label.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|A custom data validation formula.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[field](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|ID of the DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|Name of the DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|Number format of the DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|Position of the DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|Specifies if the data should be shown as a specific summary calculation.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|Specifies if all items of the DataPivotHierarchy are shown.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|Gets a DataPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|Gets a DataPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|Clears the data validation from the current range.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|Error alert when user enters invalid data.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|Specifies if data validation will be performed on blank cells.|
||[prompt](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|Data validation rule that contains different type of data validation criteria.|
||[type](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|Type of the data validation, see `Excel.DataValidationType` for details.|
||[valid](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|Represents if all cell values are valid according to the data validation rules.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|Represents the error alert message.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|Specifies whether to show an error alert dialog when a user enters invalid data.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|The data validation alert type, please see `Excel.DataValidationAlertStyle` for details.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|Represents the error alert dialog title.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|Specifies the message of the prompt.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|Specifies if a prompt is shown when a user selects a cell with data validation.|
||[title](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|Specifies the title for the prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|Custom data validation criteria.|
||[date](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|Date data validation criteria.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|Decimal data validation criteria.|
||[list](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|List data validation criteria.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|Text length data validation criteria.|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|Time data validation criteria.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|Whole number data validation criteria.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|Determines whether to allow multiple filter items.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|ID of the FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|Name of the FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|Position of the FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|Gets a FilterPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|Gets a FilterPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|Specifies whether to display the list in a cell drop-down.|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|Source of the list for data validation|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|ID of the PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|Returns the PivotItems associated with the PivotField.|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|Name of the PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|Determines whether to show all items of the PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|Subtotals of the PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|Gets a PivotField by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|Gets a PivotField by name.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[fields](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|ID of the PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|Name of the PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|Gets a PivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|Gets a PivotHierarchy by name.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|ID of the PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|Name of the PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|Specifies if the PivotItem is visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|Gets the number of PivotItems in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|Gets a PivotItem by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|Gets a PivotItem by name.|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|The Data Pivot Hierarchies of the PivotTable.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|The Pivot Hierarchies of the PivotTable.|
||[layout](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|The Row Pivot Hierarchies of the PivotTable.|
||[worksheet](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-worksheet-member)|The worksheet containing the current PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.|
||[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|Gets the number of pivot tables in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitem-member(1))|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|Gets a PivotTable by name.|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|Specifies the range reference in A1-style.|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|Represents the range reference for the specified range in the language of the user.|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|Specifies the number of cells in the range.|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|Specifies the total number of columns in the range.|
||[columnHidden](/javascript/api/excel/excel.range#excel-excel-range-columnhidden-member)|Represents if all columns in the current range are hidden.|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|Specifies the column number of the first cell in the range.|
||[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|Returns a data validation object.|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Returns a format object, encapsulating the range's font, fill, borders, alignment, and other properties.|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[formulasR1C1](/javascript/api/excel/excel.range#excel-excel-range-formulasr1c1-member)|Represents the formula in R1C1-style notation.|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|Gets the last cell within the range.|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|Gets the last column within the range.|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|Gets the last row within the range.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|Returns a `Range` object that represents the surrounding region for the top-left cell in this range.|
||[getVisibleView()](/javascript/api/excel/excel.range#excel-excel-range-getvisibleview-member(1))|Represents the visible rows of the current range.|
||[sort](/javascript/api/excel/excel.range#excel-excel-range-sort-member)|Represents the range sort of the current range.|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|The worksheet containing the current range.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|Represents the URL target for the hyperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|Represents the document reference target for the hyperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|Represents the string displayed when hovering over the hyperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|Represents the string that is displayed in the top left most cell in the range.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|ID of the RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|Name of the RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|Position of the RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|Gets a RowColumnPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|Gets a RowColumnPivotHierarchy by name.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[calculation](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|The `ShowAs` calculation to use for the PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|Specifies if the style is a built-in style.|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|Deletes this style.|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|Specifies if the formula will be hidden when the worksheet is protected.|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|Represents the horizontal alignment for the style.|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|Specifies if the style includes the auto indent, horizontal alignment, vertical alignment, wrap text, indent level, and text orientation properties.|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|Specifies if the style includes the color, color index, line style, and weight border properties.|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|Specifies if the style includes the background, bold, color, color index, font style, italic, name, size, strikethrough, subscript, superscript, and underline font properties.|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|Specifies if the style includes the number format property.|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|Specifies if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties.|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|Specifies if the style includes the formula hidden and locked protection properties.|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|An integer from 0 to 250 that indicates the indent level for the style.|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|Specifies if the object is locked when the worksheet is protected.|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|The name of the style.|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|The format code of the number format for the style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|The localized format code of the number format for the style.|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|The reading order for the style.|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|Specifies if text automatically shrinks to fit in the available column width.|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|The text orientation for the style.|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|Specifies the vertical alignment for the style.|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|Specifies if Excel wraps the text in the object.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.|
||[average](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[product](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[variance](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[ThreeArrowsGraySet](/javascript/api/excel/excel.threearrowsgrayset)|[grayDownArrow](/javascript/api/excel/excel.threearrowsgrayset#excel-excel-threearrowsgrayset-graydownarrow-member)||
||[graySideArrow](/javascript/api/excel/excel.threearrowsgrayset#excel-excel-threearrowsgrayset-graysidearrow-member)||
||[grayUpArrow](/javascript/api/excel/excel.threearrowsgrayset#excel-excel-threearrowsgrayset-grayuparrow-member)||
|[ThreeArrowsSet](/javascript/api/excel/excel.threearrowsset)|[greenUpArrow](/javascript/api/excel/excel.threearrowsset#excel-excel-threearrowsset-greenuparrow-member)||
||[redDownArrow](/javascript/api/excel/excel.threearrowsset#excel-excel-threearrowsset-reddownarrow-member)||
||[yellowSideArrow](/javascript/api/excel/excel.threearrowsset#excel-excel-threearrowsset-yellowsidearrow-member)||
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[enableEvents](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-enableevents-member)|Toggle JavaScript events in the current task pane or content add-in.|
