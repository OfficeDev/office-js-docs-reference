| Class | Fields | Description |
|:---|:---|:---|
|*global*|[createWorkbook(base64?: string)](/#excel-javascript/api/excel/-createworkbook-function(1))|Creates and opens a new workbook.|
|[BasicDataValidation](/.basicdatavalidation)|[formula1](/.basicdatavalidation#excel-javascript/api/excel/-basicdatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/.basicdatavalidation#excel-javascript/api/excel/-basicdatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/.basicdatavalidation#excel-javascript/api/excel/-basicdatavalidation-operator-member)|The operator to use for validating the data.|
|[Chart](/.chart)|[categoryLabelLevel](/.chart#excel-javascript/api/excel/-chart-categorylabellevel-member)|Specifies a chart category label level enumeration constant, referring to the level of the source category labels.|
||[displayBlanksAs](/.chart#excel-javascript/api/excel/-chart-displayblanksas-member)|Specifies the way that blank cells are plotted on a chart.|
||[onActivated](/.chart#excel-javascript/api/excel/-chart-onactivated-member)|Occurs when the chart is activated.|
||[onDeactivated](/.chart#excel-javascript/api/excel/-chart-ondeactivated-member)|Occurs when the chart is deactivated.|
||[plotArea](/.chart#excel-javascript/api/excel/-chart-plotarea-member)|Represents the plot area for the chart.|
||[plotBy](/.chart#excel-javascript/api/excel/-chart-plotby-member)|Specifies the way columns or rows are used as data series on the chart.|
||[plotVisibleOnly](/.chart#excel-javascript/api/excel/-chart-plotvisibleonly-member)|True if only visible cells are plotted.|
||[seriesNameLevel](/.chart#excel-javascript/api/excel/-chart-seriesnamelevel-member)|Specifies a chart series name level enumeration constant, referring to the level of the source series names.|
||[showDataLabelsOverMaximum](/.chart#excel-javascript/api/excel/-chart-showdatalabelsovermaximum-member)|Specifies whether to show the data labels when the value is greater than the maximum value on the value axis.|
||[style](/.chart#excel-javascript/api/excel/-chart-style-member)|Specifies the chart style for the chart.|
|[ChartActivatedEventArgs](/.chartactivatedeventargs)|[chartId](/.chartactivatedeventargs#excel-javascript/api/excel/-chartactivatedeventargs-chartid-member)|Gets the ID of the chart that is activated.|
||[type](/.chartactivatedeventargs#excel-javascript/api/excel/-chartactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.chartactivatedeventargs#excel-javascript/api/excel/-chartactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is activated.|
|[ChartAddedEventArgs](/.chartaddedeventargs)|[chartId](/.chartaddedeventargs#excel-javascript/api/excel/-chartaddedeventargs-chartid-member)|Gets the ID of the chart that is added to the worksheet.|
||[source](/.chartaddedeventargs#excel-javascript/api/excel/-chartaddedeventargs-source-member)|Gets the source of the event.|
||[type](/.chartaddedeventargs#excel-javascript/api/excel/-chartaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.chartaddedeventargs#excel-javascript/api/excel/-chartaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is added.|
|[ChartAxis](/.chartaxis)|[alignment](/.chartaxis#excel-javascript/api/excel/-chartaxis-alignment-member)|Specifies the alignment for the specified axis tick label.|
||[isBetweenCategories](/.chartaxis#excel-javascript/api/excel/-chartaxis-isbetweencategories-member)|Specifies if the value axis crosses the category axis between categories.|
||[multiLevel](/.chartaxis#excel-javascript/api/excel/-chartaxis-multilevel-member)|Specifies if an axis is multilevel.|
||[numberFormat](/.chartaxis#excel-javascript/api/excel/-chartaxis-numberformat-member)|Specifies the format code for the axis tick label.|
||[offset](/.chartaxis#excel-javascript/api/excel/-chartaxis-offset-member)|Specifies the distance between the levels of labels, and the distance between the first level and the axis line.|
||[position](/.chartaxis#excel-javascript/api/excel/-chartaxis-position-member)|Specifies the specified axis position where the other axis crosses.|
||[positionAt](/.chartaxis#excel-javascript/api/excel/-chartaxis-positionat-member)|Specifies the axis position where the other axis crosses.|
||[setPositionAt(value: number)](/.chartaxis#excel-javascript/api/excel/-chartaxis-setpositionat-member(1))|Sets the specified axis position where the other axis crosses.|
||[textOrientation](/.chartaxis#excel-javascript/api/excel/-chartaxis-textorientation-member)|Specifies the angle to which the text is oriented for the chart axis tick label.|
|[ChartAxisFormat](/.chartaxisformat)|[fill](/.chartaxisformat#excel-javascript/api/excel/-chartaxisformat-fill-member)|Specifies chart fill formatting.|
|[ChartAxisTitle](/.chartaxistitle)|[setFormula(formula: string)](/.chartaxistitle#excel-javascript/api/excel/-chartaxistitle-setformula-member(1))|A string value that represents the formula of chart axis title using A1-style notation.|
|[ChartAxisTitleFormat](/.chartaxistitleformat)|[border](/.chartaxistitleformat#excel-javascript/api/excel/-chartaxistitleformat-border-member)|Specifies the chart axis title's border format, which includes color, linestyle, and weight.|
||[fill](/.chartaxistitleformat#excel-javascript/api/excel/-chartaxistitleformat-fill-member)|Specifies the chart axis title's fill formatting.|
|[ChartBorder](/.chartborder)|[clear()](/.chartborder#excel-javascript/api/excel/-chartborder-clear-member(1))|Clear the border format of a chart element.|
|[ChartCollection](/.chartcollection)|[onActivated](/.chartcollection#excel-javascript/api/excel/-chartcollection-onactivated-member)|Occurs when a chart is activated.|
||[onAdded](/.chartcollection#excel-javascript/api/excel/-chartcollection-onadded-member)|Occurs when a new chart is added to the worksheet.|
||[onDeactivated](/.chartcollection#excel-javascript/api/excel/-chartcollection-ondeactivated-member)|Occurs when a chart is deactivated.|
||[onDeleted](/.chartcollection#excel-javascript/api/excel/-chartcollection-ondeleted-member)|Occurs when a chart is deleted.|
|[ChartDataLabel](/.chartdatalabel)|[autoText](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-autotext-member)|Specifies if the data label automatically generates appropriate text based on context.|
||[format](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-format-member)|Represents the format of chart data label.|
||[formula](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-formula-member)|String value that represents the formula of chart data label using A1-style notation.|
||[height](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-height-member)|Returns the height, in points, of the chart data label.|
||[horizontalAlignment](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-horizontalalignment-member)|Represents the horizontal alignment for chart data label.|
||[left](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-left-member)|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area.|
||[numberFormat](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-numberformat-member)|Specifies the format code for data label.|
||[text](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-text-member)|String representing the text of the data label on a chart.|
||[textOrientation](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-textorientation-member)|Represents the angle to which the text is oriented for the chart data label.|
||[top](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-top-member)|Represents the distance, in points, from the top edge of chart data label to the top of chart area.|
||[verticalAlignment](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-verticalalignment-member)|Represents the vertical alignment of chart data label.|
||[width](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-width-member)|Returns the width, in points, of the chart data label.|
|[ChartDataLabelFormat](/.chartdatalabelformat)|[border](/.chartdatalabelformat#excel-javascript/api/excel/-chartdatalabelformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartDataLabels](/.chartdatalabels)|[autoText](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-autotext-member)|Specifies if data labels automatically generate appropriate text based on context.|
||[horizontalAlignment](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-horizontalalignment-member)|Specifies the horizontal alignment for chart data label.|
||[numberFormat](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-numberformat-member)|Specifies the format code for data labels.|
||[textOrientation](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-textorientation-member)|Represents the angle to which the text is oriented for data labels.|
||[verticalAlignment](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-verticalalignment-member)|Represents the vertical alignment of chart data label.|
|[ChartDeactivatedEventArgs](/.chartdeactivatedeventargs)|[chartId](/.chartdeactivatedeventargs#excel-javascript/api/excel/-chartdeactivatedeventargs-chartid-member)|Gets the ID of the chart that is deactivated.|
||[type](/.chartdeactivatedeventargs#excel-javascript/api/excel/-chartdeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.chartdeactivatedeventargs#excel-javascript/api/excel/-chartdeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deactivated.|
|[ChartDeletedEventArgs](/.chartdeletedeventargs)|[chartId](/.chartdeletedeventargs#excel-javascript/api/excel/-chartdeletedeventargs-chartid-member)|Gets the ID of the chart that is deleted from the worksheet.|
||[source](/.chartdeletedeventargs#excel-javascript/api/excel/-chartdeletedeventargs-source-member)|Gets the source of the event.|
||[type](/.chartdeletedeventargs#excel-javascript/api/excel/-chartdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.chartdeletedeventargs#excel-javascript/api/excel/-chartdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the chart is deleted.|
|[ChartLegendEntry](/.chartlegendentry)|[height](/.chartlegendentry#excel-javascript/api/excel/-chartlegendentry-height-member)|Specifies the height of the legend entry on the chart legend.|
||[index](/.chartlegendentry#excel-javascript/api/excel/-chartlegendentry-index-member)|Specifies the index of the legend entry in the chart legend.|
||[left](/.chartlegendentry#excel-javascript/api/excel/-chartlegendentry-left-member)|Specifies the left value of a chart legend entry.|
||[top](/.chartlegendentry#excel-javascript/api/excel/-chartlegendentry-top-member)|Specifies the top of a chart legend entry.|
||[width](/.chartlegendentry#excel-javascript/api/excel/-chartlegendentry-width-member)|Represents the width of the legend entry on the chart Legend.|
|[ChartLegendFormat](/.chartlegendformat)|[border](/.chartlegendformat#excel-javascript/api/excel/-chartlegendformat-border-member)|Represents the border format, which includes color, linestyle, and weight.|
|[ChartPlotArea](/.chartplotarea)|[format](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-format-member)|Specifies the formatting of a chart plot area.|
||[height](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-height-member)|Specifies the height value of a plot area.|
||[insideHeight](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-insideheight-member)|Specifies the inside height value of a plot area.|
||[insideLeft](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-insideleft-member)|Specifies the inside left value of a plot area.|
||[insideTop](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-insidetop-member)|Specifies the inside top value of a plot area.|
||[insideWidth](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-insidewidth-member)|Specifies the inside width value of a plot area.|
||[left](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-left-member)|Specifies the left value of a plot area.|
||[position](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-position-member)|Specifies the position of a plot area.|
||[top](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-top-member)|Specifies the top value of a plot area.|
||[width](/.chartplotarea#excel-javascript/api/excel/-chartplotarea-width-member)|Specifies the width value of a plot area.|
|[ChartPlotAreaFormat](/.chartplotareaformat)|[border](/.chartplotareaformat#excel-javascript/api/excel/-chartplotareaformat-border-member)|Specifies the border attributes of a chart plot area.|
||[fill](/.chartplotareaformat#excel-javascript/api/excel/-chartplotareaformat-fill-member)|Specifies the fill format of an object, which includes background formatting information.|
|[ChartSeries](/.chartseries)|[axisGroup](/.chartseries#excel-javascript/api/excel/-chartseries-axisgroup-member)|Specifies the group for the specified series.|
||[dataLabels](/.chartseries#excel-javascript/api/excel/-chartseries-datalabels-member)|Represents a collection of all data labels in the series.|
||[explosion](/.chartseries#excel-javascript/api/excel/-chartseries-explosion-member)|Specifies the explosion value for a pie-chart or doughnut-chart slice.|
||[firstSliceAngle](/.chartseries#excel-javascript/api/excel/-chartseries-firstsliceangle-member)|Specifies the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical).|
||[invertIfNegative](/.chartseries#excel-javascript/api/excel/-chartseries-invertifnegative-member)|True if Excel inverts the pattern in the item when it corresponds to a negative number.|
||[overlap](/.chartseries#excel-javascript/api/excel/-chartseries-overlap-member)|Specifies how bars and columns are positioned.|
||[secondPlotSize](/.chartseries#excel-javascript/api/excel/-chartseries-secondplotsize-member)|Specifies the size of the secondary section of either a pie-of-pie chart or a bar-of-pie chart, as a percentage of the size of the primary pie.|
||[splitType](/.chartseries#excel-javascript/api/excel/-chartseries-splittype-member)|Specifies the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split.|
||[varyByCategories](/.chartseries#excel-javascript/api/excel/-chartseries-varybycategories-member)|True if Excel assigns a different color or pattern to each data marker.|
|[ChartTrendline](/.charttrendline)|[backwardPeriod](/.charttrendline#excel-javascript/api/excel/-charttrendline-backwardperiod-member)|Represents the number of periods that the trendline extends backward.|
||[forwardPeriod](/.charttrendline#excel-javascript/api/excel/-charttrendline-forwardperiod-member)|Represents the number of periods that the trendline extends forward.|
||[label](/.charttrendline#excel-javascript/api/excel/-charttrendline-label-member)|Represents the label of a chart trendline.|
||[showEquation](/.charttrendline#excel-javascript/api/excel/-charttrendline-showequation-member)|True if the equation for the trendline is displayed on the chart.|
||[showRSquared](/.charttrendline#excel-javascript/api/excel/-charttrendline-showrsquared-member)|True if the r-squared value for the trendline is displayed on the chart.|
|[ChartTrendlineLabel](/.charttrendlinelabel)|[autoText](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-autotext-member)|Specifies if the trendline label automatically generates appropriate text based on context.|
||[format](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-format-member)|The format of the chart trendline label.|
||[formula](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-formula-member)|String value that represents the formula of the chart trendline label using A1-style notation.|
||[height](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-height-member)|Returns the height, in points, of the chart trendline label.|
||[horizontalAlignment](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-horizontalalignment-member)|Represents the horizontal alignment of the chart trendline label.|
||[left](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-left-member)|Represents the distance, in points, from the left edge of the chart trendline label to the left edge of the chart area.|
||[numberFormat](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-numberformat-member)|String value that represents the format code for the trendline label.|
||[text](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-text-member)|String representing the text of the trendline label on a chart.|
||[textOrientation](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-textorientation-member)|Represents the angle to which the text is oriented for the chart trendline label.|
||[top](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-top-member)|Represents the distance, in points, from the top edge of the chart trendline label to the top of the chart area.|
||[verticalAlignment](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-verticalalignment-member)|Represents the vertical alignment of the chart trendline label.|
||[width](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-width-member)|Returns the width, in points, of the chart trendline label.|
|[ChartTrendlineLabelFormat](/.charttrendlinelabelformat)|[border](/.charttrendlinelabelformat#excel-javascript/api/excel/-charttrendlinelabelformat-border-member)|Specifies the border format, which includes color, linestyle, and weight.|
||[fill](/.charttrendlinelabelformat#excel-javascript/api/excel/-charttrendlinelabelformat-fill-member)|Specifies the fill format of the current chart trendline label.|
||[font](/.charttrendlinelabelformat#excel-javascript/api/excel/-charttrendlinelabelformat-font-member)|Specifies the font attributes (such as font name, font size, and color) for a chart trendline label.|
|[CustomDataValidation](/.customdatavalidation)|[formula](/.customdatavalidation#excel-javascript/api/excel/-customdatavalidation-formula-member)|A custom data validation formula.|
|[DataPivotHierarchy](/.datapivothierarchy)|[field](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-field-member)|Returns the PivotFields associated with the DataPivotHierarchy.|
||[id](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-id-member)|ID of the DataPivotHierarchy.|
||[name](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-name-member)|Name of the DataPivotHierarchy.|
||[numberFormat](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-numberformat-member)|Number format of the DataPivotHierarchy.|
||[position](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-position-member)|Position of the DataPivotHierarchy.|
||[setToDefault()](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-settodefault-member(1))|Reset the DataPivotHierarchy back to its default values.|
||[showAs](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-showas-member)|Specifies if the data should be shown as a specific summary calculation.|
||[summarizeBy](/.datapivothierarchy#excel-javascript/api/excel/-datapivothierarchy-summarizeby-member)|Specifies if all items of the DataPivotHierarchy are shown.|
|[DataPivotHierarchyCollection](/.datapivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-getitem-member(1))|Gets a DataPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-getitemornullobject-member(1))|Gets a DataPivotHierarchy by name.|
||[items](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(DataPivotHierarchy: Excel.DataPivotHierarchy)](/.datapivothierarchycollection#excel-javascript/api/excel/-datapivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[DataValidation](/.datavalidation)|[clear()](/.datavalidation#excel-javascript/api/excel/-datavalidation-clear-member(1))|Clears the data validation from the current range.|
||[errorAlert](/.datavalidation#excel-javascript/api/excel/-datavalidation-erroralert-member)|Error alert when user enters invalid data.|
||[ignoreBlanks](/.datavalidation#excel-javascript/api/excel/-datavalidation-ignoreblanks-member)|Specifies if data validation will be performed on blank cells.|
||[prompt](/.datavalidation#excel-javascript/api/excel/-datavalidation-prompt-member)|Prompt when users select a cell.|
||[rule](/.datavalidation#excel-javascript/api/excel/-datavalidation-rule-member)|Data validation rule that contains different type of data validation criteria.|
||[type](/.datavalidation#excel-javascript/api/excel/-datavalidation-type-member)|Type of the data validation, see `Excel.DataValidationType` for details.|
||[valid](/.datavalidation#excel-javascript/api/excel/-datavalidation-valid-member)|Represents if all cell values are valid according to the data validation rules.|
|[DataValidationErrorAlert](/.datavalidationerroralert)|[message](/.datavalidationerroralert#excel-javascript/api/excel/-datavalidationerroralert-message-member)|Represents the error alert message.|
||[showAlert](/.datavalidationerroralert#excel-javascript/api/excel/-datavalidationerroralert-showalert-member)|Specifies whether to show an error alert dialog when a user enters invalid data.|
||[style](/.datavalidationerroralert#excel-javascript/api/excel/-datavalidationerroralert-style-member)|The data validation alert type, please see `Excel.DataValidationAlertStyle` for details.|
||[title](/.datavalidationerroralert#excel-javascript/api/excel/-datavalidationerroralert-title-member)|Represents the error alert dialog title.|
|[DataValidationPrompt](/.datavalidationprompt)|[message](/.datavalidationprompt#excel-javascript/api/excel/-datavalidationprompt-message-member)|Specifies the message of the prompt.|
||[showPrompt](/.datavalidationprompt#excel-javascript/api/excel/-datavalidationprompt-showprompt-member)|Specifies if a prompt is shown when a user selects a cell with data validation.|
||[title](/.datavalidationprompt#excel-javascript/api/excel/-datavalidationprompt-title-member)|Specifies the title for the prompt.|
|[DataValidationRule](/.datavalidationrule)|[custom](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-custom-member)|Custom data validation criteria.|
||[date](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-date-member)|Date data validation criteria.|
||[decimal](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-decimal-member)|Decimal data validation criteria.|
||[list](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-list-member)|List data validation criteria.|
||[textLength](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-textlength-member)|Text length data validation criteria.|
||[time](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-time-member)|Time data validation criteria.|
||[wholeNumber](/.datavalidationrule#excel-javascript/api/excel/-datavalidationrule-wholenumber-member)|Whole number data validation criteria.|
|[DateTimeDataValidation](/.datetimedatavalidation)|[formula1](/.datetimedatavalidation#excel-javascript/api/excel/-datetimedatavalidation-formula1-member)|Specifies the right-hand operand when the operator property is set to a binary operator such as GreaterThan (the left-hand operand is the value the user tries to enter in the cell).|
||[formula2](/.datetimedatavalidation#excel-javascript/api/excel/-datetimedatavalidation-formula2-member)|With the ternary operators Between and NotBetween, specifies the upper bound operand.|
||[operator](/.datetimedatavalidation#excel-javascript/api/excel/-datetimedatavalidation-operator-member)|The operator to use for validating the data.|
|[FilterPivotHierarchy](/.filterpivothierarchy)|[enableMultipleFilterItems](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-enablemultiplefilteritems-member)|Determines whether to allow multiple filter items.|
||[fields](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-fields-member)|Returns the PivotFields associated with the FilterPivotHierarchy.|
||[id](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-id-member)|ID of the FilterPivotHierarchy.|
||[name](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-name-member)|Name of the FilterPivotHierarchy.|
||[position](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-position-member)|Position of the FilterPivotHierarchy.|
||[setToDefault()](/.filterpivothierarchy#excel-javascript/api/excel/-filterpivothierarchy-settodefault-member(1))|Reset the FilterPivotHierarchy back to its default values.|
|[FilterPivotHierarchyCollection](/.filterpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-getitem-member(1))|Gets a FilterPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-getitemornullobject-member(1))|Gets a FilterPivotHierarchy by name.|
||[items](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(filterPivotHierarchy: Excel.FilterPivotHierarchy)](/.filterpivothierarchycollection#excel-javascript/api/excel/-filterpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[ListDataValidation](/.listdatavalidation)|[inCellDropDown](/.listdatavalidation#excel-javascript/api/excel/-listdatavalidation-incelldropdown-member)|Specifies whether to display the list in a cell drop-down.|
||[source](/.listdatavalidation#excel-javascript/api/excel/-listdatavalidation-source-member)|Source of the list for data validation|
|[PivotField](/.pivotfield)|[id](/.pivotfield#excel-javascript/api/excel/-pivotfield-id-member)|ID of the PivotField.|
||[items](/.pivotfield#excel-javascript/api/excel/-pivotfield-items-member)|Returns the PivotItems associated with the PivotField.|
||[name](/.pivotfield#excel-javascript/api/excel/-pivotfield-name-member)|Name of the PivotField.|
||[showAllItems](/.pivotfield#excel-javascript/api/excel/-pivotfield-showallitems-member)|Determines whether to show all items of the PivotField.|
||[sortByLabels(sortBy: SortBy)](/.pivotfield#excel-javascript/api/excel/-pivotfield-sortbylabels-member(1))|Sorts the PivotField.|
||[subtotals](/.pivotfield#excel-javascript/api/excel/-pivotfield-subtotals-member)|Subtotals of the PivotField.|
|[PivotFieldCollection](/.pivotfieldcollection)|[getCount()](/.pivotfieldcollection#excel-javascript/api/excel/-pivotfieldcollection-getcount-member(1))|Gets the number of pivot fields in the collection.|
||[getItem(name: string)](/.pivotfieldcollection#excel-javascript/api/excel/-pivotfieldcollection-getitem-member(1))|Gets a PivotField by its name or ID.|
||[getItemOrNullObject(name: string)](/.pivotfieldcollection#excel-javascript/api/excel/-pivotfieldcollection-getitemornullobject-member(1))|Gets a PivotField by name.|
||[items](/.pivotfieldcollection#excel-javascript/api/excel/-pivotfieldcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotHierarchy](/.pivothierarchy)|[fields](/.pivothierarchy#excel-javascript/api/excel/-pivothierarchy-fields-member)|Returns the PivotFields associated with the PivotHierarchy.|
||[id](/.pivothierarchy#excel-javascript/api/excel/-pivothierarchy-id-member)|ID of the PivotHierarchy.|
||[name](/.pivothierarchy#excel-javascript/api/excel/-pivothierarchy-name-member)|Name of the PivotHierarchy.|
|[PivotHierarchyCollection](/.pivothierarchycollection)|[getCount()](/.pivothierarchycollection#excel-javascript/api/excel/-pivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/.pivothierarchycollection#excel-javascript/api/excel/-pivothierarchycollection-getitem-member(1))|Gets a PivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/.pivothierarchycollection#excel-javascript/api/excel/-pivothierarchycollection-getitemornullobject-member(1))|Gets a PivotHierarchy by name.|
||[items](/.pivothierarchycollection#excel-javascript/api/excel/-pivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotItem](/.pivotitem)|[id](/.pivotitem#excel-javascript/api/excel/-pivotitem-id-member)|ID of the PivotItem.|
||[isExpanded](/.pivotitem#excel-javascript/api/excel/-pivotitem-isexpanded-member)|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|
||[name](/.pivotitem#excel-javascript/api/excel/-pivotitem-name-member)|Name of the PivotItem.|
||[visible](/.pivotitem#excel-javascript/api/excel/-pivotitem-visible-member)|Specifies if the PivotItem is visible.|
|[PivotItemCollection](/.pivotitemcollection)|[getCount()](/.pivotitemcollection#excel-javascript/api/excel/-pivotitemcollection-getcount-member(1))|Gets the number of PivotItems in the collection.|
||[getItem(name: string)](/.pivotitemcollection#excel-javascript/api/excel/-pivotitemcollection-getitem-member(1))|Gets a PivotItem by its name or ID.|
||[getItemOrNullObject(name: string)](/.pivotitemcollection#excel-javascript/api/excel/-pivotitemcollection-getitemornullobject-member(1))|Gets a PivotItem by name.|
||[items](/.pivotitemcollection#excel-javascript/api/excel/-pivotitemcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/.pivotlayout)|[getColumnLabelRange()](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getcolumnlabelrange-member(1))|Returns the range where the PivotTable's column labels reside.|
||[getDataBodyRange()](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getdatabodyrange-member(1))|Returns the range where the PivotTable's data values reside.|
||[getFilterAxisRange()](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getfilteraxisrange-member(1))|Returns the range of the PivotTable's filter area.|
||[getRange()](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getrange-member(1))|Returns the range the PivotTable exists on, excluding the filter area.|
||[getRowLabelRange()](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getrowlabelrange-member(1))|Returns the range where the PivotTable's row labels reside.|
||[layoutType](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-layouttype-member)|This property indicates the PivotLayoutType of all fields on the PivotTable.|
||[showColumnGrandTotals](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-showcolumngrandtotals-member)|Specifies if the PivotTable report shows grand totals for columns.|
||[showRowGrandTotals](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-showrowgrandtotals-member)|Specifies if the PivotTable report shows grand totals for rows.|
||[subtotalLocation](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-subtotallocation-member)|This property indicates the `SubtotalLocationType` of all fields on the PivotTable.|
|[PivotTable](/.pivottable)|[columnHierarchies](/.pivottable#excel-javascript/api/excel/-pivottable-columnhierarchies-member)|The Column Pivot Hierarchies of the PivotTable.|
||[dataHierarchies](/.pivottable#excel-javascript/api/excel/-pivottable-datahierarchies-member)|The Data Pivot Hierarchies of the PivotTable.|
||[delete()](/.pivottable#excel-javascript/api/excel/-pivottable-delete-member(1))|Deletes the PivotTable.|
||[filterHierarchies](/.pivottable#excel-javascript/api/excel/-pivottable-filterhierarchies-member)|The Filter Pivot Hierarchies of the PivotTable.|
||[hierarchies](/.pivottable#excel-javascript/api/excel/-pivottable-hierarchies-member)|The Pivot Hierarchies of the PivotTable.|
||[layout](/.pivottable#excel-javascript/api/excel/-pivottable-layout-member)|The PivotLayout describing the layout and visual structure of the PivotTable.|
||[rowHierarchies](/.pivottable#excel-javascript/api/excel/-pivottable-rowhierarchies-member)|The Row Pivot Hierarchies of the PivotTable.|
|[PivotTableCollection](/.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/.pivottablecollection#excel-javascript/api/excel/-pivottablecollection-add-member(1))|Add a PivotTable based on the specified source data and insert it at the top-left cell of the destination range.|
|[Range](/.range)|[dataValidation](/.range#excel-javascript/api/excel/-range-datavalidation-member)|Returns a data validation object.|
|[RowColumnPivotHierarchy](/.rowcolumnpivothierarchy)|[fields](/.rowcolumnpivothierarchy#excel-javascript/api/excel/-rowcolumnpivothierarchy-fields-member)|Returns the PivotFields associated with the RowColumnPivotHierarchy.|
||[id](/.rowcolumnpivothierarchy#excel-javascript/api/excel/-rowcolumnpivothierarchy-id-member)|ID of the RowColumnPivotHierarchy.|
||[name](/.rowcolumnpivothierarchy#excel-javascript/api/excel/-rowcolumnpivothierarchy-name-member)|Name of the RowColumnPivotHierarchy.|
||[position](/.rowcolumnpivothierarchy#excel-javascript/api/excel/-rowcolumnpivothierarchy-position-member)|Position of the RowColumnPivotHierarchy.|
||[setToDefault()](/.rowcolumnpivothierarchy#excel-javascript/api/excel/-rowcolumnpivothierarchy-settodefault-member(1))|Reset the RowColumnPivotHierarchy back to its default values.|
|[RowColumnPivotHierarchyCollection](/.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel.PivotHierarchy)](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-add-member(1))|Adds the PivotHierarchy to the current axis.|
||[getCount()](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-getcount-member(1))|Gets the number of pivot hierarchies in the collection.|
||[getItem(name: string)](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-getitem-member(1))|Gets a RowColumnPivotHierarchy by its name or ID.|
||[getItemOrNullObject(name: string)](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|Gets a RowColumnPivotHierarchy by name.|
||[items](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(rowColumnPivotHierarchy: Excel.RowColumnPivotHierarchy)](/.rowcolumnpivothierarchycollection#excel-javascript/api/excel/-rowcolumnpivothierarchycollection-remove-member(1))|Removes the PivotHierarchy from the current axis.|
|[Runtime](/.runtime)|[enableEvents](/.runtime#excel-javascript/api/excel/-runtime-enableevents-member)|Toggle JavaScript events in the current task pane or content add-in.|
|[ShowAsRule](/.showasrule)|[baseField](/.showasrule#excel-javascript/api/excel/-showasrule-basefield-member)|The PivotField to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[baseItem](/.showasrule#excel-javascript/api/excel/-showasrule-baseitem-member)|The item to base the `ShowAs` calculation on, if applicable according to the `ShowAsCalculation` type, else `null`.|
||[calculation](/.showasrule#excel-javascript/api/excel/-showasrule-calculation-member)|The `ShowAs` calculation to use for the PivotField.|
|[Style](/.style)|[autoIndent](/.style#excel-javascript/api/excel/-style-autoindent-member)|Specifies if text is automatically indented when the text alignment in a cell is set to equal distribution.|
||[textOrientation](/.style#excel-javascript/api/excel/-style-textorientation-member)|The text orientation for the style.|
|[Subtotals](/.subtotals)|[automatic](/.subtotals#excel-javascript/api/excel/-subtotals-automatic-member)|If `Automatic` is set to `true`, then all other values will be ignored when setting the `Subtotals`.|
||[average](/.subtotals#excel-javascript/api/excel/-subtotals-average-member)||
||[count](/.subtotals#excel-javascript/api/excel/-subtotals-count-member)||
||[countNumbers](/.subtotals#excel-javascript/api/excel/-subtotals-countnumbers-member)||
||[max](/.subtotals#excel-javascript/api/excel/-subtotals-max-member)||
||[min](/.subtotals#excel-javascript/api/excel/-subtotals-min-member)||
||[product](/.subtotals#excel-javascript/api/excel/-subtotals-product-member)||
||[standardDeviation](/.subtotals#excel-javascript/api/excel/-subtotals-standarddeviation-member)||
||[standardDeviationP](/.subtotals#excel-javascript/api/excel/-subtotals-standarddeviationp-member)||
||[sum](/.subtotals#excel-javascript/api/excel/-subtotals-sum-member)||
||[variance](/.subtotals#excel-javascript/api/excel/-subtotals-variance-member)||
||[varianceP](/.subtotals#excel-javascript/api/excel/-subtotals-variancep-member)||
|[Table](/.table)|[legacyId](/.table#excel-javascript/api/excel/-table-legacyid-member)|Returns a numeric ID.|
|[TableChangedEventArgs](/.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/.tablechangedeventargs#excel-javascript/api/excel/-tablechangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a table on a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/.tablechangedeventargs#excel-javascript/api/excel/-tablechangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a table on a specific worksheet.|
|[Workbook](/.workbook)|[readOnly](/.workbook#excel-javascript/api/excel/-workbook-readonly-member)|Returns `true` if the workbook is open in read-only mode.|
|[WorkbookCreated](/.workbookcreated)|||
|[Worksheet](/.worksheet)|[onCalculated](/.worksheet#excel-javascript/api/excel/-worksheet-oncalculated-member)|Occurs when the worksheet is calculated.|
||[showGridlines](/.worksheet#excel-javascript/api/excel/-worksheet-showgridlines-member)|Specifies if gridlines are visible to the user.|
||[showHeadings](/.worksheet#excel-javascript/api/excel/-worksheet-showheadings-member)|Specifies if headings are visible to the user.|
|[WorksheetCalculatedEventArgs](/.worksheetcalculatedeventargs)|[type](/.worksheetcalculatedeventargs#excel-javascript/api/excel/-worksheetcalculatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetcalculatedeventargs#excel-javascript/api/excel/-worksheetcalculatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the calculation occurred.|
|[WorksheetChangedEventArgs](/.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/.worksheetchangedeventargs#excel-javascript/api/excel/-worksheetchangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/.worksheetchangedeventargs#excel-javascript/api/excel/-worksheetchangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a specific worksheet.|
|[WorksheetCollection](/.worksheetcollection)|[onCalculated](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-oncalculated-member)|Occurs when any worksheet in the workbook is calculated.|
