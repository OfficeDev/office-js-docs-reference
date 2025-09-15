| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[calculationEngineVersion](/.application#excel-javascript/api/excel/-application-calculationengineversion-member)|Returns the Excel calculation engine version used for the last full recalculation.|
||[calculationState](/.application#excel-javascript/api/excel/-application-calculationstate-member)|Returns the calculation state of the application.|
||[iterativeCalculation](/.application#excel-javascript/api/excel/-application-iterativecalculation-member)|Returns the iterative calculation settings.|
||[suspendScreenUpdatingUntilNextSync()](/.application#excel-javascript/api/excel/-application-suspendscreenupdatinguntilnextsync-member(1))|Suspends screen updating until the next `context.sync()` is called.|
|[AutoFilter](/.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/.autofilter#excel-javascript/api/excel/-autofilter-apply-member(1))|Applies the AutoFilter to a range.|
||[clearCriteria()](/.autofilter#excel-javascript/api/excel/-autofilter-clearcriteria-member(1))|Clears the filter criteria and sort state of the AutoFilter.|
||[criteria](/.autofilter#excel-javascript/api/excel/-autofilter-criteria-member)|An array that holds all the filter criteria in the autofiltered range.|
||[enabled](/.autofilter#excel-javascript/api/excel/-autofilter-enabled-member)|Specifies if the AutoFilter is enabled.|
||[getRange()](/.autofilter#excel-javascript/api/excel/-autofilter-getrange-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[getRangeOrNullObject()](/.autofilter#excel-javascript/api/excel/-autofilter-getrangeornullobject-member(1))|Returns the `Range` object that represents the range to which the AutoFilter applies.|
||[isDataFiltered](/.autofilter#excel-javascript/api/excel/-autofilter-isdatafiltered-member)|Specifies if the AutoFilter has filter criteria.|
||[reapply()](/.autofilter#excel-javascript/api/excel/-autofilter-reapply-member(1))|Applies the specified AutoFilter object currently on the range.|
||[remove()](/.autofilter#excel-javascript/api/excel/-autofilter-remove-member(1))|Removes the AutoFilter for the range.|
|[CellBorder](/.cellborder)|[color](/.cellborder#excel-javascript/api/excel/-cellborder-color-member)|Represents the `color` property of a single border.|
||[style](/.cellborder#excel-javascript/api/excel/-cellborder-style-member)|Represents the `style` property of a single border.|
||[tintAndShade](/.cellborder#excel-javascript/api/excel/-cellborder-tintandshade-member)|Represents the `tintAndShade` property of a single border.|
||[weight](/.cellborder#excel-javascript/api/excel/-cellborder-weight-member)|Represents the `weight` property of a single border.|
|[CellBorderCollection](/.cellbordercollection)|[bottom](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-bottom-member)|Represents the `format.borders.bottom` property.|
||[diagonalDown](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-diagonaldown-member)|Represents the `format.borders.diagonalDown` property.|
||[diagonalUp](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-diagonalup-member)|Represents the `format.borders.diagonalUp` property.|
||[horizontal](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-horizontal-member)|Represents the `format.borders.horizontal` property.|
||[left](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-left-member)|Represents the `format.borders.left` property.|
||[right](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-right-member)|Represents the `format.borders.right` property.|
||[top](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-top-member)|Represents the `format.borders.top` property.|
||[vertical](/.cellbordercollection#excel-javascript/api/excel/-cellbordercollection-vertical-member)|Represents the `format.borders.vertical` property.|
|[CellProperties](/.cellproperties)|[address](/.cellproperties#excel-javascript/api/excel/-cellproperties-address-member)|Represents the `address` property.|
||[addressLocal](/.cellproperties#excel-javascript/api/excel/-cellproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[hidden](/.cellproperties#excel-javascript/api/excel/-cellproperties-hidden-member)|Represents the `hidden` property.|
|[CellPropertiesFill](/.cellpropertiesfill)|[color](/.cellpropertiesfill#excel-javascript/api/excel/-cellpropertiesfill-color-member)|Represents the `format.fill.color` property.|
||[pattern](/.cellpropertiesfill#excel-javascript/api/excel/-cellpropertiesfill-pattern-member)|Represents the `format.fill.pattern` property.|
||[patternColor](/.cellpropertiesfill#excel-javascript/api/excel/-cellpropertiesfill-patterncolor-member)|Represents the `format.fill.patternColor` property.|
||[patternTintAndShade](/.cellpropertiesfill#excel-javascript/api/excel/-cellpropertiesfill-patterntintandshade-member)|Represents the `format.fill.patternTintAndShade` property.|
||[tintAndShade](/.cellpropertiesfill#excel-javascript/api/excel/-cellpropertiesfill-tintandshade-member)|Represents the `format.fill.tintAndShade` property.|
|[CellPropertiesFont](/.cellpropertiesfont)|[bold](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-bold-member)|Represents the `format.font.bold` property.|
||[color](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-color-member)|Represents the `format.font.color` property.|
||[italic](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-italic-member)|Represents the `format.font.italic` property.|
||[name](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-name-member)|Represents the `format.font.name` property.|
||[size](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-size-member)|Represents the `format.font.size` property.|
||[strikethrough](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-strikethrough-member)|Represents the `format.font.strikethrough` property.|
||[subscript](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-subscript-member)|Represents the `format.font.subscript` property.|
||[superscript](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-superscript-member)|Represents the `format.font.superscript` property.|
||[tintAndShade](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-tintandshade-member)|Represents the `format.font.tintAndShade` property.|
||[underline](/.cellpropertiesfont#excel-javascript/api/excel/-cellpropertiesfont-underline-member)|Represents the `format.font.underline` property.|
|[CellPropertiesFormat](/.cellpropertiesformat)|[autoIndent](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-autoindent-member)|Represents the `autoIndent` property.|
||[borders](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-borders-member)|Represents the `borders` property.|
||[fill](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-fill-member)|Represents the `fill` property.|
||[font](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-font-member)|Represents the `font` property.|
||[horizontalAlignment](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-horizontalalignment-member)|Represents the `horizontalAlignment` property.|
||[indentLevel](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-indentlevel-member)|Represents the `indentLevel` property.|
||[protection](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-protection-member)|Represents the `protection` property.|
||[readingOrder](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-readingorder-member)|Represents the `readingOrder` property.|
||[shrinkToFit](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-shrinktofit-member)|Represents the `shrinkToFit` property.|
||[textOrientation](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-textorientation-member)|Represents the `textOrientation` property.|
||[useStandardHeight](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-usestandardheight-member)|Represents the `useStandardHeight` property.|
||[useStandardWidth](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-usestandardwidth-member)|Represents the `useStandardWidth` property.|
||[verticalAlignment](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-verticalalignment-member)|Represents the `verticalAlignment` property.|
||[wrapText](/.cellpropertiesformat#excel-javascript/api/excel/-cellpropertiesformat-wraptext-member)|Represents the `wrapText` property.|
|[CellPropertiesProtection](/.cellpropertiesprotection)|[formulaHidden](/.cellpropertiesprotection#excel-javascript/api/excel/-cellpropertiesprotection-formulahidden-member)|Represents the `format.protection.formulaHidden` property.|
||[locked](/.cellpropertiesprotection#excel-javascript/api/excel/-cellpropertiesprotection-locked-member)|Represents the `format.protection.locked` property.|
|[ChangedEventDetail](/.changedeventdetail)|[valueAfter](/.changedeventdetail#excel-javascript/api/excel/-changedeventdetail-valueafter-member)|Represents the value after the change.|
||[valueBefore](/.changedeventdetail#excel-javascript/api/excel/-changedeventdetail-valuebefore-member)|Represents the value before the change.|
||[valueTypeAfter](/.changedeventdetail#excel-javascript/api/excel/-changedeventdetail-valuetypeafter-member)|Represents the type of value after the change.|
||[valueTypeBefore](/.changedeventdetail#excel-javascript/api/excel/-changedeventdetail-valuetypebefore-member)|Represents the type of value before the change.|
|[Chart](/.chart)|[activate()](/.chart#excel-javascript/api/excel/-chart-activate-member(1))|Activates the chart in the Excel UI.|
||[pivotOptions](/.chart#excel-javascript/api/excel/-chart-pivotoptions-member)|Encapsulates the options for a pivot chart.|
|[ChartAreaFormat](/.chartareaformat)|[colorScheme](/.chartareaformat#excel-javascript/api/excel/-chartareaformat-colorscheme-member)|Specifies the color scheme of the chart.|
||[roundedCorners](/.chartareaformat#excel-javascript/api/excel/-chartareaformat-roundedcorners-member)|Specifies if the chart area of the chart has rounded corners.|
|[ChartAxis](/.chartaxis)|[linkNumberFormat](/.chartaxis#excel-javascript/api/excel/-chartaxis-linknumberformat-member)|Specifies if the number format is linked to the cells.|
|[ChartBinOptions](/.chartbinoptions)|[allowOverflow](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-allowoverflow-member)|Specifies if bin overflow is enabled in a histogram chart or pareto chart.|
||[allowUnderflow](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-allowunderflow-member)|Specifies if bin underflow is enabled in a histogram chart or pareto chart.|
||[count](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-count-member)|Specifies the bin count of a histogram chart or pareto chart.|
||[overflowValue](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-overflowvalue-member)|Specifies the bin overflow value of a histogram chart or pareto chart.|
||[type](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-type-member)|Specifies the bin's type for a histogram chart or pareto chart.|
||[underflowValue](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-underflowvalue-member)|Specifies the bin underflow value of a histogram chart or pareto chart.|
||[width](/.chartbinoptions#excel-javascript/api/excel/-chartbinoptions-width-member)|Specifies the bin width value of a histogram chart or pareto chart.|
|[ChartBoxwhiskerOptions](/.chartboxwhiskeroptions)|[quartileCalculation](/.chartboxwhiskeroptions#excel-javascript/api/excel/-chartboxwhiskeroptions-quartilecalculation-member)|Specifies if the quartile calculation type of a box and whisker chart.|
||[showInnerPoints](/.chartboxwhiskeroptions#excel-javascript/api/excel/-chartboxwhiskeroptions-showinnerpoints-member)|Specifies if inner points are shown in a box and whisker chart.|
||[showMeanLine](/.chartboxwhiskeroptions#excel-javascript/api/excel/-chartboxwhiskeroptions-showmeanline-member)|Specifies if the mean line is shown in a box and whisker chart.|
||[showMeanMarker](/.chartboxwhiskeroptions#excel-javascript/api/excel/-chartboxwhiskeroptions-showmeanmarker-member)|Specifies if the mean marker is shown in a box and whisker chart.|
||[showOutlierPoints](/.chartboxwhiskeroptions#excel-javascript/api/excel/-chartboxwhiskeroptions-showoutlierpoints-member)|Specifies if outlier points are shown in a box and whisker chart.|
|[ChartDataLabel](/.chartdatalabel)|[linkNumberFormat](/.chartdatalabel#excel-javascript/api/excel/-chartdatalabel-linknumberformat-member)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ChartDataLabels](/.chartdatalabels)|[linkNumberFormat](/.chartdatalabels#excel-javascript/api/excel/-chartdatalabels-linknumberformat-member)|Specifies if the number format is linked to the cells.|
|[ChartErrorBars](/.charterrorbars)|[endStyleCap](/.charterrorbars#excel-javascript/api/excel/-charterrorbars-endstylecap-member)|Specifies if error bars have an end style cap.|
||[format](/.charterrorbars#excel-javascript/api/excel/-charterrorbars-format-member)|Specifies the formatting type of the error bars.|
||[include](/.charterrorbars#excel-javascript/api/excel/-charterrorbars-include-member)|Specifies which parts of the error bars to include.|
||[type](/.charterrorbars#excel-javascript/api/excel/-charterrorbars-type-member)|The type of range marked by the error bars.|
||[visible](/.charterrorbars#excel-javascript/api/excel/-charterrorbars-visible-member)|Specifies whether the error bars are displayed.|
|[ChartErrorBarsFormat](/.charterrorbarsformat)|[line](/.charterrorbarsformat#excel-javascript/api/excel/-charterrorbarsformat-line-member)|Represents the chart line formatting.|
|[ChartMapOptions](/.chartmapoptions)|[labelStrategy](/.chartmapoptions#excel-javascript/api/excel/-chartmapoptions-labelstrategy-member)|Specifies the series map labels strategy of a region map chart.|
||[level](/.chartmapoptions#excel-javascript/api/excel/-chartmapoptions-level-member)|Specifies the series mapping level of a region map chart.|
||[projectionType](/.chartmapoptions#excel-javascript/api/excel/-chartmapoptions-projectiontype-member)|Specifies the series projection type of a region map chart.|
|[ChartPivotOptions](/.chartpivotoptions)|[showAxisFieldButtons](/.chartpivotoptions#excel-javascript/api/excel/-chartpivotoptions-showaxisfieldbuttons-member)|Specifies whether to display the axis field buttons on a PivotChart.|
||[showLegendFieldButtons](/.chartpivotoptions#excel-javascript/api/excel/-chartpivotoptions-showlegendfieldbuttons-member)|Specifies whether to display the legend field buttons on a PivotChart.|
||[showReportFilterFieldButtons](/.chartpivotoptions#excel-javascript/api/excel/-chartpivotoptions-showreportfilterfieldbuttons-member)|Specifies whether to display the report filter field buttons on a PivotChart.|
||[showValueFieldButtons](/.chartpivotoptions#excel-javascript/api/excel/-chartpivotoptions-showvaluefieldbuttons-member)|Specifies whether to display the show value field buttons on a PivotChart.|
|[ChartSeries](/.chartseries)|[binOptions](/.chartseries#excel-javascript/api/excel/-chartseries-binoptions-member)|Encapsulates the bin options for histogram charts and pareto charts.|
||[boxwhiskerOptions](/.chartseries#excel-javascript/api/excel/-chartseries-boxwhiskeroptions-member)|Encapsulates the options for the box and whisker charts.|
||[bubbleScale](/.chartseries#excel-javascript/api/excel/-chartseries-bubblescale-member)|This can be an integer value from 0 (zero) to 300, representing the percentage of the default size.|
||[gradientMaximumColor](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmaximumcolor-member)|Specifies the color for maximum value of a region map chart series.|
||[gradientMaximumType](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmaximumtype-member)|Specifies the type for maximum value of a region map chart series.|
||[gradientMaximumValue](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmaximumvalue-member)|Specifies the maximum value of a region map chart series.|
||[gradientMidpointColor](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmidpointcolor-member)|Specifies the color for the midpoint value of a region map chart series.|
||[gradientMidpointType](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmidpointtype-member)|Specifies the type for the midpoint value of a region map chart series.|
||[gradientMidpointValue](/.chartseries#excel-javascript/api/excel/-chartseries-gradientmidpointvalue-member)|Specifies the midpoint value of a region map chart series.|
||[gradientMinimumColor](/.chartseries#excel-javascript/api/excel/-chartseries-gradientminimumcolor-member)|Specifies the color for the minimum value of a region map chart series.|
||[gradientMinimumType](/.chartseries#excel-javascript/api/excel/-chartseries-gradientminimumtype-member)|Specifies the type for the minimum value of a region map chart series.|
||[gradientMinimumValue](/.chartseries#excel-javascript/api/excel/-chartseries-gradientminimumvalue-member)|Specifies the minimum value of a region map chart series.|
||[gradientStyle](/.chartseries#excel-javascript/api/excel/-chartseries-gradientstyle-member)|Specifies the series gradient style of a region map chart.|
||[invertColor](/.chartseries#excel-javascript/api/excel/-chartseries-invertcolor-member)|Specifies the fill color for negative data points in a series.|
||[mapOptions](/.chartseries#excel-javascript/api/excel/-chartseries-mapoptions-member)|Encapsulates the options for a region map chart.|
||[parentLabelStrategy](/.chartseries#excel-javascript/api/excel/-chartseries-parentlabelstrategy-member)|Specifies the series parent label strategy area for a treemap chart.|
||[showConnectorLines](/.chartseries#excel-javascript/api/excel/-chartseries-showconnectorlines-member)|Specifies whether connector lines are shown in waterfall charts.|
||[showLeaderLines](/.chartseries#excel-javascript/api/excel/-chartseries-showleaderlines-member)|Specifies whether leader lines are displayed for each data label in the series.|
||[splitValue](/.chartseries#excel-javascript/api/excel/-chartseries-splitvalue-member)|Specifies the threshold value that separates two sections of either a pie-of-pie chart or a bar-of-pie chart.|
||[xErrorBars](/.chartseries#excel-javascript/api/excel/-chartseries-xerrorbars-member)|Represents the error bar object of a chart series.|
||[yErrorBars](/.chartseries#excel-javascript/api/excel/-chartseries-yerrorbars-member)|Represents the error bar object of a chart series.|
|[ChartTrendlineLabel](/.charttrendlinelabel)|[linkNumberFormat](/.charttrendlinelabel#excel-javascript/api/excel/-charttrendlinelabel-linknumberformat-member)|Specifies if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|
|[ColumnProperties](/.columnproperties)|[address](/.columnproperties#excel-javascript/api/excel/-columnproperties-address-member)|Represents the `address` property.|
||[addressLocal](/.columnproperties#excel-javascript/api/excel/-columnproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[columnIndex](/.columnproperties#excel-javascript/api/excel/-columnproperties-columnindex-member)|Represents the `columnIndex` property.|
|[ConditionalFormat](/.conditionalformat)|[getRanges()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-getranges-member(1))|Returns the `RangeAreas`, comprising one or more rectangular ranges, to which the conditional format is applied.|
|[DataValidation](/.datavalidation)|[getInvalidCells()](/.datavalidation#excel-javascript/api/excel/-datavalidation-getinvalidcells-member(1))|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
||[getInvalidCellsOrNullObject()](/.datavalidation#excel-javascript/api/excel/-datavalidation-getinvalidcellsornullobject-member(1))|Returns a `RangeAreas` object, comprising one or more rectangular ranges, with invalid cell values.|
|[FilterCriteria](/.filtercriteria)|[subField](/.filtercriteria#excel-javascript/api/excel/-filtercriteria-subfield-member)|The property used by the filter to do a rich filter on rich values.|
|[GeometricShape](/.geometricshape)|[id](/.geometricshape#excel-javascript/api/excel/-geometricshape-id-member)|Returns the shape identifier.|
||[shape](/.geometricshape#excel-javascript/api/excel/-geometricshape-shape-member)|Returns the `Shape` object for the geometric shape.|
|[GroupShapeCollection](/.groupshapecollection)|[getCount()](/.groupshapecollection#excel-javascript/api/excel/-groupshapecollection-getcount-member(1))|Returns the number of shapes in the shape group.|
||[getItem(key: string)](/.groupshapecollection#excel-javascript/api/excel/-groupshapecollection-getitem-member(1))|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/.groupshapecollection#excel-javascript/api/excel/-groupshapecollection-getitemat-member(1))|Gets a shape based on its position in the collection.|
||[items](/.groupshapecollection#excel-javascript/api/excel/-groupshapecollection-items-member)|Gets the loaded child items in this collection.|
|[HeaderFooter](/.headerfooter)|[centerFooter](/.headerfooter#excel-javascript/api/excel/-headerfooter-centerfooter-member)|The center footer of the worksheet.|
||[centerHeader](/.headerfooter#excel-javascript/api/excel/-headerfooter-centerheader-member)|The center header of the worksheet.|
||[leftFooter](/.headerfooter#excel-javascript/api/excel/-headerfooter-leftfooter-member)|The left footer of the worksheet.|
||[leftHeader](/.headerfooter#excel-javascript/api/excel/-headerfooter-leftheader-member)|The left header of the worksheet.|
||[rightFooter](/.headerfooter#excel-javascript/api/excel/-headerfooter-rightfooter-member)|The right footer of the worksheet.|
||[rightHeader](/.headerfooter#excel-javascript/api/excel/-headerfooter-rightheader-member)|The right header of the worksheet.|
|[HeaderFooterGroup](/.headerfootergroup)|[defaultForAllPages](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-defaultforallpages-member)|The general header/footer, used for all pages unless even/odd or first page is specified.|
||[evenPages](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-evenpages-member)|The header/footer to use for even pages, odd header/footer needs to be specified for odd pages.|
||[firstPage](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-firstpage-member)|The first page header/footer, for all other pages general or even/odd is used.|
||[oddPages](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-oddpages-member)|The header/footer to use for odd pages, even header/footer needs to be specified for even pages.|
||[state](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-state-member)|The state by which headers/footers are set.|
||[useSheetMargins](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-usesheetmargins-member)|Gets or sets a flag indicating if headers/footers are aligned with the page margins set in the page layout options for the worksheet.|
||[useSheetScale](/.headerfootergroup#excel-javascript/api/excel/-headerfootergroup-usesheetscale-member)|Gets or sets a flag indicating if headers/footers should be scaled by the page percentage scale set in the page layout options for the worksheet.|
|[Image](/.image)|[format](/.image#excel-javascript/api/excel/-image-format-member)|Returns the format of the image.|
||[id](/.image#excel-javascript/api/excel/-image-id-member)|Specifies the shape identifier for the image object.|
||[shape](/.image#excel-javascript/api/excel/-image-shape-member)|Returns the `Shape` object associated with the image.|
|[IterativeCalculation](/.iterativecalculation)|[enabled](/.iterativecalculation#excel-javascript/api/excel/-iterativecalculation-enabled-member)|True if Excel will use iteration to resolve circular references.|
||[maxChange](/.iterativecalculation#excel-javascript/api/excel/-iterativecalculation-maxchange-member)|Specifies the maximum amount of change between each iteration as Excel resolves circular references.|
||[maxIteration](/.iterativecalculation#excel-javascript/api/excel/-iterativecalculation-maxiteration-member)|Specifies the maximum number of iterations that Excel can use to resolve a circular reference.|
|[Line](/.line)|[beginArrowheadLength](/.line#excel-javascript/api/excel/-line-beginarrowheadlength-member)|Represents the length of the arrowhead at the beginning of the specified line.|
||[beginArrowheadStyle](/.line#excel-javascript/api/excel/-line-beginarrowheadstyle-member)|Represents the style of the arrowhead at the beginning of the specified line.|
||[beginArrowheadWidth](/.line#excel-javascript/api/excel/-line-beginarrowheadwidth-member)|Represents the width of the arrowhead at the beginning of the specified line.|
||[beginConnectedShape](/.line#excel-javascript/api/excel/-line-beginconnectedshape-member)|Represents the shape to which the beginning of the specified line is attached.|
||[beginConnectedSite](/.line#excel-javascript/api/excel/-line-beginconnectedsite-member)|Represents the connection site to which the beginning of a connector is connected.|
||[connectBeginShape(shape: Excel.Shape, connectionSite: number)](/.line#excel-javascript/api/excel/-line-connectbeginshape-member(1))|Attaches the beginning of the specified connector to a specified shape.|
||[connectEndShape(shape: Excel.Shape, connectionSite: number)](/.line#excel-javascript/api/excel/-line-connectendshape-member(1))|Attaches the end of the specified connector to a specified shape.|
||[connectorType](/.line#excel-javascript/api/excel/-line-connectortype-member)|Represents the connector type for the line.|
||[disconnectBeginShape()](/.line#excel-javascript/api/excel/-line-disconnectbeginshape-member(1))|Detaches the beginning of the specified connector from a shape.|
||[disconnectEndShape()](/.line#excel-javascript/api/excel/-line-disconnectendshape-member(1))|Detaches the end of the specified connector from a shape.|
||[endArrowheadLength](/.line#excel-javascript/api/excel/-line-endarrowheadlength-member)|Represents the length of the arrowhead at the end of the specified line.|
||[endArrowheadStyle](/.line#excel-javascript/api/excel/-line-endarrowheadstyle-member)|Represents the style of the arrowhead at the end of the specified line.|
||[endArrowheadWidth](/.line#excel-javascript/api/excel/-line-endarrowheadwidth-member)|Represents the width of the arrowhead at the end of the specified line.|
||[endConnectedShape](/.line#excel-javascript/api/excel/-line-endconnectedshape-member)|Represents the shape to which the end of the specified line is attached.|
||[endConnectedSite](/.line#excel-javascript/api/excel/-line-endconnectedsite-member)|Represents the connection site to which the end of a connector is connected.|
||[id](/.line#excel-javascript/api/excel/-line-id-member)|Specifies the shape identifier.|
||[isBeginConnected](/.line#excel-javascript/api/excel/-line-isbeginconnected-member)|Specifies if the beginning of the specified line is connected to a shape.|
||[isEndConnected](/.line#excel-javascript/api/excel/-line-isendconnected-member)|Specifies if the end of the specified line is connected to a shape.|
||[shape](/.line#excel-javascript/api/excel/-line-shape-member)|Returns the `Shape` object associated with the line.|
|[PageBreak](/.pagebreak)|[columnIndex](/.pagebreak#excel-javascript/api/excel/-pagebreak-columnindex-member)|Specifies the column index for the page break.|
||[delete()](/.pagebreak#excel-javascript/api/excel/-pagebreak-delete-member(1))|Deletes a page break object.|
||[getCellAfterBreak()](/.pagebreak#excel-javascript/api/excel/-pagebreak-getcellafterbreak-member(1))|Gets the first cell after the page break.|
||[rowIndex](/.pagebreak#excel-javascript/api/excel/-pagebreak-rowindex-member)|Specifies the row index for the page break.|
|[PageBreakCollection](/.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/.pagebreakcollection#excel-javascript/api/excel/-pagebreakcollection-add-member(1))|Adds a page break before the top-left cell of the range specified.|
||[getCount()](/.pagebreakcollection#excel-javascript/api/excel/-pagebreakcollection-getcount-member(1))|Gets the number of page breaks in the collection.|
||[getItem(index: number)](/.pagebreakcollection#excel-javascript/api/excel/-pagebreakcollection-getitem-member(1))|Gets a page break object via the index.|
||[items](/.pagebreakcollection#excel-javascript/api/excel/-pagebreakcollection-items-member)|Gets the loaded child items in this collection.|
||[removePageBreaks()](/.pagebreakcollection#excel-javascript/api/excel/-pagebreakcollection-removepagebreaks-member(1))|Resets all manual page breaks in the collection.|
|[PageLayout](/.pagelayout)|[blackAndWhite](/.pagelayout#excel-javascript/api/excel/-pagelayout-blackandwhite-member)|The worksheet's black and white print option.|
||[bottomMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-bottommargin-member)|The worksheet's bottom page margin to use for printing in points.|
||[centerHorizontally](/.pagelayout#excel-javascript/api/excel/-pagelayout-centerhorizontally-member)|The worksheet's center horizontally flag.|
||[centerVertically](/.pagelayout#excel-javascript/api/excel/-pagelayout-centervertically-member)|The worksheet's center vertically flag.|
||[draftMode](/.pagelayout#excel-javascript/api/excel/-pagelayout-draftmode-member)|The worksheet's draft mode option.|
||[firstPageNumber](/.pagelayout#excel-javascript/api/excel/-pagelayout-firstpagenumber-member)|The worksheet's first page number to print.|
||[footerMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-footermargin-member)|The worksheet's footer margin, in points, for use when printing.|
||[getPrintArea()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprintarea-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintAreaOrNullObject()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprintareaornullobject-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents the print area for the worksheet.|
||[getPrintTitleColumns()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprinttitlecolumns-member(1))|Gets the range object representing the title columns.|
||[getPrintTitleColumnsOrNullObject()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprinttitlecolumnsornullobject-member(1))|Gets the range object representing the title columns.|
||[getPrintTitleRows()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprinttitlerows-member(1))|Gets the range object representing the title rows.|
||[getPrintTitleRowsOrNullObject()](/.pagelayout#excel-javascript/api/excel/-pagelayout-getprinttitlerowsornullobject-member(1))|Gets the range object representing the title rows.|
||[headerMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-headermargin-member)|The worksheet's header margin, in points, for use when printing.|
||[headersFooters](/.pagelayout#excel-javascript/api/excel/-pagelayout-headersfooters-member)|Header and footer configuration for the worksheet.|
||[leftMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-leftmargin-member)|The worksheet's left margin, in points, for use when printing.|
||[orientation](/.pagelayout#excel-javascript/api/excel/-pagelayout-orientation-member)|The worksheet's orientation of the page.|
||[paperSize](/.pagelayout#excel-javascript/api/excel/-pagelayout-papersize-member)|The worksheet's paper size of the page.|
||[printComments](/.pagelayout#excel-javascript/api/excel/-pagelayout-printcomments-member)|Specifies if the worksheet's comments should be displayed when printing.|
||[printErrors](/.pagelayout#excel-javascript/api/excel/-pagelayout-printerrors-member)|The worksheet's print errors option.|
||[printGridlines](/.pagelayout#excel-javascript/api/excel/-pagelayout-printgridlines-member)|Specifies if the worksheet's gridlines will be printed.|
||[printHeadings](/.pagelayout#excel-javascript/api/excel/-pagelayout-printheadings-member)|Specifies if the worksheet's headings will be printed.|
||[printOrder](/.pagelayout#excel-javascript/api/excel/-pagelayout-printorder-member)|The worksheet's page print order option.|
||[rightMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-rightmargin-member)|The worksheet's right margin, in points, for use when printing.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/.pagelayout#excel-javascript/api/excel/-pagelayout-setprintarea-member(1))|Sets the worksheet's print area.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/.pagelayout#excel-javascript/api/excel/-pagelayout-setprintmargins-member(1))|Sets the worksheet's page margins with units.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/.pagelayout#excel-javascript/api/excel/-pagelayout-setprinttitlecolumns-member(1))|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/.pagelayout#excel-javascript/api/excel/-pagelayout-setprinttitlerows-member(1))|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|
||[topMargin](/.pagelayout#excel-javascript/api/excel/-pagelayout-topmargin-member)|The worksheet's top margin, in points, for use when printing.|
||[zoom](/.pagelayout#excel-javascript/api/excel/-pagelayout-zoom-member)|The worksheet's print zoom options.|
|[PageLayoutMarginOptions](/.pagelayoutmarginoptions)|[bottom](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-bottom-member)|Specifies the page layout bottom margin in the unit specified to use for printing.|
||[footer](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-footer-member)|Specifies the page layout footer margin in the unit specified to use for printing.|
||[header](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-header-member)|Specifies the page layout header margin in the unit specified to use for printing.|
||[left](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-left-member)|Specifies the page layout left margin in the unit specified to use for printing.|
||[right](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-right-member)|Specifies the page layout right margin in the unit specified to use for printing.|
||[top](/.pagelayoutmarginoptions#excel-javascript/api/excel/-pagelayoutmarginoptions-top-member)|Specifies the page layout top margin in the unit specified to use for printing.|
|[PageLayoutZoomOptions](/.pagelayoutzoomoptions)|[horizontalFitToPages](/.pagelayoutzoomoptions#excel-javascript/api/excel/-pagelayoutzoomoptions-horizontalfittopages-member)|Number of pages to fit horizontally.|
||[scale](/.pagelayoutzoomoptions#excel-javascript/api/excel/-pagelayoutzoomoptions-scale-member)|Print page scale value can be between 10 and 400.|
||[verticalFitToPages](/.pagelayoutzoomoptions#excel-javascript/api/excel/-pagelayoutzoomoptions-verticalfittopages-member)|Number of pages to fit vertically.|
|[PivotField](/.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/.pivotfield#excel-javascript/api/excel/-pivotfield-sortbyvalues-member(1))|Sorts the PivotField by specified values in a given scope.|
|[PivotLayout](/.pivotlayout)|[autoFormat](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-autoformat-member)|Specifies if formatting will be automatically formatted when it's refreshed or when fields are moved.|
||[getDataHierarchy(cell: Range \| string)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getdatahierarchy-member(1))|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-getpivotitems-member(1))|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|
||[preserveFormatting](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-preserveformatting-member)|Specifies if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/.pivotlayout#excel-javascript/api/excel/-pivotlayout-setautosortoncell-member(1))|Sets the PivotTable to automatically sort using the specified cell to automatically select all necessary criteria and context.|
|[PivotTable](/.pivottable)|[enableDataValueEditing](/.pivottable#excel-javascript/api/excel/-pivottable-enabledatavalueediting-member)|Specifies if the PivotTable allows values in the data body to be edited by the user.|
||[useCustomSortLists](/.pivottable#excel-javascript/api/excel/-pivottable-usecustomsortlists-member)|Specifies if the PivotTable uses custom lists when sorting.|
|[Range](/.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel.AutoFillType)](/.range#excel-javascript/api/excel/-range-autofill-member(1))|Fills a range from the current range to the destination range using the specified AutoFill logic.|
||[convertDataTypeToText()](/.range#excel-javascript/api/excel/-range-convertdatatypetotext-member(1))|Converts the range cells with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/.range#excel-javascript/api/excel/-range-converttolinkeddatatype-member(1))|Converts the range cells into linked data types in the worksheet.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/.range#excel-javascript/api/excel/-range-copyfrom-member(1))|Copies cell data or formatting from the source range or `RangeAreas` to the current range.|
||[find(text: string, criteria: Excel.SearchCriteria)](/.range#excel-javascript/api/excel/-range-find-member(1))|Finds the given string based on the criteria specified.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/.range#excel-javascript/api/excel/-range-findornullobject-member(1))|Finds the given string based on the criteria specified.|
||[flashFill()](/.range#excel-javascript/api/excel/-range-flashfill-member(1))|Does a Flash Fill to the current range.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/.range#excel-javascript/api/excel/-range-getcellproperties-member(1))|Returns a 2D array, encapsulating the data for each cell's font, fill, borders, alignment, and other properties.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/.range#excel-javascript/api/excel/-range-getcolumnproperties-member(1))|Returns a single-dimensional array, encapsulating the data for each column's font, fill, borders, alignment, and other properties.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/.range#excel-javascript/api/excel/-range-getrowproperties-member(1))|Returns a single-dimensional array, encapsulating the data for each row's font, fill, borders, alignment, and other properties.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/.range#excel-javascript/api/excel/-range-getspecialcells-member(1))|Gets the `RangeAreas` object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/.range#excel-javascript/api/excel/-range-getspecialcellsornullobject-member(1))|Gets the `RangeAreas` object, comprising one or more ranges, that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/.range#excel-javascript/api/excel/-range-gettables-member(1))|Gets a scoped collection of tables that overlap with the range.|
||[linkedDataTypeState](/.range#excel-javascript/api/excel/-range-linkeddatatypestate-member)|Represents the data type state of each cell.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/.range#excel-javascript/api/excel/-range-removeduplicates-member(1))|Removes duplicate values from the range specified by the columns.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/.range#excel-javascript/api/excel/-range-replaceall-member(1))|Finds and replaces the given string based on the criteria specified within the current range.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/.range#excel-javascript/api/excel/-range-setcellproperties-member(1))|Updates the range based on a 2D array of cell properties, encapsulating things like font, fill, borders, and alignment.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/.range#excel-javascript/api/excel/-range-setcolumnproperties-member(1))|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, and alignment.|
||[setDirty()](/.range#excel-javascript/api/excel/-range-setdirty-member(1))|Set a range to be recalculated when the next recalculation occurs.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/.range#excel-javascript/api/excel/-range-setrowproperties-member(1))|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, and alignment.|
|[RangeAreas](/.rangeareas)|[address](/.rangeareas#excel-javascript/api/excel/-rangeareas-address-member)|Returns the `RangeAreas` reference in A1-style.|
||[addressLocal](/.rangeareas#excel-javascript/api/excel/-rangeareas-addresslocal-member)|Returns the `RangeAreas` reference in the user locale.|
||[areaCount](/.rangeareas#excel-javascript/api/excel/-rangeareas-areacount-member)|Returns the number of rectangular ranges that comprise this `RangeAreas` object.|
||[areas](/.rangeareas#excel-javascript/api/excel/-rangeareas-areas-member)|Returns a collection of rectangular ranges that comprise this `RangeAreas` object.|
||[calculate()](/.rangeareas#excel-javascript/api/excel/-rangeareas-calculate-member(1))|Calculates all cells in the `RangeAreas`.|
||[cellCount](/.rangeareas#excel-javascript/api/excel/-rangeareas-cellcount-member)|Returns the number of cells in the `RangeAreas` object, summing up the cell counts of all of the individual rectangular ranges.|
||[clear(applyTo?: Excel.ClearApplyTo)](/.rangeareas#excel-javascript/api/excel/-rangeareas-clear-member(1))|Clears values, format, fill, border, and other properties on each of the areas that comprise this `RangeAreas` object.|
||[conditionalFormats](/.rangeareas#excel-javascript/api/excel/-rangeareas-conditionalformats-member)|Returns a collection of conditional formats that intersect with any cells in this `RangeAreas` object.|
||[convertDataTypeToText()](/.rangeareas#excel-javascript/api/excel/-rangeareas-convertdatatypetotext-member(1))|Converts all cells in the `RangeAreas` with data types into text.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/.rangeareas#excel-javascript/api/excel/-rangeareas-converttolinkeddatatype-member(1))|Converts all cells in the `RangeAreas` into linked data types.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/.rangeareas#excel-javascript/api/excel/-rangeareas-copyfrom-member(1))|Copies cell data or formatting from the source range or `RangeAreas` to the current `RangeAreas`.|
||[dataValidation](/.rangeareas#excel-javascript/api/excel/-rangeareas-datavalidation-member)|Returns a data validation object for all ranges in the `RangeAreas`.|
||[format](/.rangeareas#excel-javascript/api/excel/-rangeareas-format-member)|Returns a `RangeFormat` object, encapsulating the font, fill, borders, alignment, and other properties for all ranges in the `RangeAreas` object.|
||[getEntireColumn()](/.rangeareas#excel-javascript/api/excel/-rangeareas-getentirecolumn-member(1))|Returns a `RangeAreas` object that represents the entire columns of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11, H2", it returns a `RangeAreas` that represents columns "B:E, H:H").|
||[getEntireRow()](/.rangeareas#excel-javascript/api/excel/-rangeareas-getentirerow-member(1))|Returns a `RangeAreas` object that represents the entire rows of the `RangeAreas` (for example, if the current `RangeAreas` represents cells "B4:E11", it returns a `RangeAreas` that represents rows "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getintersection-member(1))|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getintersectionornullobject-member(1))|Returns the `RangeAreas` object that represents the intersection of the given ranges or `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getoffsetrangeareas-member(1))|Returns a `RangeAreas` object that is shifted by the specific row and column offset.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getspecialcells-member(1))|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getspecialcellsornullobject-member(1))|Returns a `RangeAreas` object that represents all the cells that match the specified type and value.|
||[getTables(fullyContained?: boolean)](/.rangeareas#excel-javascript/api/excel/-rangeareas-gettables-member(1))|Returns a scoped collection of tables that overlap with any range in this `RangeAreas` object.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getusedrangeareas-member(1))|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/.rangeareas#excel-javascript/api/excel/-rangeareas-getusedrangeareasornullobject-member(1))|Returns the used `RangeAreas` that comprises all the used areas of individual rectangular ranges in the `RangeAreas` object.|
||[isEntireColumn](/.rangeareas#excel-javascript/api/excel/-rangeareas-isentirecolumn-member)|Specifies if all the ranges on this `RangeAreas` object represent entire columns (e.g., "A:C, Q:Z").|
||[isEntireRow](/.rangeareas#excel-javascript/api/excel/-rangeareas-isentirerow-member)|Specifies if all the ranges on this `RangeAreas` object represent entire rows (e.g., "1:3, 5:7").|
||[setDirty()](/.rangeareas#excel-javascript/api/excel/-rangeareas-setdirty-member(1))|Sets the `RangeAreas` to be recalculated when the next recalculation occurs.|
||[style](/.rangeareas#excel-javascript/api/excel/-rangeareas-style-member)|Represents the style for all ranges in this `RangeAreas` object.|
||[worksheet](/.rangeareas#excel-javascript/api/excel/-rangeareas-worksheet-member)|Returns the worksheet for the current `RangeAreas`.|
|[RangeBorder](/.rangeborder)|[tintAndShade](/.rangeborder#excel-javascript/api/excel/-rangeborder-tintandshade-member)|Specifies a double that lightens or darkens a color for the range border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|
|[RangeBorderCollection](/.rangebordercollection)|[tintAndShade](/.rangebordercollection#excel-javascript/api/excel/-rangebordercollection-tintandshade-member)|Specifies a double that lightens or darkens a color for range borders.|
|[RangeCollection](/.rangecollection)|[getCount()](/.rangecollection#excel-javascript/api/excel/-rangecollection-getcount-member(1))|Returns the number of ranges in the `RangeCollection`.|
||[getItemAt(index: number)](/.rangecollection#excel-javascript/api/excel/-rangecollection-getitemat-member(1))|Returns the range object based on its position in the `RangeCollection`.|
||[items](/.rangecollection#excel-javascript/api/excel/-rangecollection-items-member)|Gets the loaded child items in this collection.|
|[RangeFill](/.rangefill)|[pattern](/.rangefill#excel-javascript/api/excel/-rangefill-pattern-member)|The pattern of a range.|
||[patternColor](/.rangefill#excel-javascript/api/excel/-rangefill-patterncolor-member)|The HTML color code representing the color of the range pattern, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").|
||[patternTintAndShade](/.rangefill#excel-javascript/api/excel/-rangefill-patterntintandshade-member)|Specifies a double that lightens or darkens a pattern color for the range fill.|
||[tintAndShade](/.rangefill#excel-javascript/api/excel/-rangefill-tintandshade-member)|Specifies a double that lightens or darkens a color for the range fill.|
|[RangeFont](/.rangefont)|[strikethrough](/.rangefont#excel-javascript/api/excel/-rangefont-strikethrough-member)|Specifies the strikethrough status of font.|
||[subscript](/.rangefont#excel-javascript/api/excel/-rangefont-subscript-member)|Specifies the subscript status of font.|
||[superscript](/.rangefont#excel-javascript/api/excel/-rangefont-superscript-member)|Specifies the superscript status of font.|
||[tintAndShade](/.rangefont#excel-javascript/api/excel/-rangefont-tintandshade-member)|Specifies a double that lightens or darkens a color for the range font.|
|[RangeFormat](/.rangeformat)|[autoIndent](/.rangeformat#excel-javascript/api/excel/-rangeformat-autoindent-member)|Specifies if text is automatically indented when text alignment is set to equal distribution.|
||[indentLevel](/.rangeformat#excel-javascript/api/excel/-rangeformat-indentlevel-member)|An integer from 0 to 250 that indicates the indent level.|
||[readingOrder](/.rangeformat#excel-javascript/api/excel/-rangeformat-readingorder-member)|The reading order for the range.|
||[shrinkToFit](/.rangeformat#excel-javascript/api/excel/-rangeformat-shrinktofit-member)|Specifies if text automatically shrinks to fit in the available column width.|
|[RemoveDuplicatesResult](/.removeduplicatesresult)|[removed](/.removeduplicatesresult#excel-javascript/api/excel/-removeduplicatesresult-removed-member)|Number of duplicated rows removed by the operation.|
||[uniqueRemaining](/.removeduplicatesresult#excel-javascript/api/excel/-removeduplicatesresult-uniqueremaining-member)|Number of remaining unique rows present in the resulting range.|
|[ReplaceCriteria](/.replacecriteria)|[completeMatch](/.replacecriteria#excel-javascript/api/excel/-replacecriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/.replacecriteria#excel-javascript/api/excel/-replacecriteria-matchcase-member)|Specifies if the match is case-sensitive.|
|[RowProperties](/.rowproperties)|[address](/.rowproperties#excel-javascript/api/excel/-rowproperties-address-member)|Represents the `address` property.|
||[addressLocal](/.rowproperties#excel-javascript/api/excel/-rowproperties-addresslocal-member)|Represents the `addressLocal` property.|
||[rowIndex](/.rowproperties#excel-javascript/api/excel/-rowproperties-rowindex-member)|Represents the `rowIndex` property.|
|[SearchCriteria](/.searchcriteria)|[completeMatch](/.searchcriteria#excel-javascript/api/excel/-searchcriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/.searchcriteria#excel-javascript/api/excel/-searchcriteria-matchcase-member)|Specifies if the match is case-sensitive.|
||[searchDirection](/.searchcriteria#excel-javascript/api/excel/-searchcriteria-searchdirection-member)|Specifies the search direction.|
|[SettableCellProperties](/.settablecellproperties)|[format](/.settablecellproperties#excel-javascript/api/excel/-settablecellproperties-format-member)|Represents the `format` property.|
||[hyperlink](/.settablecellproperties#excel-javascript/api/excel/-settablecellproperties-hyperlink-member)|Represents the `hyperlink` property.|
||[style](/.settablecellproperties#excel-javascript/api/excel/-settablecellproperties-style-member)|Represents the `style` property.|
|[SettableColumnProperties](/.settablecolumnproperties)|[columnHidden](/.settablecolumnproperties#excel-javascript/api/excel/-settablecolumnproperties-columnhidden-member)|Represents the `columnHidden` property.|
||[format](/.settablecolumnproperties#excel-javascript/api/excel/-settablecolumnproperties-format-member)|Represents the `format` property.|
|[SettableRowProperties](/.settablerowproperties)|[format](/.settablerowproperties#excel-javascript/api/excel/-settablerowproperties-format-member)|Represents the `format` property.|
||[rowHidden](/.settablerowproperties#excel-javascript/api/excel/-settablerowproperties-rowhidden-member)|Represents the `rowHidden` property.|
|[Shape](/.shape)|[altTextDescription](/.shape#excel-javascript/api/excel/-shape-alttextdescription-member)|Specifies the alternative description text for a `Shape` object.|
||[altTextTitle](/.shape#excel-javascript/api/excel/-shape-alttexttitle-member)|Specifies the alternative title text for a `Shape` object.|
||[connectionSiteCount](/.shape#excel-javascript/api/excel/-shape-connectionsitecount-member)|Returns the number of connection sites on this shape.|
||[delete()](/.shape#excel-javascript/api/excel/-shape-delete-member(1))|Removes the shape from the worksheet.|
||[fill](/.shape#excel-javascript/api/excel/-shape-fill-member)|Returns the fill formatting of this shape.|
||[geometricShape](/.shape#excel-javascript/api/excel/-shape-geometricshape-member)|Returns the geometric shape associated with the shape.|
||[geometricShapeType](/.shape#excel-javascript/api/excel/-shape-geometricshapetype-member)|Specifies the geometric shape type of this geometric shape.|
||[getAsImage(format: Excel.PictureFormat)](/.shape#excel-javascript/api/excel/-shape-getasimage-member(1))|Converts the shape to an image and returns the image as a Base64-encoded string.|
||[group](/.shape#excel-javascript/api/excel/-shape-group-member)|Returns the shape group associated with the shape.|
||[height](/.shape#excel-javascript/api/excel/-shape-height-member)|Specifies the height, in points, of the shape.|
||[id](/.shape#excel-javascript/api/excel/-shape-id-member)|Specifies the shape identifier.|
||[image](/.shape#excel-javascript/api/excel/-shape-image-member)|Returns the image associated with the shape.|
||[incrementLeft(increment: number)](/.shape#excel-javascript/api/excel/-shape-incrementleft-member(1))|Moves the shape horizontally by the specified number of points.|
||[incrementRotation(increment: number)](/.shape#excel-javascript/api/excel/-shape-incrementrotation-member(1))|Rotates the shape clockwise around the z-axis by the specified number of degrees.|
||[incrementTop(increment: number)](/.shape#excel-javascript/api/excel/-shape-incrementtop-member(1))|Moves the shape vertically by the specified number of points.|
||[left](/.shape#excel-javascript/api/excel/-shape-left-member)|The distance, in points, from the left side of the shape to the left side of the worksheet.|
||[level](/.shape#excel-javascript/api/excel/-shape-level-member)|Specifies the level of the specified shape.|
||[line](/.shape#excel-javascript/api/excel/-shape-line-member)|Returns the line associated with the shape.|
||[lineFormat](/.shape#excel-javascript/api/excel/-shape-lineformat-member)|Returns the line formatting of this shape.|
||[lockAspectRatio](/.shape#excel-javascript/api/excel/-shape-lockaspectratio-member)|Specifies if the aspect ratio of this shape is locked.|
||[name](/.shape#excel-javascript/api/excel/-shape-name-member)|Specifies the name of the shape.|
||[onActivated](/.shape#excel-javascript/api/excel/-shape-onactivated-member)|Occurs when the shape is activated.|
||[onDeactivated](/.shape#excel-javascript/api/excel/-shape-ondeactivated-member)|Occurs when the shape is deactivated.|
||[parentGroup](/.shape#excel-javascript/api/excel/-shape-parentgroup-member)|Specifies the parent group of this shape.|
||[rotation](/.shape#excel-javascript/api/excel/-shape-rotation-member)|Specifies the rotation, in degrees, of the shape.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/.shape#excel-javascript/api/excel/-shape-scaleheight-member(1))|Scales the height of the shape by a specified factor.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/.shape#excel-javascript/api/excel/-shape-scalewidth-member(1))|Scales the width of the shape by a specified factor.|
||[setZOrder(position: Excel.ShapeZOrder)](/.shape#excel-javascript/api/excel/-shape-setzorder-member(1))|Moves the specified shape up or down the collection's z-order, which shifts it in front of or behind other shapes.|
||[textFrame](/.shape#excel-javascript/api/excel/-shape-textframe-member)|Returns the text frame object of this shape.|
||[top](/.shape#excel-javascript/api/excel/-shape-top-member)|The distance, in points, from the top edge of the shape to the top edge of the worksheet.|
||[type](/.shape#excel-javascript/api/excel/-shape-type-member)|Returns the type of this shape.|
||[visible](/.shape#excel-javascript/api/excel/-shape-visible-member)|Specifies if the shape is visible.|
||[width](/.shape#excel-javascript/api/excel/-shape-width-member)|Specifies the width, in points, of the shape.|
||[zOrderPosition](/.shape#excel-javascript/api/excel/-shape-zorderposition-member)|Returns the position of the specified shape in the z-order, with 0 representing the bottom of the order stack.|
|[ShapeActivatedEventArgs](/.shapeactivatedeventargs)|[shapeId](/.shapeactivatedeventargs#excel-javascript/api/excel/-shapeactivatedeventargs-shapeid-member)|Gets the ID of the activated shape.|
||[type](/.shapeactivatedeventargs#excel-javascript/api/excel/-shapeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.shapeactivatedeventargs#excel-javascript/api/excel/-shapeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the shape is activated.|
|[ShapeCollection](/.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addgeometricshape-member(1))|Adds a geometric shape to the worksheet.|
||[addGroup(values: Array<string \| Shape>)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addgroup-member(1))|Groups a subset of shapes in this collection's worksheet.|
||[addImage(base64ImageString: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addimage-member(1))|Creates an image from a Base64-encoded string and adds it to the worksheet.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addline-member(1))|Adds a line to worksheet.|
||[addTextBox(text?: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-addtextbox-member(1))|Adds a text box to the worksheet with the provided text as the content.|
||[getCount()](/.shapecollection#excel-javascript/api/excel/-shapecollection-getcount-member(1))|Returns the number of shapes in the worksheet.|
||[getItem(key: string)](/.shapecollection#excel-javascript/api/excel/-shapecollection-getitem-member(1))|Gets a shape using its name or ID.|
||[getItemAt(index: number)](/.shapecollection#excel-javascript/api/excel/-shapecollection-getitemat-member(1))|Gets a shape using its position in the collection.|
||[items](/.shapecollection#excel-javascript/api/excel/-shapecollection-items-member)|Gets the loaded child items in this collection.|
|[ShapeDeactivatedEventArgs](/.shapedeactivatedeventargs)|[shapeId](/.shapedeactivatedeventargs#excel-javascript/api/excel/-shapedeactivatedeventargs-shapeid-member)|Gets the ID of the shape deactivated shape.|
||[type](/.shapedeactivatedeventargs#excel-javascript/api/excel/-shapedeactivatedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.shapedeactivatedeventargs#excel-javascript/api/excel/-shapedeactivatedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the shape is deactivated.|
|[ShapeFill](/.shapefill)|[clear()](/.shapefill#excel-javascript/api/excel/-shapefill-clear-member(1))|Clears the fill formatting of this shape.|
||[foregroundColor](/.shapefill#excel-javascript/api/excel/-shapefill-foregroundcolor-member)|Represents the shape fill foreground color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange")|
||[setSolidColor(color: string)](/.shapefill#excel-javascript/api/excel/-shapefill-setsolidcolor-member(1))|Sets the fill formatting of the shape to a uniform color.|
||[transparency](/.shapefill#excel-javascript/api/excel/-shapefill-transparency-member)|Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear).|
||[type](/.shapefill#excel-javascript/api/excel/-shapefill-type-member)|Returns the fill type of the shape.|
|[ShapeFont](/.shapefont)|[bold](/.shapefont#excel-javascript/api/excel/-shapefont-bold-member)|Represents the bold status of font.|
||[color](/.shapefont#excel-javascript/api/excel/-shapefont-color-member)|HTML color code representation of the text color (e.g., "#FF0000" represents red).|
||[italic](/.shapefont#excel-javascript/api/excel/-shapefont-italic-member)|Represents the italic status of font.|
||[name](/.shapefont#excel-javascript/api/excel/-shapefont-name-member)|Represents font name (e.g., "Calibri").|
||[size](/.shapefont#excel-javascript/api/excel/-shapefont-size-member)|Represents font size in points (e.g., 11).|
||[underline](/.shapefont#excel-javascript/api/excel/-shapefont-underline-member)|Type of underline applied to the font.|
|[ShapeGroup](/.shapegroup)|[id](/.shapegroup#excel-javascript/api/excel/-shapegroup-id-member)|Specifies the shape identifier.|
||[shape](/.shapegroup#excel-javascript/api/excel/-shapegroup-shape-member)|Returns the `Shape` object associated with the group.|
||[shapes](/.shapegroup#excel-javascript/api/excel/-shapegroup-shapes-member)|Returns the collection of `Shape` objects.|
||[ungroup()](/.shapegroup#excel-javascript/api/excel/-shapegroup-ungroup-member(1))|Ungroups any grouped shapes in the specified shape group.|
|[ShapeLineFormat](/.shapelineformat)|[color](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-color-member)|Represents the line color in HTML color format, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[dashStyle](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-dashstyle-member)|Represents the line style of the shape.|
||[style](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-style-member)|Represents the line style of the shape.|
||[transparency](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-transparency-member)|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear).|
||[visible](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-visible-member)|Specifies if the line formatting of a shape element is visible.|
||[weight](/.shapelineformat#excel-javascript/api/excel/-shapelineformat-weight-member)|Represents the weight of the line, in points.|
|[SortField](/.sortfield)|[subField](/.sortfield#excel-javascript/api/excel/-sortfield-subfield-member)|Specifies the subfield that is the target property name of a rich value to sort on.|
|[StyleCollection](/.stylecollection)|[getCount()](/.stylecollection#excel-javascript/api/excel/-stylecollection-getcount-member(1))|Gets the number of styles in the collection.|
||[getItemAt(index: number)](/.stylecollection#excel-javascript/api/excel/-stylecollection-getitemat-member(1))|Gets a style based on its position in the collection.|
|[Table](/.table)|[autoFilter](/.table#excel-javascript/api/excel/-table-autofilter-member)|Represents the `AutoFilter` object of the table.|
|[TableAddedEventArgs](/.tableaddedeventargs)|[source](/.tableaddedeventargs#excel-javascript/api/excel/-tableaddedeventargs-source-member)|Gets the source of the event.|
||[tableId](/.tableaddedeventargs#excel-javascript/api/excel/-tableaddedeventargs-tableid-member)|Gets the ID of the table that is added.|
||[type](/.tableaddedeventargs#excel-javascript/api/excel/-tableaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.tableaddedeventargs#excel-javascript/api/excel/-tableaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the table is added.|
|[TableChangedEventArgs](/.tablechangedeventargs)|[details](/.tablechangedeventargs#excel-javascript/api/excel/-tablechangedeventargs-details-member)|Gets the information about the change detail.|
|[TableCollection](/.tablecollection)|[onAdded](/.tablecollection#excel-javascript/api/excel/-tablecollection-onadded-member)|Occurs when a new table is added in a workbook.|
||[onDeleted](/.tablecollection#excel-javascript/api/excel/-tablecollection-ondeleted-member)|Occurs when the specified table is deleted in a workbook.|
|[TableDeletedEventArgs](/.tabledeletedeventargs)|[source](/.tabledeletedeventargs#excel-javascript/api/excel/-tabledeletedeventargs-source-member)|Gets the source of the event.|
||[tableId](/.tabledeletedeventargs#excel-javascript/api/excel/-tabledeletedeventargs-tableid-member)|Gets the ID of the table that is deleted.|
||[tableName](/.tabledeletedeventargs#excel-javascript/api/excel/-tabledeletedeventargs-tablename-member)|Gets the name of the table that is deleted.|
||[type](/.tabledeletedeventargs#excel-javascript/api/excel/-tabledeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.tabledeletedeventargs#excel-javascript/api/excel/-tabledeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the table is deleted.|
|[TableScopedCollection](/.tablescopedcollection)|[getCount()](/.tablescopedcollection#excel-javascript/api/excel/-tablescopedcollection-getcount-member(1))|Gets the number of tables in the collection.|
||[getFirst()](/.tablescopedcollection#excel-javascript/api/excel/-tablescopedcollection-getfirst-member(1))|Gets the first table in the collection.|
||[getItem(key: string)](/.tablescopedcollection#excel-javascript/api/excel/-tablescopedcollection-getitem-member(1))|Gets a table by name or ID.|
||[items](/.tablescopedcollection#excel-javascript/api/excel/-tablescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[TextFrame](/.textframe)|[autoSizeSetting](/.textframe#excel-javascript/api/excel/-textframe-autosizesetting-member)|The automatic sizing settings for the text frame.|
||[bottomMargin](/.textframe#excel-javascript/api/excel/-textframe-bottommargin-member)|Represents the bottom margin, in points, of the text frame.|
||[deleteText()](/.textframe#excel-javascript/api/excel/-textframe-deletetext-member(1))|Deletes all the text in the text frame.|
||[hasText](/.textframe#excel-javascript/api/excel/-textframe-hastext-member)|Specifies if the text frame contains text.|
||[horizontalAlignment](/.textframe#excel-javascript/api/excel/-textframe-horizontalalignment-member)|Represents the horizontal alignment of the text frame.|
||[horizontalOverflow](/.textframe#excel-javascript/api/excel/-textframe-horizontaloverflow-member)|Represents the horizontal overflow behavior of the text frame.|
||[leftMargin](/.textframe#excel-javascript/api/excel/-textframe-leftmargin-member)|Represents the left margin, in points, of the text frame.|
||[orientation](/.textframe#excel-javascript/api/excel/-textframe-orientation-member)|Represents the angle to which the text is oriented for the text frame.|
||[readingOrder](/.textframe#excel-javascript/api/excel/-textframe-readingorder-member)|Represents the reading order of the text frame, either left-to-right or right-to-left.|
||[rightMargin](/.textframe#excel-javascript/api/excel/-textframe-rightmargin-member)|Represents the right margin, in points, of the text frame.|
||[textRange](/.textframe#excel-javascript/api/excel/-textframe-textrange-member)|Represents the text that is attached to a shape in the text frame, and properties and methods for manipulating the text.|
||[topMargin](/.textframe#excel-javascript/api/excel/-textframe-topmargin-member)|Represents the top margin, in points, of the text frame.|
||[verticalAlignment](/.textframe#excel-javascript/api/excel/-textframe-verticalalignment-member)|Represents the vertical alignment of the text frame.|
||[verticalOverflow](/.textframe#excel-javascript/api/excel/-textframe-verticaloverflow-member)|Represents the vertical overflow behavior of the text frame.|
|[TextRange](/.textrange)|[font](/.textrange#excel-javascript/api/excel/-textrange-font-member)|Returns a `ShapeFont` object that represents the font attributes for the text range.|
||[getSubstring(start: number, length?: number)](/.textrange#excel-javascript/api/excel/-textrange-getsubstring-member(1))|Returns a TextRange object for the substring in the given range.|
||[text](/.textrange#excel-javascript/api/excel/-textrange-text-member)|Represents the plain text content of the text range.|
|[Workbook](/.workbook)|[autoSave](/.workbook#excel-javascript/api/excel/-workbook-autosave-member)|Specifies if the workbook is in AutoSave mode.|
||[calculationEngineVersion](/.workbook#excel-javascript/api/excel/-workbook-calculationengineversion-member)|Returns a number about the version of Excel Calculation Engine.|
||[chartDataPointTrack](/.workbook#excel-javascript/api/excel/-workbook-chartdatapointtrack-member)|True if all charts in the workbook are tracking the actual data points to which they are attached.|
||[getActiveChart()](/.workbook#excel-javascript/api/excel/-workbook-getactivechart-member(1))|Gets the currently active chart in the workbook.|
||[getActiveChartOrNullObject()](/.workbook#excel-javascript/api/excel/-workbook-getactivechartornullobject-member(1))|Gets the currently active chart in the workbook.|
||[getIsActiveCollabSession()](/.workbook#excel-javascript/api/excel/-workbook-getisactivecollabsession-member(1))|Returns `true` if the workbook is being edited by multiple users (through co-authoring).|
||[getSelectedRanges()](/.workbook#excel-javascript/api/excel/-workbook-getselectedranges-member(1))|Gets the currently selected one or more ranges from the workbook.|
||[isDirty](/.workbook#excel-javascript/api/excel/-workbook-isdirty-member)|Specifies if changes have been made since the workbook was last saved.|
||[onAutoSaveSettingChanged](/.workbook#excel-javascript/api/excel/-workbook-onautosavesettingchanged-member)|Occurs when the AutoSave setting is changed on the workbook.|
||[previouslySaved](/.workbook#excel-javascript/api/excel/-workbook-previouslysaved-member)|Specifies if the workbook has ever been saved locally or online.|
||[usePrecisionAsDisplayed](/.workbook#excel-javascript/api/excel/-workbook-useprecisionasdisplayed-member)|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|
|[WorkbookAutoSaveSettingChangedEventArgs](/.workbookautosavesettingchangedeventargs)|[type](/.workbookautosavesettingchangedeventargs#excel-javascript/api/excel/-workbookautosavesettingchangedeventargs-type-member)|Gets the type of the event.|
|[Worksheet](/.worksheet)|[autoFilter](/.worksheet#excel-javascript/api/excel/-worksheet-autofilter-member)|Represents the `AutoFilter` object of the worksheet.|
||[enableCalculation](/.worksheet#excel-javascript/api/excel/-worksheet-enablecalculation-member)|Determines if Excel should recalculate the worksheet when necessary.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/.worksheet#excel-javascript/api/excel/-worksheet-findall-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/.worksheet#excel-javascript/api/excel/-worksheet-findallornullobject-member(1))|Finds all occurrences of the given string based on the criteria specified and returns them as a `RangeAreas` object, comprising one or more rectangular ranges.|
||[getRanges(address?: string)](/.worksheet#excel-javascript/api/excel/-worksheet-getranges-member(1))|Gets the `RangeAreas` object, representing one or more blocks of rectangular ranges, specified by the address or name.|
||[horizontalPageBreaks](/.worksheet#excel-javascript/api/excel/-worksheet-horizontalpagebreaks-member)|Gets the horizontal page break collection for the worksheet.|
||[onFormatChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onformatchanged-member)|Occurs when format changed on a specific worksheet.|
||[pageLayout](/.worksheet#excel-javascript/api/excel/-worksheet-pagelayout-member)|Gets the `PageLayout` object of the worksheet.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/.worksheet#excel-javascript/api/excel/-worksheet-replaceall-member(1))|Finds and replaces the given string based on the criteria specified within the current worksheet.|
||[shapes](/.worksheet#excel-javascript/api/excel/-worksheet-shapes-member)|Returns the collection of all the Shape objects on the worksheet.|
||[verticalPageBreaks](/.worksheet#excel-javascript/api/excel/-worksheet-verticalpagebreaks-member)|Gets the vertical page break collection for the worksheet.|
|[WorksheetChangedEventArgs](/.worksheetchangedeventargs)|[details](/.worksheetchangedeventargs#excel-javascript/api/excel/-worksheetchangedeventargs-details-member)|Represents the information about the change detail.|
|[WorksheetCollection](/.worksheetcollection)|[onChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onchanged-member)|Occurs when any worksheet in the workbook is changed.|
||[onFormatChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onformatchanged-member)|Occurs when any worksheet in the workbook has a format changed.|
||[onSelectionChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onselectionchanged-member)|Occurs when the selection changes on any worksheet.|
|[WorksheetFormatChangedEventArgs](/.worksheetformatchangedeventargs)|[address](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[getRange(ctx: Excel.RequestContext)](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-getrange-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Gets the range that represents the changed area of a specific worksheet.|
||[source](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-source-member)|Gets the source of the event.|
||[type](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetformatchangedeventargs#excel-javascript/api/excel/-worksheetformatchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
|[WorksheetSearchCriteria](/.worksheetsearchcriteria)|[completeMatch](/.worksheetsearchcriteria#excel-javascript/api/excel/-worksheetsearchcriteria-completematch-member)|Specifies if the match needs to be complete or partial.|
||[matchCase](/.worksheetsearchcriteria#excel-javascript/api/excel/-worksheetsearchcriteria-matchcase-member)|Specifies if the match is case-sensitive.|
