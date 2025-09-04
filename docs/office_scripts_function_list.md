# Complete Office Scripts API Function Reference

This document provides a comprehensive list of all available functions for manipulating Office products through Office Scripts. These functions are specifically for Excel automation and do not include core TypeScript/JavaScript functions.

## Workbook Object

### Workbook Access & Properties

- `getActiveCell()` - Gets the currently active cell
- `getActiveChart()` - Gets the currently active chart (returns undefined if none)
- `getActiveSlicer()` - Gets the currently active slicer (returns undefined if none)
- `getActiveWorksheet()` - Gets the currently active worksheet
- `getApplication()` - Gets the Excel application instance
- `getAutoSave()` - Gets AutoSave mode status
- `getName()` - Gets the workbook name
- `getIsDirty()` - Checks if workbook has unsaved changes
- `getPreviouslySaved()` - Checks if workbook was ever saved
- `getReadOnly()` - Checks if workbook is in read-only mode
- `getProperties()` - Gets workbook properties object

### Workbook Navigation

- `getFirstWorksheet(visibleOnly?)` - Gets first worksheet
- `getLastWorksheet(visibleOnly?)` - Gets last worksheet
- `getSelectedRange()` - Gets currently selected single range
- `getSelectedRanges()` - Gets all currently selected ranges

### Workbook Structure Management

- `addWorksheet(name?)` - Adds new worksheet
- `getWorksheet(name)` - Gets worksheet by name
- `getWorksheets()` - Gets all worksheets

### Workbook Data Objects

- `getTables()` - Gets all tables in workbook
- `getTable(name)` - Gets table by name
- `getCharts()` - Gets all charts in workbook
- `getChart(name)` - Gets chart by name
- `getPivotTables()` - Gets all pivot tables
- `getPivotTable(name)` - Gets pivot table by name
- `getSlicers()` - Gets all slicers
- `getSlicer(key)` - Gets slicer by name/ID

### Workbook Comments & Protection

- `getComments()` - Gets all comments
- `getComment(commentId)` - Gets comment by ID
- `getCommentByCell(cellAddress)` - Gets comment from specific cell
- `getCommentByReplyId(replyId)` - Gets comment by reply ID
- `getProtection()` - Gets workbook protection object

### Workbook Styles & Formatting

- `getDefaultTableStyle()` - Gets default table style
- `getDefaultPivotTableStyle()` - Gets default pivot table style
- `getDefaultSlicerStyle()` - Gets default slicer style
- `getDefaultTimelineStyle()` - Gets default timeline style
- `getPredefinedCellStyle(name)` - Gets predefined cell style by name
- `getPredefinedCellStyles()` - Gets all predefined cell styles

### Workbook Bindings & XML

- `getBinding(id)` - Gets binding by ID
- `getBindings()` - Gets all bindings
- `getCustomXmlPart(id)` - Gets custom XML part by ID
- `getCustomXmlParts()` - Gets all custom XML parts
- `getCustomXmlPartByNamespace(namespaceUri)` - Gets XML parts by namespace

### Workbook Calculations & External Links

- `getCalculationEngineVersion()` - Gets calculation engine version
- `getChartDataPointTrack()` - Gets chart data point tracking setting
- `getLinkedWorkbooks()` - Gets all linked workbooks
- `getLinkedWorkbookByUrl(key)` - Gets linked workbook by URL
- `getLinkedWorkbookRefreshMode()` - Gets refresh mode for linked workbooks
- `breakAllLinksToLinkedWorkbooks()` - Breaks all external links

### Workbook Named Items & Queries

- `getNamedItem(name)` - Gets named item by name
- `getNames()` - Gets workbook-scoped named items
- `getQueries()` - Gets Power Query queries
- `getQuery(key)` - Gets query by name

## Worksheet Object

### Worksheet Access & Properties

- `getName()` - Gets worksheet name
- `getPosition()` - Gets zero-based position in workbook
- `getTabColor()` - Gets tab color
- `getStandardHeight()` - Gets standard row height
- `getStandardWidth()` - Gets standard column width
- `getVisibility()` - Gets worksheet visibility state
- `getShowGridlines()` - Gets gridlines visibility
- `getShowHeadings()` - Gets headings visibility
- `getShowDataTypeIcons()` - Gets data type icons visibility

### Worksheet Navigation

- `activate()` - Activates the worksheet
- `getNext(visibleOnly?)` - Gets next worksheet
- `getPrevious(visibleOnly?)` - Gets previous worksheet

### Worksheet Ranges

- `getRange(address)` - Gets range by A1-style address
- `getRangeByIndexes(startRow, startColumn, rowCount, columnCount)` - Gets range by indices
- `getRanges(address)` - Gets multiple ranges
- `getCell(row, column)` - Gets single cell by coordinates
- `getUsedRange(valuesOnly?)` - Gets used range
- `getActiveNamedSheetView()` - Gets active named sheet view

### Worksheet Data Objects

- `getTables()` - Gets all tables
- `getTable(name)` - Gets table by name
- `addTable(address, hasHeaders)` - Creates new table
- `getCharts()` - Gets all charts
- `getChart(name)` - Gets chart by name
- `addChart(chartType, sourceData, seriesBy?)` - Creates new chart
- `getPivotTables()` - Gets all pivot tables
- `getPivotTable(name)` - Gets pivot table by name

### Worksheet Shapes

- `getShapes()` - Gets all shapes
- `getShape(key)` - Gets shape by name/ID
- `addGeometricShape(geometricShapeType)` - Adds geometric shape
- `addLine(startLeft, startTop, endLeft, endTop, connectorType?)` - Adds line
- `addTextBox(text?)` - Adds text box

### Worksheet Comments

- `getComments()` - Gets all comments
- `getComment(commentId)` - Gets comment by ID
- `addComment(cellAddress, content, contentType?)` - Adds new comment

### Worksheet Protection & Filtering

- `getProtection()` - Gets worksheet protection
- `getAutoFilter()` - Gets AutoFilter object
- `getSlicers()` - Gets all slicers
- `getSlicer(key)` - Gets slicer by name/ID

### Worksheet Page Layout

- `getPageLayout()` - Gets page layout object
- `addHorizontalPageBreak(pageBreakRange)` - Adds horizontal page break
- `addVerticalPageBreak(pageBreakRange)` - Adds vertical page break

### Worksheet Operations

- `calculate(markAllDirty)` - Calculates all cells
- `copy(positionType?, relativeTo?)` - Copies worksheet
- `delete()` - Deletes worksheet
- `findAll(text, criteria)` - Finds all matching text
- `replaceAll(text, replacement, criteria)` - Replaces all matching text

### Worksheet Custom Properties

- `addWorksheetCustomProperty(key, value)` - Adds custom property
- `getWorksheetCustomProperties()` - Gets all custom properties

### Worksheet Named Sheet Views

- `getNamedSheetView(key)` - Gets named sheet view
- `getNamedSheetViews()` - Gets all named sheet views

## Range Object

### Range Access & Properties

- `getAddress()` - Gets A1-style address
- `getAddressLocal()` - Gets localized address
- `getCellCount()` - Gets number of cells
- `getRowCount()` - Gets number of rows
- `getColumnCount()` - Gets number of columns
- `getRowIndex()` - Gets zero-based row index
- `getColumnIndex()` - Gets zero-based column index
- `getHeight()` - Gets height in points
- `getWidth()` - Gets width in points
- `getTop()` - Gets distance from top of worksheet
- `getLeft()` - Gets distance from left of worksheet
- `getWorksheet()` - Gets containing worksheet

### Range Navigation

- `getCell(row, column)` - Gets single cell by relative position
- `getColumn(column)` - Gets specific column
- `getRow(row)` - Gets specific row
- `getEntireColumn()` - Gets entire column(s)
- `getEntireRow()` - Gets entire row(s)
- `getOffsetRange(rowOffset, columnOffset)` - Gets offset range
- `getResizedRange(deltaRows, deltaColumns)` - Gets resized range
- `getAbsoluteResizedRange(numRows, numColumns)` - Gets absolute resized range
- `getBoundingRect(anotherRange)` - Gets bounding rectangle with another range
- `getRangeEdge(direction, activeCell?)` - Gets edge range in direction
- `getExtendedRange(direction, activeCell?)` - Gets extended range to edge
- `getUsedRange(valuesOnly?)` - Gets used portion of range

### Range Values & Formulas

- `getValue()` - Gets single cell value
- `getValues()` - Gets all values as 2D array
- `setValue(value)` - Sets single cell value
- `setValues(values)` - Sets all values from 2D array
- `getValueType()` - Gets data type of single cell
- `getValueTypes()` - Gets data types of all cells
- `getText()` - Gets text representation of single cell
- `getTexts()` - Gets text representations as 2D array
- `getFormula()` - Gets formula of single cell
- `getFormulas()` - Gets all formulas as 2D array
- `setFormula(formula)` - Sets formula for single cell
- `setFormulas(formulas)` - Sets all formulas from 2D array
- `getFormulaLocal()` - Gets localized formula
- `getFormulasLocal()` - Gets all localized formulas
- `setFormulaLocal(formulaLocal)` - Sets localized formula
- `setFormulasLocal(formulasLocal)` - Sets all localized formulas
- `getFormulaR1C1()` - Gets R1C1-style formula
- `getFormulasR1C1()` - Gets all R1C1-style formulas
- `setFormulaR1C1(formulaR1C1)` - Sets R1C1-style formula
- `setFormulasR1C1(formulasR1C1)` - Sets all R1C1-style formulas

### Range Formatting

- `getFormat()` - Gets RangeFormat object
- `getNumberFormat()` - Gets number format
- `getNumberFormats()` - Gets all number formats
- `setNumberFormat(numberFormat)` - Sets number format
- `setNumberFormats(numberFormats)` - Sets all number formats
- `getNumberFormatLocal()` - Gets localized number format
- `getNumberFormatsLocal()` - Gets all localized number formats
- `getNumberFormatCategory()` - Gets number format category
- `getNumberFormatCategories()` - Gets all number format categories

### Range Visibility & State

- `getHidden()` - Gets hidden state
- `getRowHidden()` - Gets row hidden state
- `getColumnHidden()` - Gets column hidden state
- `setRowHidden(rowHidden)` - Sets row hidden state
- `setColumnHidden(columnHidden)` - Sets column hidden state
- `getHasSpill()` - Checks for spill borders
- `getIsEntireColumn()` - Checks if range is entire column(s)
- `getIsEntireRow()` - Checks if range is entire row(s)

### Range Operations

- `clear(applyTo?)` - Clears range content/formatting
- `copy(destinationRange?, copyType?, skipBlanks?, transpose?)` - Copies range
- `copyFrom(sourceRange, copyType?, skipBlanks?, transpose?)` - Copies from source
- `cut(destinationRange?)` - Cuts range
- `delete(shift)` - Deletes range and shifts cells
- `insert(shift)` - Inserts range and shifts cells
- `merge(across?)` - Merges cells
- `unmerge()` - Unmerges cells
- `select()` - Selects range in UI

### Range Data Manipulation

- `moveTo(destinationRange)` - Moves range to destination
- `removeDuplicates(columns, includesHeader)` - Removes duplicate rows
- `replaceAll(text, replacement, criteria)` - Replaces text
- `autoFill(destinationRange, autoFillType?)` - Auto fills range
- `flashFill()` - Performs Flash Fill

### Range Analysis

- `find(text, criteria)` - Finds text in range
- `findAll(text, criteria)` - Finds all instances of text
- `getSpecialCells(cellType, cellValueType?)` - Gets special cells
- `getIntersection(anotherRange)` - Gets intersection with another range
- `getColumnsAfter(count)` - Gets columns after range
- `getColumnsBefore(count)` - Gets columns before range
- `getRowsAbove(count)` - Gets rows above range
- `getRowsBelow(count)` - Gets rows below range

### Range Dependencies

- `getPrecedents()` - Gets precedent cells
- `getDirectPrecedents()` - Gets direct precedent cells
- `getDependents()` - Gets dependent cells
- `getDirectDependents()` - Gets direct dependent cells

### Range Grouping & Sorting

- `group(groupOption)` - Groups rows/columns
- `ungroup(groupOption)` - Ungroups rows/columns
- `hideGroupDetails(groupOption)` - Hides group details
- `showGroupDetails(groupOption)` - Shows group details
- `getSort()` - Gets RangeSort object

### Range Data Validation & Conditional Formatting

- `getDataValidation()` - Gets data validation object
- `getConditionalFormats()` - Gets conditional formats
- `getConditionalFormat(id)` - Gets conditional format by ID

### Range Controls & Pivot Tables

- `getControl()` - Gets cell control
- `setControl(control)` - Sets cell control
- `getPivotTables(fullyContained?)` - Gets pivot tables in range

### Range Views

- `getVisibleView()` - Gets visible portion of range

## Table Object

### Table Properties

- `getName()` - Gets table name
- `setName(name)` - Sets table name
- `getId()` - Gets unique table ID
- `getLegacyId()` - Gets numeric ID
- `getRowCount()` - Gets number of rows

### Table Structure

- `getRange()` - Gets entire table range
- `getHeaderRowRange()` - Gets header row range
- `getRangeBetweenHeaderAndTotal()` - Gets data body range
- `getTotalRowRange()` - Gets total row range
- `getColumns()` - Gets all columns
- `getColumn(key)` - Gets column by name/ID
- `getColumnById(key)` - Gets column by ID
- `getColumnByName(key)` - Gets column by name

### Table Data Management

- `addRow(index?, values?)` - Adds single row
- `addRows(index?, values?)` - Adds multiple rows
- `addColumn(index?, values?, name?)` - Adds column
- `deleteRowsAt(index, count)` - Deletes rows at index
- `resize(newRange)` - Resizes table to new range

### Table Operations

- `convertToRange()` - Converts table to normal range
- `delete()` - Deletes table
- `reapplyFilters()` - Reapplies all filters
- `clearFilters()` - Clears all filters

### Table Formatting & Display

- `getShowHeaders()` - Gets header row visibility
- `setShowHeaders(showHeaders)` - Sets header row visibility
- `getShowTotals()` - Gets total row visibility
- `setShowTotals(showTotals)` - Sets total row visibility
- `getShowBandedRows()` - Gets banded rows setting
- `setShowBandedRows(showBandedRows)` - Sets banded rows
- `getShowBandedColumns()` - Gets banded columns setting
- `setShowBandedColumns(showBandedColumns)` - Sets banded columns
- `getShowFilterButton()` - Gets filter button visibility
- `setShowFilterButton(showFilterButton)` - Sets filter button visibility
- `getHighlightFirstColumn()` - Gets first column highlighting
- `setHighlightFirstColumn(highlightFirstColumn)` - Sets first column highlighting
- `getHighlightLastColumn()` - Gets last column highlighting
- `setHighlightLastColumn(highlightLastColumn)` - Sets last column highlighting

### Table Styles

- `getPredefinedTableStyle()` - Gets predefined table style
- `setPredefinedTableStyle(predefinedTableStyle)` - Sets predefined table style

### Table Sorting & Filtering

- `getSort()` - Gets TableSort object
- `getAutoFilter()` - Gets AutoFilter object

## TableColumn Object

### Column Properties

- `getName()` - Gets column name
- `setName(name)` - Sets column name
- `getId()` - Gets unique column ID
- `getIndex()` - Gets zero-based column index

### Column Ranges

- `getRange()` - Gets entire column range
- `getHeaderRowRange()` - Gets header cell range
- `getRangeBetweenHeaderAndTotal()` - Gets data range
- `getTotalRowRange()` - Gets total cell range

### Column Operations

- `delete()` - Deletes column from table
- `getFilter()` - Gets column filter

## Chart Object

### Chart Properties

- `getName()` - Gets chart name
- `setName(name)` - Sets chart name
- `getChartType()` - Gets chart type
- `setChartType(chartType)` - Sets chart type
- `getHeight()` - Gets chart height
- `setHeight(height)` - Sets chart height
- `getWidth()` - Gets chart width
- `setWidth(width)` - Sets chart width
- `getTop()` - Gets top position
- `setTop(top)` - Sets top position
- `getLeft()` - Gets left position
- `setLeft(left)` - Sets left position

### Chart Data

- `getSeries()` - Gets all chart series
- `addChartSeries(name?, index?)` - Adds new chart series
- `setData(sourceData, seriesBy?)` - Sets chart data source
- `getPlotBy()` - Gets how data is plotted
- `setPlotBy(plotBy)` - Sets how data is plotted
- `getPlotVisibleOnly()` - Gets visible cells only setting
- `setPlotVisibleOnly(plotVisibleOnly)` - Sets visible cells only

### Chart Display

- `getTitle()` - Gets chart title object
- `getAxes()` - Gets chart axes
- `getLegend()` - Gets chart legend
- `getDataLabels()` - Gets data labels
- `getPlotArea()` - Gets plot area
- `getDisplayBlanksAs()` - Gets blank cell display method
- `setDisplayBlanksAs(displayBlanksAs)` - Sets blank cell display method

### Chart Operations

- `activate()` - Activates chart in UI
- `delete()` - Deletes chart
- `setPosition(startCell, endCell?)` - Positions chart relative to cells

### Chart Advanced Properties

- `getCategoryLabelLevel()` - Gets category label level
- `setCategoryLabelLevel(categoryLabelLevel)` - Sets category label level
- `getSeriesNameLevel()` - Gets series name level
- `setSeriesNameLevel(seriesNameLevel)` - Sets series name level
- `getStyle()` - Gets chart style
- `setStyle(style)` - Sets chart style
- `getPivotOptions()` - Gets pivot chart options
- `getShowAllFieldButtons()` - Gets field buttons visibility
- `setShowAllFieldButtons(showAllFieldButtons)` - Sets field buttons visibility
- `getShowDataLabelsOverMaximum()` - Gets data labels over max setting
- `setShowDataLabelsOverMaximum(showDataLabelsOverMaximum)` - Sets data labels over max

## Shape Object

### Shape Properties

- `getName()` - Gets shape name
- `setName(name)` - Sets shape name
- `getType()` - Gets shape type
- `getGeometricShapeType()` - Gets geometric shape type
- `setGeometricShapeType(geometricShapeType)` - Sets geometric shape type
- `getVisible()` - Gets visibility
- `setVisible(visible)` - Sets visibility

### Shape Position & Size

- `getHeight()` - Gets height in points
- `setHeight(height)` - Sets height in points
- `getWidth()` - Gets width in points
- `setWidth(width)` - Sets width in points
- `getTop()` - Gets top position
- `setTop(top)` - Sets top position
- `getLeft()` - Gets left position
- `setLeft(left)` - Sets left position
- `getRotation()` - Gets rotation in degrees
- `setRotation(rotation)` - Sets rotation in degrees

### Shape Operations

- `copyTo(destinationSheet?)` - Copies shape to worksheet
- `delete()` - Removes shape from worksheet
- `incrementLeft(increment)` - Moves shape horizontally
- `incrementTop(increment)` - Moves shape vertically
- `incrementRotation(increment)` - Rotates shape incrementally

### Shape Scaling & Ordering

- `scaleHeight(scaleFactor, scaleType?, scaleFrom?)` - Scales height
- `scaleWidth(scaleFactor, scaleType?, scaleFrom?)` - Scales width
- `getZOrderPosition()` - Gets z-order position
- `setZOrder(position)` - Sets z-order position

### Shape Properties & Attributes

- `getLockAspectRatio()` - Gets aspect ratio lock state
- `setLockAspectRatio(lockAspectRatio)` - Sets aspect ratio lock
- `getPlacement()` - Gets cell attachment method
- `setPlacement(placement)` - Sets cell attachment method
- `getAltTextTitle()` - Gets alt text title
- `setAltTextTitle(altTextTitle)` - Sets alt text title
- `getAltTextDescription()` - Gets alt text description
- `setAltTextDescription(altTextDescription)` - Sets alt text description

### Shape Text & Formatting

- `getTextFrame()` - Gets text frame object
- `getConnectionSiteCount()` - Gets connection sites count
- `getDisplayName()` - Gets display name
- `getImageAsBase64(format)` - Converts shape to base64 image

## Comment Object

### Comment Properties

- `getContent()` - Gets comment content (plain text)
- `setContent(content)` - Sets comment content
- `getContentType()` - Gets content type
- `getId()` - Gets comment identifier
- `getAuthorEmail()` - Gets author email
- `getAuthorName()` - Gets author name
- `getCreationDate()` - Gets creation date
- `getLocation()` - Gets cell location

### Comment State

- `getResolved()` - Gets resolved status
- `setResolved(resolved)` - Sets resolved status
- `getRichContent()` - Gets rich content for parsing
- `getMentions()` - Gets mentioned entities

### Comment Replies

- `getReplies()` - Gets all replies
- `getCommentReply(commentReplyId)` - Gets reply by ID
- `addCommentReply(content, contentType?)` - Adds new reply

### Comment Operations

- `delete()` - Deletes comment and all replies
- `updateMentions(contentWithMentions)` - Updates content with mentions

## CommentReply Object

### Reply Properties

- `getContent()` - Gets reply content
- `setContent(content)` - Sets reply content
- `getContentType()` - Gets content type
- `getId()` - Gets reply identifier
- `getAuthorEmail()` - Gets author email
- `getAuthorName()` - Gets author name
- `getCreationDate()` - Gets creation date
- `getLocation()` - Gets cell location

### Reply State & Relations

- `getResolved()` - Gets resolved status
- `getParentComment()` - Gets parent comment
- `getRichContent()` - Gets rich content
- `getMentions()` - Gets mentioned entities

### Reply Operations

- `delete()` - Deletes reply
- `updateMentions(contentWithMentions)` - Updates content with mentions

## RangeFormat Object

### Font Formatting

- `getFont()` - Gets RangeFont object

### Fill Formatting

- `getFill()` - Gets RangeFill object

### Border Formatting

- `getBorders()` - Gets RangeBorderCollection object

### Alignment & Protection

- `getHorizontalAlignment()` - Gets horizontal alignment
- `setHorizontalAlignment(horizontalAlignment)` - Sets horizontal alignment
- `getVerticalAlignment()` - Gets vertical alignment
- `setVerticalAlignment(verticalAlignment)` - Sets vertical alignment
- `getWrapText()` - Gets text wrap setting
- `setWrapText(wrapText)` - Sets text wrap
- `getIndentLevel()` - Gets indent level
- `setIndentLevel(indentLevel)` - Sets indent level
- `getReadingOrder()` - Gets reading order
- `setReadingOrder(readingOrder)` - Sets reading order
- `getShrinkToFit()` - Gets shrink to fit setting
- `setShrinkToFit(shrinkToFit)` - Sets shrink to fit
- `getTextOrientation()` - Gets text orientation
- `setTextOrientation(textOrientation)` - Sets text orientation

### Format Operations

- `autofitColumns()` - Auto fits column widths
- `autofitRows()` - Auto fits row heights

## Filter Object

### Basic Filtering

- `apply(criteria)` - Applies filter criteria
- `clear()` - Clears filter

### Item-based Filtering

- `applyBottomItemsFilter(count)` - Filters bottom N items
- `applyTopItemsFilter(count)` - Filters top N items
- `applyBottomPercentFilter(percent)` - Filters bottom N percent
- `applyTopPercentFilter(percent)` - Filters top N percent

### Color & Icon Filtering

- `applyCellColorFilter(color)` - Filters by cell color
- `applyFontColorFilter(color)` - Filters by font color
- `applyIconFilter(icon)` - Filters by icon

### Value-based Filtering

- `applyValuesFilter(values)` - Filters by specific values
- `applyCustomFilter(criteria1, criteria2?, oper?)` - Applies custom filter
- `applyDynamicFilter(dynamicCriteria)` - Applies dynamic filter

### Filter Properties

- `getCriteria()` - Gets current filter criteria

## AutoFilter Object

### AutoFilter Operations

- `apply(range, columnIndex?, criteria?)` - Applies AutoFilter
- `clearCriteria()` - Clears all filter criteria
- `getRange()` - Gets filtered range
- `remove()` - Removes AutoFilter

## RangeSort Object

### Sorting Operations

- `apply(fields, matchCase?, hasHeaders?, orientation?, method?)` - Applies sort

## TableSort Object

### Table Sorting

- `apply(fields, matchCase?, method?)` - Applies sort to table
- `clear()` - Clears sort state
- `reapply()` - Reapplies current sort
- `getFields()` - Gets current sort fields
- `getMatchCase()` - Gets case sensitivity setting
- `getMethod()` - Gets sort method

## PivotTable Object

### PivotTable Properties

- `getName()` - Gets pivot table name
- `setName(name)` - Sets pivot table name
- `getId()` - Gets pivot table ID

### PivotTable Structure

- `getLayout()` - Gets pivot table layout
- `getWorksheet()` - Gets containing worksheet

### PivotTable Operations

- `delete()` - Deletes pivot table
- `refresh()` - Refreshes pivot table data
- `refreshOnOpen()` - Gets refresh on open setting
- `setRefreshOnOpen(refreshOnOpen)` - Sets refresh on open

## Slicer Object

### Slicer Properties

- `getName()` - Gets slicer name
- `setName(name)` - Sets slicer name
- `getId()` - Gets slicer ID
- `getCaption()` - Gets slicer caption
- `setCaption(caption)` - Sets slicer caption

### Slicer Position & Size

- `getHeight()` - Gets height
- `setHeight(height)` - Sets height
- `getWidth()` - Gets width
- `setWidth(width)` - Sets width
- `getTop()` - Gets top position
- `setTop(top)` - Sets top position
- `getLeft()` - Gets left position
- `setLeft(left)` - Sets left position

### Slicer Operations

- `delete()` - Deletes slicer
- `clearFilters()` - Clears all slicer filters
- `getSlicerItems()` - Gets slicer items
- `getWorksheet()` - Gets containing worksheet

## NamedItem Object

### Named Item Properties

- `getName()` - Gets name
- `setName(name)` - Sets name
- `getType()` - Gets named item type
- `getValue()` - Gets value
- `setValue(value)` - Sets value
- `getFormula()` - Gets formula
- `setFormula(formula)` - Sets formula

### Named Item Operations

- `delete()` - Deletes named item
- `getRange()` - Gets range (if applicable)

## Application Object

### Application Properties

- `getCalculationMode()` - Gets calculation mode
- `setCalculationMode(calculationMode)` - Sets calculation mode
- `getCalculationState()` - Gets calculation state

### Application Operations

- `calculate(calculationType)` - Performs calculation
- `suspendApiCalculationUntilNextSync()` - Suspends calculations
- `suspendScreenUpdates()` - Suspends screen updates
- `enableScreenUpdates()` - Enables screen updates

## TextRange Object

### Text Content

- `getText()` - Gets plain text content
- `setText(text)` - Sets plain text content
- `getSubstring(start, length?)` - Gets text substring

### Text Formatting

- `getFont()` - Gets ShapeFont object for text formatting

## TextFrame Object

### Text Frame Properties

- `getTextRange()` - Gets TextRange object
- `getAutoSize()` - Gets auto-size setting
- `setAutoSize(autoSize)` - Sets auto-size behavior
- `getHasText()` - Checks if frame has text

### Text Frame Margins

- `getBottomMargin()` - Gets bottom margin
- `setBottomMargin(bottomMargin)` - Sets bottom margin
- `getTopMargin()` - Gets top margin
- `setTopMargin(topMargin)` - Sets top margin
- `getLeftMargin()` - Gets left margin
- `setLeftMargin(leftMargin)` - Sets left margin
- `getRightMargin()` - Gets right margin
- `setRightMargin(rightMargin)` - Sets right margin

### Text Frame Alignment

- `getHorizontalAlignment()` - Gets horizontal alignment
- `setHorizontalAlignment(horizontalAlignment)` - Sets horizontal alignment
- `getVerticalAlignment()` - Gets vertical alignment
- `setVerticalAlignment(verticalAlignment)` - Sets vertical alignment
- `getHorizontalOverflow()` - Gets horizontal overflow behavior
- `setHorizontalOverflow(horizontalOverflow)` - Sets horizontal overflow
- `getVerticalOverflow()` - Gets vertical overflow behavior
- `setVerticalOverflow(verticalOverflow)` - Sets vertical overflow

### Text Frame Properties

- `getReadingOrder()` - Gets reading order
- `setReadingOrder(readingOrder)` - Sets reading order
- `getOrientation()` - Gets text orientation
- `setOrientation(orientation)` - Sets text orientation

### Text Frame Operations

- `deleteText()` - Deletes all text from frame

## ChartSeries Object

### Series Properties

- `getName()` - Gets series name
- `setName(name)` - Sets series name
- `getChartType()` - Gets series chart type
- `setChartType(chartType)` - Sets series chart type

### Series Data

- `getValues()` - Gets series values
- `setValues(values)` - Sets series values
- `getXAxisValues()` - Gets X-axis values
- `setXAxisValues(xAxisValues)` - Sets X-axis values
- `getBubbleSizes()` - Gets bubble sizes (bubble charts)
- `setBubbleSizes(bubbleSizes)` - Sets bubble sizes

### Series Formatting

- `getFormat()` - Gets series format
- `getDataLabels()` - Gets data labels
- `getMarkerStyle()` - Gets marker style
- `setMarkerStyle(markerStyle)` - Sets marker style
- `getMarkerSize()` - Gets marker size
- `setMarkerSize(markerSize)` - Sets marker size

### Series Display Properties

- `getSmooth()` - Gets smooth line setting
- `setSmooth(smooth)` - Sets smooth line
- `getHasDataLabels()` - Gets data labels visibility
- `setHasDataLabels(hasDataLabels)` - Sets data labels visibility
- `getOverlap()` - Gets column overlap
- `setOverlap(overlap)` - Sets column overlap
- `getGapWidth()` - Gets gap width
- `setGapWidth(gapWidth)` - Sets gap width

### Series Operations

- `delete()` - Deletes series

## PageLayout Object

### Page Setup

- `getOrientation()` - Gets page orientation
- `setOrientation(orientation)` - Sets page orientation
- `getPaperSize()` - Gets paper size
- `setPaperSize(paperSize)` - Sets paper size

### Page Margins

- `getTopMargin()` - Gets top margin
- `setTopMargin(topMargin)` - Sets top margin
- `getBottomMargin()` - Gets bottom margin
- `setBottomMargin(bottomMargin)` - Sets bottom margin
- `getLeftMargin()` - Gets left margin
- `setLeftMargin(leftMargin)` - Sets left margin
- `getRightMargin()` - Gets right margin
- `setRightMargin(rightMargin)` - Sets right margin
- `getHeaderMargin()` - Gets header margin
- `setHeaderMargin(headerMargin)` - Sets header margin
- `getFooterMargin()` - Gets footer margin
- `setFooterMargin(footerMargin)` - Sets footer margin

### Page Headers & Footers

- `getPrintTitleColumns()` - Gets print title columns
- `setPrintTitleColumns(printTitleColumns)` - Sets print title columns
- `getPrintTitleRows()` - Gets print title rows
- `setPrintTitleRows(printTitleRows)` - Sets print title rows

### Page Display

- `getCenterHorizontally()` - Gets center horizontally setting
- `setCenterHorizontally(centerHorizontally)` - Sets center horizontally
- `getCenterVertically()` - Gets center vertically setting
- `setCenterVertically(centerVertically)` - Sets center vertically
- `getDraftMode()` - Gets draft mode setting
- `setDraftMode(draftMode)` - Sets draft mode
- `getFirstPageNumber()` - Gets first page number
- `setFirstPageNumber(firstPageNumber)` - Sets first page number

### Page Printing

- `getPrintOrder()` - Gets print order
- `setPrintOrder(printOrder)` - Sets print order
- `getPrintComments()` - Gets print comments setting
- `setPrintComments(printComments)` - Sets print comments
- `getPrintErrors()` - Gets print errors setting
- `setPrintErrors(printErrors)` - Sets print errors
- `getPrintGridlines()` - Gets print gridlines setting
- `setPrintGridlines(printGridlines)` - Sets print gridlines
- `getPrintHeadings()` - Gets print headings setting
- `setPrintHeadings(printHeadings)` - Sets print headings

### Page Zoom & Quality

- `getZoom()` - Gets zoom percentage
- `setZoom(zoom)` - Sets zoom percentage
- `getFitToPagesWide()` - Gets fit to pages wide
- `setFitToPagesWide(fitToPagesWide)` - Sets fit to pages wide
- `getFitToPagesTall()` - Gets fit to pages tall
- `setFitToPagesTall(fitToPagesTall)` - Sets fit to pages tall
- `getPrintQuality()` - Gets print quality
- `setPrintQuality(printQuality)` - Sets print quality

## WorksheetProtection Object

### Protection State

- `getProtected()` - Gets protection status
- `protect(options?, password?)` - Enables protection
- `unprotect(password?)` - Disables protection

### Protection Options

- `getAllowAutoFilter()` - Gets AutoFilter permission
- `getAllowDeleteColumns()` - Gets delete columns permission
- `getAllowDeleteRows()` - Gets delete rows permission
- `getAllowFormatCells()` - Gets format cells permission
- `getAllowFormatColumns()` - Gets format columns permission
- `getAllowFormatRows()` - Gets format rows permission
- `getAllowInsertColumns()` - Gets insert columns permission
- `getAllowInsertRows()` - Gets insert rows permission
- `getAllowInsertHyperlinks()` - Gets insert hyperlinks permission
- `getAllowPivotTables()` - Gets pivot tables permission
- `getAllowSort()` - Gets sort permission

### Protection Ranges

- `getAllowEditRanges()` - Gets all allowed edit ranges
- `getAllowEditRange(key)` - Gets allowed edit range by key
- `addAllowEditRange(address, password?)` - Adds allowed edit range

## DataValidation Object

### Validation Rules

- `getRule()` - Gets validation rule
- `setRule(rule)` - Sets validation rule
- `clear()` - Clears validation

### Validation Properties

- `getErrorAlert()` - Gets error alert settings
- `setErrorAlert(errorAlert)` - Sets error alert
- `getInputMessage()` - Gets input message
- `setInputMessage(inputMessage)` - Sets input message
- `getPromptTitle()` - Gets prompt title
- `setPromptTitle(promptTitle)` - Sets prompt title
- `getPromptMessage()` - Gets prompt message
- `setPromptMessage(promptMessage)` - Sets prompt message

### Validation State

- `getValid()` - Checks if range values are valid
- `getIgnoreBlanks()` - Gets ignore blanks setting
- `setIgnoreBlanks(ignoreBlanks)` - Sets ignore blanks

## ConditionalFormat Object

### Conditional Format Types

- `getCellValue()` - Gets cell value conditional format
- `getColorScale()` - Gets color scale conditional format
- `getDataBar()` - Gets data bar conditional format
- `getIconSet()` - Gets icon set conditional format
- `getPreset()` - Gets preset conditional format
- `getTextComparison()` - Gets text comparison conditional format
- `getTopBottom()` - Gets top/bottom conditional format

### Conditional Format Properties

- `getId()` - Gets conditional format ID
- `getPriority()` - Gets priority
- `setPriority(priority)` - Sets priority
- `getStopIfTrue()` - Gets stop if true setting
- `setStopIfTrue(stopIfTrue)` - Sets stop if true
- `getType()` - Gets conditional format type

### Conditional Format Operations

- `delete()` - Deletes conditional format

## RangeBorder Object

### Border Properties

- `getColor()` - Gets border color
- `setColor(color)` - Sets border color
- `getStyle()` - Gets border style
- `setStyle(style)` - Sets border style
- `getTintAndShade()` - Gets tint and shade
- `setTintAndShade(tintAndShade)` - Sets tint and shade
- `getWeight()` - Gets border weight
- `setWeight(weight)` - Sets border weight

## RangeFill Object

### Fill Properties

- `getColor()` - Gets fill color
- `setColor(color)` - Sets fill color
- `getPattern()` - Gets fill pattern
- `setPattern(pattern)` - Sets fill pattern
- `getPatternColor()` - Gets pattern color
- `setPatternColor(patternColor)` - Sets pattern color
- `getPatternTintAndShade()` - Gets pattern tint and shade
- `setPatternTintAndShade(patternTintAndShade)` - Sets pattern tint and shade
- `getTintAndShade()` - Gets tint and shade
- `setTintAndShade(tintAndShade)` - Sets tint and shade

### Fill Operations

- `clear()` - Clears fill formatting

## RangeFont Object

### Font Properties

- `getName()` - Gets font name
- `setName(name)` - Sets font name
- `getSize()` - Gets font size
- `setSize(size)` - Sets font size
- `getColor()` - Gets font color
- `setColor(color)` - Sets font color
- `getBold()` - Gets bold setting
- `setBold(bold)` - Sets bold
- `getItalic()` - Gets italic setting
- `setItalic(italic)` - Sets italic
- `getUnderline()` - Gets underline setting
- `setUnderline(underline)` - Sets underline
- `getStrikethrough()` - Gets strikethrough setting
- `setStrikethrough(strikethrough)` - Sets strikethrough
- `getSubscript()` - Gets subscript setting
- `setSubscript(subscript)` - Sets subscript
- `getSuperscript()` - Gets superscript setting
- `setSuperscript(superscript)` - Sets superscript
- `getTintAndShade()` - Gets tint and shade
- `setTintAndShade(tintAndShade)` - Sets tint and shade

## ChartTitle Object

### Title Properties

- `getText()` - Gets title text
- `setText(text)` - Sets title text
- `getVisible()` - Gets title visibility
- `setVisible(visible)` - Sets title visibility

### Title Formatting

- `getFormat()` - Gets title format
- `getHorizontalAlignment()` - Gets horizontal alignment
- `setHorizontalAlignment(horizontalAlignment)` - Sets horizontal alignment
- `getVerticalAlignment()` - Gets vertical alignment
- `setVerticalAlignment(verticalAlignment)` - Sets vertical alignment

### Title Position

- `getTop()` - Gets top position
- `setTop(top)` - Sets top position
- `getLeft()` - Gets left position
- `setLeft(left)` - Sets left position
- `getHeight()` - Gets height
- `getWidth()` - Gets width

### Title Operations

- `setFormula(formula)` - Sets title formula

## ChartLegend Object

### Legend Properties

- `getVisible()` - Gets legend visibility
- `setVisible(visible)` - Sets legend visibility
- `getPosition()` - Gets legend position
- `setPosition(position)` - Sets legend position

### Legend Formatting

- `getFormat()` - Gets legend format
- `getHeight()` - Gets height
- `setHeight(height)` - Sets height
- `getWidth()` - Gets width
- `setWidth(width)` - Sets width
- `getTop()` - Gets top position
- `setTop(top)` - Sets top position
- `getLeft()` - Gets left position
- `setLeft(left)` - Sets left position

### Legend Display

- `getOverlay()` - Gets overlay setting
- `setOverlay(overlay)` - Sets overlay
- `getShowShadow()` - Gets shadow visibility
- `setShowShadow(showShadow)` - Sets shadow visibility

## ChartDataLabels Object

### Data Labels Properties

- `getVisible()` - Gets data labels visibility
- `setVisible(visible)` - Sets data labels visibility
- `getPosition()` - Gets data labels position
- `setPosition(position)` - Sets data labels position
- `getShowSeriesName()` - Gets series name display
- `setShowSeriesName(showSeriesName)` - Sets series name display
- `getShowCategoryName()` - Gets category name display
- `setShowCategoryName(showCategoryName)` - Sets category name display
- `getShowValue()` - Gets value display
- `setShowValue(showValue)` - Sets value display
- `getShowPercentage()` - Gets percentage display
- `setShowPercentage(showPercentage)` - Sets percentage display

### Data Labels Formatting

- `getFormat()` - Gets data labels format
- `getNumberFormat()` - Gets number format
- `setNumberFormat(numberFormat)` - Sets number format
- `getSeparator()` - Gets label separator
- `setSeparator(separator)` - Sets label separator
- `getShowBubbleSize()` - Gets bubble size display
- `setShowBubbleSize(showBubbleSize)` - Sets bubble size display
- `getShowLeaderLines()` - Gets leader lines display
- `setShowLeaderLines(showLeaderLines)` - Sets leader lines display

## WorkbookProtection Object

### Workbook Protection State

- `getProtected()` - Gets protection status
- `protect(password?)` - Enables workbook protection
- `unprotect(password?)` - Disables workbook protection

## CustomXmlPart Object

### XML Properties

- `getId()` - Gets XML part ID
- `getNamespaceUri()` - Gets namespace URI

### XML Operations

- `delete()` - Deletes XML part
- `getXml()` - Gets XML content
- `setXml(xml)` - Sets XML content

## Binding Object

### Binding Properties

- `getId()` - Gets binding ID
- `getType()` - Gets binding type

### Binding Operations

- `delete()` - Deletes binding
- `getRange()` - Gets bound range
- `getText()` - Gets bound text
- `getTable()` - Gets bound table
- `getMatrix()` - Gets bound matrix

## AllowEditRange Object

### Edit Range Properties

- `getAddress()` - Gets range address
- `setAddress(address)` - Sets range address
- `getTitle()` - Gets range title
- `setTitle(title)` - Sets range title
- `getIsPasswordProtected()` - Gets password protection status

### Edit Range Protection

- `setPassword(password?)` - Sets password
- `pauseProtection(password?)` - Pauses protection temporarily

### Edit Range Operations

- `delete()` - Deletes allowed edit range

## RangeAreas Object

### Areas Properties

- `getAddress()` - Gets areas address
- `getAddressLocal()` - Gets localized address
- `getCellCount()` - Gets total cell count
- `getAreaCount()` - Gets number of areas

### Areas Formatting

- `getFormat()` - Gets format for all areas
- `getConditionalFormats()` - Gets conditional formats

### Areas Operations

- `clear(applyTo?)` - Clears content/formatting
- `copyFrom(sourceRange, copyType?, skipBlanks?, transpose?)` - Copies from source

### Areas Data

- `getValues()` - Gets values from all areas
- `setValues(values)` - Sets values for all areas
- `getFormulas()` - Gets formulas from all areas
- `setFormulas(formulas)` - Sets formulas for all areas

### Areas Analysis

- `calculate()` - Calculates all areas
- `getSpecialCells(cellType, cellValueType?)` - Gets special cells across areas

## WorksheetCustomProperty Object

### Custom Property Access

- `getKey()` - Gets property key
- `getValue()` - Gets property value
- `setValue(value)` - Sets property value

### Custom Property Operations

- `delete()` - Deletes custom property

## NamedSheetView Object

### Sheet View Properties

- `getName()` - Gets view name
- `activate()` - Activates the view
- `duplicate(name?)` - Duplicates the view

### Sheet View Operations

- `delete()` - Deletes the view

## Query Object (Power Query)

### Query Properties

- `getName()` - Gets query name
- `getRefreshDate()` - Gets last refresh date
- `getRowsLoadedCount()` - Gets loaded rows count

### Query Operations

- `refresh()` - Refreshes query data

## WorkbookProperties Object

### Document Properties

- `getTitle()` - Gets document title
- `setTitle(title)` - Sets document title
- `getAuthor()` - Gets document author
- `setAuthor(author)` - Sets document author
- `getSubject()` - Gets document subject
- `setSubject(subject)` - Sets document subject
- `getKeywords()` - Gets document keywords
- `setKeywords(keywords)` - Sets document keywords
- `getComments()` - Gets document comments
- `setComments(comments)` - Sets document comments
- `getCategory()` - Gets document category
- `setCategory(category)` - Sets document category
- `getManager()` - Gets manager
- `setManager(manager)` - Sets manager
- `getCompany()` - Gets company
- `setCompany(company)` - Sets company
- `getLastAuthor()` - Gets last author
- `getRevisionNumber()` - Gets revision number
- `setRevisionNumber(revisionNumber)` - Sets revision number

### Custom Properties

- `getCustomProperty(key)` - Gets custom property by key
- `addCustomProperty(key, value)` - Adds custom property

## SlicerItem Object

### Slicer Item Properties

- `getName()` - Gets item name
- `getIsSelected()` - Gets selection state
- `setIsSelected(isSelected)` - Sets selection state
- `getHasData()` - Gets data availability

### Slicer Item Operations

- `select()` - Selects the item

## PageBreak Object

### Page Break Properties

- `getColumnIndex()` - Gets column index (vertical breaks)
- `getRowIndex()` - Gets row index (horizontal breaks)

### Page Break Operations

- `delete()` - Deletes page break

## Special Methods for Collections

### Search & Find Operations

- `findAll(text, criteria)` - Finds all text occurrences in worksheet
- `find(text, criteria)` - Finds first text occurrence in range
- `getSpecialCells(cellType, cellValueType?)` - Gets cells meeting special criteria

### Bulk Operations

- `autoFill(destinationRange, autoFillType?)` - Auto fills data patterns
- `flashFill()` - Performs Flash Fill operation
- `removeDuplicates(columns, includesHeader)` - Removes duplicate rows
- `group(groupOption)` - Groups rows or columns
- `ungroup(groupOption)` - Ungroups rows or columns

### Advanced Range Operations

- `getIntersection(anotherRange)` - Gets intersection of ranges
- `getColumnsAfter(count)` - Gets columns after current range
- `getColumnsBefore(count)` - Gets columns before current range
- `getRowsAbove(count)` - Gets rows above current range
- `getRowsBelow(count)` - Gets rows below current range

## Notes

- All functions return either void, specific object types, or primitive types (string, number, boolean)
- Optional parameters are marked with `?`
- Many get/set pairs exist for properties
- Most objects support `delete()` method for removal
- Index parameters are typically zero-based
- Address parameters accept A1-style notation (e.g., "A1:C3")
- Range parameters can often accept either Range objects or string addresses
- Collection methods typically return arrays of objects
- Error handling may throw specific Excel error types for invalid operations

## PivotTable Object

### PivotTable Structure Management

- `addRowHierarchy(pivotHierarchy)` - Adds hierarchy to row axis
- `addColumnHierarchy(pivotHierarchy)` - Adds hierarchy to column axis
- `addDataHierarchy(pivotHierarchy)` - Adds hierarchy to data axis
- `addFilterHierarchy(pivotHierarchy)` - Adds hierarchy to filter axis
- `removeRowHierarchy(rowColumnPivotHierarchy)` - Removes row hierarchy
- `removeColumnHierarchy(rowColumnPivotHierarchy)` - Removes column hierarchy
- `removeDataHierarchy(dataPivotHierarchy)` - Removes data hierarchy
- `removeFilterHierarchy(filterPivotHierarchy)` - Removes filter hierarchy

### PivotTable Properties

- `getName()` - Gets pivot table name
- `setName(name)` - Sets pivot table name
- `getId()` - Gets pivot table ID
- `getDataSourceString()` - Gets data source string representation
- `getDataSourceType()` - Gets data source type
- `getUseCustomSortLists()` - Gets custom sort lists usage
- `setUseCustomSortLists(useCustomSortLists)` - Sets custom sort lists usage

### PivotTable Hierarchies Access

- `getRowHierarchies()` - Gets all row hierarchies
- `getRowHierarchy(name)` - Gets row hierarchy by name
- `getColumnHierarchies()` - Gets all column hierarchies
- `getColumnHierarchy(name)` - Gets column hierarchy by name
- `getDataHierarchies()` - Gets all data hierarchies
- `getDataHierarchy(name)` - Gets data hierarchy by name
- `getFilterHierarchies()` - Gets all filter hierarchies
- `getFilterHierarchy(name)` - Gets filter hierarchy by name
- `getHierarchies()` - Gets all hierarchies
- `getHierarchy(name)` - Gets hierarchy by name

### PivotTable Settings

- `getAllowMultipleFiltersPerField()` - Gets multiple filters setting
- `setAllowMultipleFiltersPerField(allowMultipleFiltersPerField)` - Sets multiple filters
- `getEnableDataValueEditing()` - Gets data value editing setting
- `setEnableDataValueEditing(enableDataValueEditing)` - Sets data value editing

### PivotTable Operations

- `refresh()` - Refreshes pivot table data
- `delete()` - Deletes pivot table
- `getLayout()` - Gets PivotLayout object
- `getWorksheet()` - Gets containing worksheet

## PivotLayout Object

### PivotLayout Range Access

- `getRange()` - Gets pivot table range (excluding filter area)
- `getBodyAndTotalRange()` - Gets data values range
- `getColumnLabelRange()` - Gets column labels range
- `getRowLabelRange()` - Gets row labels range
- `getFilterAxisRange()` - Gets filter area range
- `getDataHierarchy(cell)` - Gets data hierarchy for specific cell

### PivotLayout Display Settings

- `getLayoutType()` - Gets layout type
- `setLayoutType(layoutType)` - Sets layout type
- `getShowFieldHeaders()` - Gets field headers visibility
- `setShowFieldHeaders(showFieldHeaders)` - Sets field headers visibility
- `getShowColumnGrandTotals()` - Gets column grand totals visibility
- `setShowColumnGrandTotals(showColumnGrandTotals)` - Sets column grand totals
- `getShowRowGrandTotals()` - Gets row grand totals visibility
- `setShowRowGrandTotals(showRowGrandTotals)` - Sets row grand totals
- `getSubtotalLocation()` - Gets subtotal location
- `setSubtotalLocation(subtotalLocation)` - Sets subtotal location

### PivotLayout Formatting

- `getAutoFormat()` - Gets auto format setting
- `setAutoFormat(autoFormat)` - Sets auto format
- `getPreserveFormatting()` - Gets preserve formatting setting
- `setPreserveFormatting(preserveFormatting)` - Sets preserve formatting
- `getFillEmptyCells()` - Gets fill empty cells setting
- `setFillEmptyCells(fillEmptyCells)` - Sets fill empty cells
- `getEmptyCellText()` - Gets empty cell text
- `setEmptyCellText(emptyCellText)` - Sets empty cell text
- `getEnableFieldList()` - Gets field list visibility
- `setEnableFieldList(enableFieldList)` - Sets field list visibility

### PivotLayout Advanced Operations

- `displayBlankLineAfterEachItem(display)` - Sets blank line display
- `repeatAllItemLabels(repeatLabels)` - Sets label repetition
- `setAutoSortOnCell(cell, sortBy)` - Sets auto sort based on cell

### PivotLayout Accessibility

- `getAltTextTitle()` - Gets alt text title
- `setAltTextTitle(altTextTitle)` - Sets alt text title
- `getAltTextDescription()` - Gets alt text description
- `setAltTextDescription(altTextDescription)` - Sets alt text description

## PivotField Object

### PivotField Properties

- `getName()` - Gets field name
- `setName(name)` - Sets field name
- `getId()` - Gets field ID
- `getShowAllItems()` - Gets show all items setting
- `setShowAllItems(showAllItems)` - Sets show all items
- `getSubtotals()` - Gets subtotals configuration
- `setSubtotals(subtotals)` - Sets subtotals configuration

### PivotField Items

- `getItems()` - Gets all pivot items
- `getPivotItem(name)` - Gets pivot item by name

### PivotField Filtering

- `applyFilter(filter)` - Applies filter to field
- `clearAllFilters()` - Clears all filters
- `clearFilter(filterType)` - Clears specific filter type
- `getFilters()` - Gets all applied filters
- `isFiltered(filterType)` - Checks if filter type is applied

### PivotField Sorting

- `sortByLabels(sortBy)` - Sorts by labels
- `sortByValues(sortBy, valuesHierarchy, pivotItemScope?)` - Sorts by values

## PivotHierarchy Object

### PivotHierarchy Properties

- `getName()` - Gets hierarchy name
- `setName(name)` - Sets hierarchy name
- `getId()` - Gets hierarchy ID

### PivotHierarchy Fields

- `getFields()` - Gets associated pivot fields
- `getPivotField(name)` - Gets pivot field by name

## FilterPivotHierarchy Object

### Filter Hierarchy Properties

- `getName()` - Gets hierarchy name
- `setName(name)` - Sets hierarchy name
- `getId()` - Gets hierarchy ID
- `getPosition()` - Gets hierarchy position
- `setPosition(position)` - Sets hierarchy position
- `getEnableMultipleFilterItems()` - Gets multiple filter items setting
- `setEnableMultipleFilterItems(enableMultipleFilterItems)` - Sets multiple filter items

### Filter Hierarchy Fields

- `getFields()` - Gets associated pivot fields
- `getPivotField(name)` - Gets pivot field by name

### Filter Hierarchy Operations

- `setToDefault()` - Resets hierarchy to default values

## Conditional Format Objects

### ConditionalFormat Base Object

- `getId()` - Gets conditional format ID
- `getType()` - Gets conditional format type
- `getPriority()` - Gets priority/index
- `setPriority(priority)` - Sets priority
- `getStopIfTrue()` - Gets stop if true setting
- `setStopIfTrue(stopIfTrue)` - Sets stop if true
- `getRange()` - Gets applied range (single range)
- `getRanges()` - Gets applied ranges (multiple ranges)
- `delete()` - Deletes conditional format

### ConditionalFormat Type Converters

- `changeRuleToCellValue(properties)` - Changes to cell value type
- `changeRuleToColorScale()` - Changes to color scale type
- `changeRuleToDataBar()` - Changes to data bar type
- `changeRuleToIconSet()` - Changes to icon set type
- `changeRuleToPresetCriteria(properties)` - Changes to preset criteria type
- `changeRuleToTopBottom(properties)` - Changes to top/bottom type
- `changeRuleToContainsText(properties)` - Changes to text comparison type
- `changeRuleToCustom(formula)` - Changes to custom type

### ConditionalFormat Type Getters

- `getCellValue()` - Gets cell value conditional format
- `getColorScale()` - Gets color scale conditional format
- `getDataBar()` - Gets data bar conditional format
- `getIconSet()` - Gets icon set conditional format
- `getPreset()` - Gets preset criteria conditional format
- `getTopBottom()` - Gets top/bottom conditional format
- `getTextComparison()` - Gets text comparison conditional format
- `getCustom()` - Gets custom conditional format

### CellValueConditionalFormat Object

- `getFormat()` - Gets format object
- `getRule()` - Gets cell value rule
- `setRule(rule)` - Sets cell value rule

### ColorScaleConditionalFormat Object

- `getCriteria()` - Gets color scale criteria
- `setCriteria(criteria)` - Sets color scale criteria
- `getThreeColorScale()` - Gets three color scale setting
- `setThreeColorScale(threeColorScale)` - Sets three color scale

### DataBarConditionalFormat Object

- `getAxisColor()` - Gets axis color
- `setAxisColor(axisColor)` - Sets axis color
- `getAxisFormat()` - Gets axis format
- `setAxisFormat(axisFormat)` - Sets axis format
- `getBarDirection()` - Gets bar direction
- `setBarDirection(barDirection)` - Sets bar direction
- `getLowerBoundRule()` - Gets lower bound rule
- `setLowerBoundRule(lowerBoundRule)` - Sets lower bound rule
- `getUpperBoundRule()` - Gets upper bound rule
- `setUpperBoundRule(upperBoundRule)` - Sets upper bound rule
- `getNegativeFormat()` - Gets negative value format
- `getPositiveFormat()` - Gets positive value format
- `getShowDataBarOnly()` - Gets show data bar only setting
- `setShowDataBarOnly(showDataBarOnly)` - Sets show data bar only

### IconSetConditionalFormat Object

- `getCriteria()` - Gets icon criteria array
- `setCriteria(criteria)` - Sets icon criteria array
- `getReverseIconOrder()` - Gets reverse icon order setting
- `setReverseIconOrder(reverseIconOrder)` - Sets reverse icon order
- `getShowIconOnly()` - Gets show icon only setting
- `setShowIconOnly(showIconOnly)` - Sets show icon only
- `getStyle()` - Gets icon set style
- `setStyle(style)` - Sets icon set style

### TopBottomConditionalFormat Object

- `getFormat()` - Gets format object
- `getRule()` - Gets top/bottom rule
- `setRule(rule)` - Sets top/bottom rule

### PresetCriteriaConditionalFormat Object

- `getFormat()` - Gets format object
- `getRule()` - Gets preset criteria rule
- `setRule(rule)` - Sets preset criteria rule

### TextConditionalFormat Object

- `getFormat()` - Gets format object
- `getRule()` - Gets text rule
- `setRule(rule)` - Sets text rule

### CustomConditionalFormat Object

- `getFormat()` - Gets format object
- `getRule()` - Gets custom rule

## ConditionalFormatRule Objects

### ConditionalFormatRule Base

- `getFormula()` - Gets rule formula
- `setFormula(formula)` - Sets rule formula

## Additional Specialized Objects

### RangeHyperlink Object

- `getAddress()` - Gets hyperlink address
- `setAddress(address)` - Sets hyperlink address
- `getDocumentReference()` - Gets document reference
- `setDocumentReference(documentReference)` - Sets document reference
- `getScreenTip()` - Gets screen tip text
- `setScreenTip(screenTip)` - Sets screen tip text
- `getTextToDisplay()` - Gets display text
- `setTextToDisplay(textToDisplay)` - Sets display text

### WorksheetFreezePanes Object

- `freezeAt(frozenRange)` - Freezes panes at range
- `freezeColumns(count)` - Freezes first N columns
- `freezeRows(count)` - Freezes first N rows
- `getLocation()` - Gets freeze location
- `unfreeze()` - Unfreezes all panes

### CellControl Objects

- `getType()` - Gets control type (for all control types)

### DocumentProperty Object

- `getKey()` - Gets property key
- `getValue()` - Gets property value
- `setValue(value)` - Sets property value
- `getType()` - Gets property type

## Chart Axis Objects

### ChartAxis Object

- `getAlignment()` - Gets axis alignment
- `setAlignment(alignment)` - Sets axis alignment
- `getBaseTimeUnit()` - Gets base time unit
- `setBaseTimeUnit(baseTimeUnit)` - Sets base time unit
- `getCategoryType()` - Gets category type
- `setCategoryType(categoryType)` - Sets category type
- `getDisplayUnit()` - Gets display unit
- `setDisplayUnit(displayUnit)` - Sets display unit
- `getFormat()` - Gets axis format
- `getHeight()` - Gets axis height
- `getIsBetweenCategories()` - Gets between categories setting
- `setIsBetweenCategories(isBetweenCategories)` - Sets between categories
- `getLinkNumberFormat()` - Gets link number format setting
- `setLinkNumberFormat(linkNumberFormat)` - Sets link number format
- `getLogBase()` - Gets logarithmic base
- `setLogBase(logBase)` - Sets logarithmic base
- `getMajorTimeUnit()` - Gets major time unit
- `setMajorTimeUnit(majorTimeUnit)` - Sets major time unit
- `getMaximum()` - Gets maximum value
- `setMaximum(maximum)` - Sets maximum value
- `getMinimum()` - Gets minimum value
- `setMinimum(minimum)` - Sets minimum value
- `getMinorTimeUnit()` - Gets minor time unit
- `setMinorTimeUnit(minorTimeUnit)` - Sets minor time unit
- `getMultiLevel()` - Gets multi-level setting
- `setMultiLevel(multiLevel)` - Sets multi-level
- `getNumberFormat()` - Gets number format
- `setNumberFormat(numberFormat)` - Sets number format
- `getOffset()` - Gets axis offset
- `setOffset(offset)` - Sets axis offset
- `getPosition()` - Gets axis position
- `setPosition(position)` - Sets axis position
- `getReversePlotOrder()` - Gets reverse plot order
- `setReversePlotOrder(reversePlotOrder)` - Sets reverse plot order
- `getScaleType()` - Gets scale type
- `setScaleType(scaleType)` - Sets scale type
- `getShowDisplayUnitLabel()` - Gets display unit label visibility
- `setShowDisplayUnitLabel(showDisplayUnitLabel)` - Sets display unit label
- `getTickLabelPosition()` - Gets tick label position
- `setTickLabelPosition(tickLabelPosition)` - Sets tick label position
- `getTickLabelSpacing()` - Gets tick label spacing
- `setTickLabelSpacing(tickLabelSpacing)` - Sets tick label spacing
- `getTickMarkSpacing()` - Gets tick mark spacing
- `setTickMarkSpacing(tickMarkSpacing)` - Sets tick mark spacing
- `getTitle()` - Gets axis title
- `getTop()` - Gets top position
- `getType()` - Gets axis type
- `getVisible()` - Gets axis visibility
- `setVisible(visible)` - Sets axis visibility
- `getWidth()` - Gets axis width

### ChartAxisTitle Object

- `getFormat()` - Gets title format
- `getText()` - Gets title text
- `setText(text)` - Sets title text
- `getTextOrientation()` - Gets text orientation
- `setTextOrientation(textOrientation)` - Sets text orientation
- `getVisible()` - Gets title visibility
- `setVisible(visible)` - Sets title visibility

## Missing Worksheet Methods

### Worksheet Creation Methods

- `addPivotTable(name, source, destination)` - Creates new pivot table

## Additional Range Methods

### Range Advanced Operations

- `getLinkedDataTypeState()` - Gets linked data type state
- `setLinkedDataTypeState(linkedDataTypeState)` - Sets linked data type state
- `getPredefinedCellStyle()` - Gets predefined cell style
- `setPredefinedCellStyle(predefinedCellStyle)` - Sets predefined cell style
- `setDirty()` - Marks range as dirty for recalculation

### Range Hyperlinks

- `getHyperlink()` - Gets hyperlink object
- `setHyperlink(hyperlink)` - Sets hyperlink

## Range Border Collection

### RangeBorderCollection Object

- `getItem(index)` - Gets border by index
- `getCount()` - Gets border count

## ShapeGroup Object

### Shape Group Properties

- `getId()` - Gets group ID
- `getShape()` - Gets corresponding Shape object

### Shape Group Operations

- `ungroup()` - Ungroups the shapes

This reference now covers ALL manipulation functions available in Office Scripts for Excel automation. Each function provides programmatic access to Excel features without requiring manual user interface interaction.
