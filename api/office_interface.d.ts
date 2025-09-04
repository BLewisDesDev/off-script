// office_script_interface.d

declare namespace ExcelScript {
  // ================================
  // CORE WORKBOOK INTERFACE
  // ================================
  interface Workbook {
    // Worksheet management
    getWorksheets(): Worksheet[];
    getWorksheet(name: string): Worksheet | undefined;
    addWorksheet(name?: string): Worksheet;
    getActiveWorksheet(): Worksheet;
    getFirstWorksheet(visibleOnly?: boolean): Worksheet;
    getLastWorksheet(visibleOnly?: boolean): Worksheet;

    // Cell and range access
    getActiveCell(): Range;
    getSelectedRange(): Range;
    getSelectedRanges(): RangeAreas;

    // Data objects
    getTables(): Table[];
    getTable(name: string): Table | undefined;
    getPivotTables(): PivotTable[];
    getPivotTable(name: string): PivotTable | undefined;
    getCharts(): Chart[];
    getChart(name: string): Chart | undefined;
    getSlicers(): Slicer[];
    getSlicer(key: string): Slicer | undefined;

    // Named items and queries
    getNames(): NamedItem[];
    getNamedItem(name: string): NamedItem | undefined;
    getQueries(): Query[];
    getQuery(key: string): Query | undefined;

    // Comments
    getComments(): Comment[];
    getComment(commentId: string): Comment | undefined;
    getCommentByCell(cellAddress: string): Comment;
    getCommentByReplyId(replyId: string): Comment;

    // Properties and metadata
    getProperties(): WorkbookProperties;
    getName(): string;
    getAutoSave(): boolean;
    getIsDirty(): boolean;
    getPreviouslySaved(): boolean;
    getReadOnly(): boolean;

    // Application and protection
    getApplication(): Application;
    getProtection(): WorkbookProtection;

    // Styles
    getDefaultTableStyle(): TableStyle;
    getDefaultPivotTableStyle(): PivotTableStyle;
    getDefaultSlicerStyle(): SlicerStyle;
    getDefaultTimelineStyle(): TimelineStyle;
    getPredefinedCellStyle(name: string): Style | undefined;
    getPredefinedCellStyles(): Style[];

    // Bindings and XML
    getBinding(id: string): Binding | undefined;
    getBindings(): Binding[];
    getCustomXmlPart(id: string): CustomXmlPart | undefined;
    getCustomXmlParts(): CustomXmlPart[];
    getCustomXmlPartByNamespace(namespaceUri: string): CustomXmlPart[];

    // External links
    getLinkedWorkbooks(): LinkedWorkbook[];
    getLinkedWorkbookByUrl(key: string): LinkedWorkbook | undefined;
    getLinkedWorkbookRefreshMode(): LinkedWorkbookRefreshMode;
    breakAllLinksToLinkedWorkbooks(): void;

    // Calculations
    getCalculationEngineVersion(): number;
    getChartDataPointTrack(): boolean;

    // Operations
    save(): void;
  }

  // ================================
  // WORKSHEET INTERFACE
  // ================================
  interface Worksheet {
    // Basic properties
    getName(): string;
    setName(name: string): void;
    getPosition(): number;
    setPosition(position: number): void;
    getVisibility(): SheetVisibility;
    setVisibility(visibility: SheetVisibility): void;
    getTabColor(): string;
    setTabColor(tabColor: string): void;

    // Display settings
    getShowGridlines(): boolean;
    setShowGridlines(showGridlines: boolean): void;
    getShowHeadings(): boolean;
    setShowHeadings(showHeadings: boolean): void;
    getShowDataTypeIcons(): boolean;
    setShowDataTypeIcons(showDataTypeIcons: boolean): void;
    getStandardHeight(): number;
    getStandardWidth(): number;

    // Range access
    getRange(address: string): Range;
    getRangeByIndexes(
      startRow: number,
      startColumn: number,
      rowCount: number,
      columnCount: number
    ): Range;
    getRanges(address: string): RangeAreas;
    getCell(row: number, column: number): Range;
    getUsedRange(valuesOnly?: boolean): Range | undefined;

    // Navigation
    activate(): void;
    getNext(visibleOnly?: boolean): Worksheet | undefined;
    getPrevious(visibleOnly?: boolean): Worksheet | undefined;

    // Data objects
    getTables(): Table[];
    getTable(name: string): Table | undefined;
    addTable(address: string | Range, hasHeaders: boolean): Table;
    getPivotTables(): PivotTable[];
    getPivotTable(name: string): PivotTable | undefined;
    addPivotTable(
      name: string,
      source: Range | Table,
      destination: Range
    ): PivotTable;
    getCharts(): Chart[];
    getChart(name: string): Chart | undefined;
    addChart(
      type: ChartType,
      sourceData: Range,
      seriesBy?: ChartSeriesBy
    ): Chart;

    // Shapes
    getShapes(): Shape[];
    getShape(key: string): Shape | undefined;
    addGeometricShape(geometricShapeType: GeometricShapeType): Shape;
    addLine(
      startLeft: number,
      startTop: number,
      endLeft: number,
      endTop: number,
      connectorType?: ConnectorType
    ): Shape;
    addTextBox(text?: string): Shape;

    // Comments
    getComments(): Comment[];
    getComment(commentId: string): Comment | undefined;
    addComment(
      cellAddress: Range | string,
      content: CommentRichContent | string,
      contentType?: ContentType
    ): Comment;

    // Protection and filtering
    getProtection(): WorksheetProtection;
    getAutoFilter(): AutoFilter;
    getSlicers(): Slicer[];
    getSlicer(key: string): Slicer | undefined;

    // Page layout
    getPageLayout(): PageLayout;
    addHorizontalPageBreak(pageBreakRange: Range | string): PageBreak;
    addVerticalPageBreak(pageBreakRange: Range | string): PageBreak;

    // Custom properties
    addWorksheetCustomProperty(
      key: string,
      value: string
    ): WorksheetCustomProperty;
    getWorksheetCustomProperties(): WorksheetCustomProperty[];

    // Named sheet views
    getNamedSheetView(key: string): NamedSheetView | undefined;
    getNamedSheetViews(): NamedSheetView[];
    getActiveNamedSheetView(): NamedSheetView;

    // Operations
    calculate(markAllDirty: boolean): void;
    copy(
      positionType?: WorksheetPositionType,
      relativeTo?: Worksheet
    ): Worksheet;
    delete(): void;
    findAll(text: string, criteria: WorksheetSearchCriteria): RangeAreas;
    replaceAll(
      text: string,
      replacement: string,
      criteria: ReplaceCriteria
    ): number;
  }

  // ================================
  // RANGE INTERFACE
  // ================================
  interface Range {
    // Address and location
    getAddress(): string;
    getAddressLocal(): string;
    getCellCount(): number;
    getRowCount(): number;
    getColumnCount(): number;
    getRowIndex(): number;
    getColumnIndex(): number;
    getHeight(): number;
    getWidth(): number;
    getTop(): number;
    getLeft(): number;
    getWorksheet(): Worksheet;

    // Values and formulas
    getValue(): string | number | boolean | Date;
    setValue(value: string | number | boolean | Date): void;
    getValues(): (string | number | boolean | Date)[][];
    setValues(values: (string | number | boolean | Date)[][]): void;
    getText(): string;
    getTexts(): string[][];
    getValueType(): RangeValueType;
    getValueTypes(): RangeValueType[][];

    getFormula(): string;
    setFormula(formula: string): void;
    getFormulas(): string[][];
    setFormulas(formulas: string[][]): void;
    getFormulaLocal(): string;
    setFormulaLocal(formulaLocal: string): void;
    getFormulasLocal(): string[][];
    setFormulasLocal(formulasLocal: string[][]): void;
    getFormulaR1C1(): string;
    setFormulaR1C1(formulaR1C1: string): void;
    getFormulasR1C1(): string[][];
    setFormulasR1C1(formulasR1C1: string[][]): void;

    // Formatting
    getFormat(): RangeFormat;
    getNumberFormat(): string;
    setNumberFormat(format: string): void;
    getNumberFormats(): string[][];
    setNumberFormats(numberFormats: string[][]): void;
    getNumberFormatLocal(): string;
    getNumberFormatsLocal(): string[][];
    getNumberFormatCategory(): string;
    getNumberFormatCategories(): string[][];

    // Visibility and state
    getHidden(): boolean | null;
    getRowHidden(): boolean | null;
    setRowHidden(rowHidden: boolean): void;
    getColumnHidden(): boolean | null;
    setColumnHidden(columnHidden: boolean): void;
    getHasSpill(): boolean | null;
    getIsEntireColumn(): boolean;
    getIsEntireRow(): boolean;

    // Navigation and relationships
    getCell(row: number, column: number): Range;
    getColumn(column: number): Range;
    getRow(row: number): Range;
    getEntireColumn(): Range;
    getEntireRow(): Range;
    getOffsetRange(rowOffset: number, columnOffset: number): Range;
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    getAbsoluteResizedRange(numRows: number, numColumns: number): Range;
    getBoundingRect(anotherRange: Range): Range;
    getRangeEdge(direction: KeyboardDirection, activeCell?: Range): Range;
    getExtendedRange(direction: KeyboardDirection, activeCell?: Range): Range;
    getUsedRange(valuesOnly?: boolean): Range | undefined;
    getIntersection(anotherRange: Range): Range | undefined;
    getColumnsAfter(count: number): Range;
    getColumnsBefore(count: number): Range;
    getRowsAbove(count: number): Range;
    getRowsBelow(count: number): Range;

    // Dependencies
    getPrecedents(): WorkbookRangeAreas;
    getDirectPrecedents(): WorkbookRangeAreas;
    getDependents(): WorkbookRangeAreas;
    getDirectDependents(): WorkbookRangeAreas;

    // Operations
    select(): void;
    clear(applyTo?: ClearApplyTo): void;
    delete(shift: DeleteShiftDirection): void;
    insert(shift: InsertShiftDirection): Range;
    merge(across?: boolean): void;
    unmerge(): void;
    copy(
      destinationRange?: Range,
      copyType?: RangeCopyType,
      skipBlanks?: boolean,
      transpose?: boolean
    ): void;
    copyFrom(
      sourceRange: Range,
      copyType?: RangeCopyType,
      skipBlanks?: boolean,
      transpose?: boolean
    ): void;
    cut(destinationRange?: Range): void;
    moveTo(destinationRange: Range): void;

    // Data manipulation
    removeDuplicates(
      columns: number[],
      includesHeader: boolean
    ): RemoveDuplicatesResult;
    replaceAll(
      text: string,
      replacement: string,
      criteria: ReplaceCriteria
    ): number;
    autoFill(destinationRange: Range, autoFillType?: AutoFillType): void;
    flashFill(): void;

    // Search and analysis
    find(text: string, criteria: SearchCriteria): Range | undefined;
    findAll(text: string, criteria: SearchCriteria): RangeAreas;
    getSpecialCells(
      cellType: SpecialCellType,
      cellValueType?: SpecialCellValueType
    ): RangeAreas;

    // Grouping
    group(groupOption: GroupOption): void;
    ungroup(groupOption: GroupOption): void;
    hideGroupDetails(groupOption: GroupOption): void;
    showGroupDetails(groupOption: GroupOption): void;

    // Sorting and data validation
    getSort(): RangeSort;
    getDataValidation(): DataValidation;

    // Conditional formatting
    getConditionalFormats(): ConditionalFormat[];
    getConditionalFormat(id: string): ConditionalFormat | undefined;
    addConditionalFormat(type: ConditionalFormatType): ConditionalFormat;

    // Controls and pivot tables
    getControl(): CellControl;
    setControl(control: CellControl): void;
    getPivotTables(fullyContained?: boolean): PivotTable[];

    // Views and hyperlinks
    getVisibleView(): RangeView;
    getHyperlink(): RangeHyperlink;
    setHyperlink(hyperlink: RangeHyperlink): void;

    // Advanced properties
    getPredefinedCellStyle(): string;
    setPredefinedCellStyle(predefinedCellStyle: string): void;
    setDirty(): void;
  }

  // ================================
  // TABLE INTERFACE
  // ================================
  interface Table {
    // Properties
    getName(): string;
    setName(name: string): void;
    getId(): string;
    getLegacyId(): number;
    getRowCount(): number;

    // Range access
    getRange(): Range;
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getTotalRowRange(): Range;

    // Column management
    getColumns(): TableColumn[];
    getColumn(key: string | number): TableColumn | undefined;
    getColumnById(key: number): TableColumn | undefined;
    getColumnByName(key: string): TableColumn | undefined;
    addColumn(
      index?: number,
      values?: (string | number | boolean)[],
      name?: string
    ): TableColumn;

    // Row management
    addRow(index?: number, values?: (string | number | boolean)[]): void;
    addRows(index?: number, values?: (string | number | boolean)[][]): void;
    deleteRowsAt(index: number, count: number): void;

    // Display properties
    getShowHeaders(): boolean;
    setShowHeaders(showHeaders: boolean): void;
    getShowTotals(): boolean;
    setShowTotals(showTotals: boolean): void;
    getShowBandedRows(): boolean;
    setShowBandedRows(showBandedRows: boolean): void;
    getShowBandedColumns(): boolean;
    setShowBandedColumns(showBandedColumns: boolean): void;
    getShowFilterButton(): boolean;
    setShowFilterButton(showFilterButton: boolean): void;
    getHighlightFirstColumn(): boolean;
    setHighlightFirstColumn(highlightFirstColumn: boolean): void;
    getHighlightLastColumn(): boolean;
    setHighlightLastColumn(highlightLastColumn: boolean): void;

    // Style and operations
    getPredefinedTableStyle(): string;
    setPredefinedTableStyle(predefinedTableStyle: string): void;
    getSort(): TableSort;
    getAutoFilter(): AutoFilter;
    getWorksheet(): Worksheet;

    // Operations
    convertToRange(): Range;
    delete(): void;
    resize(newRange: Range): void;
    reapplyFilters(): void;
    clearFilters(): void;
  }

  // ================================
  // TABLE COLUMN INTERFACE
  // ================================
  interface TableColumn {
    getName(): string;
    setName(name: string): void;
    getId(): number;
    getIndex(): number;
    getRange(): Range;
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getTotalRowRange(): Range;
    getFilter(): Filter;
    delete(): void;
  }

  // ================================
  // CHART INTERFACE
  // ================================
  interface Chart {
    // Properties
    getName(): string;
    setName(name: string): void;
    getChartType(): ChartType;
    setChartType(chartType: ChartType): void;
    getHeight(): number;
    setHeight(height: number): void;
    getWidth(): number;
    setWidth(width: number): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;

    // Chart elements
    getTitle(): ChartTitle;
    getLegend(): ChartLegend;
    getAxes(): ChartAxes;
    getDataLabels(): ChartDataLabels;
    getPlotArea(): ChartPlotArea;
    getSeries(): ChartSeries[];
    addChartSeries(name?: string, index?: number): ChartSeries;

    // Data and display
    setData(sourceData: Range, seriesBy?: ChartSeriesBy): void;
    getPlotBy(): ChartPlotBy;
    setPlotBy(plotBy: ChartPlotBy): void;
    getPlotVisibleOnly(): boolean;
    setPlotVisibleOnly(plotVisibleOnly: boolean): void;
    getDisplayBlanksAs(): ChartDisplayBlanksAs;
    setDisplayBlanksAs(displayBlanksAs: ChartDisplayBlanksAs): void;

    // Advanced properties
    getCategoryLabelLevel(): number;
    setCategoryLabelLevel(categoryLabelLevel: number): void;
    getSeriesNameLevel(): number;
    setSeriesNameLevel(seriesNameLevel: number): void;
    getStyle(): number;
    setStyle(style: number): void;
    getPivotOptions(): ChartPivotOptions;
    getShowAllFieldButtons(): boolean;
    setShowAllFieldButtons(showAllFieldButtons: boolean): void;
    getShowDataLabelsOverMaximum(): boolean;
    setShowDataLabelsOverMaximum(showDataLabelsOverMaximum: boolean): void;

    // Operations
    activate(): void;
    delete(): void;
    setPosition(startCell: Range | string, endCell?: Range | string): void;
    getWorksheet(): Worksheet;
  }

  // ================================
  // PIVOTTABLE INTERFACE
  // ================================
  interface PivotTable {
    // Properties
    getName(): string;
    setName(name: string): void;
    getId(): string;
    getDataSourceString(): string;
    getDataSourceType(): DataSourceType;
    getUseCustomSortLists(): boolean;
    setUseCustomSortLists(useCustomSortLists: boolean): void;

    // Hierarchy management
    addRowHierarchy(pivotHierarchy: PivotHierarchy): RowColumnPivotHierarchy;
    addColumnHierarchy(pivotHierarchy: PivotHierarchy): RowColumnPivotHierarchy;
    addDataHierarchy(pivotHierarchy: PivotHierarchy): DataPivotHierarchy;
    addFilterHierarchy(pivotHierarchy: PivotHierarchy): FilterPivotHierarchy;
    removeRowHierarchy(rowColumnPivotHierarchy: RowColumnPivotHierarchy): void;
    removeColumnHierarchy(
      rowColumnPivotHierarchy: RowColumnPivotHierarchy
    ): void;
    removeDataHierarchy(dataPivotHierarchy: DataPivotHierarchy): void;
    removeFilterHierarchy(filterPivotHierarchy: FilterPivotHierarchy): void;

    // Hierarchy access
    getRowHierarchies(): RowColumnPivotHierarchy[];
    getRowHierarchy(name: string): RowColumnPivotHierarchy | undefined;
    getColumnHierarchies(): RowColumnPivotHierarchy[];
    getColumnHierarchy(name: string): RowColumnPivotHierarchy | undefined;
    getDataHierarchies(): DataPivotHierarchy[];
    getDataHierarchy(name: string): DataPivotHierarchy | undefined;
    getFilterHierarchies(): FilterPivotHierarchy[];
    getFilterHierarchy(name: string): FilterPivotHierarchy | undefined;
    getHierarchies(): PivotHierarchy[];
    getHierarchy(name: string): PivotHierarchy | undefined;

    // Settings
    getAllowMultipleFiltersPerField(): boolean;
    setAllowMultipleFiltersPerField(
      allowMultipleFiltersPerField: boolean
    ): void;
    getEnableDataValueEditing(): boolean;
    setEnableDataValueEditing(enableDataValueEditing: boolean): void;

    // Operations
    refresh(): void;
    delete(): void;
    getLayout(): PivotLayout;
    getWorksheet(): Worksheet;
  }

  // ================================
  // SHAPE INTERFACE
  // ================================
  interface Shape {
    // Properties
    getName(): string;
    setName(name: string): void;
    getType(): ShapeType;
    getGeometricShapeType(): GeometricShapeType;
    setGeometricShapeType(geometricShapeType: GeometricShapeType): void;
    getVisible(): boolean;
    setVisible(visible: boolean): void;

    // Position and size
    getHeight(): number;
    setHeight(height: number): void;
    getWidth(): number;
    setWidth(width: number): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;
    getRotation(): number;
    setRotation(rotation: number): void;

    // Advanced properties
    getLockAspectRatio(): boolean;
    setLockAspectRatio(lockAspectRatio: boolean): void;
    getPlacement(): Placement;
    setPlacement(placement: Placement): void;
    getAltTextTitle(): string;
    setAltTextTitle(altTextTitle: string): void;
    getAltTextDescription(): string;
    setAltTextDescription(altTextDescription: string): void;

    // Text
    getTextFrame(): TextFrame;

    // Operations
    copyTo(destinationSheet?: Worksheet | string): Shape;
    delete(): void;
    incrementLeft(increment: number): void;
    incrementTop(increment: number): void;
    incrementRotation(increment: number): void;
    scaleHeight(
      scaleFactor: number,
      scaleType?: ShapeScaleType,
      scaleFrom?: ShapeScaleFrom
    ): void;
    scaleWidth(
      scaleFactor: number,
      scaleType?: ShapeScaleType,
      scaleFrom?: ShapeScaleFrom
    ): void;
    getZOrderPosition(): number;
    setZOrder(position: ShapeZOrder): void;
    getConnectionSiteCount(): number;
    getDisplayName(): string;
    getImageAsBase64(format: PictureFormat): string;
  }

  // ================================
  // COMMENT INTERFACE
  // ================================
  interface Comment {
    // Properties
    getContent(): string;
    setContent(content: string): void;
    getContentType(): ContentType;
    getId(): string;
    getAuthorEmail(): string;
    getAuthorName(): string;
    getCreationDate(): Date;
    getLocation(): Range;

    // State
    getResolved(): boolean;
    setResolved(resolved: boolean): void;
    getRichContent(): string;
    getMentions(): CommentMention[];

    // Replies
    getReplies(): CommentReply[];
    getCommentReply(commentReplyId: string): CommentReply | undefined;
    addCommentReply(
      content: CommentRichContent | string,
      contentType?: ContentType
    ): CommentReply;

    // Operations
    delete(): void;
    updateMentions(contentWithMentions: CommentRichContent): void;
  }

  // ================================
  // FORMATTING INTERFACES
  // ================================
  interface RangeFormat {
    getFill(): RangeFill;
    getFont(): RangeFont;
    getBorders(): RangeBorderCollection;
    getHorizontalAlignment(): HorizontalAlignment;
    setHorizontalAlignment(alignment: HorizontalAlignment): void;
    getVerticalAlignment(): VerticalAlignment;
    setVerticalAlignment(alignment: VerticalAlignment): void;
    getWrapText(): boolean;
    setWrapText(wrapText: boolean): void;
    getIndentLevel(): number;
    setIndentLevel(indentLevel: number): void;
    getReadingOrder(): ReadingOrder;
    setReadingOrder(readingOrder: ReadingOrder): void;
    getShrinkToFit(): boolean;
    setShrinkToFit(shrinkToFit: boolean): void;
    getTextOrientation(): number;
    setTextOrientation(textOrientation: number): void;
    autofitColumns(): void;
    autofitRows(): void;
    // ADD THIS MISSING METHOD:
    setRowHeight(rowHeight: number): void;
    getRowHeight(): number;

    // Also add these related methods that are missing:
    setColumnWidth(columnWidth: number): void;
    getColumnWidth(): number;
    setUseStandardHeight(useStandardHeight: boolean): void;
    getUseStandardHeight(): boolean;
    setUseStandardWidth(useStandardWidth: boolean): void;
    getUseStandardWidth(): boolean;

    // Additional missing RangeFormat methods:
    setAutoIndent(autoIndent: boolean): void;
    getAutoIndent(): boolean;
    adjustIndent(amount: number): void;
    setRangeBorderTintAndShade(rangeBorderTintAndShade: number): void;
    getRangeBorderTintAndShade(): number;

    // Missing border access method:
    getRangeBorder(index: BorderIndex): RangeBorder;
  }

  interface RangeFill {
    getColor(): string;
    setColor(color: string): void;
    getPattern(): FillPattern;
    setPattern(pattern: FillPattern): void;
    getPatternColor(): string;
    setPatternColor(patternColor: string): void;
    getPatternTintAndShade(): number;
    setPatternTintAndShade(patternTintAndShade: number): void;
    getTintAndShade(): number;
    setTintAndShade(tintAndShade: number): void;
    clear(): void;
  }

  interface RangeFont {
    getName(): string;
    setName(name: string): void;
    getSize(): number;
    setSize(size: number): void;
    getBold(): boolean;
    setBold(bold: boolean): void;
    getItalic(): boolean;
    setItalic(italic: boolean): void;
    getColor(): string;
    setColor(color: string): void;
    getUnderline(): RangeFontUnderlineStyle;
    setUnderline(underline: RangeFontUnderlineStyle): void;
    getStrikethrough(): boolean;
    setStrikethrough(strikethrough: boolean): void;
    getSubscript(): boolean;
    setSubscript(subscript: boolean): void;
    getSuperscript(): boolean;
    setSuperscript(superscript: boolean): void;
    getTintAndShade(): number;
    setTintAndShade(tintAndShade: number): void;
  }

  interface RangeBorderCollection {
    getTop(): RangeBorder;
    getBottom(): RangeBorder;
    getLeft(): RangeBorder;
    getRight(): RangeBorder;
    getDiagonalDown(): RangeBorder;
    getDiagonalUp(): RangeBorder;
    getHorizontal(): RangeBorder;
    getVertical(): RangeBorder;
    getItem(index: BorderIndex): RangeBorder;
    getCount(): number;
  }

  interface RangeBorder {
    getStyle(): BorderLineStyle;
    setStyle(style: BorderLineStyle): void;
    getColor(): string;
    setColor(color: string): void;
    getTintAndShade(): number;
    setTintAndShade(tintAndShade: number): void;
    getWeight(): BorderWeight;
    setWeight(weight: BorderWeight): void;
  }

  // ================================
  // CONDITIONAL FORMATTING
  // ================================
  interface ConditionalFormat {
    getId(): string;
    getType(): ConditionalFormatType;
    getPriority(): number;
    setPriority(priority: number): void;
    getStopIfTrue(): boolean;
    setStopIfTrue(stopIfTrue: boolean): void;
    getRange(): Range | undefined;
    getRanges(): RangeAreas;
    delete(): void;

    // Type converters
    changeRuleToCellValue(properties: ConditionalCellValueRule): void;
    changeRuleToColorScale(): void;
    changeRuleToDataBar(): void;
    changeRuleToIconSet(): void;
    changeRuleToPresetCriteria(properties: ConditionalPresetCriteriaRule): void;
    changeRuleToTopBottom(properties: ConditionalTopBottomRule): void;
    changeRuleToContainsText(properties: ConditionalTextComparisonRule): void;
    changeRuleToCustom(formula: string): void;

    // Type getters
    getCellValue(): CellValueConditionalFormat | undefined;
    getColorScale(): ColorScaleConditionalFormat | undefined;
    getDataBar(): DataBarConditionalFormat | undefined;
    getIconSet(): IconSetConditionalFormat | undefined;
    getPreset(): PresetCriteriaConditionalFormat | undefined;
    getTopBottom(): TopBottomConditionalFormat | undefined;
    getTextComparison(): TextConditionalFormat | undefined;
    getCustom(): CustomConditionalFormat | undefined;
  }

  // ================================
  // PIVOT LAYOUT INTERFACE
  // ================================
  interface PivotLayout {
    // Range access
    getRange(): Range;
    getBodyAndTotalRange(): Range;
    getColumnLabelRange(): Range;
    getRowLabelRange(): Range;
    getFilterAxisRange(): Range;
    getDataHierarchy(cell: Range): DataPivotHierarchy;

    // Display settings
    getLayoutType(): PivotLayoutType;
    setLayoutType(layoutType: PivotLayoutType): void;
    getShowFieldHeaders(): boolean;
    setShowFieldHeaders(showFieldHeaders: boolean): void;
    getShowColumnGrandTotals(): boolean;
    setShowColumnGrandTotals(showColumnGrandTotals: boolean): void;
    getShowRowGrandTotals(): boolean;
    setShowRowGrandTotals(showRowGrandTotals: boolean): void;
    getSubtotalLocation(): SubtotalLocationType;
    setSubtotalLocation(subtotalLocation: SubtotalLocationType): void;

    // Formatting
    getAutoFormat(): boolean;
    setAutoFormat(autoFormat: boolean): void;
    getPreserveFormatting(): boolean;
    setPreserveFormatting(preserveFormatting: boolean): void;
    getFillEmptyCells(): boolean;
    setFillEmptyCells(fillEmptyCells: boolean): void;
    getEmptyCellText(): string;
    setEmptyCellText(emptyCellText: string): void;
    getEnableFieldList(): boolean;
    setEnableFieldList(enableFieldList: boolean): void;

    // Advanced operations
    displayBlankLineAfterEachItem(display: boolean): void;
    repeatAllItemLabels(repeatLabels: boolean): void;
    setAutoSortOnCell(cell: Range, sortBy: SortBy): void;

    // Accessibility
    getAltTextTitle(): string;
    setAltTextTitle(altTextTitle: string): void;
    getAltTextDescription(): string;
    setAltTextDescription(altTextDescription: string): void;
  }

  // ================================
  // PIVOT FIELD INTERFACE
  // ================================
  interface PivotField {
    getName(): string;
    setName(name: string): void;
    getId(): string;
    getShowAllItems(): boolean;
    setShowAllItems(showAllItems: boolean): void;
    getSubtotals(): Subtotals;
    setSubtotals(subtotals: Subtotals): void;

    // Items
    getItems(): PivotItem[];
    getPivotItem(name: string): PivotItem | undefined;

    // Filtering
    applyFilter(filter: PivotFilters): void;
    clearAllFilters(): void;
    clearFilter(filterType: PivotFilterType): void;
    getFilters(): PivotFilters;
    isFiltered(filterType?: PivotFilterType): boolean;

    // Sorting
    sortByLabels(sortBy: SortBy): void;
    sortByValues(
      sortBy: SortBy,
      valuesHierarchy: DataPivotHierarchy,
      pivotItemScope?: PivotItem[]
    ): void;
  }

  // ================================
  // PIVOT HIERARCHY INTERFACES
  // ================================
  interface PivotHierarchy {
    getName(): string;
    setName(name: string): void;
    getId(): string;
    getFields(): PivotField[];
    getPivotField(name: string): PivotField | undefined;
  }

  interface RowColumnPivotHierarchy extends PivotHierarchy {
    getPosition(): number;
    setPosition(position: number): void;
    getFields(): PivotField[];
  }

  interface DataPivotHierarchy extends PivotHierarchy {
    getPosition(): number;
    setPosition(position: number): void;
    getNumberFormat(): string;
    setNumberFormat(numberFormat: string): void;
    getShowAs(): ShowAsRule;
    setShowAs(showAs: ShowAsRule): void;
    getSummarizeBy(): AggregationFunction;
    setSummarizeBy(summarizeBy: AggregationFunction): void;
  }

  interface FilterPivotHierarchy extends PivotHierarchy {
    getPosition(): number;
    setPosition(position: number): void;
    getEnableMultipleFilterItems(): boolean;
    setEnableMultipleFilterItems(enableMultipleFilterItems: boolean): void;
    setToDefault(): void;
  }

  // ================================
  // TEXT AND SHAPE TEXT INTERFACES
  // ================================
  interface TextFrame {
    getTextRange(): TextRange;
    getAutoSize(): ShapeAutoSize;
    setAutoSize(autoSize: ShapeAutoSize): void;
    getHasText(): boolean;
    getBottomMargin(): number;
    setBottomMargin(bottomMargin: number): void;
    getTopMargin(): number;
    setTopMargin(topMargin: number): void;
    getLeftMargin(): number;
    setLeftMargin(leftMargin: number): void;
    getRightMargin(): number;
    setRightMargin(rightMargin: number): void;
    getHorizontalAlignment(): ShapeTextHorizontalAlignment;
    setHorizontalAlignment(
      horizontalAlignment: ShapeTextHorizontalAlignment
    ): void;
    getVerticalAlignment(): ShapeTextVerticalAlignment;
    setVerticalAlignment(verticalAlignment: ShapeTextVerticalAlignment): void;
    getHorizontalOverflow(): ShapeTextHorizontalOverflow;
    setHorizontalOverflow(
      horizontalOverflow: ShapeTextHorizontalOverflow
    ): void;
    getVerticalOverflow(): ShapeTextVerticalOverflow;
    setVerticalOverflow(verticalOverflow: ShapeTextVerticalOverflow): void;
    getReadingOrder(): ShapeTextReadingOrder;
    setReadingOrder(readingOrder: ShapeTextReadingOrder): void;
    getOrientation(): ShapeTextOrientation;
    setOrientation(orientation: ShapeTextOrientation): void;
    deleteText(): void;
  }

  interface TextRange {
    getText(): string;
    setText(text: string): void;
    getFont(): ShapeFont;
    getSubstring(start: number, length?: number): TextRange;
  }

  interface ShapeFont {
    getName(): string;
    setName(name: string): void;
    getSize(): number;
    setSize(size: number): void;
    getBold(): boolean;
    setBold(bold: boolean): void;
    getItalic(): boolean;
    setItalic(italic: boolean): void;
    getColor(): string;
    setColor(color: string): void;
  }

  // ================================
  // COMMENT REPLY INTERFACE
  // ================================
  interface CommentReply {
    getContent(): string;
    setContent(content: string): void;
    getContentType(): ContentType;
    getId(): string;
    getAuthorEmail(): string;
    getAuthorName(): string;
    getCreationDate(): Date;
    getLocation(): Range;
    getResolved(): boolean;
    getParentComment(): Comment;
    getRichContent(): string;
    getMentions(): CommentMention[];
    delete(): void;
    updateMentions(contentWithMentions: CommentRichContent): void;
  }

  // ================================
  // PROTECTION INTERFACES
  // ================================
  interface WorksheetProtection {
    getProtected(): boolean;
    protect(options?: WorksheetProtectionOptions, password?: string): void;
    unprotect(password?: string): void;
    getAllowAutoFilter(): boolean;
    getAllowDeleteColumns(): boolean;
    getAllowDeleteRows(): boolean;
    getAllowFormatCells(): boolean;
    getAllowFormatColumns(): boolean;
    getAllowFormatRows(): boolean;
    getAllowInsertColumns(): boolean;
    getAllowInsertRows(): boolean;
    getAllowInsertHyperlinks(): boolean;
    getAllowPivotTables(): boolean;
    getAllowSort(): boolean;
    getAllowEditRanges(): AllowEditRange[];
    getAllowEditRange(key: string): AllowEditRange | undefined;
    addAllowEditRange(
      address: string | Range,
      password?: string
    ): AllowEditRange;
  }

  interface WorkbookProtection {
    getProtected(): boolean;
    protect(password?: string): void;
    unprotect(password?: string): void;
  }

  interface AllowEditRange {
    getAddress(): string;
    setAddress(address: string): void;
    getTitle(): string;
    setTitle(title: string): void;
    getIsPasswordProtected(): boolean;
    setPassword(password?: string): void;
    pauseProtection(password?: string): void;
    delete(): void;
  }

  // ================================
  // DATA VALIDATION INTERFACE
  // ================================
  interface DataValidation {
    getRule(): DataValidationRule;
    setRule(rule: DataValidationRule): void;
    clear(): void;
    getErrorAlert(): DataValidationErrorAlert;
    setErrorAlert(errorAlert: DataValidationErrorAlert): void;
    getInputMessage(): DataValidationInputMessage;
    setInputMessage(inputMessage: DataValidationInputMessage): void;
    getPromptTitle(): string;
    setPromptTitle(promptTitle: string): void;
    getPromptMessage(): string;
    setPromptMessage(promptMessage: string): void;
    getValid(): boolean;
    getIgnoreBlanks(): boolean;
    setIgnoreBlanks(ignoreBlanks: boolean): void;
  }

  // ================================
  // FILTERING INTERFACES
  // ================================
  interface AutoFilter {
    apply(range: Range, columnIndex?: number, criteria?: FilterCriteria): void;
    clearCriteria(): void;
    remove(): void;
    getRange(): Range;
  }

  interface Filter {
    apply(criteria: FilterCriteria): void;
    clear(): void;
    getCriteria(): FilterCriteria;
    applyBottomItemsFilter(count: number): void;
    applyTopItemsFilter(count: number): void;
    applyBottomPercentFilter(percent: number): void;
    applyTopPercentFilter(percent: number): void;
    applyCellColorFilter(color: string): void;
    applyFontColorFilter(color: string): void;
    applyIconFilter(icon: Icon): void;
    applyValuesFilter(values: Array<string | FilterDatetime>): void;
    applyCustomFilter(
      criteria1: string,
      criteria2?: string,
      oper?: FilterOperator2
    ): void;
    applyDynamicFilter(dynamicCriteria: DynamicFilterCriteria): void;
  }

  // ================================
  // SORTING INTERFACES
  // ================================
  interface RangeSort {
    apply(
      fields: SortField[],
      matchCase?: boolean,
      hasHeaders?: boolean,
      orientation?: SortOrientation,
      method?: SortMethod
    ): void;
  }

  interface TableSort {
    apply(fields: SortField[], matchCase?: boolean, method?: SortMethod): void;
    clear(): void;
    reapply(): void;
    getFields(): SortField[];
    getMatchCase(): boolean;
    getMethod(): SortMethod;
  }

  // ================================
  // CHART RELATED INTERFACES
  // ================================
  interface ChartTitle {
    getText(): string;
    setText(text: string): void;
    getVisible(): boolean;
    setVisible(visible: boolean): void;
    getFormat(): ChartTitleFormat;
    getHorizontalAlignment(): ChartTextHorizontalAlignment;
    setHorizontalAlignment(
      horizontalAlignment: ChartTextHorizontalAlignment
    ): void;
    getVerticalAlignment(): ChartTextVerticalAlignment;
    setVerticalAlignment(verticalAlignment: ChartTextVerticalAlignment): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;
    getHeight(): number;
    getWidth(): number;
    setFormula(formula: string): void;
  }

  interface ChartLegend {
    getVisible(): boolean;
    setVisible(visible: boolean): void;
    getPosition(): ChartLegendPosition;
    setPosition(position: ChartLegendPosition): void;
    getFormat(): ChartLegendFormat;
    getHeight(): number;
    setHeight(height: number): void;
    getWidth(): number;
    setWidth(width: number): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;
    getOverlay(): boolean;
    setOverlay(overlay: boolean): void;
    getShowShadow(): boolean;
    setShowShadow(showShadow: boolean): void;
  }

  interface ChartAxes {
    getCategoryAxis(): ChartAxis;
    getValueAxis(): ChartAxis;
    getSecondaryValueAxis(): ChartAxis;
  }

  interface ChartAxis {
    getAlignment(): ChartTickLabelAlignment;
    setAlignment(alignment: ChartTickLabelAlignment): void;
    getBaseTimeUnit(): ChartAxisTimeUnit;
    setBaseTimeUnit(baseTimeUnit: ChartAxisTimeUnit): void;
    getCategoryType(): ChartAxisCategoryType;
    setCategoryType(categoryType: ChartAxisCategoryType): void;
    getDisplayUnit(): ChartAxisDisplayUnit;
    setDisplayUnit(displayUnit: ChartAxisDisplayUnit): void;
    getFormat(): ChartAxisFormat;
    getHeight(): number;
    getIsBetweenCategories(): boolean;
    setIsBetweenCategories(isBetweenCategories: boolean): void;
    getLinkNumberFormat(): boolean;
    setLinkNumberFormat(linkNumberFormat: boolean): void;
    getLogBase(): number;
    setLogBase(logBase: number): void;
    getMajorTimeUnit(): ChartAxisTimeUnit;
    setMajorTimeUnit(majorTimeUnit: ChartAxisTimeUnit): void;
    getMaximum(): number;
    setMaximum(maximum: number): void;
    getMinimum(): number;
    setMinimum(minimum: number): void;
    getMinorTimeUnit(): ChartAxisTimeUnit;
    setMinorTimeUnit(minorTimeUnit: ChartAxisTimeUnit): void;
    getMultiLevel(): boolean;
    setMultiLevel(multiLevel: boolean): void;
    getNumberFormat(): string;
    setNumberFormat(numberFormat: string): void;
    getOffset(): number;
    setOffset(offset: number): void;
    getPosition(): ChartAxisPosition;
    setPosition(position: ChartAxisPosition): void;
    getReversePlotOrder(): boolean;
    setReversePlotOrder(reversePlotOrder: boolean): void;
    getScaleType(): ChartAxisScaleType;
    setScaleType(scaleType: ChartAxisScaleType): void;
    getShowDisplayUnitLabel(): boolean;
    setShowDisplayUnitLabel(showDisplayUnitLabel: boolean): void;
    getTickLabelPosition(): ChartAxisTickLabelPosition;
    setTickLabelPosition(tickLabelPosition: ChartAxisTickLabelPosition): void;
    getTickLabelSpacing(): number;
    setTickLabelSpacing(tickLabelSpacing: number): void;
    getTickMarkSpacing(): number;
    setTickMarkSpacing(tickMarkSpacing: number): void;
    getTitle(): ChartAxisTitle;
    getTop(): number;
    getType(): ChartAxisType;
    getVisible(): boolean;
    setVisible(visible: boolean): void;
    getWidth(): number;
  }

  interface ChartAxisTitle {
    getFormat(): ChartAxisTitleFormat;
    getText(): string;
    setText(text: string): void;
    getTextOrientation(): number;
    setTextOrientation(textOrientation: number): void;
    getVisible(): boolean;
    setVisible(visible: boolean): void;
  }

  interface ChartSeries {
    getName(): string;
    setName(name: string): void;
    getChartType(): ChartType;
    setChartType(chartType: ChartType): void;
    getValues(): Range;
    setValues(values: Range): void;
    getXAxisValues(): Range;
    setXAxisValues(xAxisValues: Range): void;
    getBubbleSizes(): Range;
    setBubbleSizes(bubbleSizes: Range): void;
    getFormat(): ChartSeriesFormat;
    getDataLabels(): ChartDataLabels;
    getMarkerStyle(): ChartMarkerStyle;
    setMarkerStyle(markerStyle: ChartMarkerStyle): void;
    getMarkerSize(): number;
    setMarkerSize(markerSize: number): void;
    getSmooth(): boolean;
    setSmooth(smooth: boolean): void;
    getHasDataLabels(): boolean;
    setHasDataLabels(hasDataLabels: boolean): void;
    getOverlap(): number;
    setOverlap(overlap: number): void;
    getGapWidth(): number;
    setGapWidth(gapWidth: number): void;
    delete(): void;
  }

  interface ChartDataLabels {
    getVisible(): boolean;
    setVisible(visible: boolean): void;
    getPosition(): ChartDataLabelPosition;
    setPosition(position: ChartDataLabelPosition): void;
    getShowSeriesName(): boolean;
    setShowSeriesName(showSeriesName: boolean): void;
    getShowCategoryName(): boolean;
    setShowCategoryName(showCategoryName: boolean): void;
    getShowValue(): boolean;
    setShowValue(showValue: boolean): void;
    getShowPercentage(): boolean;
    setShowPercentage(showPercentage: boolean): void;
    getFormat(): ChartDataLabelFormat;
    getNumberFormat(): string;
    setNumberFormat(numberFormat: string): void;
    getSeparator(): string;
    setSeparator(separator: string): void;
    getShowBubbleSize(): boolean;
    setShowBubbleSize(showBubbleSize: boolean): void;
    getShowLeaderLines(): boolean;
    setShowLeaderLines(showLeaderLines: boolean): void;
  }

  // ================================
  // CONDITIONAL FORMAT SPECIFIC INTERFACES
  // ================================
  interface CellValueConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalCellValueRule;
    setRule(rule: ConditionalCellValueRule): void;
  }

  interface ColorScaleConditionalFormat {
    getCriteria(): ConditionalColorScaleCriteria;
    setCriteria(criteria: ConditionalColorScaleCriteria): void;
    getThreeColorScale(): boolean;
  }

  interface DataBarConditionalFormat {
    getAxisColor(): string;
    setAxisColor(axisColor: string): void;
    getAxisFormat(): ConditionalDataBarAxisFormat;
    setAxisFormat(axisFormat: ConditionalDataBarAxisFormat): void;
    getBarDirection(): ConditionalDataBarDirection;
    setBarDirection(barDirection: ConditionalDataBarDirection): void;
    getLowerBoundRule(): ConditionalDataBarRule;
    setLowerBoundRule(lowerBoundRule: ConditionalDataBarRule): void;
    getUpperBoundRule(): ConditionalDataBarRule;
    setUpperBoundRule(upperBoundRule: ConditionalDataBarRule): void;
    getNegativeFormat(): ConditionalDataBarNegativeColorFormat;
    getPositiveFormat(): ConditionalDataBarPositiveColorFormat;
    getShowDataBarOnly(): boolean;
    setShowDataBarOnly(showDataBarOnly: boolean): void;
  }

  interface IconSetConditionalFormat {
    getCriteria(): ConditionalIconCriterion[];
    setCriteria(criteria: ConditionalIconCriterion[]): void;
    getReverseIconOrder(): boolean;
    setReverseIconOrder(reverseIconOrder: boolean): void;
    getShowIconOnly(): boolean;
    setShowIconOnly(showIconOnly: boolean): void;
    getStyle(): IconSet;
    setStyle(style: IconSet): void;
  }

  interface TopBottomConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalTopBottomRule;
    setRule(rule: ConditionalTopBottomRule): void;
  }

  interface PresetCriteriaConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalPresetCriteriaRule;
    setRule(rule: ConditionalPresetCriteriaRule): void;
  }

  interface TextConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalTextComparisonRule;
    setRule(rule: ConditionalTextComparisonRule): void;
  }

  interface CustomConditionalFormat {
    getFormat(): ConditionalRangeFormat;
    getRule(): ConditionalFormatRule;
  }

  // ================================
  // PAGE LAYOUT INTERFACE
  // ================================
  interface PageLayout {
    // Orientation and size
    getOrientation(): PageOrientation;
    setOrientation(orientation: PageOrientation): void;
    getPaperSize(): PaperType;
    setPaperSize(paperSize: PaperType): void;

    // Margins
    getTopMargin(): number;
    setTopMargin(topMargin: number): void;
    getBottomMargin(): number;
    setBottomMargin(bottomMargin: number): void;
    getLeftMargin(): number;
    setLeftMargin(leftMargin: number): void;
    getRightMargin(): number;
    setRightMargin(rightMargin: number): void;
    getHeaderMargin(): number;
    setHeaderMargin(headerMargin: number): void;
    getFooterMargin(): number;
    setFooterMargin(footerMargin: number): void;

    // Print settings
    getPrintTitleColumns(): string;
    setPrintTitleColumns(printTitleColumns: string): void;
    getPrintTitleRows(): string;
    setPrintTitleRows(printTitleRows: string): void;
    getCenterHorizontally(): boolean;
    setCenterHorizontally(centerHorizontally: boolean): void;
    getCenterVertically(): boolean;
    setCenterVertically(centerVertically: boolean): void;
    getDraftMode(): boolean;
    setDraftMode(draftMode: boolean): void;
    getFirstPageNumber(): number;
    setFirstPageNumber(firstPageNumber: number): void;
    getPrintOrder(): PrintOrder;
    setPrintOrder(printOrder: PrintOrder): void;
    getPrintComments(): PrintComments;
    setPrintComments(printComments: PrintComments): void;
    getPrintErrors(): PrintErrorType;
    setPrintErrors(printErrors: PrintErrorType): void;
    getPrintGridlines(): boolean;
    setPrintGridlines(printGridlines: boolean): void;
    getPrintHeadings(): boolean;
    setPrintHeadings(printHeadings: boolean): void;
    getZoom(): number;
    setZoom(zoom: number): void;
    getFitToPagesWide(): number;
    setFitToPagesWide(fitToPagesWide: number): void;
    getFitToPagesTall(): number;
    setFitToPagesTall(fitToPagesTall: number): void;
    getPrintQuality(): number;
    setPrintQuality(printQuality: number): void;
  }

  // ================================
  // APPLICATION INTERFACE
  // ================================
  interface Application {
    getCalculationMode(): CalculationMode;
    setCalculationMode(calculationMode: CalculationMode): void;
    getCalculationState(): CalculationState;
    calculate(calculationType: CalculationType): void;
    suspendApiCalculationUntilNextSync(): void;
    suspendScreenUpdates(): void;
    enableScreenUpdates(): void;
  }

  // ================================
  // ADDITIONAL CORE INTERFACES
  // ================================
  interface NamedItem {
    getName(): string;
    setName(name: string): void;
    getType(): NamedItemType;
    getValue(): string | number | boolean;
    setValue(value: string | number | boolean): void;
    getFormula(): string;
    setFormula(formula: string): void;
    getRange(): Range;
    delete(): void;
  }

  interface WorkbookProperties {
    getTitle(): string;
    setTitle(title: string): void;
    getAuthor(): string;
    setAuthor(author: string): void;
    getSubject(): string;
    setSubject(subject: string): void;
    getKeywords(): string;
    setKeywords(keywords: string): void;
    getComments(): string;
    setComments(comments: string): void;
    getCategory(): string;
    setCategory(category: string): void;
    getManager(): string;
    setManager(manager: string): void;
    getCompany(): string;
    setCompany(company: string): void;
    getLastAuthor(): string;
    getRevisionNumber(): number;
    setRevisionNumber(revisionNumber: number): void;
    getCustomProperty(key: string): DocumentProperty;
    addCustomProperty(key: string, value: any): DocumentProperty;
  }

  interface RangeAreas {
    getAddress(): string;
    getAddressLocal(): string;
    getCellCount(): number;
    getAreaCount(): number;
    getFormat(): RangeFormat;
    getConditionalFormats(): ConditionalFormat[];
    clear(applyTo?: ClearApplyTo): void;
    copyFrom(
      sourceRange: Range | RangeAreas,
      copyType?: RangeCopyType,
      skipBlanks?: boolean,
      transpose?: boolean
    ): void;
    getValues(): (string | number | boolean)[][];
    setValues(values: (string | number | boolean)[][]): void;
    getFormulas(): string[][];
    setFormulas(formulas: string[][]): void;
    calculate(): void;
    getSpecialCells(
      cellType: SpecialCellType,
      cellValueType?: SpecialCellValueType
    ): RangeAreas;
  }

  interface WorkbookRangeAreas {
    getAddress(): string;
    getAreas(): Range[];
  }

  interface RangeView {
    getCellAddresses(): string[][];
    getColumnCount(): number;
    getFormulas(): string[][];
    getFormulasLocal(): string[][];
    getFormulasR1C1(): string[][];
    getIndex(): number;
    getNumberFormat(): string[][];
    getRange(): Range;
    getRowCount(): number;
    getRows(): RangeView[];
    getText(): string[][];
    getValues(): (string | number | boolean)[][];
    getValueTypes(): RangeValueType[][];
    setFormulas(formulas: string[][]): void;
    setFormulasLocal(formulasLocal: string[][]): void;
    setFormulasR1C1(formulasR1C1: string[][]): void;
    setNumberFormat(numberFormat: string[][]): void;
    setValues(values: (string | number | boolean)[][]): void;
  }

  interface Slicer {
    getName(): string;
    setName(name: string): void;
    getId(): string;
    getCaption(): string;
    setCaption(caption: string): void;
    getHeight(): number;
    setHeight(height: number): void;
    getWidth(): number;
    setWidth(width: number): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;
    delete(): void;
    clearFilters(): void;
    getSlicerItems(): SlicerItem[];
    getWorksheet(): Worksheet;
  }

  interface SlicerItem {
    getName(): string;
    getIsSelected(): boolean;
    setIsSelected(isSelected: boolean): void;
    getHasData(): boolean;
    select(): void;
  }

  // ================================
  // RULE AND CRITERIA INTERFACES
  // ================================
  interface FilterCriteria {
    filterOn: FilterOn;
    values?: string[];
    criterion1?: string;
    criterion2?: string;
    operator?: FilterOperator;
  }

  interface SortField {
    key: number;
    ascending?: boolean;
    color?: string;
    dataOption?: SortDataOption;
    icon?: Icon;
    sortOn?: SortOn;
    subField?: string;
  }

  interface SearchCriteria {
    completeMatch?: boolean;
    matchCase?: boolean;
    searchDirection?: SearchDirection;
  }

  interface ReplaceCriteria {
    completeMatch?: boolean;
    matchCase?: boolean;
  }

  interface WorksheetSearchCriteria {
    completeMatch?: boolean;
    matchCase?: boolean;
  }

  interface ConditionalCellValueRule {
    formula1: string;
    formula2?: string;
    operator: ConditionalCellValueOperator;
  }

  interface ConditionalColorScaleCriteria {
    minimum: ConditionalColorScaleCriterion;
    midpoint?: ConditionalColorScaleCriterion;
    maximum: ConditionalColorScaleCriterion;
  }

  interface ConditionalColorScaleCriterion {
    color?: string;
    formula?: string;
    type: ConditionalFormatColorCriterionType;
  }

  interface ConditionalDataBarRule {
    formula?: string;
    type: ConditionalFormatRuleType;
  }

  interface ConditionalIconCriterion {
    formula?: string;
    operator?: ConditionalIconCriterionOperator;
    type?: ConditionalFormatIconRuleType;
    customIcon?: Icon;
  }

  interface ConditionalTopBottomRule {
    rank: number;
    type: ConditionalTopBottomCriterionType;
  }

  interface ConditionalPresetCriteriaRule {
    criterion: ConditionalFormatPresetCriterion;
  }

  interface ConditionalTextComparisonRule {
    operator: ConditionalTextOperator;
    text: string;
  }

  interface ConditionalFormatRule {
    getFormula(): string;
    setFormula(formula: string): void;
  }

  interface DataValidationRule {
    wholeNumber?: BasicDataValidation;
    decimal?: BasicDataValidation;
    textLength?: BasicDataValidation;
    list?: ListDataValidation;
    date?: DateTimeDataValidation;
    time?: DateTimeDataValidation;
    custom?: CustomDataValidation;
  }

  interface PivotFilters {
    dateFilter?: PivotDateFilter;
    labelFilter?: PivotLabelFilter;
    manualFilter?: PivotManualFilter;
    valueFilter?: PivotValueFilter;
  }

  // ================================
  // SUPPORTING INTERFACES
  // ================================
  interface RemoveDuplicatesResult {
    getRemoved(): number;
    getUniqueRemaining(): number;
  }

  interface RangeHyperlink {
    getAddress(): string;
    setAddress(address: string): void;
    getDocumentReference(): string;
    setDocumentReference(documentReference: string): void;
    getScreenTip(): string;
    setScreenTip(screenTip: string): void;
    getTextToDisplay(): string;
    setTextToDisplay(textToDisplay: string): void;
  }

  interface WorksheetFreezePanes {
    freezeAt(frozenRange: Range): void;
    freezeColumns(count: number): void;
    freezeRows(count: number): void;
    getLocation(): Range;
    unfreeze(): void;
  }

  interface WorksheetCustomProperty {
    getKey(): string;
    getValue(): string;
    setValue(value: string): void;
    delete(): void;
  }

  interface NamedSheetView {
    getName(): string;
    activate(): void;
    duplicate(name?: string): NamedSheetView;
    delete(): void;
  }

  interface DocumentProperty {
    getKey(): string;
    getValue(): any;
    setValue(value: any): void;
    getType(): DocumentPropertyType;
  }

  interface CustomXmlPart {
    delete(): void;
    getId(): string;
    getNamespaceUri(): string;
    getXml(): string;
    setXml(xml: string): void;
  }

  interface Binding {
    getId(): string;
    getType(): BindingType;
    delete(): void;
    getRange(): Range;
    getText(): string;
    getTable(): Table;
    getMatrix(): any[][];
  }

  interface Query {
    getName(): string;
    getRefreshDate(): Date;
    getRowsLoadedCount(): number;
    refresh(): void;
  }

  interface PageBreak {
    getColumnIndex(): number;
    getRowIndex(): number;
    delete(): void;
  }

  // ================================
  // RICH CONTENT INTERFACES
  // ================================
  interface CommentRichContent {
    richContent: string;
    mentions?: CommentMention[];
  }

  interface CommentMention {
    email: string;
    id: number;
    name: string;
  }

  // ================================
  // CELL CONTROL INTERFACES
  // ================================
}

// ================================
// CELL CONTROL INTERFACES
// ================================
interface CellControl {
  getType(): CellControlType;
}

interface UnknownCellControl extends CellControl {
  type: CellControlType.unknown;
}

interface CheckboxCellControl extends CellControl {
  type: CellControlType.checkbox;
  getChecked(): boolean;
  setChecked(checked: boolean): void;
}

interface EmptyCellControl extends CellControl {
  type: CellControlType.empty;
}

// ================================
// FORMAT INTERFACES - CHART FORMATTING
// ================================
interface ChartTitleFormat {
  getBorder(): ChartBorder;
  getFill(): ChartFill;
  getFont(): ChartFont;
}

interface ChartLegendFormat {
  getBorder(): ChartBorder;
  getFill(): ChartFill;
  getFont(): ChartFont;
}

interface ChartAxisFormat {
  getFill(): ChartFill;
  getFont(): ChartFont;
  getLine(): ChartLineFormat;
}

interface ChartAxisTitleFormat {
  getBorder(): ChartBorder;
  getFill(): ChartFill;
  getFont(): ChartFont;
}

interface ChartDataLabelFormat {
  getBorder(): ChartBorder;
  getFill(): ChartFill;
  getFont(): ChartFont;
}

interface ChartSeriesFormat {
  getFill(): ChartFill;
  getLine(): ChartLineFormat;
}

interface ChartPlotArea {
  getFormat(): ChartPlotAreaFormat;
  getHeight(): number;
  setHeight(height: number): void;
  getWidth(): number;
  setWidth(width: number): void;
  getTop(): number;
  setTop(top: number): void;
  getLeft(): number;
  setLeft(left: number): void;
  getInsideHeight(): number;
  getInsideWidth(): number;
  getInsideTop(): number;
  getInsideLeft(): number;
}

interface ChartPlotAreaFormat {
  getBorder(): ChartBorder;
  getFill(): ChartFill;
}

interface ChartBorder {
  getColor(): string;
  setColor(color: string): void;
  getLineStyle(): ChartLineStyle;
  setLineStyle(lineStyle: ChartLineStyle): void;
  getWeight(): number;
  setWeight(weight: number): void;
  clear(): void;
}

interface ChartFill {
  getType(): ChartFillType;
  setSolidColor(color: string): void;
  clear(): void;
}

interface ChartFont {
  getBold(): boolean;
  setBold(bold: boolean): void;
  getColor(): string;
  setColor(color: string): void;
  getItalic(): boolean;
  setItalic(italic: boolean): void;
  getName(): string;
  setName(name: string): void;
  getSize(): number;
  setSize(size: number): void;
  getUnderline(): ChartUnderlineStyle;
  setUnderline(underline: ChartUnderlineStyle): void;
}

interface ChartLineFormat {
  getColor(): string;
  setColor(color: string): void;
  getLineStyle(): ChartLineStyle;
  setLineStyle(lineStyle: ChartLineStyle): void;
  getWeight(): number;
  setWeight(weight: number): void;
  clear(): void;
}

interface ChartPivotOptions {
  getShowAxisFieldButtons(): boolean;
  setShowAxisFieldButtons(showAxisFieldButtons: boolean): void;
  getShowLegendFieldButtons(): boolean;
  setShowLegendFieldButtons(showLegendFieldButtons: boolean): void;
  getShowReportFilterFieldButtons(): boolean;
  setShowReportFilterFieldButtons(showReportFilterFieldButtons: boolean): void;
  getShowValueFieldButtons(): boolean;
  setShowValueFieldButtons(showValueFieldButtons: boolean): void;
}

// ================================
// CONDITIONAL FORMAT SUPPORT INTERFACES
// ================================
interface ConditionalRangeFormat {
  getNumberFormat(): string;
  setNumberFormat(numberFormat: string): void;
  getBorders(): ConditionalRangeBorderCollection;
  getFill(): ConditionalRangeFill;
  getFont(): ConditionalRangeFont;
}

interface ConditionalRangeBorderCollection {
  getTop(): ConditionalRangeBorder;
  getBottom(): ConditionalRangeBorder;
  getLeft(): ConditionalRangeBorder;
  getRight(): ConditionalRangeBorder;
}

interface ConditionalRangeBorder {
  getColor(): string;
  setColor(color: string): void;
  getStyle(): BorderLineStyle;
  setStyle(style: BorderLineStyle): void;
}

interface ConditionalRangeFill {
  getColor(): string;
  setColor(color: string): void;
  clear(): void;
}

interface ConditionalRangeFont {
  getBold(): boolean;
  setBold(bold: boolean): void;
  getColor(): string;
  setColor(color: string): void;
  getItalic(): boolean;
  setItalic(italic: boolean): void;
  getStrikethrough(): boolean;
  setStrikethrough(strikethrough: boolean): void;
  getUnderline(): ConditionalRangeFontUnderlineStyle;
  setUnderline(underline: ConditionalRangeFontUnderlineStyle): void;
}

interface ConditionalDataBarPositiveColorFormat {
  getBorderColor(): string;
  setBorderColor(borderColor: string): void;
  getFillColor(): string;
  setFillColor(fillColor: string): void;
  getGradientFill(): boolean;
  setGradientFill(gradientFill: boolean): void;
}

interface ConditionalDataBarNegativeColorFormat {
  getBorderColor(): string;
  setBorderColor(borderColor: string): void;
  getFillColor(): string;
  setFillColor(fillColor: string): void;
  getMatchPositiveBorderColor(): boolean;
  setMatchPositiveBorderColor(matchPositiveBorderColor: boolean): void;
  getMatchPositiveFillColor(): boolean;
  setMatchPositiveFillColor(matchPositiveFillColor: boolean): void;
}

// ================================
// DATA VALIDATION SUPPORT INTERFACES
// ================================
interface DataValidationErrorAlert {
  getMessage(): string;
  setMessage(message: string): void;
  getShowAlert(): boolean;
  setShowAlert(showAlert: boolean): void;
  getStyle(): DataValidationAlertStyle;
  setStyle(style: DataValidationAlertStyle): void;
  getTitle(): string;
  setTitle(title: string): void;
}

interface DataValidationInputMessage {
  getMessage(): string;
  setMessage(message: string): void;
  getShowInputMessage(): boolean;
  setShowInputMessage(showInputMessage: boolean): void;
  getTitle(): string;
  setTitle(title: string): void;
}

interface BasicDataValidation {
  formula1: string;
  formula2?: string;
  operator: DataValidationOperator;
}

interface ListDataValidation {
  inCellDropDown?: boolean;
  source: string;
}

interface CustomDataValidation {
  formula: string;
}

interface DateTimeDataValidation {
  formula1: string;
  formula2?: string;
  operator: DataValidationOperator;
}

// ================================
// PIVOT SPECIFIC INTERFACES
// ================================
interface PivotItem {
  getName(): string;
  setName(name: string): void;
  getId(): string;
  getIsExpanded(): boolean;
  setIsExpanded(isExpanded: boolean): void;
  getVisible(): boolean;
  setVisible(visible: boolean): void;
}

interface PivotDateFilter {
  comparator?: PivotFilterComparator;
  condition: DateFilterCondition;
  exclusive?: boolean;
  lowerBound?: string;
  upperBound?: string;
  wholeDays?: boolean;
}

interface PivotLabelFilter {
  comparator?: PivotFilterComparator;
  condition: LabelFilterCondition;
  exclusive?: boolean;
  lowerBound?: string;
  upperBound?: string;
}

interface PivotManualFilter {
  selectedItems?: string[];
}

interface PivotValueFilter {
  comparator?: PivotFilterComparator;
  condition: ValueFilterCondition;
  exclusive?: boolean;
  lowerBound?: number;
  upperBound?: number;
  selectionType?: PivotValueSelectionType;
  threshold?: number;
}

interface ShowAsRule {
  baseField?: PivotField;
  baseItem?: PivotItem;
  calculation: ShowAsCalculation;
}

interface Subtotals {
  automatic?: boolean;
  average?: boolean;
  count?: boolean;
  countNumbers?: boolean;
  max?: boolean;
  min?: boolean;
  product?: boolean;
  standardDeviation?: boolean;
  standardDeviationP?: boolean;
  sum?: boolean;
  variance?: boolean;
  varianceP?: boolean;
}

// ================================
// LINKED WORKBOOK INTERFACES
// ================================
interface LinkedWorkbook {
  getLinkedWorkbookUrl(): string;
}

// ================================
// STYLE INTERFACES
// ================================
interface TableStyle {
  // Style properties for tables
  getName(): string;
  getReadOnly(): boolean;
}

interface PivotTableStyle {
  // Style properties for pivot tables
  getName(): string;
  getReadOnly(): boolean;
}

interface SlicerStyle {
  // Style properties for slicers
  getName(): string;
  getReadOnly(): boolean;
}

interface TimelineStyle {
  // Style properties for timelines
  getName(): string;
  getReadOnly(): boolean;
  setName(name: string): void;
}

interface Style {
  // Style properties for cells
  getName(): string;
  getReadOnly(): boolean;
}

// ================================
// SHAPE GROUP INTERFACE
// ================================
interface ShapeGroup {
  getId(): string;
  getShape(): Shape;
  ungroup(): void;
}

// ================================
// FILTER RELATED SUPPORT
// ================================
interface FilterDatetime {
  date: string;
  specificity: FilterDatetimeSpecificity;
}

interface Icon {
  index: number;
  set: IconSet;
}

// ================================
// PROTECTION OPTIONS
// ================================
interface WorksheetProtectionOptions {
  allowAutoFilter?: boolean;
  allowDeleteColumns?: boolean;
  allowDeleteRows?: boolean;
  allowFormatCells?: boolean;
  allowFormatColumns?: boolean;
  allowFormatRows?: boolean;
  allowInsertColumns?: boolean;
  allowInsertRows?: boolean;
  allowInsertHyperlinks?: boolean;
  allowPivotTables?: boolean;
  allowSort?: boolean;
  selectionMode?: ProtectionSelectionMode;
}

// ================================
// ENUMS AND CONSTANTS
// ================================
enum CellControlType {
  checkbox = 'Checkbox',
  empty = 'Empty',
  unknown = 'Unknown',
}

enum SheetVisibility {
  visible = 'Visible',
  hidden = 'Hidden',
  veryHidden = 'VeryHidden',
}

enum RangeValueType {
  unknown = 'Unknown',
  empty = 'Empty',
  string = 'String',
  integer = 'Integer',
  double = 'Double',
  boolean = 'Boolean',
  error = 'Error',
  richValue = 'RichValue',
}

enum ChartType {
  invalid = 'Invalid',
  columnClustered = 'ColumnClustered',
  columnStacked = 'ColumnStacked',
  columnStacked100 = 'ColumnStacked100',
  line = 'Line',
  lineStacked = 'LineStacked',
  lineStacked100 = 'LineStacked100',
  lineMarkers = 'LineMarkers',
  lineMarkersStacked = 'LineMarkersStacked',
  lineMarkersStacked100 = 'LineMarkersStacked100',
  pie = 'Pie',
  pieExploded = 'PieExploded',
  pieOfPie = 'PieOfPie',
  barOfPie = 'BarOfPie',
  barClustered = 'BarClustered',
  barStacked = 'BarStacked',
  barStacked100 = 'BarStacked100',
  area = 'Area',
  areaStacked = 'AreaStacked',
  areaStacked100 = 'AreaStacked100',
  doughnut = 'Doughnut',
  doughnutExploded = 'DoughnutExploded',
  radar = 'Radar',
  radarMarkers = 'RadarMarkers',
  radarFilled = 'RadarFilled',
  surface = 'Surface',
  surfaceWireframe = 'SurfaceWireframe',
  surfaceTopView = 'SurfaceTopView',
  surfaceTopViewWireframe = 'SurfaceTopViewWireframe',
  bubble = 'Bubble',
  bubble3DEffect = 'Bubble3DEffect',
  stockHLC = 'StockHLC',
  stockOHLC = 'StockOHLC',
  stockVHLC = 'StockVHLC',
  stockVOHLC = 'StockVOHLC',
  cylinderColClustered = 'CylinderColClustered',
  cylinderColStacked = 'CylinderColStacked',
  cylinderColStacked100 = 'CylinderColStacked100',
  cylinderBarClustered = 'CylinderBarClustered',
  cylinderBarStacked = 'CylinderBarStacked',
  cylinderBarStacked100 = 'CylinderBarStacked100',
  cylinderCol = 'CylinderCol',
  coneColClustered = 'ConeColClustered',
  coneColStacked = 'ConeColStacked',
  coneColStacked100 = 'ConeColStacked100',
  coneBarClustered = 'ConeBarClustered',
  coneBarStacked = 'ConeBarStacked',
  coneBarStacked100 = 'ConeBarStacked100',
  coneCol = 'ConeCol',
  pyramidColClustered = 'PyramidColClustered',
  pyramidColStacked = 'PyramidColStacked',
  pyramidColStacked100 = 'PyramidColStacked100',
  pyramidBarClustered = 'PyramidBarClustered',
  pyramidBarStacked = 'PyramidBarStacked',
  pyramidBarStacked100 = 'PyramidBarStacked100',
  pyramidCol = 'PyramidCol',
  histogram = 'Histogram',
  pareto = 'Pareto',
  boxWhisker = 'BoxWhisker',
  waterfall = 'Waterfall',
  funnel = 'Funnel',
  treemap = 'Treemap',
  sunburst = 'Sunburst',
  combo = 'Combo',
  regionMap = 'RegionMap',
}

enum ChartSeriesBy {
  auto = 'Auto',
  columns = 'Columns',
  rows = 'Rows',
}

enum ChartPlotBy {
  rows = 'Rows',
  columns = 'Columns',
}

enum ChartDisplayBlanksAs {
  notPlotted = 'NotPlotted',
  zero = 'Zero',
  interplotted = 'Interplotted',
}

enum GeometricShapeType {
  lineInverse = 'LineInverse',
  triangle = 'Triangle',
  rightTriangle = 'RightTriangle',
  rectangle = 'Rectangle',
  diamond = 'Diamond',
  hexagon = 'Hexagon',
  octagon = 'Octagon',
  plus = 'Plus',
  star = 'Star',
  arrow = 'Arrow',
  thickArrow = 'ThickArrow',
  homePlate = 'HomePlate',
  cube = 'Cube',
  balloon = 'Balloon',
  seal = 'Seal',
  arc = 'Arc',
  line = 'Line',
  plaque = 'Plaque',
  can = 'Can',
  donut = 'Donut',
  textSimple = 'TextSimple',
  textOctagon = 'TextOctagon',
  textHexagon = 'TextHexagon',
  textCurve = 'TextCurve',
  textWave = 'TextWave',
  textRing = 'TextRing',
  textOnCurve = 'TextOnCurve',
  textOnRing = 'TextOnRing',
  straightConnector = 'StraightConnector',
  bentConnector = 'BentConnector',
  curvedConnector = 'CurvedConnector',
  leftArrow = 'LeftArrow',
  downArrow = 'DownArrow',
  upArrow = 'UpArrow',
  leftRightArrow = 'LeftRightArrow',
  upDownArrow = 'UpDownArrow',
  irregular = 'IrregularSeal1',
  lightning = 'Lightning',
  heart = 'Heart',
  frame = 'Frame',
  halfFrame = 'HalfFrame',
  corner = 'Corner',
  diagonalStripe = 'DiagonalStripe',
  chord = 'Chord',
  moon = 'Moon',
  prohibitedSign = 'NoSmoking',
  blockArc = 'BlockArc',
  smileyFace = 'SmileyFace',
  verticalScroll = 'VerticalScroll',
  horizontalScroll = 'HorizontalScroll',
  circularArrow = 'CircularArrow',
  notchedRightArrow = 'NotchedRightArrow',
  bentUpArrow = 'BentUpArrow',
  leftUpArrow = 'LeftUpArrow',
  leftCircularArrow = 'LeftCircularArrow',
  leftRightCircularArrow = 'LeftRightCircularArrow',
  quadArrow = 'QuadArrow',
  leftArrowCallout = 'LeftArrowCallout',
  rightArrowCallout = 'RightArrowCallout',
  upArrowCallout = 'UpArrowCallout',
  downArrowCallout = 'DownArrowCallout',
  leftRightArrowCallout = 'LeftRightArrowCallout',
  upDownArrowCallout = 'UpDownArrowCallout',
  quadArrowCallout = 'QuadArrowCallout',
  bentArrow = 'BentArrow',
  uturnArrow = 'UturnArrow',
  leftBrace = 'LeftBrace',
  rightBrace = 'RightBrace',
  leftBracket = 'LeftBracket',
  rightBracket = 'RightBracket',
  callout1 = 'Callout1',
  callout2 = 'Callout2',
  callout3 = 'Callout3',
  accentCallout1 = 'AccentCallout1',
  accentCallout2 = 'AccentCallout2',
  accentCallout3 = 'AccentCallout3',
  borderCallout1 = 'BorderCallout1',
  borderCallout2 = 'BorderCallout2',
  borderCallout3 = 'BorderCallout3',
  accentBorderCallout1 = 'AccentBorderCallout1',
  accentBorderCallout2 = 'AccentBorderCallout2',
  accentBorderCallout3 = 'AccentBorderCallout3',
  wedgeRectangle = 'WedgeRectCallout',
  wedgeRRectangle = 'WedgeRRectCallout',
  wedgeEllipse = 'WedgeEllipseCallout',
  cloudCallout = 'CloudCallout',
  cloud = 'Cloud',
  ribbon = 'Ribbon',
  ribbon2 = 'Ribbon2',
  ellipseRibbon = 'EllipseRibbon',
  ellipseRibbon2 = 'EllipseRibbon2',
  leftRightRibbon = 'LeftRightRibbon',
  verticalRibbon = 'VerticalRibbon',
  leftCircularArrow2 = 'LeftCircularArrow',
  notchedCircularArrow = 'NotchedCircularArrow',
  bentCircularArrow = 'BentCircularArrow',
  leftRightCircularArrow2 = 'LeftRightCircularArrow',
  circle = 'Circle',
  ellipse = 'Ellipse',
}

enum ConnectorType {
  straight = 'Straight',
  elbow = 'Elbow',
  curve = 'Curve',
}

enum ShapeType {
  unsupported = 'Unsupported',
  image = 'Image',
  geometricShape = 'GeometricShape',
  group = 'Group',
  line = 'Line',
}

enum Placement {
  absolute = 'Absolute',
  oneCell = 'OneCell',
  twoCell = 'TwoCell',
}

enum ShapeScaleType {
  currentSize = 'CurrentSize',
  originalSize = 'OriginalSize',
  relativeToOriginalSize = 'RelativeToOriginalSize',
}

enum ShapeScaleFrom {
  scaleFromTopLeft = 'ScaleFromTopLeft',
  scaleFromMiddle = 'ScaleFromMiddle',
  scaleFromBottomRight = 'ScaleFromBottomRight',
}

enum ShapeZOrder {
  bringToFront = 'BringToFront',
  bringForward = 'BringForward',
  sendToBack = 'SendToBack',
  sendBackward = 'SendBackward',
}

enum PictureFormat {
  bmp = 'BMP',
  gif = 'GIF',
  jpeg = 'JPEG',
  png = 'PNG',
  svg = 'SVG',
}

enum ShapeAutoSize {
  autoSizeNone = 'AutoSizeNone',
  autoSizeMixed = 'AutoSizeMixed',
  autoSizeTextToFitShape = 'AutoSizeTextToFitShape',
  autoSizeShapeToFitText = 'AutoSizeShapeToFitText',
}

enum ShapeTextHorizontalAlignment {
  left = 'Left',
  center = 'Center',
  right = 'Right',
  justify = 'Justify',
  justifyLow = 'JustifyLow',
  distributed = 'Distributed',
  thaiDistributed = 'ThaiDistributed',
}

enum ShapeTextVerticalAlignment {
  top = 'Top',
  middle = 'Middle',
  bottom = 'Bottom',
  justify = 'Justify',
  distributed = 'Distributed',
}

enum ShapeTextHorizontalOverflow {
  overflow = 'Overflow',
  clip = 'Clip',
}

enum ShapeTextVerticalOverflow {
  overflow = 'Overflow',
  ellipsis = 'Ellipsis',
  clip = 'Clip',
}

enum ShapeTextReadingOrder {
  leftToRight = 'LeftToRight',
  rightToLeft = 'RightToLeft',
  context = 'Context',
}

enum ShapeTextOrientation {
  horizontal = 'Horizontal',
  verticalFarEast = 'VerticalFarEast',
  vertical = 'Vertical',
  vertical270 = 'Vertical270',
  wordArtVertical = 'WordArtVertical',
  wordArtVerticalRightToLeft = 'WordArtVerticalRightToLeft',
}

enum ContentType {
  plain = 'Plain',
  mention = 'Mention',
}

// ================================
// ALIGNMENT AND FORMATTING ENUMS
// ================================
enum HorizontalAlignment {
  general = 'General',
  left = 'Left',
  center = 'Center',
  right = 'Right',
  fill = 'Fill',
  justify = 'Justify',
  centerAcrossSelection = 'CenterAcrossSelection',
  distributed = 'Distributed',
}

enum VerticalAlignment {
  top = 'Top',
  center = 'Center',
  bottom = 'Bottom',
  justify = 'Justify',
  distributed = 'Distributed',
}

enum ReadingOrder {
  context = 'Context',
  leftToRight = 'LeftToRight',
  rightToLeft = 'RightToLeft',
}

enum FillPattern {
  none = 'None',
  solid = 'Solid',
  gray50 = 'Gray50',
  gray75 = 'Gray75',
  gray25 = 'Gray25',
  horizontal = 'Horizontal',
  vertical = 'Vertical',
  down = 'Down',
  up = 'Up',
  checker = 'Checker',
  semiGray75 = 'SemiGray75',
  lightHorizontal = 'LightHorizontal',
  lightVertical = 'LightVertical',
  lightDown = 'LightDown',
  lightUp = 'LightUp',
  grid = 'Grid',
  crissCross = 'CrissCross',
  gray16 = 'Gray16',
  gray8 = 'Gray8',
  linearGradient = 'LinearGradient',
  rectangularGradient = 'RectangularGradient',
}

enum RangeFontUnderlineStyle {
  none = 'None',
  single = 'Single',
  double = 'Double',
  singleAccountingUnderline = 'SingleAccountingUnderline',
  doubleAccountingUnderline = 'DoubleAccountingUnderline',
}

enum BorderLineStyle {
  none = 'None',
  continuous = 'Continuous',
  dash = 'Dash',
  dashDot = 'DashDot',
  dashDotDot = 'DashDotDot',
  dot = 'Dot',
  double = 'Double',
  slantDashDot = 'SlantDashDot',
}

enum BorderWeight {
  hairline = 'Hairline',
  thin = 'Thin',
  medium = 'Medium',
  thick = 'Thick',
}

enum BorderIndex {
  edgeTop = 'EdgeTop',
  edgeBottom = 'EdgeBottom',
  edgeLeft = 'EdgeLeft',
  edgeRight = 'EdgeRight',
  insideVertical = 'InsideVertical',
  insideHorizontal = 'InsideHorizontal',
  diagonalDown = 'DiagonalDown',
  diagonalUp = 'DiagonalUp',
}

// ================================
// OPERATION AND DIRECTION ENUMS
// ================================
enum ClearApplyTo {
  all = 'All',
  formats = 'Formats',
  contents = 'Contents',
  hyperlinks = 'Hyperlinks',
  removeHyperlinks = 'RemoveHyperlinks',
}

enum DeleteShiftDirection {
  up = 'Up',
  left = 'Left',
}

enum InsertShiftDirection {
  down = 'Down',
  right = 'Right',
}

enum KeyboardDirection {
  up = 'Up',
  down = 'Down',
  left = 'Left',
  right = 'Right',
}

enum RangeCopyType {
  all = 'All',
  formulas = 'Formulas',
  values = 'Values',
  formats = 'Formats',
  formulasAndNumberFormats = 'FormulasAndNumberFormats',
  valuesAndNumberFormats = 'ValuesAndNumberFormats',
}

enum AutoFillType {
  fillDefault = 'FillDefault',
  fillCopy = 'FillCopy',
  fillSeries = 'FillSeries',
  fillFormats = 'FillFormats',
  fillValues = 'FillValues',
  fillDays = 'FillDays',
  fillWeekdays = 'FillWeekdays',
  fillMonths = 'FillMonths',
  fillYears = 'FillYears',
  linearTrend = 'LinearTrend',
  growthTrend = 'GrowthTrend',
  flashFill = 'FlashFill',
}

enum GroupOption {
  byRows = 'ByRows',
  byColumns = 'ByColumns',
}

enum SearchDirection {
  forward = 'Forward',
  backwards = 'Backwards',
}

enum SpecialCellType {
  conditionalFormats = 'ConditionalFormats',
  dataValidations = 'DataValidations',
  blanks = 'Blanks',
  constants = 'Constants',
  formulas = 'Formulas',
  sameConditionalFormat = 'SameConditionalFormat',
  sameDataValidation = 'SameDataValidation',
  visible = 'Visible',
}

enum SpecialCellValueType {
  errors = 'Errors',
  errorValues = 'ErrorValues',
  logicalValues = 'LogicalValues',
  numbers = 'Numbers',
  text = 'Text',
}

// ================================
// CONDITIONAL FORMATTING ENUMS
// ================================
enum ConditionalFormatType {
  custom = 'Custom',
  dataBar = 'DataBar',
  colorScale = 'ColorScale',
  iconSet = 'IconSet',
  topBottom = 'TopBottom',
  presetCriteria = 'PresetCriteria',
  containsText = 'ContainsText',
  cellValue = 'CellValue',
}

enum ConditionalCellValueOperator {
  invalid = 'Invalid',
  between = 'Between',
  notBetween = 'NotBetween',
  equalTo = 'EqualTo',
  notEqualTo = 'NotEqualTo',
  greaterThan = 'GreaterThan',
  lessThan = 'LessThan',
  greaterThanOrEqual = 'GreaterThanOrEqual',
  lessThanOrEqual = 'LessThanOrEqual',
}

enum ConditionalFormatColorCriterionType {
  invalid = 'Invalid',
  lowestValue = 'LowestValue',
  highestValue = 'HighestValue',
  number = 'Number',
  percent = 'Percent',
  formula = 'Formula',
  percentile = 'Percentile',
}

enum ConditionalFormatRuleType {
  invalid = 'Invalid',
  automaticMin = 'AutomaticMin',
  automaticMax = 'AutomaticMax',
  lowestValue = 'LowestValue',
  highestValue = 'HighestValue',
  number = 'Number',
  percent = 'Percent',
  formula = 'Formula',
  percentile = 'Percentile',
}

enum ConditionalFormatIconRuleType {
  invalid = 'Invalid',
  number = 'Number',
  percent = 'Percent',
  formula = 'Formula',
  percentile = 'Percentile',
}

enum ConditionalIconCriterionOperator {
  invalid = 'Invalid',
  greaterThan = 'GreaterThan',
  greaterThanOrEqual = 'GreaterThanOrEqual',
}

enum ConditionalTopBottomCriterionType {
  invalid = 'Invalid',
  topItems = 'TopItems',
  topPercent = 'TopPercent',
  bottomItems = 'BottomItems',
  bottomPercent = 'BottomPercent',
}

enum ConditionalFormatPresetCriterion {
  invalid = 'Invalid',
  blanks = 'Blanks',
  nonBlanks = 'NonBlanks',
  errors = 'Errors',
  nonErrors = 'NonErrors',
  yesterday = 'Yesterday',
  today = 'Today',
  tomorrow = 'Tomorrow',
  lastSevenDays = 'LastSevenDays',
  lastWeek = 'LastWeek',
  thisWeek = 'ThisWeek',
  nextWeek = 'NextWeek',
  lastMonth = 'LastMonth',
  thisMonth = 'ThisMonth',
  nextMonth = 'NextMonth',
  aboveAverage = 'AboveAverage',
  belowAverage = 'BelowAverage',
  equalOrAboveAverage = 'EqualOrAboveAverage',
  equalOrBelowAverage = 'EqualOrBelowAverage',
  oneStdDevAboveAverage = 'OneStdDevAboveAverage',
  oneStdDevBelowAverage = 'OneStdDevBelowAverage',
  twoStdDevAboveAverage = 'TwoStdDevAboveAverage',
  twoStdDevBelowAverage = 'TwoStdDevBelowAverage',
  threeStdDevAboveAverage = 'ThreeStdDevAboveAverage',
  threeStdDevBelowAverage = 'ThreeStdDevBelowAverage',
  uniqueValues = 'UniqueValues',
  duplicateValues = 'DuplicateValues',
}

enum ConditionalTextOperator {
  invalid = 'Invalid',
  contains = 'Contains',
  notContains = 'NotContains',
  beginsWith = 'BeginsWith',
  endsWith = 'EndsWith',
}

enum ConditionalDataBarDirection {
  context = 'Context',
  leftToRight = 'LeftToRight',
  rightToLeft = 'RightToLeft',
}

enum ConditionalDataBarAxisFormat {
  automatic = 'Automatic',
  none = 'None',
  cellMidPoint = 'CellMidPoint',
}

enum ConditionalRangeFontUnderlineStyle {
  none = 'None',
  single = 'Single',
  double = 'Double',
}

enum IconSet {
  invalid = 'Invalid',
  threeArrows = 'ThreeArrows',
  threeArrowsGray = 'ThreeArrowsGray',
  threeFlags = 'ThreeFlags',
  threeTrafficLights1 = 'ThreeTrafficLights1',
  threeTrafficLights2 = 'ThreeTrafficLights2',
  threeSigns = 'ThreeSigns',
  threeSymbols = 'ThreeSymbols',
  threeSymbols2 = 'ThreeSymbols2',
  fourArrows = 'FourArrows',
  fourArrowsGray = 'FourArrowsGray',
  fourRedToBlack = 'FourRedToBlack',
  fourRating = 'FourRating',
  fourTrafficLights = 'FourTrafficLights',
  fiveArrows = 'FiveArrows',
  fiveArrowsGray = 'FiveArrowsGray',
  fiveRating = 'FiveRating',
  fiveQuarters = 'FiveQuarters',
  threeStars = 'ThreeStars',
  threeTriangles = 'ThreeTriangles',
  fiveBoxes = 'FiveBoxes',
}

// ================================
// CHART SPECIFIC ENUMS
// ================================
enum ChartLegendPosition {
  invalid = 'Invalid',
  top = 'Top',
  bottom = 'Bottom',
  left = 'Left',
  right = 'Right',
  corner = 'Corner',
  custom = 'Custom',
}

enum ChartDataLabelPosition {
  invalid = 'Invalid',
  none = 'None',
  center = 'Center',
  insideEnd = 'InsideEnd',
  insideBase = 'InsideBase',
  outsideEnd = 'OutsideEnd',
  left = 'Left',
  right = 'Right',
  top = 'Top',
  bottom = 'Bottom',
  bestFit = 'BestFit',
  callout = 'Callout',
}

enum ChartMarkerStyle {
  invalid = 'Invalid',
  automatic = 'Automatic',
  none = 'None',
  square = 'Square',
  diamond = 'Diamond',
  triangle = 'Triangle',
  x = 'X',
  star = 'Star',
  dot = 'Dot',
  dash = 'Dash',
  circle = 'Circle',
  plus = 'Plus',
  picture = 'Picture',
}

enum ChartLineStyle {
  none = 'None',
  continuous = 'Continuous',
  dash = 'Dash',
  dashDot = 'DashDot',
  dashDotDot = 'DashDotDot',
  dot = 'Dot',
  grey25 = 'Grey25',
  grey50 = 'Grey50',
  grey75 = 'Grey75',
  automatic = 'Automatic',
  roundDot = 'RoundDot',
}

enum ChartFillType {
  noFill = 'NoFill',
  automatic = 'Automatic',
  solidColor = 'SolidColor',
  gradient = 'Gradient',
  pattern = 'Pattern',
  pictureAndTexture = 'PictureAndTexture',
}

enum ChartUnderlineStyle {
  none = 'None',
  single = 'Single',
}

enum ChartTextHorizontalAlignment {
  center = 'Center',
  left = 'Left',
  right = 'Right',
  justify = 'Justify',
  distributed = 'Distributed',
}

enum ChartTextVerticalAlignment {
  center = 'Center',
  bottom = 'Bottom',
  top = 'Top',
  justify = 'Justify',
  distributed = 'Distributed',
}

enum ChartAxisCategoryType {
  automatic = 'Automatic',
  textAxis = 'TextAxis',
  dateAxis = 'DateAxis',
}

enum ChartAxisDisplayUnit {
  none = 'None',
  hundreds = 'Hundreds',
  thousands = 'Thousands',
  tenThousands = 'TenThousands',
  hundredThousands = 'HundredThousands',
  millions = 'Millions',
  tenMillions = 'TenMillions',
  hundredMillions = 'HundredMillions',
  billions = 'Billions',
  trillions = 'Trillions',
  custom = 'Custom',
}

enum ChartAxisPosition {
  automatic = 'Automatic',
  maximum = 'Maximum',
  minimum = 'Minimum',
  custom = 'Custom',
}

enum ChartAxisScaleType {
  linear = 'Linear',
  logarithmic = 'Logarithmic',
}

enum ChartAxisTickLabelPosition {
  nextToAxis = 'NextToAxis',
  high = 'High',
  low = 'Low',
  none = 'None',
}

enum ChartAxisTickMarkType {
  none = 'None',
  cross = 'Cross',
  inside = 'Inside',
  outside = 'Outside',
}

enum ChartAxisTimeUnit {
  days = 'Days',
  months = 'Months',
  years = 'Years',
}

enum ChartAxisType {
  invalid = 'Invalid',
  category = 'Category',
  value = 'Value',
  series = 'Series',
}

enum ChartTickLabelAlignment {
  center = 'Center',
  left = 'Left',
  right = 'Right',
}

// ================================
// FILTERING ENUMS
// ================================
enum FilterOn {
  bottomItems = 'BottomItems',
  bottomPercent = 'BottomPercent',
  cellColor = 'CellColor',
  dynamic = 'Dynamic',
  fontColor = 'FontColor',
  values = 'Values',
  topItems = 'TopItems',
  topPercent = 'TopPercent',
  icon = 'Icon',
  custom = 'Custom',
}

enum FilterOperator {
  and = 'And',
  or = 'Or',
}

enum FilterOperator2 {
  equals = 'Equals',
  greaterThan = 'GreaterThan',
  greaterThanOrEqualTo = 'GreaterThanOrEqualTo',
  lessThan = 'LessThan',
  lessThanOrEqualTo = 'LessThanOrEqualTo',
  notEqual = 'NotEqual',
  beginsWith = 'BeginsWith',
  endsWith = 'EndsWith',
  contains = 'Contains',
  doesNotContain = 'DoesNotContain',
}

enum DynamicFilterCriteria {
  unknown = 'Unknown',
  aboveAverage = 'AboveAverage',
  allDatesInPeriodApril = 'AllDatesInPeriodApril',
  allDatesInPeriodAugust = 'AllDatesInPeriodAugust',
  allDatesInPeriodDecember = 'AllDatesInPeriodDecember',
  allDatesInPeriodFebruary = 'AllDatesInPeriodFebruary',
  allDatesInPeriodJanuary = 'AllDatesInPeriodJanuary',
  allDatesInPeriodJuly = 'AllDatesInPeriodJuly',
  allDatesInPeriodJune = 'AllDatesInPeriodJune',
  allDatesInPeriodMarch = 'AllDatesInPeriodMarch',
  allDatesInPeriodMay = 'AllDatesInPeriodMay',
  allDatesInPeriodNovember = 'AllDatesInPeriodNovember',
  allDatesInPeriodOctober = 'AllDatesInPeriodOctober',
  allDatesInPeriodQuarter1 = 'AllDatesInPeriodQuarter1',
  allDatesInPeriodQuarter2 = 'AllDatesInPeriodQuarter2',
  allDatesInPeriodQuarter3 = 'AllDatesInPeriodQuarter3',
  allDatesInPeriodQuarter4 = 'AllDatesInPeriodQuarter4',
  allDatesInPeriodSeptember = 'AllDatesInPeriodSeptember',
  belowAverage = 'BelowAverage',
  lastMonth = 'LastMonth',
  lastQuarter = 'LastQuarter',
  lastWeek = 'LastWeek',
  lastYear = 'LastYear',
  nextMonth = 'NextMonth',
  nextQuarter = 'NextQuarter',
  nextWeek = 'NextWeek',
  nextYear = 'NextYear',
  thisMonth = 'ThisMonth',
  thisQuarter = 'ThisQuarter',
  thisWeek = 'ThisWeek',
  thisYear = 'ThisYear',
  today = 'Today',
  tomorrow = 'Tomorrow',
  yearToDate = 'YearToDate',
  yesterday = 'Yesterday',
}

enum FilterDatetimeSpecificity {
  year = 'Year',
  month = 'Month',
  day = 'Day',
  hour = 'Hour',
  minute = 'Minute',
  second = 'Second',
}

// ================================
// SORTING ENUMS
// ================================
enum SortOrientation {
  rows = 'Rows',
  columns = 'Columns',
}

enum SortMethod {
  pinYin = 'PinYin',
  strokeCount = 'StrokeCount',
}

enum SortOn {
  value = 'Value',
  cellColor = 'CellColor',
  fontColor = 'FontColor',
  icon = 'Icon',
}

enum SortDataOption {
  normal = 'Normal',
  textAsNumber = 'TextAsNumber',
}

enum SortBy {
  ascending = 'Ascending',
  descending = 'Descending',
}

// ================================
// PIVOT TABLE ENUMS
// ================================
enum PivotLayoutType {
  compact = 'Compact',
  tabular = 'Tabular',
  outline = 'Outline',
}

enum SubtotalLocationType {
  atTop = 'AtTop',
  atBottom = 'AtBottom',
  off = 'Off',
}

enum AggregationFunction {
  automatic = 'Automatic',
  sum = 'Sum',
  count = 'Count',
  average = 'Average',
  max = 'Max',
  min = 'Min',
  product = 'Product',
  countNumbers = 'CountNumbers',
  standardDeviation = 'StandardDeviation',
  standardDeviationP = 'StandardDeviationP',
  variance = 'Variance',
  varianceP = 'VarianceP',
  unknown = 'Unknown',
}

enum ShowAsCalculation {
  none = 'None',
  percentOfGrandTotal = 'PercentOfGrandTotal',
  percentOfColumnTotal = 'PercentOfColumnTotal',
  percentOfRowTotal = 'PercentOfRowTotal',
  percentOf = 'PercentOf',
  percentOfParentRowTotal = 'PercentOfParentRowTotal',
  percentOfParentColumnTotal = 'PercentOfParentColumnTotal',
  percentOfParentTotal = 'PercentOfParentTotal',
  percentRunningTotal = 'PercentRunningTotal',
  rankAscending = 'RankAscending',
  rankDescending = 'RankDescending',
  differenceFrom = 'DifferenceFrom',
  percentDifferenceFrom = 'PercentDifferenceFrom',
  runningTotal = 'RunningTotal',
  index = 'Index',
}

enum PivotFilterType {
  unknown = 'Unknown',
  value = 'Value',
  manual = 'Manual',
  label = 'Label',
  date = 'Date',
}

enum DateFilterCondition {
  unknown = 'Unknown',
  equals = 'Equals',
  before = 'Before',
  after = 'After',
  between = 'Between',
  tomorrow = 'Tomorrow',
  today = 'Today',
  yesterday = 'Yesterday',
  nextWeek = 'NextWeek',
  thisWeek = 'ThisWeek',
  lastWeek = 'LastWeek',
  nextMonth = 'NextMonth',
  thisMonth = 'ThisMonth',
  lastMonth = 'LastMonth',
  nextQuarter = 'NextQuarter',
  thisQuarter = 'ThisQuarter',
  lastQuarter = 'LastQuarter',
  nextYear = 'NextYear',
  thisYear = 'ThisYear',
  lastYear = 'LastYear',
  yearToDate = 'YearToDate',
  allDatesInPeriodQuarter1 = 'AllDatesInPeriodQuarter1',
  allDatesInPeriodQuarter2 = 'AllDatesInPeriodQuarter2',
  allDatesInPeriodQuarter3 = 'AllDatesInPeriodQuarter3',
  allDatesInPeriodQuarter4 = 'AllDatesInPeriodQuarter4',
  allDatesInPeriodJanuary = 'AllDatesInPeriodJanuary',
  allDatesInPeriodFebruary = 'AllDatesInPeriodFebruary',
  allDatesInPeriodMarch = 'AllDatesInPeriodMarch',
  allDatesInPeriodApril = 'AllDatesInPeriodApril',
  allDatesInPeriodMay = 'AllDatesInPeriodMay',
  allDatesInPeriodJune = 'AllDatesInPeriodJune',
  allDatesInPeriodJuly = 'AllDatesInPeriodJuly',
  allDatesInPeriodAugust = 'AllDatesInPeriodAugust',
  allDatesInPeriodSeptember = 'AllDatesInPeriodSeptember',
  allDatesInPeriodOctober = 'AllDatesInPeriodOctober',
  allDatesInPeriodNovember = 'AllDatesInPeriodNovember',
  allDatesInPeriodDecember = 'AllDatesInPeriodDecember',
}

enum LabelFilterCondition {
  unknown = 'Unknown',
  equals = 'Equals',
  beginsWith = 'BeginsWith',
  endsWith = 'EndsWith',
  contains = 'Contains',
  between = 'Between',
  greaterThan = 'GreaterThan',
  greaterThanOrEqualTo = 'GreaterThanOrEqualTo',
  lessThan = 'LessThan',
  lessThanOrEqualTo = 'LessThanOrEqualTo',
}

enum ValueFilterCondition {
  unknown = 'Unknown',
  equals = 'Equals',
  greaterThan = 'GreaterThan',
  greaterThanOrEqualTo = 'GreaterThanOrEqualTo',
  lessThan = 'LessThan',
  lessThanOrEqualTo = 'LessThanOrEqualTo',
  between = 'Between',
  topN = 'TopN',
  bottomN = 'BottomN',
}

enum PivotFilterComparator {
  equals = 'Equals',
  doesNotEqual = 'DoesNotEqual',
  beginsWith = 'BeginsWith',
  doesNotBeginWith = 'DoesNotBeginWith',
  endsWith = 'EndsWith',
  doesNotEndWith = 'DoesNotEndWith',
  contains = 'Contains',
  doesNotContain = 'DoesNotContain',
  greaterThan = 'GreaterThan',
  greaterThanOrEqualTo = 'GreaterThanOrEqualTo',
  lessThan = 'LessThan',
  lessThanOrEqualTo = 'LessThanOrEqualTo',
  between = 'Between',
  notBetween = 'NotBetween',
}

enum PivotValueSelectionType {
  item = 'Item',
  percent = 'Percent',
}

// ================================
// DATA VALIDATION ENUMS
// ================================
enum DataValidationOperator {
  between = 'Between',
  notBetween = 'NotBetween',
  equalTo = 'EqualTo',
  notEqualTo = 'NotEqualTo',
  greaterThan = 'GreaterThan',
  lessThan = 'LessThan',
  greaterThanOrEqualTo = 'GreaterThanOrEqualTo',
  lessThanOrEqualTo = 'LessThanOrEqualTo',
}

enum DataValidationAlertStyle {
  stop = 'Stop',
  warning = 'Warning',
  information = 'Information',
}

enum DataValidationType {
  none = 'None',
  wholeNumber = 'WholeNumber',
  decimal = 'Decimal',
  list = 'List',
  date = 'Date',
  time = 'Time',
  textLength = 'TextLength',
  custom = 'Custom',
  inconsistentFormula = 'InconsistentFormula',
  inconsistentTable = 'InconsistentTable',
  mixedCriteria = 'MixedCriteria',
}

// ================================
// APPLICATION AND CALCULATION ENUMS
// ================================
enum CalculationMode {
  automatic = 'Automatic',
  automaticExceptTables = 'AutomaticExceptTables',
  manual = 'Manual',
}

enum CalculationState {
  done = 'Done',
  calculating = 'Calculating',
  pending = 'Pending',
}

enum CalculationType {
  recalculate = 'Recalculate',
  full = 'Full',
  fullRebuild = 'FullRebuild',
}

// ================================
// PAGE LAYOUT ENUMS
// ================================
enum PageOrientation {
  portrait = 'Portrait',
  landscape = 'Landscape',
}

enum PaperType {
  letter = 'Letter',
  letterSmall = 'LetterSmall',
  tabloid = 'Tabloid',
  ledger = 'Ledger',
  legal = 'Legal',
  statement = 'Statement',
  executive = 'Executive',
  a3 = 'A3',
  a4 = 'A4',
  a4Small = 'A4Small',
  a5 = 'A5',
  b4 = 'B4',
  b5 = 'B5',
  folio = 'Folio',
  quatro = 'Quatro',
  paper10x14 = 'Paper10x14',
  paper11x17 = 'Paper11x17',
  note = 'Note',
  envelope9 = 'Envelope9',
  envelope10 = 'Envelope10',
  envelope11 = 'Envelope11',
  envelope12 = 'Envelope12',
  envelope14 = 'Envelope14',
  cSheet = 'CSheet',
  dSheet = 'DSheet',
  eSheet = 'ESheet',
  envelopeDL = 'EnvelopeDL',
  envelopeC5 = 'EnvelopeC5',
  envelopeC3 = 'EnvelopeC3',
  envelopeC4 = 'EnvelopeC4',
  envelopeC6 = 'EnvelopeC6',
  envelopeC65 = 'EnvelopeC65',
  envelopeB4 = 'EnvelopeB4',
  envelopeB5 = 'EnvelopeB5',
  envelopeB6 = 'EnvelopeB6',
  envelopeItaly = 'EnvelopeItaly',
  envelopeMonarch = 'EnvelopeMonarch',
  envelopePersonal = 'EnvelopePersonal',
  fanfoldUS = 'FanfoldUS',
  fanfoldStdGerman = 'FanfoldStdGerman',
  fanfoldLegalGerman = 'FanfoldLegalGerman',
}

enum PrintOrder {
  downThenOver = 'DownThenOver',
  overThenDown = 'OverThenDown',
}

enum PrintComments {
  noComments = 'NoComments',
  endSheet = 'EndSheet',
  inPlace = 'InPlace',
}

enum PrintErrorType {
  asDisplayed = 'AsDisplayed',
  blank = 'Blank',
  dash = 'Dash',
  notAvailable = 'NotAvailable',
}

// ================================
// NAMED ITEM ENUMS
// ================================
enum NamedItemType {
  string = 'String',
  integer = 'Integer',
  double = 'Double',
  boolean = 'Boolean',
  range = 'Range',
  error = 'Error',
  array = 'Array',
}

enum NamedItemScope {
  worksheet = 'Worksheet',
  workbook = 'Workbook',
}

// ================================
// WORKSHEET ENUMS
// ================================
enum WorksheetPositionType {
  none = 'None',
  before = 'Before',
  after = 'After',
  beginning = 'Beginning',
  end = 'End',
}

// ================================
// PROTECTION ENUMS
// ================================
enum ProtectionSelectionMode {
  normal = 'Normal',
  unlocked = 'Unlocked',
  none = 'None',
}

// ================================
// BINDING ENUMS
// ================================
enum BindingType {
  range = 'Range',
  table = 'Table',
  text = 'Text',
}

// ================================
// LINKED WORKBOOK ENUMS
// ================================
enum LinkedWorkbookRefreshMode {
  file = 'File',
  prompt = 'Prompt',
  never = 'Never',
}

enum DataSourceType {
  unknown = 'Unknown',
  external = 'External',
  consolidation = 'Consolidation',
  scenario = 'Scenario',
}

// ================================
// DOCUMENT PROPERTY ENUMS
// ================================
enum DocumentPropertyType {
  number = 'Number',
  boolean = 'Boolean',
  date = 'Date',
  string = 'String',
  float = 'Float',
}

// End of ExcelScript namespace

// ================================
// MAIN FUNCTION TYPE
// ================================

/**
 * Global main function that must be present in every Office Script
 */
declare function main(workbook: ExcelScript.Workbook): void | Promise<void>;

// ================================
// MISSING GLOBAL OFFICESCRIPT NAMESPACE
// ================================
declare namespace OfficeScript {
  // Global functions
  function sendMail(mailProperties: MailProperties): void;
  function saveCopyAs(filename: string): void;
  function convertToPdf(): string;
  function downloadFile(options: { name: string; content: string }): void;

  // Metadata namespace
  namespace Metadata {
    function getScriptName(): string;
  }

  // Email interfaces
  interface MailProperties {
    to?: string | string[];
    cc?: string | string[];
    bcc?: string | string[];
    subject?: string;
    content?: string;
    contentType?: EmailContentType;
    importance?: EmailImportance;
    attachments?: EmailAttachment | EmailAttachment[];
  }

  interface EmailAttachment {
    name: string;
    content: string;
  }

  enum EmailContentType {
    text = 'text',
    html = 'html',
  }

  enum EmailImportance {
    low = 'low',
    normal = 'normal',
    high = 'high',
  }
}

// ================================
// MISSING WORKBOOK INTERFACE METHODS
// ================================
declare namespace ExcelScript {
  interface Workbook {
    // Additional missing methods
    getChartDataPointTrack(): boolean;
    setChartDataPointTrack(chartDataPointTrack: boolean): void;
    getCalculationEngineVersion(): number;
    breakAllLinksToLinkedWorkbooks(): void;

    // Refresh methods
    refreshAllDataConnections(): void;
    refreshAllPivotTables(): void;
  }

  // ================================
  // MISSING WORKSHEET INTERFACE METHODS
  // ================================
  interface Worksheet {
    // Page break methods
    addHorizontalPageBreak(pageBreakRange: Range | string): PageBreak;
    addVerticalPageBreak(pageBreakRange: Range | string): PageBreak;
    getHorizontalPageBreaks(): PageBreak[];
    getVerticalPageBreaks(): PageBreak[];

    // Freeze panes
    getFreezePanes(): WorksheetFreezePanes;

    // Additional view methods
    getEnableCalculation(): boolean;
    setEnableCalculation(enableCalculation: boolean): void;
    getEnableFormatConditionsCalculation(): boolean;
    setEnableFormatConditionsCalculation(
      enableFormatConditionsCalculation: boolean
    ): void;

    // Selection methods
    getRangeByRectangle(
      topLeftCell: string | Range,
      bottomRightCell: string | Range
    ): Range;

    // Advanced operations
    replaceAll(
      text: string,
      replacement: string,
      criteria: ReplaceCriteria
    ): number;
    findAll(text: string, criteria: WorksheetSearchCriteria): RangeAreas;
  }

  // ================================
  // MISSING RANGE INTERFACE METHODS
  // ================================
  interface Range {
    // Dependency tracking
    getPrecedents(): WorkbookRangeAreas;
    getDirectPrecedents(): WorkbookRangeAreas;
    getDependents(): WorkbookRangeAreas;
    getDirectDependents(): WorkbookRangeAreas;

    // Additional methods
    setDirty(): void;
    getSurroundingRegion(): Range;

    // Advanced selection
    getRangeEdge(direction: KeyboardDirection, activeCell?: Range): Range;
    getExtendedRange(direction: KeyboardDirection, activeCell?: Range): Range;

    // Additional value methods
    getValueType(): RangeValueType;
    getValueTypes(): RangeValueType[][];

    // Advanced formula methods
    getFormulaLocal(): string;
    setFormulaLocal(formulaLocal: string): void;
    getFormulasLocal(): string[][];
    setFormulasLocal(formulasLocal: string[][]): void;
    getFormulaR1C1(): string;
    setFormulaR1C1(formulaR1C1: string): void;
    getFormulasR1C1(): string[][];
    setFormulasR1C1(formulasR1C1: string[][]): void;

    // Number format category
    getNumberFormatCategory(): string;
    getNumberFormatCategories(): string[][];
  }

  // ================================
  // MISSING TABLE INTERFACE METHODS
  // ================================
  interface Table {
    // Additional table operations
    resize(newRange: Range | string): void;
    reapplyFilters(): void;
    clearFilters(): void;

    // Table style methods
    getTableStyle(): TableStyle;
    setTableStyle(tableStyle: TableStyle): void;
  }

  // ================================
  // MISSING CHART INTERFACE METHODS
  // ================================
  interface Chart {
    // Additional chart methods
    getSeriesNameLevel(): number;
    setSeriesNameLevel(seriesNameLevel: number): void;
    getCategoryLabelLevel(): number;
    setCategoryLabelLevel(categoryLabelLevel: number): void;

    // Chart data methods
    setData(sourceData: Range, seriesBy?: ChartSeriesBy): void;
    getPlotBy(): ChartPlotBy;
    setPlotBy(plotBy: ChartPlotBy): void;
    getPlotVisibleOnly(): boolean;
    setPlotVisibleOnly(plotVisibleOnly: boolean): void;
    getDisplayBlanksAs(): ChartDisplayBlanksAs;
    setDisplayBlanksAs(displayBlanksAs: ChartDisplayBlanksAs): void;

    // Chart options
    getShowAllFieldButtons(): boolean;
    setShowAllFieldButtons(showAllFieldButtons: boolean): void;
    getShowDataLabelsOverMaximum(): boolean;
    setShowDataLabelsOverMaximum(showDataLabelsOverMaximum: boolean): void;
  }

  // ================================
  // MISSING SUPPORT INTERFACES
  // ================================
  interface WorksheetFreezePanes {
    freezeAt(frozenRange: Range): void;
    freezeColumns(count: number): void;
    freezeRows(count: number): void;
    getLocation(): Range;
    unfreeze(): void;
  }

  interface PageBreak {
    getColumnIndex(): number;
    getRowIndex(): number;
    delete(): void;
  }

  interface LinkedWorkbook {
    getLinkedWorkbookUrl(): string;
    breakAllLinks(): void;
    refreshAllLinks(): void;
    getRefreshDate(): Date;
    getRefreshMode(): LinkedWorkbookRefreshMode;
    setRefreshMode(refreshMode: LinkedWorkbookRefreshMode): void;
  }

  interface Query {
    getName(): string;
    getRefreshDate(): Date;
    getRowsLoadedCount(): number;
    refresh(): void;
    delete(): void;
    getConnection(): WorkbookConnection;
  }

  interface WorkbookConnection {
    getName(): string;
    getDescription(): string;
    delete(): void;
    refresh(): void;
    refreshWithDisplayAlerts(displayAlerts: boolean): void;
  }

  interface CustomXmlPart {
    delete(): void;
    getId(): string;
    getNamespaceUri(): string;
    getXml(): string;
    setXml(xml: string): void;
  }

  interface Binding {
    getId(): string;
    getType(): BindingType;
    delete(): void;
    getRange(): Range;
    getText(): string;
    getTable(): Table;
    getMatrix(): any[][];
  }

  // Additional chart interfaces
  interface ChartPlotArea {
    getFormat(): ChartPlotAreaFormat;
    getHeight(): number;
    setHeight(height: number): void;
    getWidth(): number;
    setWidth(width: number): void;
    getTop(): number;
    setTop(top: number): void;
    getLeft(): number;
    setLeft(left: number): void;
    getInsideHeight(): number;
    getInsideWidth(): number;
    getInsideTop(): number;
    getInsideLeft(): number;
  }

  interface ChartPlotAreaFormat {
    getBorder(): ChartBorder;
    getFill(): ChartFill;
  }

  interface ChartPivotOptions {
    getShowAxisFieldButtons(): boolean;
    setShowAxisFieldButtons(showAxisFieldButtons: boolean): void;
    getShowLegendFieldButtons(): boolean;
    setShowLegendFieldButtons(showLegendFieldButtons: boolean): void;
    getShowReportFilterFieldButtons(): boolean;
    setShowReportFilterFieldButtons(
      showReportFilterFieldButtons: boolean
    ): void;
    getShowValueFieldButtons(): boolean;
    setShowValueFieldButtons(showValueFieldButtons: boolean): void;
  }

  // Timeline interfaces
  interface Timeline {
    getName(): string;
    setName(name: string): void;
    delete(): void;
  }

  interface TimelineStyle {
    getName(): string;
    getReadOnly(): boolean;
    setName(name: string): void;
    duplicate(): TimelineStyle;
    delete(): void;
  }

  // Additional formatting interfaces
  interface ConditionalDataBarPositiveColorFormat {
    getBorderColor(): string;
    setBorderColor(borderColor: string): void;
    getFillColor(): string;
    setFillColor(fillColor: string): void;
    getGradientFill(): boolean;
    setGradientFill(gradientFill: boolean): void;
  }

  interface ConditionalDataBarNegativeColorFormat {
    getBorderColor(): string;
    setBorderColor(borderColor: string): void;
    getFillColor(): string;
    setFillColor(fillColor: string): void;
    getMatchPositiveBorderColor(): boolean;
    setMatchPositiveBorderColor(matchPositiveBorderColor: boolean): void;
    getMatchPositiveFillColor(): boolean;
    setMatchPositiveFillColor(matchPositiveFillColor: boolean): void;
  }

  // ================================
  // MISSING ENUMS
  // ================================
  enum LinkedWorkbookRefreshMode {
    file = 'File',
    prompt = 'Prompt',
    never = 'Never',
  }

  enum DataSourceType {
    unknown = 'Unknown',
    external = 'External',
    consolidation = 'Consolidation',
    scenario = 'Scenario',
  }

  enum DocumentPropertyType {
    number = 'Number',
    boolean = 'Boolean',
    date = 'Date',
    string = 'String',
    float = 'Float',
  }

  enum BindingType {
    range = 'Range',
    table = 'Table',
    text = 'Text',
  }

  enum WorksheetPositionType {
    none = 'None',
    before = 'Before',
    after = 'After',
    beginning = 'Beginning',
    end = 'End',
  }

  enum ProtectionSelectionMode {
    normal = 'Normal',
    unlocked = 'Unlocked',
    none = 'None',
  }

  enum FilterDatetimeSpecificity {
    year = 'Year',
    month = 'Month',
    day = 'Day',
    hour = 'Hour',
    minute = 'Minute',
    second = 'Second',
  }

  // Chart-specific missing enums
  enum ChartAxisGroup {
    primary = 'Primary',
    secondary = 'Secondary',
  }

  enum ChartBinType {
    category = 'Category',
    auto = 'Auto',
    binWidth = 'BinWidth',
    binCount = 'BinCount',
  }

  enum ChartBoxQuartileCalculation {
    inclusive = 'Inclusive',
    exclusive = 'Exclusive',
  }

  enum ChartColorScheme {
    colorful = 'Colorful',
    monochromatic = 'Monochromatic',
  }

  enum ChartDataTableType {
    none = 'None',
    withLegendKeys = 'WithLegendKeys',
    withoutLegendKeys = 'WithoutLegendKeys',
  }

  enum ChartErrorBarsInclude {
    both = 'Both',
    minusValues = 'MinusValues',
    plusValues = 'PlusValues',
  }

  enum ChartErrorBarsType {
    fixedValue = 'FixedValue',
    percent = 'Percent',
    standardDeviation = 'StandardDeviation',
    standardError = 'StandardError',
    custom = 'Custom',
  }

  enum ChartGradientStyle {
    linear = 'Linear',
    radial = 'Radial',
    rectangular = 'Rectangular',
    path = 'Path',
  }

  enum ChartMapAreaLevel {
    automatic = 'Automatic',
    dataOnly = 'DataOnly',
    city = 'City',
    county = 'County',
    state = 'State',
    country = 'Country',
    continent = 'Continent',
    world = 'World',
  }

  enum ChartMapLabelStrategy {
    none = 'None',
    bestFit = 'BestFit',
    showAll = 'ShowAll',
  }

  enum ChartMapProjectionType {
    automatic = 'Automatic',
    mercator = 'Mercator',
    miller = 'Miller',
    robinson = 'Robinson',
    albers = 'Albers',
  }

  enum ChartParentLabelStrategy {
    none = 'None',
    banner = 'Banner',
    overlapping = 'Overlapping',
  }

  enum ChartPlotAreaPosition {
    automatic = 'Automatic',
    custom = 'Custom',
  }

  enum ChartSplitType {
    splitByPosition = 'SplitByPosition',
    splitByValue = 'SplitByValue',
    splitByPercentValue = 'SplitByPercentValue',
    splitByCustomSplit = 'SplitByCustomSplit',
  }

  enum ChartTickLabelPosition {
    nextToAxis = 'NextToAxis',
    high = 'High',
    low = 'Low',
    none = 'None',
  }

  enum ChartTrendlineType {
    linear = 'Linear',
    exponential = 'Exponential',
    logarithmic = 'Logarithmic',
    movingAverage = 'MovingAverage',
    polynomial = 'Polynomial',
    power = 'Power',
  }

  // Data validation missing enums
  enum DataValidationType {
    none = 'None',
    wholeNumber = 'WholeNumber',
    decimal = 'Decimal',
    list = 'List',
    date = 'Date',
    time = 'Time',
    textLength = 'TextLength',
    custom = 'Custom',
    inconsistentFormula = 'InconsistentFormula',
    inconsistentTable = 'InconsistentTable',
    mixedCriteria = 'MixedCriteria',
  }

  enum DataValidationErrorStyle {
    stop = 'Stop',
    warning = 'Warning',
    information = 'Information',
  }

  // Slicer missing enums
  enum SlicerSortType {
    dataSourceOrder = 'DataSourceOrder',
    ascending = 'Ascending',
    descending = 'Descending',
  }

  enum SlicerStyle {
    slicerStyleLight1 = 'SlicerStyleLight1',
    slicerStyleLight2 = 'SlicerStyleLight2',
    slicerStyleLight3 = 'SlicerStyleLight3',
    slicerStyleLight4 = 'SlicerStyleLight4',
    slicerStyleLight5 = 'SlicerStyleLight5',
    slicerStyleLight6 = 'SlicerStyleLight6',
    slicerStyleOther1 = 'SlicerStyleOther1',
    slicerStyleOther2 = 'SlicerStyleOther2',
    slicerStyleDark1 = 'SlicerStyleDark1',
    slicerStyleDark2 = 'SlicerStyleDark2',
    slicerStyleDark3 = 'SlicerStyleDark3',
    slicerStyleDark4 = 'SlicerStyleDark4',
    slicerStyleDark5 = 'SlicerStyleDark5',
    slicerStyleDark6 = 'SlicerStyleDark6',
  }

  // Table style missing enums
  enum TableStyleType {
    tableStyleLight = 'TableStyleLight',
    tableStyleMedium = 'TableStyleMedium',
    tableStyleDark = 'TableStyleDark',
  }

  enum PivotTableStyleType {
    pivotStyleLight = 'PivotStyleLight',
    pivotStyleMedium = 'PivotStyleMedium',
    pivotStyleDark = 'PivotStyleDark',
  }

  // Additional formatting enums
  enum RangeBorderLineStyle {
    none = 'None',
    continuous = 'Continuous',
    dash = 'Dash',
    dashDot = 'DashDot',
    dashDotDot = 'DashDotDot',
    dot = 'Dot',
    double = 'Double',
    slantDashDot = 'SlantDashDot',
  }

  enum ArrowheadLength {
    short = 'Short',
    medium = 'Medium',
    long = 'Long',
  }

  enum ArrowheadStyle {
    none = 'None',
    triangle = 'Triangle',
    stealth = 'Stealth',
    diamond = 'Diamond',
    oval = 'Oval',
    open = 'Open',
  }

  enum ArrowheadWidth {
    narrow = 'Narrow',
    medium = 'Medium',
    wide = 'Wide',
  }

  // Built-in style enum
  enum BuiltInStyle {
    normal = 'Normal',
    comma = 'Comma',
    currency = 'Currency',
    percent = 'Percent',
    wholeComma = 'WholeComma',
    wholeCurrency = 'WholeCurrency',
    hlink = 'Hlink',
    hlinkFollowed = 'HlinkFollowed',
    note = 'Note',
    warningText = 'WarningText',
    title = 'Title',
    heading1 = 'Heading1',
    heading2 = 'Heading2',
    heading3 = 'Heading3',
    heading4 = 'Heading4',
    input = 'Input',
    output = 'Output',
    calculation = 'Calculation',
    checkCell = 'CheckCell',
    linkedCell = 'LinkedCell',
    total = 'Total',
    good = 'Good',
    bad = 'Bad',
    neutral = 'Neutral',
    accent1 = 'Accent1',
    accent2 = 'Accent2',
    accent3 = 'Accent3',
    accent4 = 'Accent4',
    accent5 = 'Accent5',
    accent6 = 'Accent6',
    explanatoryText = 'ExplanatoryText',
  }

  // ================================
  // MISSING INTERFACE EXTENSIONS
  // ================================

  // Additional methods for existing interfaces that were missed
  interface Application {
    // Additional calculation methods
    getIterativeCalculation(): boolean;
    setIterativeCalculation(iterativeCalculation: boolean): void;
    getMaxIterations(): number;
    setMaxIterations(maxIterations: number): void;
    getMaxChange(): number;
    setMaxChange(maxChange: number): void;
  }

  interface WorkbookProperties {
    // Additional properties
    getCreationDate(): Date;
    getLastSaveTime(): Date;
  }

  interface Slicer {
    // Additional slicer methods
    getStyle(): string;
    setStyle(style: string): void;
    getSortType(): SlicerSortType;
    setSortType(sortType: SlicerSortType): void;
    getColumnWidth(): number;
    setColumnWidth(columnWidth: number): void;
    getNumberOfColumns(): number;
    setNumberOfColumns(numberOfColumns: number): void;
  }
}

// ================================
// FINAL MISSING COMPONENTS FOR OFFICE SCRIPTS
// ================================

declare namespace ExcelScript {
  // ================================
  // MISSING WORKBOOK METHODS
  // ================================
  interface Workbook {
    // Additional chart and slicer access
    getActiveChart(): Chart | undefined;
    getActiveSlicer(): Slicer | undefined;

    // Additional file operations
    refreshAllDataConnections(): void;
    refreshAllPivotTables(): void;

    // Advanced calculation settings
    getEnableEvents(): boolean;
    setEnableEvents(enableEvents: boolean): void;
    recalculate(): void;
  }

  // ================================
  // MISSING WORKSHEET METHODS
  // ================================
  interface Worksheet {
    // Display settings that were missed
    getEnableCalculation(): boolean;
    setEnableCalculation(enableCalculation: boolean): void;
    getEnableFormatConditionsCalculation(): boolean;
    setEnableFormatConditionsCalculation(
      enableFormatConditionsCalculation: boolean
    ): void;

    // Page break collections
    getHorizontalPageBreaks(): PageBreak[];
    getVerticalPageBreaks(): PageBreak[];

    // Advanced range selection
    getRangeByRectangle(
      topLeftCell: string | Range,
      bottomRightCell: string | Range
    ): Range;
  }

  // ================================
  // MISSING RANGE METHODS
  // ================================
  interface Range {
    // Region and edge detection methods that were missed
    getSurroundingRegion(): Range;
    getCurrentRegion(): Range;

    // Advanced calculation dependencies
    showPrecedents(): void;
    showDependents(): void;

    // Additional text and value methods
    getDisplayValue(): string;
    getDisplayValues(): string[][];

    // Advanced number formatting
    getNumberFormatCategoriesLocal(): string[][];
    getNumberFormatLocal(): string;
    getNumberFormatsLocal(): string[][];
  }

  // ================================
  // MISSING TABLE METHODS
  // ================================
  interface Table {
    // Table refresh operations
    refresh(): void;

    // Table style access
    getTableStyle(): TableStyle;
    setTableStyle(tableStyle: TableStyle | string): void;

    // Table data source
    getDataSource(): string;
  }

  // ================================
  // MISSING CHART METHODS
  // ================================
  interface Chart {
    // Chart refresh and update
    refresh(): void;
    update(): void;

    // Chart image export
    getImageAsBase64(
      format?: PictureFormat,
      width?: number,
      height?: number
    ): string;

    // Chart data range
    fullSeriesCollection(seriesIndex: number): ChartSeries;
  }

  // ================================
  // MISSING PIVOTTABLE METHODS
  // ================================
  interface PivotTable {
    // PivotTable refresh settings
    getRefreshOnOpen(): boolean;
    setRefreshOnOpen(refreshOnOpen: boolean): void;

    // PivotTable cache
    getCacheIndex(): number;

    // PivotTable layout options
    getCompactLayoutRowHeader(): string;
    setCompactLayoutRowHeader(compactLayoutRowHeader: string): void;
    getCompactLayoutColumnHeader(): string;
    setCompactLayoutColumnHeader(compactLayoutColumnHeader: string): void;

    // PivotTable data source settings
    getChangeDataSource(): boolean;
    setChangeDataSource(changeDataSource: boolean): void;
  }

  // ================================
  // MISSING SLICER METHODS
  // ================================
  interface Slicer {
    // Slicer style and formatting
    getStyle(): string;
    setStyle(style: string): void;
    getSortType(): SlicerSortType;
    setSortType(sortType: SlicerSortType): void;

    // Slicer layout
    getColumnWidth(): number;
    setColumnWidth(columnWidth: number): void;
    getNumberOfColumns(): number;
    setNumberOfColumns(numberOfColumns: number): void;
    getRowHeight(): number;
    setRowHeight(rowHeight: number): void;

    // Slicer data
    getSourceName(): string;
  }

  // ================================
  // MISSING SHAPE METHODS
  // ================================
  interface Shape {
    // Shape duplication and grouping
    duplicate(): Shape;
    group(shapes: Shape[]): ShapeGroup;
    ungroup(): Shape[];

    // Shape ordering
    bringToFront(): void;
    sendToBack(): void;
    bringForward(): void;
    sendBackward(): void;

    // Shape image methods (for image shapes)
    getPictureFormat(): PictureFormat | undefined;

    // Shape line connections
    getConnectorFormat(): ConnectorFormat | undefined;

    // Shape effects
    getShadowFormat(): ShadowFormat | undefined;
    getReflectionFormat(): ReflectionFormat | undefined;
    getGlowFormat(): GlowFormat | undefined;
  }

  // ================================
  // MISSING APPLICATION METHODS
  // ================================
  interface Application {
    // Iterative calculation settings
    getIterativeCalculation(): boolean;
    setIterativeCalculation(iterativeCalculation: boolean): void;
    getMaxIterations(): number;
    setMaxIterations(maxIterations: number): void;
    getMaxChange(): number;
    setMaxChange(maxChange: number): void;

    // Application events
    getEnableEvents(): boolean;
    setEnableEvents(enableEvents: boolean): void;

    // Screen updating
    getScreenUpdating(): boolean;
    setScreenUpdating(screenUpdating: boolean): void;

    // Automatic calculation
    getAutomaticCalculation(): boolean;
    setAutomaticCalculation(automaticCalculation: boolean): void;
  }

  // ================================
  // MISSING SUPPORT INTERFACES
  // ================================
  interface ShapeGroup {
    getId(): string;
    getShape(): Shape;
    ungroup(): Shape[];
    getGroupItems(): Shape[];
  }

  interface ConnectorFormat {
    getBeginConnected(): boolean;
    getEndConnected(): boolean;
    getBeginConnectedShape(): Shape | undefined;
    getEndConnectedShape(): Shape | undefined;
    getBeginConnectionSite(): number;
    getEndConnectionSite(): number;
  }

  interface ShadowFormat {
    getVisible(): boolean;
    setVisible(visible: boolean): void;
    getType(): ShadowType;
    setType(type: ShadowType): void;
    getColor(): string;
    setColor(color: string): void;
    getTransparency(): number;
    setTransparency(transparency: number): void;
    getSize(): number;
    setSize(size: number): void;
    getBlur(): number;
    setBlur(blur: number): void;
    getOffsetX(): number;
    setOffsetX(offsetX: number): void;
    getOffsetY(): number;
    setOffsetY(offsetY: number): void;
  }

  interface ReflectionFormat {
    getType(): ReflectionType;
    setType(type: ReflectionType): void;
    getTransparency(): number;
    setTransparency(transparency: number): void;
    getSize(): number;
    setSize(size: number): void;
    getDistance(): number;
    setDistance(distance: number): void;
    getBlur(): number;
    setBlur(blur: number): void;
  }

  interface GlowFormat {
    getRadius(): number;
    setRadius(radius: number): void;
    getColor(): string;
    setColor(color: string): void;
    getTransparency(): number;
    setTransparency(transparency: number): void;
  }

  // Enhanced WorkbookProperties interface
  interface WorkbookProperties {
    // Additional built-in properties
    getCreationDate(): Date;
    getLastSaveTime(): Date;
    getLastPrintDate(): Date;
    getTemplate(): string;
    setTemplate(template: string): void;
    getHyperlinkBase(): string;
    setHyperlinkBase(hyperlinkBase: string): void;

    // Security properties
    getSecurity(): number;
    setSecurity(security: number): void;
  }

  // Enhanced AutoFilter interface
  interface AutoFilter {
    // AutoFilter criteria management
    getCriteria(): FilterCriteria;
    setCriteria(criteria: FilterCriteria): void;

    // AutoFilter state
    getEnabled(): boolean;
    setEnabled(enabled: boolean): void;
  }

  // Enhanced Filter interface
  interface Filter {
    // Additional filter methods
    applyColorFilter(color: string): void;
    applyFillColorFilter(color: string): void;
    applyFontColorFilter(color: string): void;

    // Filter state
    getFilterType(): FilterOn;
  }

  // ================================
  // MISSING ENUMS
  // ================================
  enum ShadowType {
    outerShadow = 'OuterShadow',
    innerShadow = 'InnerShadow',
  }

  enum ReflectionType {
    none = 'None',
    full = 'Full',
    half = 'Half',
    tight = 'Tight',
    touching = 'Touching',
  }

  enum SlicerSortType {
    dataSourceOrder = 'DataSourceOrder',
    ascending = 'Ascending',
    descending = 'Descending',
  }

  // Additional chart enums that were missed
  enum ChartErrorBarsInclude {
    both = 'Both',
    minusValues = 'MinusValues',
    plusValues = 'PlusValues',
  }

  enum ChartErrorBarsType {
    fixedValue = 'FixedValue',
    percent = 'Percent',
    standardDeviation = 'StandardDeviation',
    standardError = 'StandardError',
    custom = 'Custom',
  }

  enum ChartTrendlineType {
    linear = 'Linear',
    exponential = 'Exponential',
    logarithmic = 'Logarithmic',
    movingAverage = 'MovingAverage',
    polynomial = 'Polynomial',
    power = 'Power',
  }

  enum ChartDataTableType {
    none = 'None',
    withLegendKeys = 'WithLegendKeys',
    withoutLegendKeys = 'WithoutLegendKeys',
  }

  enum ChartGradientStyle {
    linear = 'Linear',
    radial = 'Radial',
    rectangular = 'Rectangular',
    path = 'Path',
  }

  // Range selection enums
  enum RangeUnderlineStyle {
    none = 'None',
    single = 'Single',
    double = 'Double',
    singleAccountingUnderline = 'SingleAccountingUnderline',
    doubleAccountingUnderline = 'DoubleAccountingUnderline',
  }

  enum HorizontalAlignment {
    general = 'General',
    left = 'Left',
    center = 'Center',
    right = 'Right',
    fill = 'Fill',
    justify = 'Justify',
    centerAcrossSelection = 'CenterAcrossSelection',
    distributed = 'Distributed',
  }

  enum RangeVerticalAlignment {
    top = 'Top',
    center = 'Center',
    bottom = 'Bottom',
    justify = 'Justify',
    distributed = 'Distributed',
  }

  // Additional table enums
  enum TableStyleType {
    tableStyleLight = 'TableStyleLight',
    tableStyleMedium = 'TableStyleMedium',
    tableStyleDark = 'TableStyleDark',
  }

  enum PivotTableStyleType {
    pivotStyleLight = 'PivotStyleLight',
    pivotStyleMedium = 'PivotStyleMedium',
    pivotStyleDark = 'PivotStyleDark',
  }

  // Shape effect enums
  enum ShapeEffectType {
    none = 'None',
    preset = 'Preset',
    custom = 'Custom',
  }

  // ================================
  // GLOBAL TYPE ALIASES
  // ================================

  // Union types for cell values
  type CellValue = string | number | boolean;
  type CellValueRange = CellValue[][];

  // Union types for range addresses
  type RangeAddress = string | Range;

  // Union types for table references
  type TableReference = string | Table;

  // Union types for worksheet references
  type WorksheetReference = string | Worksheet;
}

// ================================
// ADDITIONAL GLOBAL FUNCTIONS
// ================================

// Built-in JavaScript Math object is available
declare const Math: Math;

// Built-in JavaScript Date object is available
declare const Date: DateConstructor;

// Built-in JavaScript Array methods are available
declare interface Array<T> {
  filter(
    predicate: (value: T, index: number, array: T[]) => unknown,
    thisArg?: any
  ): T[];
  map<U>(
    callbackfn: (value: T, index: number, array: T[]) => U,
    thisArg?: any
  ): U[];
  reduce<U>(
    callbackfn: (
      previousValue: U,
      currentValue: T,
      currentIndex: number,
      array: T[]
    ) => U,
    initialValue: U
  ): U;
  forEach(
    callbackfn: (value: T, index: number, array: T[]) => void,
    thisArg?: any
  ): void;
  find(
    predicate: (value: T, index: number, obj: T[]) => unknown,
    thisArg?: any
  ): T | undefined;
  some(
    predicate: (value: T, index: number, array: T[]) => unknown,
    thisArg?: any
  ): boolean;
  every(
    predicate: (value: T, index: number, array: T[]) => unknown,
    thisArg?: any
  ): boolean;
  sort(compareFn?: (a: T, b: T) => number): this;
  reverse(): T[];
  slice(start?: number, end?: number): T[];
  splice(start: number, deleteCount?: number, ...items: T[]): T[];
  push(...items: T[]): number;
  pop(): T | undefined;
  shift(): T | undefined;
  unshift(...items: T[]): number;
  indexOf(searchElement: T, fromIndex?: number): number;
  includes(searchElement: T, fromIndex?: number): boolean;
  join(separator?: string): string;
}

// Built-in JavaScript String methods are available
declare interface String {
  charAt(pos: number): string;
  charCodeAt(index: number): number;
  concat(...strings: string[]): string;
  indexOf(searchString: string, position?: number): number;
  lastIndexOf(searchString: string, position?: number): number;
  localeCompare(that: string): number;
  match(regexp: string | RegExp): RegExpMatchArray | null;
  replace(searchValue: string | RegExp, replaceValue: string): string;
  search(regexp: string | RegExp): number;
  slice(start?: number, end?: number): string;
  split(separator: string | RegExp, limit?: number): string[];
  substring(start: number, end?: number): string;
  toLowerCase(): string;
  toUpperCase(): string;
  trim(): string;
  length: number;
}

// Built-in JavaScript Number methods are available
declare interface Number {
  toString(radix?: number): string;
  toFixed(fractionDigits?: number): string;
  toExponential(fractionDigits?: number): string;
  toPrecision(precision?: number): string;
  valueOf(): number;
}

// Built-in JavaScript Object methods are available
declare interface Object {
  toString(): string;
  valueOf(): Object;
  hasOwnProperty(v: PropertyKey): boolean;
}

declare interface ObjectConstructor {
  keys(o: object): string[];
  values<T>(o: { [s: string]: T } | ArrayLike<T>): T[];
  entries<T>(o: { [s: string]: T } | ArrayLike<T>): [string, T][];
}

declare const Object: ObjectConstructor;

// ================================
// PERFORMANCE AND DEBUG HELPERS
// ================================

// Console object for debugging (should be removed in production)
declare const console: {
  log(...data: any[]): void;
  error(...data: any[]): void;
  warn(...data: any[]): void;
  info(...data: any[]): void;
  debug(...data: any[]): void;
  trace(...data: any[]): void;
  assert(condition?: boolean, ...data: any[]): void;
  time(label?: string): void;
  timeEnd(label?: string): void;
  count(label?: string): void;
  countReset(label?: string): void;
  clear(): void;
  dir(obj: any, options?: any): void;
  dirxml(...data: any[]): void;
  group(...data: any[]): void;
  groupCollapsed(...data: any[]): void;
  groupEnd(): void;
  table(tabularData: any, properties?: string[]): void;
};

// ================================
// ASYNC/AWAIT SUPPORT
// ================================

// Promise type for async operations
declare interface Promise<T> {
  then<TResult1 = T, TResult2 = never>(
    onfulfilled?:
      | ((value: T) => TResult1 | PromiseLike<TResult1>)
      | undefined
      | null,
    onrejected?:
      | ((reason: any) => TResult2 | PromiseLike<TResult2>)
      | undefined
      | null
  ): Promise<TResult1 | TResult2>;

  catch<TResult = never>(
    onrejected?:
      | ((reason: any) => TResult | PromiseLike<TResult>)
      | undefined
      | null
  ): Promise<T | TResult>;

  finally(onfinally?: (() => void) | undefined | null): Promise<T>;
}

declare interface PromiseConstructor {
  all<T extends readonly unknown[] | []>(
    values: T
  ): Promise<{ -readonly [P in keyof T]: Awaited<T[P]> }>;
  allSettled<T extends readonly unknown[] | []>(
    values: T
  ): Promise<{ -readonly [P in keyof T]: PromiseSettledResult<Awaited<T[P]>> }>;
  race<T extends readonly unknown[] | []>(
    values: T
  ): Promise<Awaited<T[number]>>;
  resolve(): Promise<void>;
  resolve<T>(value: T | PromiseLike<T>): Promise<T>;
  reject<T = never>(reason?: any): Promise<T>;
}

declare const Promise: PromiseConstructor;

// Fetch API for external calls (when async main function is used)
declare function fetch(
  input: string,
  init?: {
    method?: string;
    headers?: { [key: string]: string };
    body?: string;
  }
): Promise<{
  ok: boolean;
  status: number;
  statusText: string;
  headers: { get(name: string): string | null };
  text(): Promise<string>;
  json(): Promise<any>;
}>;

// ================================
// ERROR TYPES
// ================================
declare class Error {
  name: string;
  message: string;
  stack?: string;
  constructor(message?: string);
}

declare class TypeError extends Error {}
declare class RangeError extends Error {}
declare class ReferenceError extends Error {}
declare class SyntaxError extends Error {}

// ================================
// UTILITY TYPES FOR BETTER TYPE SAFETY
// ================================
declare namespace ExcelScript {
  // Utility type for optional properties
  type Optional<T, K extends keyof T> = Omit<T, K> & Partial<Pick<T, K>>;

  // Utility type for readonly properties
  type ReadOnly<T> = {
    readonly [P in keyof T]: T[P];
  };

  // Utility type for making all properties required
  type Required<T> = {
    [P in keyof T]-?: T[P];
  };

  // Utility type for deep partial
  type DeepPartial<T> = {
    [P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P];
  };
}
