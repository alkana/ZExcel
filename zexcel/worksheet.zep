namespace ZExcel;

class Worksheet implements IComparable
{
    /* Break types */
    const BREAK_NONE   = 0;
    const BREAK_ROW    = 1;
    const BREAK_COLUMN = 2;

    /* Sheet state */
    const SHEETSTATE_VISIBLE    = "visible";
    const SHEETSTATE_HIDDEN     = "hidden";
    const SHEETSTATE_VERYHIDDEN = "veryHidden";

    /**
     * Invalid characters in sheet title
     *
     * @var array
     */
    private static _invalidCharacters = ["*", ":", "/", "\\", "?", "[", "]"];

    /**
     * Parent spreadsheet
     *
     * @var PHPExcel
     */
    private parent;

    /**
     * Cacheable collection of cells
     *
     * @var \ZExcel\CachedObjectStorage\xxx
     */
    private cellCollection = null;

    /**
     * Collection of row dimensions
     *
     * @var \ZExcel\Worksheet\RowDimension[]
     */
    private rowDimensions = [];

    /**
     * Default row dimension
     *
     * @var \ZExcel\Worksheet\RowDimension
     */
    private _defaultRowDimension = null;

    /**
     * Collection of column dimensions
     *
     * @var \ZExcel\Worksheet\ColumnDimension[]
     */
    private columnDimensions = [];

    /**
     * Default column dimension
     *
     * @var \ZExcel\Worksheet\ColumnDimension
     */
    private _defaultColumnDimension = null;

    /**
     * Collection of drawings
     *
     * @var \ZExcel\Worksheet\BaseDrawing[]
     */
    private drawingCollection = null;

    /**
     * Collection of Chart objects
     *
     * @var \ZExcel\Chart[]
     */
    private chartCollection = [];

    /**
     * Worksheet title
     *
     * @var string
     */
    private title;

    /**
     * Sheet state
     *
     * @var string
     */
    private _sheetState;

    /**
     * Page setup
     *
     * @var \ZExcel\Worksheet\PageSetup
     */
    private _pageSetup;

    /**
     * Page margins
     *
     * @var \ZExcel\Worksheet\PageMargins
     */
    private _pageMargins;

    /**
     * Page header/footer
     *
     * @var \ZExcel\Worksheet\HeaderFooter
     */
    private _headerFooter;

    /**
     * Sheet view
     *
     * @var \ZExcel\Worksheet\SheetView
     */
    private _sheetView;

    /**
     * Protection
     *
     * @var \ZExcel\Worksheet\Protection
     */
    private protection;

    /**
     * Collection of styles
     *
     * @var \ZExcel\Style[]
     */
    private _styles = [];

    /**
     * Conditional styles. Indexed by cell coordinate, e.g. "A1"
     *
     * @var array
     */
    private conditionalStylesCollection = [];

    /**
     * Is the current cell collection sorted already?
     *
     * @var boolean
     */
    private cellCollectionIsSorted = false;

    /**
     * Collection of breaks
     *
     * @var array
     */
    private breaks = [];

    /**
     * Collection of merged cell ranges
     *
     * @var array
     */
    private mergeCells = [];

    /**
     * Collection of protected cell ranges
     *
     * @var array
     */
    private protectedCells = [];

    /**
     * Autofilter Range and selection
     *
     * @var \ZExcel\Worksheet\AutoFilter
     */
    private autoFilter = null;

    /**
     * Freeze pane
     *
     * @var string
     */
    private freezePane = "";

    /**
     * Show gridlines?
     *
     * @var boolean
     */
    private _showGridlines = true;

    /**
    * Print gridlines?
    *
    * @var boolean
    */
    private _printGridlines = false;

    /**
    * Show row and column headers?
    *
    * @var boolean
    */
    private _showRowColHeaders = true;

    /**
     * Show summary below? (Row/Column outline)
     *
     * @var boolean
     */
    private _showSummaryBelow = true;

    /**
     * Show summary right? (Row/Column outline)
     *
     * @var boolean
     */
    private _showSummaryRight = true;

    /**
     * Collection of comments
     *
     * @var \ZExcel\Comment[]
     */
    private comments = [];

    /**
     * Active cell. (Only one!)
     *
     * @var string
     */
    private _activeCell = "A1";

    /**
     * Selected cells
     *
     * @var string
     */
    private _selectedCells = "A1";

    /**
     * Cached highest column
     *
     * @var string
     */
    private cachedHighestColumn = "A";

    /**
     * Cached highest row
     *
     * @var int
     */
    private cachedHighestRow = 1;

    /**
     * Right-to-left?
     *
     * @var boolean
     */
    private _rightToLeft = false;

    /**
     * Hyperlinks. Indexed by cell coordinate, e.g. "A1"
     *
     * @var array
     */
    private _hyperlinkCollection = [];

    /**
     * Data validation objects. Indexed by cell coordinate, e.g. "A1"
     *
     * @var array
     */
    private _dataValidationCollection = [];

    /**
     * Tab color
     *
     * @var \ZExcel\Style\Color
     */
    private _tabColor;

    /**
     * Dirty flag
     *
     * @var boolean
     */
    private dirty    = true;

    /**
     * Hash
     *
     * @var string
     */
    private hash    = null;

    /**
    * CodeName
    *
    * @var string
    */
    private _codeName = null;

    /**
     * Create a new worksheet
     *
     * @param PHPExcel        pParent
     * @param string        pTitle
     */
    public function __construct(<\ZExcel\ZExcel> pParent = null, var pTitle = "Worksheet")
    {
        // Set parent and title
        let this->parent = pParent;
        
        this->setTitle(pTitle, false);
        
        // setTitle can change pTitle
        this->setCodeName(this->getTitle());
        this->setSheetState(\ZExcel\Worksheet::SHEETSTATE_VISIBLE);

        let this->cellCollection = \ZExcel\CachedObjectStorageFactory::getInstance(this);

        // Set page setup
        let this->_pageSetup  = new \ZExcel\Worksheet\PageSetup();

        // Set page margins
        let this->_pageMargins = new \ZExcel\Worksheet\PageMargins();

        // Set page header/footer
        let this->_headerFooter = new \ZExcel\Worksheet\HeaderFooter();

        // Set sheet view
        let this->_sheetView = new \ZExcel\Worksheet\SheetView();

        // Drawing collection
        let this->drawingCollection = new \ArrayObject();

        // Chart collection
        let this->chartCollection = new \ArrayObject();

        // Protection
        let this->protection = new \ZExcel\Worksheet\Protection();

        // Default row dimension
        let this->_defaultRowDimension = new \ZExcel\Worksheet\RowDimension(null);

        // Default column dimension
        let this->_defaultColumnDimension = new \ZExcel\Worksheet\ColumnDimension(null);

        let this->autoFilter = new \ZExcel\Worksheet\AutoFilter(null, this);
    }


    /**
     * Disconnect all cells from this \ZExcel\Worksheet object,
     *    typically so that the worksheet object can be unset
     *
     */
    public function disconnectCells()
    {
        if (this->cellCollection !== null){
            this->cellCollection->unsetWorksheetCells();
            let this->cellCollection = null;
        }
        // detach ourself from the workbook, so that it can then delete this worksheet successfully
        let this->parent = null;
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    public function __destruct()
    {
        \ZExcel\Calculation::getInstance(this->parent)->clearCalculationCacheForWorksheet(this->title);

        this->disconnectCells();
    }

   /**
     * Return the cache controller for the cell collection
     *
     * @return \ZExcel\CachedObjectStorage\ICache
     */
    public function getCellCacheController()
    {
        return this->cellCollection;
    }    //    function getCellCacheController()


    /**
     * Get array of invalid characters for sheet title
     *
     * @return array
     */
    public static function getInvalidCharacters()
    {
        return self::_invalidCharacters;
    }

    /**
     * Check sheet code name for valid Excel syntax
     *
     * @FIXME Must be private (phpunit test error)
     *
     * @param string pValue The string to check
     * @return string The valid string
     * @throws Exception
     */
    public static function _checkSheetCodeName(var pValue)
    {
        var charCount = 0, invalidCharacter;
        string checkPValue;
        
        let charCount = \ZExcel\Shared\Stringg::CountCharacters(pValue);
        
        if (charCount == 0) {
            throw new \ZExcel\Exception("Sheet code name cannot be empty.");
        }
        
        // @FIXEME Segment_fault str_replace with array
        let checkPValue = pValue;
        for invalidCharacter in self::getInvalidCharacters() {
            let checkPValue = str_replace(invalidCharacter, "", checkPValue);
        }
        
        if (pValue !== checkPValue) {
            throw new \ZExcel\Exception("Invalid character found in sheet code name");
        }
        
        if (charCount > 31) {
            throw new \ZExcel\Exception("Maximum 31 characters allowed in sheet code name.");
        }
        
        return pValue;
    }

   /**
     * Check sheet title for valid Excel syntax
     *
     * @FIXME Must be private (problem on zephir access)
     *
     * @param string pValue The string to check
     * @return string The valid string
     * @throws \ZExcel\Exception
     */
    public static function _checkSheetTitle(pValue)
    {
        if (str_replace(self::_invalidCharacters, "", pValue) !== false) {
            return new \ZExcel\Exception("Invalid character found in sheet title");
        }
        
        // Maximum 31 characters allowed for sheet title
        if (\ZExcel\Shared\Stringg::CountCharacters(pValue) > 31) {
            throw new \ZExcel\Exception("Maximum 31 characters allowed in sheet title.");
        }
        
        return pValue;
    }

    /**
     * Get collection of cells
     *
     * @param boolean pSorted Also sort the cell collection?
     * @return \ZExcel\Cell[]
     */
    public function getCellCollection(boolean pSorted = true)
    {
        if (pSorted) {
            // Re-order cell collection
            return this->sortCellCollection();
        }
        
        if (this->cellCollection !== null) {
            return this->cellCollection->getCellList();
        }
        
        return [];
    }

    /**
     * Sort collection of cells
     *
     * @return \ZExcel\Worksheet
     */
    public function sortCellCollection()
    {
        if (this->cellCollection !== null) {
            return this->cellCollection->getSortedCellList();
        }
        return [];
    }

    /**
     * Get collection of row dimensions
     *
     * @return \ZExcel\Worksheet\RowDimension[]
     */
    public function getRowDimensions()
    {
        return this->rowDimensions;
    }

    /**
     * Get default row dimension
     *
     * @return \ZExcel\Worksheet\RowDimension
     */
    public function getDefaultRowDimension()
    {
        return this->_defaultRowDimension;
    }

    /**
     * Get collection of column dimensions
     *
     * @return \ZExcel\Worksheet\ColumnDimension[]
     */
    public function getColumnDimensions()
    {
        return this->columnDimensions;
    }

    /**
     * Get default column dimension
     *
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function getDefaultColumnDimension()
    {
        return this->_defaultColumnDimension;
    }

    /**
     * Get collection of drawings
     *
     * @return \ZExcel\Worksheet\BaseDrawing[]
     */
    public function getDrawingCollection()
    {
        return this->drawingCollection;
    }

    /**
     * Get collection of charts
     *
     * @return \ZExcel\Chart[]
     */
    public function getChartCollection()
    {
        return this->chartCollection;
    }

    /**
     * Add chart
     *
     * @param \ZExcel\Chart pChart
     * @param int|null iChartIndex Index where chart should go (0,1,..., or null for last)
     * @return \ZExcel\Chart
     */
    public function addChart(<\ZExcel\Chart> pChart = null, var iChartIndex = null)
    {
        pChart->setWorksheet(this);
        
        if (is_null(iChartIndex)) {
            let this->chartCollection[] = pChart;
        } else {
            // Insert the chart at the requested index
            array_splice(this->chartCollection, iChartIndex, 0, [pChart]);
        }

        return pChart;
    }

    /**
     * Return the count of charts on this worksheet
     *
     * @return int        The number of charts
     */
    public function getChartCount()
    {
        return count(this->chartCollection);
    }

    /**
     * Get a chart by its index position
     *
     * @param string index Chart index position
     * @return false|\ZExcel\Chart
     * @throws \ZExcel\Exception
     */
    public function getChartByIndex(index = null)
    {
        var chartCount;
        
        let chartCount = count(this->chartCollection);
        
        if (chartCount == 0) {
            return false;
        }
        if (is_null(index)) {
            let index = chartCount - 1;
        }
        if (!isset(this->chartCollection[index])) {
            return false;
        }

        return this->chartCollection[index];
    }

    /**
     * Return an array of the names of charts on this worksheet
     *
     * @return string[] The names of charts
     * @throws \ZExcel\Exception
     */
    public function getChartNames()
    {
        var chart;
        array chartNames = [];
        
        for chart in this->chartCollection {
            let chartNames[] = chart->getName();
        }
        
        return chartNames;
    }

    /**
     * Get a chart by name
     *
     * @param string chartName Chart name
     * @return false|\ZExcel\Chart
     * @throws \ZExcel\Exception
     */
    public function getChartByName(var chartName = "")
    {
        var chartCount, index, chart;
        
        let chartCount = count(this->chartCollection);
        
        if (chartCount == 0) {
            return false;
        }
        
        for index, chart in this->chartCollection {
            if chart->getName() == chartName {
                return this->chartCollection[index];
            }
        }
        
        return false;
    }

    /**
     * Refresh column dimensions
     *
     * @return \ZExcel\Worksheet
     */
    public function refreshColumnDimensions()
    {
        var currentColumnDimensions, objColumnDimension;
        array newColumnDimensions = [];
        
        let currentColumnDimensions = this->getColumnDimensions();

        for objColumnDimension in currentColumnDimensions {
            let newColumnDimensions[objColumnDimension->getColumnIndex()] = objColumnDimension;
        }

        let this->columnDimensions = newColumnDimensions;

        return this;
    }

    /**
     * Refresh row dimensions
     *
     * @return \ZExcel\Worksheet
     */
    public function refreshRowDimensions()
    {
        var currentRowDimensions, objRowDimension;
        array newRowDimensions = [];
        
        let currentRowDimensions = this->getRowDimensions();

        for objRowDimension in currentRowDimensions {
            let newRowDimensions[objRowDimension->getRowIndex()] = objRowDimension;
        }

        let this->rowDimensions = newRowDimensions;

        return this;
    }

    /**
     * Calculate worksheet dimension
     *
     * @return string  String containing the dimension of this worksheet
     */
    public function calculateWorksheetDimension()
    {
        // Return
        return "A1:" . this->getHighestColumn() . this->getHighestRow();
    }

    /**
     * Calculate worksheet data dimension
     *
     * @return string  String containing the dimension of this worksheet that actually contain data
     */
    public function calculateWorksheetDataDimension()
    {
        // Return
        return "A1:" . this->getHighestDataColumn() . this->getHighestDataRow();
    }

    /**
     * Calculate widths for auto-size columns
     *
     * @param  boolean  calculateMergeCells  Calculate merge cell width
     * @return \ZExcel\Worksheet;
     */
    public function calculateColumnWidths(boolean calculateMergeCells = false)
    {
        // initialize autoSizes array
        var colDimension, cells, cell, cellID, cellValue, cellReference, columnIndex, width;
        array autoSizes = [], isMergeCell = [];
        
        for colDimension in this->getColumnDimensions() {
            if (colDimension->getAutoSize()) {
                let autoSizes[colDimension->getColumnIndex()] = -1;
            }
        }

        // There is only something to do if there are some auto-size columns
        if (!empty(autoSizes)) {
            // build list of cells references that participate in a merge
            for cells in this->getMergeCells() {
                for cellReference in \ZExcel\Cell::extractAllCellReferencesInRange(cells) {
                    let isMergeCell[cellReference] = true;
                }
            }

            // loop through all cells in the worksheet
            for cellID in this->getCellCollection(false) {
                
                let cell = this->getCell(cellID, false);
                
                if (cell !== null && isset(autoSizes[this->cellCollection->getCurrentColumn()])) {
                    // Determine width if cell does not participate in a merge
                    if (!isset(isMergeCell[this->cellCollection->getCurrentAddress()])) {
                        // Calculated value
                        // To formatted string
                        let cellValue = \ZExcel\Style\NumberFormat::toFormattedString(
                            cell->getCalculatedValue(),
                            this->getParent()->getCellXfByIndex(cell->getXfIndex())->getNumberFormat()->getFormatCode()
                        );

                        let autoSizes[this->cellCollection->getCurrentColumn()] = max(
                            (float) autoSizes[this->cellCollection->getCurrentColumn()],
                            (float) \ZExcel\Shared\Font::calculateColumnWidth(
                                this->getParent()->getCellXfByIndex(cell->getXfIndex())->getFont(),
                                cellValue,
                                this->getParent()->getCellXfByIndex(cell->getXfIndex())->getAlignment()->getTextRotation(),
                                this->getDefaultStyle()->getFont()
                            )
                        );
                    }
                }
            }

            // adjust column widths
            for columnIndex, width in autoSizes {
                if (width == -1) {
                    let width = this->getDefaultColumnDimension()->getWidth();
                }
                
                this->getColumnDimension(columnIndex)->setWidth(width);
            }
        }

        return this;
    }

    /**
     * Get parent
     *
     * @return PHPExcel
     */
    public function getParent()
    {
        return this->parent;
    }

    /**
     * Re-bind parent
     *
     * @param PHPExcel parent
     * @return \ZExcel\Worksheet
     */
    public function rebindParent(<ZExcel> parent)
    {
        var namedRanges, namedRange;
        
        if (this->parent !== null) {
            let namedRanges = this->parent->getNamedRanges();
            
            for namedRange in namedRanges {
                parent->addNamedRange(namedRange);
            }

            this->parent->removeSheetByIndex(this->parent->getIndex(this));
        }
        
        let this->parent = parent;

        return this;
    }

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle() -> string
    {
        return this->title;
    }

    /**
     * Set title
     *
     * @param string pValue String containing the dimension of this worksheet
     * @param string updateFormulaCellReferences boolean Flag indicating whether cell references in formulae should
     *            be updated to reflect the new sheet name.
     *          This should be left as the default true, unless you are
     *          certain that no formula cells on any worksheet contain
     *          references to this worksheet
     * @return \ZExcel\Worksheet
     */
    public function setTitle(var pValue = "Worksheet", var updateFormulaCellReferences = true)
    {
        var oldTitle, newTitle, altTitle, i;
        
        // Is this a "rename" or not?
        if (this->getTitle() == pValue) {
            return this;
        }
        
        self::_checkSheetTitle(pValue);
        
        let oldTitle = this->getTitle();
        
        if (is_object(this->parent) && this->parent instanceof \ZExcel\ZExcel) {
            // Is there already such sheet name?
            if (this->parent->sheetNameExists(pValue)) {
                // Use name, but append with lowest possible integer

                if (\ZExcel\Shared\Stringg::CountCharacters(pValue) > 29) {
                    let pValue = \ZExcel\Shared\Stringg::Substring(pValue,0,29);
                }
                
                let i = 1;
                
                while (this->parent->sheetNameExists(pValue . " " . i)) {
                    let i = i + 1;
                    
                    if (i == 10) {
                        if (\ZExcel\Shared\Stringg::CountCharacters(pValue) > 28) {
                            let pValue = \ZExcel\Shared\Stringg::Substring(pValue,0,28);
                        }
                    } elseif (i == 100) {
                        if (\ZExcel\Shared\Stringg::CountCharacters(pValue) > 27) {
                            let pValue = \ZExcel\Shared\Stringg::Substring(pValue,0,27);
                        }
                    }
                }

                let altTitle = pValue . " " . i;
                
                return this->setTitle(altTitle, updateFormulaCellReferences);
            }
        }
        
        let this->title = pValue;
        let this->dirty = true;
        
        if (is_object(this->parent) && this->parent instanceof \ZExcel\ZExcel) {
            let newTitle = this->getTitle();
            
            \ZExcel\Calculation::getInstance(this->parent)->renameCalculationCacheForWorksheet(oldTitle, newTitle);
            
            if (updateFormulaCellReferences == true) {
                \ZExcel\ReferenceHelper::getInstance()->updateNamedFormulas(this->parent, oldTitle, newTitle);
            }
        }
        
        return this;
    }

    /**
     * Get sheet state
     *
     * @return string Sheet state (visible, hidden, veryHidden)
     */
    public function getSheetState() {
        return this->_sheetState;
    }

    /**
     * Set sheet state
     *
     * @param string value Sheet state (visible, hidden, veryHidden)
     * @return \ZExcel\Worksheet
     */
    public function setSheetState(value = \ZExcel\Worksheet::SHEETSTATE_VISIBLE) {
        let this->_sheetState = value;
        return this;
    }

    /**
     * Get page setup
     *
     * @return \ZExcel\Worksheet\PageSetup
     */
    public function getPageSetup()
    {
        return this->_pageSetup;
    }

    /**
     * Set page setup
     *
     * @param \ZExcel\Worksheet\PageSetup    pValue
     * @return \ZExcel\Worksheet
     */
    public function setPageSetup(<\ZExcel\Worksheet\PageSetup> pValue)
    {
        let this->_pageSetup = pValue;
        return this;
    }

    /**
     * Get page margins
     *
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function getPageMargins()
    {
        return this->_pageMargins;
    }

    /**
     * Set page margins
     *
     * @param \ZExcel\Worksheet\PageMargins    pValue
     * @return \ZExcel\Worksheet
     */
    public function setPageMargins(<\ZExcel\Worksheet\PageMargins> pValue)
    {
        let this->_pageMargins = pValue;
        return this;
    }

    /**
     * Get page header/footer
     *
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function getHeaderFooter()
    {
        return this->_headerFooter;
    }

    /**
     * Set page header/footer
     *
     * @param \ZExcel\Worksheet\HeaderFooter    pValue
     * @return \ZExcel\Worksheet
     */
    public function setHeaderFooter(<\ZExcel\Worksheet\HeaderFooter> pValue)
    {
        let this->_headerFooter = pValue;
        return this;
    }

    /**
     * Get sheet view
     *
     * @return \ZExcel\Worksheet\SheetView
     */
    public function getSheetView()
    {
        return this->_sheetView;
    }

    /**
     * Set sheet view
     *
     * @param \ZExcel\Worksheet\SheetView    pValue
     * @return \ZExcel\Worksheet
     */
    public function setSheetView(<\ZExcel\Worksheet\SheetView> pValue)
    {
        let this->_sheetView = pValue;
        return this;
    }

    /**
     * Get Protection
     *
     * @return \ZExcel\Worksheet\Protection
     */
    public function getProtection()
    {
        return this->protection;
    }

    /**
     * Set Protection
     *
     * @param \ZExcel\Worksheet\Protection    pValue
     * @return \ZExcel\Worksheet
     */
    public function setProtection(<\ZExcel\Worksheet\Protection> pValue)
    {
        let this->protection = pValue;
        let this->dirty = true;

        return this;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     row        Return the data highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return string Highest column name
     */
    public function getHighestColumn(row = null)
    {
        if (row == null) {
            return this->cachedHighestColumn;
        }
        return this->getHighestDataColumn(row);
    }

    /**
     * Get highest worksheet column that contains data
     *
     * @param   string     row        Return the highest data column for the specified row,
     *                                     or the highest data column of any row if no row number is passed
     * @return string Highest column name that contains data
     */
    public function getHighestDataColumn(row = null)
    {
        return this->cellCollection->getHighestColumn(row);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     column     Return the highest data row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return int Highest row number
     */
    public function getHighestRow(column = null)
    {
        if (column == null) {
            return this->cachedHighestRow;
        }
        return this->getHighestDataRow(column);
    }

    /**
     * Get highest worksheet row that contains data
     *
     * @param   string     column     Return the highest data row for the specified column,
     *                                     or the highest data row of any column if no column letter is passed
     * @return string Highest row number that contains data
     */
    public function getHighestDataRow(column = null)
    {
        return this->cellCollection->getHighestRow(column);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        return this->cellCollection->getHighestRowAndColumn();
    }

    /**
     * Set a cell value
     *
     * @param string pCoordinate Coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param bool returnCell   Return the worksheet (false, default) or the cell (true)
     * @return \ZExcel\Worksheet|\ZExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValue(var pCoordinate = "A1", var pValue = null, boolean returnCell = false)
    {
        var cell;
        
        let cell = this->getCell(strtoupper(pCoordinate))->setValue(pValue);
        
        if (returnCell === true) {
            return cell;
        }
        
        return this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell (A = 0)
     * @param string pRow Numeric row coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return \ZExcel\Worksheet|\ZExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueByColumnAndRow(var pColumn = 0, var pRow = 1, var pValue = null, boolean returnCell = false)
    {
        var cell;
        
        let cell = this->getCellByColumnAndRow(pColumn, pRow)->setValue(pValue);
        
        if (returnCell === true) {
            return cell;
        }
        
        return this;
    }

    /**
     * Set a cell value
     *
     * @param string pCoordinate Coordinate of the cell
     * @param mixed  pValue Value of the cell
     * @param string pDataType Explicit data type
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return \ZExcel\Worksheet|\ZExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicit(var pCoordinate = "A1", var pValue = null, var pDataType = \ZExcel\Cell\DataType::TYPE_STRING, boolean returnCell = false)
    {
        var cell;
        
        let cell = this->getCell(strtoupper(pCoordinate))->setValueExplicit(pValue, pDataType);
        
        if (returnCell === true) {
            return cell;
        }
        
        return this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param string pDataType Explicit data type
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return \ZExcel\Worksheet|\ZExcel\Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicitByColumnAndRow(var pColumn = 0, var pRow = 1, var pValue = null, var pDataType = \ZExcel\Cell\DataType::TYPE_STRING, boolean returnCell = false)
    {
        var cell;
        
        let cell = this->getCellByColumnAndRow(pColumn, pRow)->setValueExplicit(pValue, pDataType);
        
        if (returnCell === true) {
            return cell;
        }
        
        return this;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param string pCoordinate    Coordinate of the cell
     * @param boolean createIfNotExists  Flag indicating whether a new cell should be created if it doesn"t
     *                                       already exist, or a null should be returned instead
     * @throws \ZExcel\Exception
     * @return \ZExcel\Cell Cell that was found
     */
    public function getCell(var pCoordinate = "A1", boolean createIfNotExists = true)
    {
        var worksheetReference, namedRange;
        
        // Check cell collection
        if (this->cellCollection->isDataSet(strtoupper(pCoordinate))) {
            return this->cellCollection->getCacheData(pCoordinate);
        }
        
        // Worksheet reference?
        if (strpos(pCoordinate, "!") !== false) {
            let worksheetReference = \ZExcel\Worksheet::extractSheetTitle(pCoordinate, true);
            
            return this->parent->getSheetByName(worksheetReference[0])->getCell(strtoupper(worksheetReference[1]), createIfNotExists);
        }

        // Named range?
        if ((!preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_CELLREF . "$/i", pCoordinate)) && (preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_NAMEDRANGE . "$/i", pCoordinate))) {
            let namedRange = \ZExcel\NamedRange::resolveRange(pCoordinate, this);
            
            if (namedRange !== null) {
                let pCoordinate = namedRange->getRange();
                
                return namedRange->getWorksheet()->getCell(pCoordinate, createIfNotExists);
            }
        }

        // Uppercase coordinate
        let pCoordinate = strtoupper(pCoordinate);

        if (strpos(pCoordinate, ":") !== false || strpos(pCoordinate, ",") !== false) {
            throw new \ZExcel\Exception("Cell coordinate can not be a range of cells.");
        } elseif (strpos(pCoordinate, "$") !== false) {
            throw new \ZExcel\Exception("Cell coordinate must not be absolute.");
        }

        // Create new cell object, if required
        return createIfNotExists ? this->createNewCell(pCoordinate) : null;
    }

    /**
     * Get cell at a specific coordinate by using numeric cell coordinates
     *
     * @param  string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @return \ZExcel\Cell Cell that was found
     */
    public function getCellByColumnAndRow(var pColumn = 0, var pRow = 1, boolean createIfNotExists = true)
    {
        var columnLetter, coordinate;
        
        let columnLetter = \ZExcel\Cell::stringFromColumnIndex(pColumn);
        let coordinate = columnLetter . pRow;

        if (this->cellCollection->isDataSet(coordinate)) {
            return this->cellCollection->getCacheData(coordinate);
        }

        // Create new cell object, if required
        if (createIfNotExists === true) {
            return this->createNewCell(coordinate);
        }
        
        return null;
    }

    /**
     * Create a new cell at the specified coordinate
     *
     * @param string pCoordinate    Coordinate of the cell
     * @return \ZExcel\Cell Cell that was created
     */
    private function createNewCell(pCoordinate)
    {
        var cell, aCoordinates, rowDimension, columnDimension;
        
        let cell = new \ZExcel\Cell(null, \ZExcel\Cell\DataType::TYPE_NULL, this);
        
        let this->cellCollectionIsSorted = false;

        // Coordinates
        let aCoordinates = \ZExcel\Cell::coordinateFromString(pCoordinate);
        
        if (\ZExcel\Cell::columnIndexFromString(this->cachedHighestColumn) < \ZExcel\Cell::columnIndexFromString(aCoordinates[0])) {
            let this->cachedHighestColumn = aCoordinates[0];
        }
        
        let this->cachedHighestRow = max(this->cachedHighestRow, aCoordinates[1]);

        // Cell needs appropriate xfIndex from dimensions records
        // but don"t create dimension records if they don"t already exist
        let rowDimension    = this->getRowDimension(aCoordinates[1], false);
        let columnDimension = this->getColumnDimension(aCoordinates[0], false);

        if (rowDimension !== null && rowDimension->getXfIndex() > 0) {
            // then there is a row dimension with explicit style, assign it to the cell
            cell->setXfIndex(rowDimension->getXfIndex());
        } else {
            if (columnDimension !== null && columnDimension->getXfIndex() > 0) {
	            // then there is a column dimension, assign it to the cell
	            cell->setXfIndex(columnDimension->getXfIndex());
	        }
        }
        
        this->cellCollection->addCacheData(pCoordinate, cell);

        return cell;
    }
    
    /**
     * Does the cell at a specific coordinate exist?
     *
     * @param string pCoordinate  Coordinate of the cell
     * @throws \ZExcel\Exception
     * @return boolean
     */
    public function cellExists(var pCoordinate = "A1")
    {
        var worksheetReference, namedRange;
        
       // Worksheet reference?
        if (strpos(pCoordinate, "!") !== false) {
            let worksheetReference = \ZExcel\Worksheet::extractSheetTitle(pCoordinate, true);
            
            return this->parent->getSheetByName(worksheetReference[0])->cellExists(strtoupper(worksheetReference[1]));
        }

        // Named range?
        if ((!preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_CELLREF . "/i", pCoordinate)) && (preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_NAMEDRANGE . "/i", pCoordinate))) {
            let namedRange = \ZExcel\NamedRange::resolveRange(pCoordinate, this);
            
            if (namedRange !== null) {
                let pCoordinate = namedRange->getRange();
                
                if (this->getHashCode() != namedRange->getWorksheet()->getHashCode()) {
                    if (!namedRange->getLocalOnly()) {
                        return namedRange->getWorksheet()->cellExists(pCoordinate);
                    } else {
                        throw new \ZExcel\Exception("Named range " . namedRange->getName() . " is not accessible from within sheet " . this->getTitle());
                    }
                }
            } else {
                return false;
            }
        }

        // Uppercase coordinate
        let pCoordinate = strtoupper(pCoordinate);

        if (strpos(pCoordinate, ":") !== false || strpos(pCoordinate, ",") !== false) {
            throw new \ZExcel\Exception("Cell coordinate can not be a range of cells.");
        } elseif (strpos(pCoordinate, "") !== false) {
            throw new \ZExcel\Exception("Cell coordinate must not be absolute.");
        } else {
            // Coordinates
            \ZExcel\Cell::coordinateFromString(pCoordinate);

            // Cell exists?
            return this->cellCollection->isDataSet(pCoordinate);
        }
    }

    /**
     * Cell at a specific coordinate by using numeric cell coordinates exists?
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @return boolean
     */
    public function cellExistsByColumnAndRow(pColumn = 0, pRow = 1)
    {
        return this->cellExists(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Get row dimension at a specific row
     *
     * @param int pRow Numeric index of the row
     * @return \ZExcel\Worksheet\RowDimension
     */
    public function getRowDimension(int pRow = 1, boolean create = true)
    {
        // Get row dimension
        if (!isset(this->rowDimensions[pRow])) {
            if (!create) {
                return null;
            }
            
            let this->rowDimensions[pRow] = new \ZExcel\Worksheet\RowDimension(pRow);

            let this->cachedHighestRow = max(this->cachedHighestRow, pRow);
        }
        
        return this->rowDimensions[pRow];
    }

    /**
     * Get column dimension at a specific column
     *
     * @param string pColumn String index of the column
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function getColumnDimension(string pColumn = "A", boolean create = true)
    {
        // Uppercase coordinate
        let pColumn = strtoupper(pColumn);

        // Fetch dimensions
        if (!isset(this->columnDimensions[pColumn])) {
            if (!create) {
                return null;
            }
            
            let this->columnDimensions[pColumn] = new \ZExcel\Worksheet\ColumnDimension(pColumn);

            if (\ZExcel\Cell::columnIndexFromString(this->cachedHighestColumn) < \ZExcel\Cell::columnIndexFromString(pColumn)) {
                let this->cachedHighestColumn = pColumn;
            }
        }
        
        return this->columnDimensions[pColumn];
    }

    /**
     * Get column dimension at a specific column by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function getColumnDimensionByColumn(pColumn = 0)
    {
        return this->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(pColumn));
    }

    /**
     * Get styles
     *
     * @return \ZExcel\Style[]
     */
    public function getStyles()
    {
        return this->_styles;
    }

    /**
     * Get default style of workbook.
     *
     * @deprecated
     * @return \ZExcel\Style
     * @throws \ZExcel\Exception
     */
    public function getDefaultStyle()
    {
        return this->parent->getDefaultStyle();
    }

    /**
     * Set default style - should only be used by \ZExcel\IReader implementations!
     *
     * @deprecated
     * @param \ZExcel\Style pValue
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function setDefaultStyle(<\ZExcel\Style> pValue) -> <\ZExcel\Worksheet>
    {
        this->parent->getDefaultStyle()->applyFromArray([
            "font": [
                "name": pValue->getFont()->getName(),
                "size": pValue->getFont()->getSize()
            ]
        ]);
        
        return this;
    }

    /**
     * Get style for cell
     *
     * @param string pCellCoordinate Cell coordinate (or range) to get style for
     * @return \ZExcel\Style
     * @throws \ZExcel\Exception
     */
    public function getStyle(pCellCoordinate = "A1")
    {
        // set this sheet as active
        this->parent->setActiveSheetIndex(this->parent->getIndex(this));

        // set cell coordinate as active
        this->setSelectedCells(strtoupper(pCellCoordinate));

        return this->parent->getCellXfSupervisor();
    }

    /**
     * Get conditional styles for a cell
     *
     * @param string pCoordinate
     * @return \ZExcel\Style\Conditional[]
     */
    public function getConditionalStyles(var pCoordinate = "A1")
    {
        let pCoordinate = strtoupper(pCoordinate);
        
        if (!isset(this->conditionalStylesCollection[pCoordinate])) {
            let this->conditionalStylesCollection[pCoordinate] = [];
        }
        
        return this->conditionalStylesCollection[pCoordinate];
    }

    /**
     * Do conditional styles exist for this cell?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function conditionalStylesExists(pCoordinate = "A1")
    {
        if (isset(this->conditionalStylesCollection[strtoupper(pCoordinate)])) {
            return true;
        }
        return false;
    }

    /**
     * Removes conditional styles for a cell
     *
     * @param string pCoordinate
     * @return \ZExcel\Worksheet
     */
    public function removeConditionalStyles(pCoordinate = "A1")
    {
        unset(this->conditionalStylesCollection[strtoupper(pCoordinate)]);
        return this;
    }

    /**
     * Get collection of conditional styles
     *
     * @return array
     */
    public function getConditionalStylesCollection()
    {
        return this->conditionalStylesCollection;
    }

    /**
     * Set conditional styles
     *
     * @param pCoordinate string E.g. "A1"
     * @param pValue \ZExcel\Style\Conditional[]
     * @return \ZExcel\Worksheet
     */
    public function setConditionalStyles(var pCoordinate = "A1", var pValue) -> <\ZExcel\Worksheet>
    {
        let this->conditionalStylesCollection[strtoupper(pCoordinate)] = pValue;
        
        return this;
    }

    /**
     * Get style for cell by using numeric cell coordinates
     *
     * @param int pColumn  Numeric column coordinate of the cell
     * @param int pRow Numeric row coordinate of the cell
     * @param int pColumn2 Numeric column coordinate of the range cell
     * @param int pRow2 Numeric row coordinate of the range cell
     * @return \ZExcel\Style
     */
    public function getStyleByColumnAndRow(var pColumn = 0, var pRow = 1, var pColumn2 = null, var pRow2 = null)
    {
        var cellRange;
        
        if (!is_null(pColumn2) && !is_null(pRow2)) {
            let cellRange = \ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow . ":" . \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2;
            
            return this->getStyle(cellRange);
        }

        return this->getStyle(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Set shared cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @deprecated
     * @param \ZExcel\Style pSharedCellStyle Cell style to share
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function setSharedStyle(<\ZExcel\Style> pSharedCellStyle = null, var pRange = "")
    {
        this->duplicateStyle(pSharedCellStyle, pRange);
        
        return this;
    }

    /**
     * Duplicate cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param \ZExcel\Style pCellStyle Cell style to duplicate
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function duplicateStyle(<\ZExcel\Style> pCellStyle = null, string pRange = "") -> <\ZExcel\Worksheet>
    {
        var style, workbook, existingStyle, xfIndex, rangeStart, rangeEnd, tmp, col, row;
        
        // make sure we have a real style and not supervisor
        let style = pCellStyle->getIsSupervisor() ? pCellStyle->getSharedComponent() : pCellStyle;

        // Add the style to the workbook if necessary
        let workbook = this->parent;
        let existingStyle = this->parent->getCellXfByHashCode(pCellStyle->getHashCode());
        if (existingStyle) {
            // there is already such cell Xf in our collection
            let xfIndex = existingStyle->getIndex();
        } else {
            // we don"t have such a cell Xf, need to add
            workbook->addCellXf(pCellStyle);
            let xfIndex = pCellStyle->getIndex();
        }

        // Calculate range outer borders
        let tmp = \ZExcel\Cell::rangeBoundaries(pRange . ":" . pRange);
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];

        // Make sure we can loop upwards on rows and columns
        if (rangeStart[0] > rangeEnd[0] && rangeStart[1] > rangeEnd[1]) {
            let tmp = rangeStart;
            let rangeStart = rangeEnd;
            let rangeEnd = tmp;
        }

        // Loop through cells and apply styles
        for col in range(rangeStart[0], rangeEnd[0]) {
            for row in range(rangeStart[1], rangeEnd[1]) {
                this->getCell(\ZExcel\Cell::stringFromColumnIndex(col - 1) . row)->setXfIndex(xfIndex);
            }
        }

        return this;
    }

    /**
     * Duplicate conditional style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param    array of \ZExcel\Style\Conditional    pCellStyle    Cell style to duplicate
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function duplicateConditionalStyle(array pCellStyle = null, string pRange = "") -> <\ZExcel\Worksheet>
    {
        var cellStyle, rangeStart, rangeEnd, tmp, col, row;
        
        for cellStyle in pCellStyle {
            if (!(cellStyle instanceof \ZExcel\Style\Conditional)) {
                throw new \ZExcel\Exception("Style is not a conditional style");
            }
        }

        // Calculate range outer borders
        let tmp = \ZExcel\Cell::rangeBoundaries(pRange . ":" . pRange);
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];

        // Make sure we can loop upwards on rows and columns
        if (rangeStart[0] > rangeEnd[0] && rangeStart[1] > rangeEnd[1]) {
            let tmp = rangeStart;
            let rangeStart = rangeEnd;
            let rangeEnd = tmp;
        }

        // Loop through cells and apply styles
        for col in range(rangeStart[0], rangeEnd[0]) {
            for row in range(rangeStart[1], rangeEnd[1]) {
                this->setConditionalStyles(\ZExcel\Cell::stringFromColumnIndex(col - 1) . row, pCellStyle);
            }
        }

        return this;
    }

    /**
     * Duplicate cell style array to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range,
     * if they are in the styles array. For example, if you decide to set a range of
     * cells to font bold, only include font bold in the styles array.
     *
     * @deprecated
     * @param array pStyles Array containing style information
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param boolean pAdvanced Advanced mode for setting borders.
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function duplicateStyleArray(var pStyles = null, string pRange = "", boolean pAdvanced = true) -> <\ZExcel\Worksheet>
    {
        this->getStyle(pRange)->applyFromArray(pStyles, pAdvanced);
        
        return this;
    }

    /**
     * Set break on a cell
     *
     * @param string pCell Cell coordinate (e.g. A1)
     * @param int pBreak Break type (type of \ZExcel\Worksheet::BREAK_*)
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function setBreak(string pCell = "A1", var pBreak = \ZExcel\Worksheet::BREAK_NONE) -> <\ZExcel\Worksheet>
    {
        // Uppercase coordinate
        let pCell = strtoupper(pCell);

        if (pCell != "") {
            if (pBreak == \ZExcel\Worksheet::BREAK_NONE) {
                if (isset(this->breaks[pCell])) {
                    unset(this->breaks[pCell]);
                }
            } else {
                let this->breaks[pCell] = pBreak;
            }
        } else {
            throw new \ZExcel\Exception("No cell coordinate specified.");
        }

        return this;
    }

    /**
     * Set break on a cell by using numeric cell coordinates
     *
     * @param integer pColumn Numeric column coordinate of the cell
     * @param integer pRow Numeric row coordinate of the cell
     * @param  integer pBreak Break type (type of \ZExcel\Worksheet::BREAK_*)
     * @return \ZExcel\Worksheet
     */
    public function setBreakByColumnAndRow(int pColumn = 0, int pRow = 1, var pBreak = \ZExcel\Worksheet::BREAK_NONE)
    {
        return this->setBreak(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow, pBreak);
    }

    /**
     * Get breaks
     *
     * @return array[]
     */
    public function getBreaks()
    {
        return this->breaks;
    }

    /**
     * Set merge on a cell range
     *
     * @param string pRange  Cell range (e.g. A1:E1)
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function mergeCells(string pRange = "A1:A1") -> <\ZExcel\Worksheet>
    {
        var aReferences, upperLeft, count, i;
        
        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        if (strpos(pRange, ":") !== false) {
            let this->mergeCells[pRange] = pRange;

            // make sure cells are created

            // get the cells in the range
            let aReferences = \ZExcel\Cell::extractAllCellReferencesInRange(pRange);

            // create upper left cell if it does not already exist
            let upperLeft = aReferences[0];
            
            if (!this->cellExists(upperLeft)) {
                this->getCell(upperLeft)->setValueExplicit(null, \ZExcel\Cell\DataType::TYPE_NULL);
            }

            // Blank out the rest of the cells in the range (if they exist)
            let count = count(aReferences);
            
            for i in range(1, count - 1) {
                if (this->cellExists(aReferences[i])) {
                    this->getCell(aReferences[i])->setValueExplicit(null, \ZExcel\Cell\DataType::TYPE_NULL);
                }
            }
        } else {
            throw new \ZExcel\Exception("Merge must be set on a range of cells.");
        }

        return this;
    }

    /**
     * Set merge on a cell range by using numeric cell coordinates
     *
     * @param int pColumn1    Numeric column coordinate of the first cell
     * @param int pRow1        Numeric row coordinate of the first cell
     * @param int pColumn2    Numeric column coordinate of the last cell
     * @param int pRow2        Numeric row coordinate of the last cell
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function mergeCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1) -> <\ZExcel\Worksheet>
    {
        var cellRange;
        
        let cellRange = \ZExcel\Cell::stringFromColumnIndex(pColumn1) . pRow1 . ":" . \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2;
        
        return this->mergeCells(cellRange);
    }

    /**
     * Remove merge on a cell range
     *
     * @param    string            pRange        Cell range (e.g. A1:E1)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function unmergeCells(string pRange = "A1:A1") -> <\ZExcel\Worksheet>
    {
        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        if (strpos(pRange, ":") !== false) {
            if (isset(this->mergeCells[pRange])) {
                unset(this->mergeCells[pRange]);
            } else {
                throw new \ZExcel\Exception("Cell range " . pRange . " not known as merged.");
            }
        } else {
            throw new \ZExcel\Exception("Merge can only be removed from a range of cells.");
        }

        return this;
    }

    /**
     * Remove merge on a cell range by using numeric cell coordinates
     *
     * @param int pColumn1    Numeric column coordinate of the first cell
     * @param int pRow1        Numeric row coordinate of the first cell
     * @param int pColumn2    Numeric column coordinate of the last cell
     * @param int pRow2        Numeric row coordinate of the last cell
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function unmergeCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1) -> <\ZExcel\Worksheet>
    {
        var cellRange;
        
        let cellRange = \ZExcel\Cell::stringFromColumnIndex(pColumn1) . pRow1 . ":" . \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2;
        
        return this->unmergeCells(cellRange);
    }

    /**
     * Get merge cells array.
     *
     * @return array[]
     */
    public function getMergeCells()
    {
        return this->mergeCells;
    }

    /**
     * Set merge cells array for the entire sheet. Use instead mergeCells() to merge
     * a single cell range.
     *
     * @param array
     */
    public function setMergeCells(array pValue = []) -> <\ZExcel\Worksheet>
    {
        let this->mergeCells = pValue;

        return this;
    }

    /**
     * Set protection on a cell range
     *
     * @param    string            pRange                Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @param    string            pPassword            Password to unlock the protection
     * @param    boolean        pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function protectCells(string pRange = "A1", string pPassword = "", boolean pAlreadyHashed = false) -> <\ZExcel\Worksheet>
    {
        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        if (!pAlreadyHashed) {
            let pPassword = \ZExcel\Shared\PasswordHasher::hashPassword(pPassword);
        }
        let this->protectedCells[pRange] = pPassword;

        return this;
    }

    /**
     * Set protection on a cell range by using numeric cell coordinates
     *
     * @param int  pColumn1            Numeric column coordinate of the first cell
     * @param int  pRow1                Numeric row coordinate of the first cell
     * @param int  pColumn2            Numeric column coordinate of the last cell
     * @param int  pRow2                Numeric row coordinate of the last cell
     * @param string pPassword            Password to unlock the protection
     * @param    boolean pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function protectCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1, string pPassword = "", boolean pAlreadyHashed = false) -> <\ZExcel\Worksheet>
    {
        var cellRange;
        
        let cellRange = \ZExcel\Cell::stringFromColumnIndex(pColumn1) . pRow1 . ":" . \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2;
        
        return this->protectCells(cellRange, pPassword, pAlreadyHashed);
    }

    /**
     * Remove protection on a cell range
     *
     * @param    string            pRange        Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function unprotectCells(string pRange = "A1") -> <\ZExcel\Worksheet>
    {
        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        if (isset(this->protectedCells[pRange])) {
            unset(this->protectedCells[pRange]);
        } else {
            throw new \ZExcel\Exception("Cell range " . pRange . " not known as protected.");
        }
        return this;
    }

    /**
     * Remove protection on a cell range by using numeric cell coordinates
     *
     * @param int  pColumn1            Numeric column coordinate of the first cell
     * @param int  pRow1                Numeric row coordinate of the first cell
     * @param int  pColumn2            Numeric column coordinate of the last cell
     * @param int pRow2                Numeric row coordinate of the last cell
     * @param string pPassword            Password to unlock the protection
     * @param    boolean pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function unprotectCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1, string pPassword = "", boolean pAlreadyHashed = false)
    {
        var cellRange;
        
        let cellRange = \ZExcel\Cell::stringFromColumnIndex(pColumn1) . pRow1 . ":" . \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2;
        
        return this->unprotectCells(cellRange);
    }

    /**
     * Get protected cells
     *
     * @return array[]
     */
    public function getProtectedCells()
    {
        return this->protectedCells;
    }

    /**
     *    Get Autofilter
     *
     *    @return \ZExcel\Worksheet\AutoFilter
     */
    public function getAutoFilter()
    {
        return this->autoFilter;
    }

    /**
     *    Set AutoFilter
     *
     *    @param    \ZExcel\Worksheet\AutoFilter|string   pValue
     *            A simple string containing a Cell range like "A1:E10" is permitted for backward compatibility
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet
     */
    public function setAutoFilter(pValue)
    {
        var pRange;
        
        let pRange = strtoupper(pValue);

        if (is_string(pValue)) {
            this->autoFilter->setRange(pValue);
        } elseif(is_object(pValue) && (pValue instanceof \ZExcel\Worksheet\AutoFilter)) {
            let this->autoFilter = pValue;
        }
        return this;
    }

    /**
     *    Set Autofilter Range by using numeric cell coordinates
     *
     *    @param  integer  pColumn1    Numeric column coordinate of the first cell
     *    @param  integer  pRow1       Numeric row coordinate of the first cell
     *    @param  integer  pColumn2    Numeric column coordinate of the second cell
     *    @param  integer  pRow2       Numeric row coordinate of the second cell
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet
     */
    public function setAutoFilterByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1)
    {
        return this->setAutoFilter(
            \ZExcel\Cell::stringFromColumnIndex(pColumn1) . pRow1
            . ":" .
            \ZExcel\Cell::stringFromColumnIndex(pColumn2) . pRow2
        );
    }

    /**
     * Remove autofilter
     *
     * @return \ZExcel\Worksheet
     */
    public function removeAutoFilter()
    {
        this->autoFilter->setRange(null);
        
        return this;
    }

    /**
     * Get Freeze Pane
     *
     * @return string
     */
    public function getFreezePane()
    {
        return this->freezePane;
    }

    /**
     * Freeze Pane
     *
     * @param    string        pCell        Cell (i.e. A2)
     *                                    Examples:
     *                                        A2 will freeze the rows above cell A2 (i.e row 1)
     *                                        B1 will freeze the columns to the left of cell B1 (i.e column A)
     *                                        B2 will freeze the rows above and to the left of cell A2
     *                                            (i.e row 1 and column A)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function freezePane(string pCell = "") -> <\ZExcel\Worksheet>
    {
        // Uppercase coordinate
        let pCell = strtoupper(pCell);
        
        if (strpos(pCell, ":") === false && strpos(pCell, ",") === false) {
            let this->freezePane = pCell;
        } else {
            throw new \ZExcel\Exception("Freeze pane can not be set on a range of cells.");
        }
        
        return this;
    }

    /**
     * Freeze Pane by using numeric cell coordinates
     *
     * @param int pColumn    Numeric column coordinate of the cell
     * @param int pRow        Numeric row coordinate of the cell
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function freezePaneByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->freezePane(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Unfreeze Pane
     *
     * @return \ZExcel\Worksheet
     */
    public function unfreezePane()
    {
        return this->freezePane("");
    }

    /**
     * Insert a new row, updating all possible related data
     *
     * @param int pBefore    Insert before this one
     * @param int pNumRows    Number of rows to insert
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function insertNewRowBefore(pBefore = 1, pNumRows = 1) -> <\ZExcel\Worksheet>
    {
        var objReferenceHelper;
        
        if (pBefore >= 1) {
            let objReferenceHelper = \ZExcel\ReferenceHelper::getInstance();
            objReferenceHelper->insertNewBefore("A" . pBefore, 0, pNumRows, this);
        } else {
            throw new \ZExcel\Exception("Rows can only be inserted before at least row 1.");
        }
        
        return this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int pBefore    Insert before this one
     * @param int pNumCols    Number of columns to insert
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function insertNewColumnBefore(string pBefore = "A", int pNumCols = 1) -> <\ZExcel\Worksheet>
    {
        var objReferenceHelper;
        
        if (!is_numeric(pBefore)) {
            let objReferenceHelper = \ZExcel\ReferenceHelper::getInstance();
            objReferenceHelper->insertNewBefore(pBefore . "1", pNumCols, 0, this);
        } else {
            throw new \ZExcel\Exception("Column references should not be numeric.");
        }
        
        return this;
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int pBefore    Insert before this one (numeric column coordinate of the cell)
     * @param int pNumCols    Number of columns to insert
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function insertNewColumnBeforeByIndex(int pBefore = 0, int pNumCols = 1)
    {
        if (pBefore >= 0) {
            return this->insertNewColumnBefore(\ZExcel\Cell::stringFromColumnIndex(pBefore), pNumCols);
        } else {
            throw new \ZExcel\Exception("Columns can only be inserted before at least column A (0).");
        }
    }

    /**
     * Delete a row, updating all possible related data
     *
     * @param int pRow        Remove starting with this one
     * @param int pNumRows    Number of rows to remove
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function removeRow(int pRow = 1, int pNumRows = 1) -> <\ZExcel\Worksheet>
    {
        var highestRow, objReferenceHelper, r;
        
        if (pRow >= 1) {
            let highestRow = this->getHighestDataRow();
            let objReferenceHelper = \ZExcel\ReferenceHelper::getInstance();
            
            objReferenceHelper->insertNewBefore("A" . (pRow + pNumRows), 0, -pNumRows, this);
            
            for r in range(0, pNumRows - 1) {
                this->getCellCacheController()->removeRow(highestRow);
                let highestRow = highestRow - 1;
            }
        } else {
            throw new \ZExcel\Exception("Rows to be deleted should at least start from row 1.");
        }
        
        return this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param string    pColumn     Remove starting with this one
     * @param int       pNumCols    Number of columns to remove
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function removeColumn(var pColumn = "A", int pNumCols = 1)
    {
        var highestColumn, objReferenceHelper, c;
        
        if (!is_numeric(pColumn)) {
            let highestColumn = this->getHighestDataColumn();
            let pColumn = \ZExcel\Cell::stringFromColumnIndex(\ZExcel\Cell::columnIndexFromString(pColumn) - 1 + pNumCols);
            let objReferenceHelper = \ZExcel\ReferenceHelper::getInstance();
            
            objReferenceHelper->insertNewBefore(pColumn . "1", -pNumCols, 0, this);
            
            for c in range(0, pNumCols - 1) {
                this->getCellCacheController()->removeColumn(highestColumn);
                let highestColumn = \ZExcel\Cell::stringFromColumnIndex(\ZExcel\Cell::columnIndexFromString(highestColumn) - 2);
            }
        } else {
            throw new \ZExcel\Exception("Column references should not be numeric.");
        }
        return this;
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int pColumn    Remove starting with this one (numeric column coordinate of the cell)
     * @param int pNumCols    Number of columns to remove
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function removeColumnByIndex(int pColumn = 0, int pNumCols = 1)
    {
        if (pColumn >= 0) {
            return this->removeColumn(\ZExcel\Cell::stringFromColumnIndex(pColumn), pNumCols);
        } else {
            throw new \ZExcel\Exception("Columns to be deleted should at least start from column 0");
        }
    }

    /**
     * Show gridlines?
     *
     * @return boolean
     */
    public function getShowGridlines()
    {
        return this->_showGridlines;
    }

    /**
     * Set show gridlines
     *
     * @param boolean pValue    Show gridlines (true/false)
     * @return \ZExcel\Worksheet
     */
    public function setShowGridlines(boolean pValue = false)
    {
        let this->_showGridlines = pValue;
        
        return this;
    }

    /**
    * Print gridlines?
    *
    * @return boolean
    */
    public function getPrintGridlines()
    {
        return this->_printGridlines;
    }

    /**
    * Set print gridlines
    *
    * @param boolean pValue Print gridlines (true/false)
    * @return \ZExcel\Worksheet
    */
    public function setPrintGridlines(boolean pValue = false)
    {
        let this->_printGridlines = pValue;
        
        return this;
    }

    /**
    * Show row and column headers?
    *
    * @return boolean
    */
    public function getShowRowColHeaders()
    {
        return this->_showRowColHeaders;
    }

    /**
    * Set show row and column headers
    *
    * @param boolean pValue Show row and column headers (true/false)
    * @return \ZExcel\Worksheet
    */
    public function setShowRowColHeaders(boolean pValue = false)
    {
        let this->_showRowColHeaders = pValue;
        
        return this;
    }

    /**
     * Show summary below? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryBelow()
    {
        return this->_showSummaryBelow;
    }

    /**
     * Set show summary below
     *
     * @param boolean pValue    Show summary below (true/false)
     * @return \ZExcel\Worksheet
     */
    public function setShowSummaryBelow(boolean pValue = true) {
        let this->_showSummaryBelow = pValue;
        
        return this;
    }

    /**
     * Show summary right? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryRight() {
        return this->_showSummaryRight;
    }

    /**
     * Set show summary right
     *
     * @param boolean pValue    Show summary right (true/false)
     * @return \ZExcel\Worksheet
     */
    public function setShowSummaryRight(boolean pValue = true)
    {
        let this->_showSummaryRight = pValue;
        
        return this;
    }

    /**
     * Get comments
     *
     * @return \ZExcel\Comment[]
     */
    public function getComments()
    {
        return this->comments;
    }

    /**
     * Set comments array for the entire sheet.
     *
     * @param array of \ZExcel\Comment
     * @return \ZExcel\Worksheet
     */
    public function setComments(array pValue = []) -> <\ZExcel\Worksheet>
    {
        let this->comments = pValue;

        return this;
    }

    /**
     * Get comment for cell
     *
     * @param string pCellCoordinate    Cell coordinate to get comment for
     * @return \ZExcel\Comment
     * @throws \ZExcel\Exception
     */
    public function getComment(string pCellCoordinate = "A1")
    {
        var newComment;
        
        // Uppercase coordinate
        let pCellCoordinate = strtoupper(pCellCoordinate);

        if (strpos(pCellCoordinate, ":") !== false || strpos(pCellCoordinate, ",") !== false) {
            throw new \ZExcel\Exception("Cell coordinate string can not be a range of cells.");
        } elseif (strpos(pCellCoordinate, "") !== false) {
            throw new \ZExcel\Exception("Cell coordinate string must not be absolute.");
        } elseif (pCellCoordinate == "") {
            throw new \ZExcel\Exception("Cell coordinate can not be zero-length string.");
        } else {
            // Check if we already have a comment for this cell.
            // If not, create a new comment.
            if (isset(this->comments[pCellCoordinate])) {
                return this->comments[pCellCoordinate];
            } else {
                let newComment = new \ZExcel\Comment();
                let this->comments[pCellCoordinate] = newComment;
                
                return newComment;
            }
        }
    }

    /**
     * Get comment for cell by using numeric cell coordinates
     *
     * @param int pColumn    Numeric column coordinate of the cell
     * @param int pRow        Numeric row coordinate of the cell
     * @return \ZExcel\Comment
     */
    public function getCommentByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->getComment(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Get selected cell
     *
     * @deprecated
     * @return string
     */
    public function getSelectedCell()
    {
        return this->getSelectedCells();
    }

    /**
     * Get active cell
     *
     * @return string Example: "A1"
     */
    public function getActiveCell()
    {
        return this->_activeCell;
    }

    /**
     * Get selected cells
     *
     * @return string
     */
    public function getSelectedCells()
    {
        return this->_selectedCells;
    }

    /**
     * Selected cell
     *
     * @param    string        pCoordinate    Cell (i.e. A1)
     * @return \ZExcel\Worksheet
     */
    public function setSelectedCell(pCoordinate = "A1")
    {
        return this->setSelectedCells(pCoordinate);
    }

    /**
     * Select a range of cells.
     *
     * @param    string        pCoordinate    Cell range, examples: "A1", "B2:G5", "A:C", "3:6"
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function setSelectedCells(string pCoordinate = "A1")
    {
        var pCoordinate, first;
        // Uppercase coordinate
        let pCoordinate = strtoupper(pCoordinate);

        // Convert "A" to "A:A"
        let pCoordinate = preg_replace("/^([A-Z]+)$/", "${1}:${1}", pCoordinate);

        // Convert "1" to "1:1"
        let pCoordinate = preg_replace("/^([0-9]+)$/", "${1}:${1}", pCoordinate);

        // Convert "A:C" to "A1:C1048576"
        let pCoordinate = preg_replace("/^([A-Z]+):([A-Z]+)$/", "${1}1:${2}1048576", pCoordinate);

        // Convert "1:3" to "A1:XFD3"
        let pCoordinate = preg_replace("/^([0-9]+):([0-9]+)$/", "A${1}:XFD${2}", pCoordinate);

        if (strpos(pCoordinate, ":") !== false || strpos(pCoordinate, ",") !== false) {
            let first = \ZExcel\Cell::splitRange(pCoordinate);
            
            let this->_activeCell = first[0][0];
        } else {
            let this->_activeCell = pCoordinate;
        }
        
        let this->_selectedCells = pCoordinate;
        
        return this;
    }

    /**
     * Selected cell by using numeric cell coordinates
     *
     * @param int pColumn Numeric column coordinate of the cell
     * @param int pRow Numeric row coordinate of the cell
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function setSelectedCellByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->setSelectedCells(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Get right-to-left
     *
     * @return boolean
     */
    public function getRightToLeft() {
        return this->_rightToLeft;
    }

    /**
     * Set right-to-left
     *
     * @param boolean value    Right-to-left true/false
     * @return \ZExcel\Worksheet
     */
    public function setRightToLeft(boolean value = false) {
        let this->_rightToLeft = value;
        
        return this;
    }

    /**
     * Fill worksheet from values in array
     *
     * @param array source Source array
     * @param mixed nullValue Value in source array that stands for blank cell
     * @param string startCell Insert array starting from this cell address as the top left coordinate
     * @param boolean strictNullComparison Apply strict comparison when testing for null values in the array
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet
     */
    public function fromArray(var source = null, var nullValue = null, string startCell = "A1", boolean strictNullComparison = false) -> <\ZExcel\Worksheet>
    {
        var startColumn, startRow, rowData, currentColumn, cellValue, tmp;
        
        if (is_array(source)) {
            //    Convert a 1-D array to 2-D (for ease of looping)
            if (!is_array(end(source))) {
                let source = [source];
            }

            // start coordinate
            let tmp = \ZEXcel\Cell::coordinateFromString(startCell);
            let startColumn = tmp[0];
            let startRow = tmp[1];

            // Loop through source
            for rowData in source {
                let currentColumn = startColumn;
                
                for cellValue in rowData {
                    if (strictNullComparison) {
                        if (cellValue !== nullValue) {
                            // Set cell value
                            this->getCell(currentColumn . startRow)->setValue(cellValue);
                        }
                    } else {
                        if (cellValue != nullValue) {
                            // Set cell value
                            this->getCell(currentColumn . startRow)->setValue(cellValue);
                        }
                    }
                    
                    let currentColumn = currentColumn + 1;
                }
                
                let startRow = startRow + 1;
            }
        } else {
            throw new \ZEXcel\Exception("Parameter $source should be an array.");
        }
        
        return this;
    }

    /**
     * Create array from a range of cells
     *
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param mixed nullValue Value returned in the array entry if a cell doesn"t exist
     * @param boolean calculateFormulas Should formulas be calculated?
     * @param boolean formatData Should formatting be applied to cell values?
     * @param boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function rangeToArray(string pRange = "A1", var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false)
    {
        var returnCellReg, rangeStart, rangeEnd, minCol, minRow, maxCol, maxRow, r, c, row, col, rRef, cRef, cell, tmp, style;
        array returnValue = [];
        
        //    Identify the range that we need to extract from the worksheet
        let tmp = \ZExcel\Cell::rangeBoundaries(pRange);
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];
        let minCol = \ZExcel\Cell::stringFromColumnIndex(rangeStart[0] - 1);
        let minRow = rangeStart[1];
        let maxCol = \ZExcel\Cell::stringFromColumnIndex(rangeEnd[0] - 1);
        let maxRow = rangeEnd[1];

        let maxCol = maxCol + 1;
        // Loop through rows
        let r = -1;
        
        for row in range(minRow, maxRow) {
            if (returnCellRef) {
                let rRef = row;
            } else {
                let r = r + 1;
                let rRef = r;
            }
            
            let c = -1;
            
            // Loop through columns in the current row
            let col = minCol;
            
            while (col != maxCol) {
                
                if (returnCellReg) {
                    let cRef = col;
                } else {
                    let c = c + 1;
                    let cRef = c;
                }
                
                //    Using getCell() will create a new cell if it doesn"t already exist. We don"t want that to happen
                //        so we test and retrieve directly against cellCollection
                if (this->cellCollection->isDataSet(col . row)) {
                    // Cell exists
                    let cell = this->cellCollection->getCacheData(col . row);
                    
                    if (cell->getValue() !== null) {
                        if (is_object(cell->getValue()) && cell->getValue() instanceof \ZExcel\RichText) {
                            let returnValue[rRef][cRef] = cell->getValue()->getPlainText();
                        } else {
                            if (calculateFormulas) {
                                let returnValue[rRef][cRef] = cell->getCalculatedValue();
                            } else {
                                let returnValue[rRef][cRef] = cell->getValue();
                            }
                        }

                        if (formatData) {
                            let style = this->parent->getCellXfByIndex(cell->getXfIndex());
                            let returnValue[rRef][cRef] = \ZExcel\Style\NumberFormat::toFormattedString(
                                returnValue[rRef][cRef],
                                (style && style->getNumberFormat()) ? style->getNumberFormat()->getFormatCode() : \ZExcel\Style\NumberFormat::FORMAT_GENERAL
                            );
                        }
                    } else {
                        // Cell holds a NULL
                        let returnValue[rRef][cRef] = nullValue;
                    }
                } else {
                    // Cell doesn"t exist
                    let returnValue[rRef][cRef] = nullValue;
                }
                
                let col = col + 1;
            }
        }

        // Return
        return returnValue;
    }


    /**
     * Create array from a range of cells
     *
     * @param  string pNamedRange Name of the Named Range
     * @param  mixed  nullValue Value returned in the array entry if a cell doesn"t exist
     * @param  boolean calculateFormulas  Should formulas be calculated?
     * @param  boolean formatData  Should formatting be applied to cell values?
     * @param  boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                                True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     * @throws \ZExcel\Exception
     */
    public function namedRangeToArray(string pNamedRange = "", var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false)
    {
        var namedRange, pWorkSheet, pCellRange;
        
        let namedRange = \ZExcel\NamedRange::resolveRange(pNamedRange, this);
        
        if (namedRange !== null) {
            let pWorkSheet = namedRange->getWorksheet();
            let pCellRange = namedRange->getRange();

            return pWorkSheet->rangeToArray(pCellRange, nullValue, calculateFormulas, formatData, returnCellRef);
        }

        throw new \ZExcel\Exception("Named Range " . pNamedRange . " does not exist.");
    }


    /**
     * Create array from worksheet
     *
     * @param mixed nullValue Value returned in the array entry if a cell doesn"t exist
     * @param boolean calculateFormulas Should formulas be calculated?
     * @param boolean formatData  Should formatting be applied to cell values?
     * @param boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function toArray(var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false)
    {
        var maxCol, maxRow;
        
        // Garbage collect...
        this->garbageCollect();

        //    Identify the range that we need to extract from the worksheet
        let maxCol = this->getHighestColumn();
        let maxRow = this->getHighestRow();
        // Return
        return this->rangeToArray("A1:" . maxCol . maxRow, nullValue, calculateFormulas, formatData, returnCellRef);
    }

    /**
     * Get row iterator
     *
     * @param   integer   startRow   The row number at which to start iterating
     * @param   integer   endRow     The row number at which to stop iterating
     *
     * @return \ZExcel\Worksheet\RowIterator
     */
    public function getRowIterator(int startRow = 1, int endRow = null)
    {
        return new \ZExcel\Worksheet\RowIterator(this, startRow, endRow);
    }

    /**
     * Get column iterator
     *
     * @param   string   startColumn The column address at which to start iterating
     * @param   string   endColumn   The column address at which to stop iterating
     *
     * @return \ZExcel\Worksheet\ColumnIterator
     */
    public function getColumnIterator(string startColumn = "A", string endColumn = null)
    {
        return new \ZExcel\Worksheet\ColumnIterator(this, startColumn, endColumn);
    }

    /**
     * Run PHPExcel garabage collector.
     *
     * @return \ZExcel\Worksheet
     */
    public function garbageCollect() -> <\ZExcel\Worksheet>
    {
        var colRow, highestRow, highestColumn, dimension;
        
        // Flush cache
        this->cellCollection->getCacheData("A1");

        // Lookup highest column and highest row if cells are cleaned
        let colRow = this->cellCollection->getHighestRowAndColumn();
        let highestRow = colRow["row"];
        let highestColumn = \ZExcel\Cell::columnIndexFromString(colRow["column"]);

        // Loop through column dimensions
        for dimension in this->columnDimensions {
            let highestColumn = max(highestColumn, \ZExcel\Cell::columnIndexFromString(dimension->getColumnIndex()));
        }

        // Loop through row dimensions
        for dimension in this->rowDimensions {
            let highestRow = max(highestRow, dimension->getRowIndex());
        }

        // Cache values
        if (highestColumn < 0) {
            let this->cachedHighestColumn = "A";
        } else {
            let highestColumn = highestColumn - 1;
            let this->cachedHighestColumn = \ZExcel\Cell::stringFromColumnIndex(highestColumn);
        }
        
        let this->cachedHighestRow = highestRow;

        // Return
        return this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        if (this->dirty) {
            let this->hash = md5(this->title . this->autoFilter . (this->protection->isProtectionEnabled() ? "t" : "f") . get_class(this));
            let this->dirty = false;
        }
        
        return this->hash;
    }

    /**
     * Extract worksheet title from range.
     *
     * Example: extractSheetTitle("testSheet!A1") ==> "A1"
     * Example: extractSheetTitle(""testSheet 1"!A1", true) ==> array("testSheet 1", "A1");
     *
     * @param string pRange    Range to extract title from
     * @param bool returnRange    Return range? (see example)
     * @return mixed
     */
    public static function extractSheetTitle(var pRange, returnRange = false)
    {
        var sep;
        
        let sep = strpos(pRange, "!");
        
        // Sheet title included?
        if (sep === false) {
            return "";
        }

        if (returnRange) {
            return [trim(substr(pRange, 0, sep), "\""), substr(pRange, sep + 1)];
        }

        return substr(pRange, sep + 1);
    }

    /**
     * Get hyperlink
     *
     * @param string pCellCoordinate    Cell coordinate to get hyperlink for
     */
    public function getHyperlink(string pCellCoordinate = "A1")
    {
        // return hyperlink if we already have one
        if (isset(this->_hyperlinkCollection[pCellCoordinate])) {
            return this->_hyperlinkCollection[pCellCoordinate];
        }

        // else create hyperlink
        let this->_hyperlinkCollection[pCellCoordinate] = new \ZExcel\Cell\Hyperlink();
        
        return this->_hyperlinkCollection[pCellCoordinate];
    }

    /**
     * Set hyperlnk
     *
     * @param string pCellCoordinate    Cell coordinate to insert hyperlink
     * @param    \ZExcel\Cell\Hyperlink    pHyperlink
     * @return \ZExcel\Worksheet
     */
    public function setHyperlink(string pCellCoordinate = "A1", <\ZExcel\Cell\Hyperlink> pHyperlink = null)
    {
        if (pHyperlink === null) {
            unset(this->_hyperlinkCollection[pCellCoordinate]);
        } else {
            let this->_hyperlinkCollection[pCellCoordinate] = pHyperlink;
        }
        
        return this;
    }

    /**
     * Hyperlink at a specific coordinate exists?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function hyperlinkExists(string pCoordinate = "A1")
    {
        return isset(this->_hyperlinkCollection[pCoordinate]);
    }

    /**
     * Get collection of hyperlinks
     *
     * @return \ZExcel\Cell\Hyperlink[]
     */
    public function getHyperlinkCollection()
    {
        return this->_hyperlinkCollection;
    }

    /**
     * Get data validation
     *
     * @param string pCellCoordinate Cell coordinate to get data validation for
     */
    public function getDataValidation(string pCellCoordinate = "A1")
    {
        // return data validation if we already have one
        if (isset(this->_dataValidationCollection[pCellCoordinate])) {
            return this->_dataValidationCollection[pCellCoordinate];
        }

        // else create data validation
        let this->_dataValidationCollection[pCellCoordinate] = new \ZExcel\Cell\DataValidation();
        
        return this->_dataValidationCollection[pCellCoordinate];
    }

    /**
     * Set data validation
     *
     * @param string pCellCoordinate    Cell coordinate to insert data validation
     * @param    \ZExcel\Cell\DataValidation    pDataValidation
     * @return \ZExcel\Worksheet
     */
    public function setDataValidation(string pCellCoordinate = "A1", <\ZExcel\Cell\DataValidation> pDataValidation = null)
    {
        if (pDataValidation === null) {
            unset(this->_dataValidationCollection[pCellCoordinate]);
        } else {
            let this->_dataValidationCollection[pCellCoordinate] = pDataValidation;
        }
        return this;
    }

    /**
     * Data validation at a specific coordinate exists?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function dataValidationExists(string pCoordinate = "A1")
    {
        return isset(this->_dataValidationCollection[pCoordinate]);
    }

    /**
     * Get collection of data validations
     *
     * @return \ZExcel\Cell\DataValidation[]
     */
    public function getDataValidationCollection()
    {
        return this->_dataValidationCollection;
    }

    /**
     * Accepts a range, returning it as a range that falls within the current highest row and column of the worksheet
     *
     * @param string range
     * @return string Adjusted range value
     */
    public function shrinkRangeToFit(var range)
    {
        var maxCol, maxRow, k, rangeSet, rangeBlocks, rangeBoundaries, stRange;
        
        let maxCol = this->getHighestColumn();
        let maxRow = this->getHighestRow();
        let maxCol = \ZExcel\Cell::columnIndexFromString(maxCol);

        let rangeBlocks = explode(" ", range);
        for k, rangeSet in rangeBlocks {
            let rangeBoundaries = \ZExcel\Cell::getRangeBoundaries(rangeSet);

            if (\ZExcel\Cell::columnIndexFromString(rangeBoundaries[0][0]) > maxCol) {
                let rangeBoundaries[0][0] = \ZExcel\Cell::stringFromColumnIndex(maxCol);
            }
            if (rangeBoundaries[0][1] > maxRow) {
                let rangeBoundaries[0][1] = maxRow;
            }
            if (\ZExcel\Cell::columnIndexFromString(rangeBoundaries[1][0]) > maxCol) {
                let rangeBoundaries[1][0] = \ZExcel\Cell::stringFromColumnIndex(maxCol);
            }
            if (rangeBoundaries[1][1] > maxRow) {
                let rangeBoundaries[1][1] = maxRow;
            }
            
            let rangeBlocks[k] = rangeBoundaries[0][0].rangeBoundaries[0][1].":".rangeBoundaries[1][0].rangeBoundaries[1][1];
        }
        
        let stRange = implode(" ", rangeBlocks);

        return stRange;
    }

    /**
     * Get tab color
     *
     * @return \ZExcel\Style\Color
     */
    public function getTabColor()
    {
        if (this->_tabColor === null) {
            let this->_tabColor = new \ZExcel\Style\Color();
        }

        return this->_tabColor;
    }

    /**
     * Reset tab color
     *
     * @return \ZExcel\Worksheet
     */
    public function resetTabColor()
    {
        let this->_tabColor = null;

        return this;
    }

    /**
     * Tab color set?
     *
     * @return boolean
     */
    public function isTabColorSet()
    {
        return (this->_tabColor !== null);
    }

    /**
     * Copy worksheet (!= clone!)
     *
     * @return \ZExcel\Worksheet
     */
    public function copy() {
        var copied;
        
        let copied = clone this;

        return copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        var vars, key, val, newCollection, newAutoFilter;
        
        let vars = get_object_vars(this);
        
        for key, val in vars {
            if (key == "parent") {
                continue;
            }

            if (is_object(val) || (is_array(val))) {
                if (key == "cellCollection") {
                    let newCollection = clone this->cellCollection;
                    
                    newCollection->copyCellCollection(this);
                    
                    let this->cellCollection = newCollection;
                } elseif (key == "drawingCollection") {
                    let newCollection = clone this->drawingCollection;
                    let this->drawingCollection = newCollection;
                } elseif ((key == "autoFilter") && is_object(this->autoFilter) && (this->autoFilter instanceof \ZExcel\Worksheet\AutoFilter)) {
                    let newAutoFilter = clone this->autoFilter;
                    let this->autoFilter = newAutoFilter;
                    this->autoFilter->setParent(this);
                } else {
                    let this->{key} = unserialize(serialize(val));
                }
            }
        }
    }

    /**
     * Define the code name of the sheet
     *
     * @param null|string Same rule as Title minus space not allowed (but, like Excel, change silently space to underscore)
     * @return objWorksheet
     * @throws \ZExcel\Exception
    */
    public function setCodeName(string pValue = null){
        var oldCodeName;
        
        let oldCodeName = this->getCodeName();
        let pValue = str_replace(" ", "_", pValue);
        
        if (oldCodeName === pValue) {
            return this;
        }
        
        self::_checkSheetCodeName(pValue);
        
        // if (isset(this->parent) && this->parent->sheetCodeNameExists(pValue)) {
        //    throw new \ZExcel\Exception("Code name already exists.");
        // }
        
        // let this->_codeName = pValue;
        
        return this;
    }
    /**
     * Return the code name of the sheet
     *
     * @return null|string
    */
    public function getCodeName(){
        return this->_codeName;
    }
    /**
     * Sheet has a code name ?
     * @return boolean
    */
    public function hasCodeName(){
        return !(is_null(this->_codeName));
    }
}
