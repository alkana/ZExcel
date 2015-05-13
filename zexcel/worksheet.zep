namespace ZExcel;

class Worksheet implements IComparable
{
	/* Break types */
    const BREAK_NONE   = 0;
    const BREAK_ROW    = 1;
    const BREAK_COLUMN = 2;

    /* Sheet state */
    const SHEETSTATE_VISIBLE    = 'visible';
    const SHEETSTATE_HIDDEN     = 'hidden';
    const SHEETSTATE_VERYHIDDEN = 'veryHidden';

    /**
     * Invalid characters in sheet title
     *
     * @var array
     */
    private static _invalidCharacters = ['*', ':', '/', '\\', '?', '[', ']'];

    /**
     * Parent spreadsheet
     *
     * @var PHPExcel
     */
    private _parent;

    /**
     * Cacheable collection of cells
     *
     * @var PHPExcel_CachedObjectStorage_xxx
     */
    private _cellCollection = null;

    /**
     * Collection of row dimensions
     *
     * @var PHPExcel_Worksheet_RowDimension[]
     */
    private _rowDimensions = [];

    /**
     * Default row dimension
     *
     * @var PHPExcel_Worksheet_RowDimension
     */
    private _defaultRowDimension = null;

    /**
     * Collection of column dimensions
     *
     * @var PHPExcel_Worksheet_ColumnDimension[]
     */
    private _columnDimensions = [];

    /**
     * Default column dimension
     *
     * @var PHPExcel_Worksheet_ColumnDimension
     */
    private _defaultColumnDimension = null;

    /**
     * Collection of drawings
     *
     * @var PHPExcel_Worksheet_BaseDrawing[]
     */
    private _drawingCollection = null;

    /**
     * Collection of Chart objects
     *
     * @var PHPExcel_Chart[]
     */
    private _chartCollection = [];

    /**
     * Worksheet title
     *
     * @var string
     */
    private _title;

    /**
     * Sheet state
     *
     * @var string
     */
    private _sheetState;

    /**
     * Page setup
     *
     * @var PHPExcel_Worksheet_PageSetup
     */
    private _pageSetup;

    /**
     * Page margins
     *
     * @var PHPExcel_Worksheet_PageMargins
     */
    private _pageMargins;

    /**
     * Page header/footer
     *
     * @var PHPExcel_Worksheet_HeaderFooter
     */
    private _headerFooter;

    /**
     * Sheet view
     *
     * @var PHPExcel_Worksheet_SheetView
     */
    private _sheetView;

    /**
     * Protection
     *
     * @var PHPExcel_Worksheet_Protection
     */
    private _protection;

    /**
     * Collection of styles
     *
     * @var PHPExcel_Style[]
     */
    private _styles = [];

    /**
     * Conditional styles. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private _conditionalStylesCollection = [];

    /**
     * Is the current cell collection sorted already?
     *
     * @var boolean
     */
    private _cellCollectionIsSorted = false;

    /**
     * Collection of breaks
     *
     * @var array
     */
    private _breaks = [];

    /**
     * Collection of merged cell ranges
     *
     * @var array
     */
    private _mergeCells = [];

    /**
     * Collection of protected cell ranges
     *
     * @var array
     */
    private _protectedCells = [];

    /**
     * Autofilter Range and selection
     *
     * @var PHPExcel_Worksheet_AutoFilter
     */
    private _autoFilter = NULL;

    /**
     * Freeze pane
     *
     * @var string
     */
    private _freezePane = '';

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
     * @var PHPExcel_Comment[]
     */
    private _comments = [];

    /**
     * Active cell. (Only one!)
     *
     * @var string
     */
    private _activeCell = 'A1';

    /**
     * Selected cells
     *
     * @var string
     */
    private _selectedCells = 'A1';

    /**
     * Cached highest column
     *
     * @var string
     */
    private _cachedHighestColumn = 'A';

    /**
     * Cached highest row
     *
     * @var int
     */
    private _cachedHighestRow = 1;

    /**
     * Right-to-left?
     *
     * @var boolean
     */
    private _rightToLeft = false;

    /**
     * Hyperlinks. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private _hyperlinkCollection = [];

    /**
     * Data validation objects. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private _dataValidationCollection = [];

    /**
     * Tab color
     *
     * @var PHPExcel_Style_Color
     */
    private _tabColor;

    /**
     * Dirty flag
     *
     * @var boolean
     */
    private _dirty    = true;

    /**
     * Hash
     *
     * @var string
     */
    private _hash    = null;

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
    public function __construct(<\ZExcel\ZExcel> pParent = null, string pTitle = 'Worksheet')
    {
        // Set parent and title
        let this->_parent = pParent;
        this->setTitle(pTitle, false);
        
        // setTitle can change pTitle
        this->setCodeName(this->getTitle());
        this->setSheetState(\ZExcel\Worksheet::SHEETSTATE_VISIBLE);

        let this->_cellCollection = \ZExcel\CachedObjectStorageFactory::getInstance(this);

        // Set page setup
        let this->_pageSetup = new \ZExcel\Worksheet\PageSetup();

        // Set page margins
        let this->_pageMargins = new \ZExcel\Worksheet\PageMargins();

        // Set page header/footer
        let this->_headerFooter = new \ZExcel\Worksheet\HeaderFooter();

        // Set sheet view
        let this->_sheetView = new \ZExcel\Worksheet\SheetView();

        // Drawing collection
        let this->_drawingCollection = new \ArrayObject();

        // Chart collection
        let this->_chartCollection = new \ArrayObject();

        // Protection
        let this->_protection = new \ZExcel\Worksheet\Protection();

        // Default row dimension
        let this->_defaultRowDimension = new \ZExcel\Worksheet\RowDimension(null);

        // Default column dimension
        let this->_defaultColumnDimension = new \ZExcel\Worksheet\ColumnDimension(null);

        let this->_autoFilter = new \ZExcel\Worksheet\AutoFilter(null, this);
    }

    /**
     * Disconnect all cells from this PHPExcel_Worksheet object,
     * typically so that the worksheet object can be unset
     *
     */
    public function disconnectCells()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    public function __destruct() {
    	var calc;
    	
        let calc = \PHPExcel\Calculation::getInstance(this->_parent);
        calc->clearCalculationCacheForWorksheet(this->_title);

        this->disconnectCells();
    }

   /**
     * Return the cache controller for the cell collection
     *
     * @return PHPExcel_CachedObjectStorage_xxx
     */
    public function getCellCacheController() {
        return this->_cellCollection;
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
     * @param string pValue The string to check
     * @return string The valid string
     * @throws Exception
     */
    private static function _checkSheetCodeName(pValue)
    {
        throw new \Exception("Not implemented yet!");
    }

   /**
     * Check sheet title for valid Excel syntax
     *
     * @param string pValue The string to check
     * @return string The valid string
     * @throws PHPExcel_Exception
     */
    private static function _checkSheetTitle(pValue)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get collection of cells
     *
     * @param boolean pSorted Also sort the cell collection?
     * @return PHPExcel_Cell[]
     */
    public function getCellCollection(pSorted = true)
    {
        if (pSorted) {
            // Re-order cell collection
            return this->sortCellCollection();
        }
        if (this->_cellCollection !== NULL) {
            return this->_cellCollection->getCellList();
        }
        return [];
    }

    /**
     * Sort collection of cells
     *
     * @return PHPExcel_Worksheet
     */
    public function sortCellCollection()
    {
        if (this->_cellCollection !== NULL) {
            return this->_cellCollection->getSortedCellList();
        }
        return [];
    }

    /**
     * Get collection of row dimensions
     *
     * @return PHPExcel_Worksheet_RowDimension[]
     */
    public function getRowDimensions()
    {
        return this->_rowDimensions;
    }

    /**
     * Get default row dimension
     *
     * @return PHPExcel_Worksheet_RowDimension
     */
    public function getDefaultRowDimension()
    {
        return this->_defaultRowDimension;
    }

    /**
     * Get collection of column dimensions
     *
     * @return PHPExcel_Worksheet_ColumnDimension[]
     */
    public function getColumnDimensions()
    {
        return this->_columnDimensions;
    }

    /**
     * Get default column dimension
     *
     * @return PHPExcel_Worksheet_ColumnDimension
     */
    public function getDefaultColumnDimension()
    {
        return this->_defaultColumnDimension;
    }

    /**
     * Get collection of drawings
     *
     * @return PHPExcel_Worksheet_BaseDrawing[]
     */
    public function getDrawingCollection()
    {
        return this->_drawingCollection;
    }

    /**
     * Get collection of charts
     *
     * @return PHPExcel_Chart[]
     */
    public function getChartCollection()
    {
        return this->_chartCollection;
    }

    /**
     * Add chart
     *
     * @param PHPExcel_Chart pChart
     * @param int|null iChartIndex Index where chart should go (0,1,..., or null for last)
     * @return PHPExcel_Chart
     */
    public function addChart(<\ZExcel\Chart> pChart = null, int iChartIndex = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Return the count of charts on this worksheet
     *
     * @return int        The number of charts
     */
    public function getChartCount()
    {
        return count(this->_chartCollection);
    }

    /**
     * Get a chart by its index position
     *
     * @param string index Chart index position
     * @return false|PHPExcel_Chart
     * @throws PHPExcel_Exception
     */
    public function getChartByIndex(index = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Return an array of the names of charts on this worksheet
     *
     * @return string[] The names of charts
     * @throws PHPExcel_Exception
     */
    public function getChartNames()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get a chart by name
     *
     * @param string chartName Chart name
     * @return false|PHPExcel_Chart
     * @throws PHPExcel_Exception
     */
    public function getChartByName(chartName = '')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Refresh column dimensions
     *
     * @return PHPExcel_Worksheet
     */
    public function refreshColumnDimensions()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Refresh row dimensions
     *
     * @return PHPExcel_Worksheet
     */
    public function refreshRowDimensions()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Calculate worksheet dimension
     *
     * @return string  String containing the dimension of this worksheet
     */
    public function calculateWorksheetDimension()
    {
        // Return
        return 'A1' . ':' .  this->getHighestColumn() . this->getHighestRow();
    }

    /**
     * Calculate worksheet data dimension
     *
     * @return string  String containing the dimension of this worksheet that actually contain data
     */
    public function calculateWorksheetDataDimension()
    {
        // Return
        return 'A1' . ':' .  this->getHighestDataColumn() . this->getHighestDataRow();
    }

    /**
     * Calculate widths for auto-size columns
     *
     * @param  boolean  calculateMergeCells  Calculate merge cell width
     * @return PHPExcel_Worksheet;
     */
    public function calculateColumnWidths(calculateMergeCells = false)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get parent
     *
     * @return PHPExcel
     */
    public function getParent() {
        return this->_parent;
    }

    /**
     * Re-bind parent
     *
     * @param PHPExcel parent
     * @return PHPExcel_Worksheet
     */
    public function rebindParent(<\ZExcel\ZExcel> parent) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle()
    {
        return this->_title;
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
     * @return PHPExcel_Worksheet
     */
    public function setTitle(pValue = 'Worksheet', updateFormulaCellReferences = true)
    {
        throw new \Exception("Not implemented yet!");
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
     * @return PHPExcel_Worksheet
     */
    public function setSheetState(string value = \ZExcel\Worksheet::SHEETSTATE_VISIBLE) -> <\ZExcel\Worksheet>
	{
        let this->_sheetState = value;
        return this;
    }

    /**
     * Get page setup
     *
     * @return PHPExcel_Worksheet_PageSetup
     */
    public function getPageSetup()
    {
        return this->_pageSetup;
    }

    /**
     * Set page setup
     *
     * @param PHPExcel_Worksheet_PageSetup    pValue
     * @return PHPExcel_Worksheet
     */
    public function setPageSetup(<\ZExcel\Worksheet\PageSetup> pValue)
    {
        let this->_pageSetup = pValue;
        return this;
    }

    /**
     * Get page margins
     *
     * @return PHPExcel_Worksheet_PageMargins
     */
    public function getPageMargins()
    {
        return this->_pageMargins;
    }

    /**
     * Set page margins
     *
     * @param PHPExcel_Worksheet_PageMargins    pValue
     * @return PHPExcel_Worksheet
     */
    public function setPageMargins(<\ZExcel\Worksheet\PageMargins> pValue)
    {
        let this->_pageMargins = pValue;
        return this;
    }

    /**
     * Get page header/footer
     *
     * @return PHPExcel_Worksheet_HeaderFooter
     */
    public function getHeaderFooter()
    {
        return this->_headerFooter;
    }

    /**
     * Set page header/footer
     *
     * @param PHPExcel_Worksheet_HeaderFooter    pValue
     * @return PHPExcel_Worksheet
     */
    public function setHeaderFooter(<\ZExcel\Worksheet\HeaderFooter> pValue)
    {
        let this->_headerFooter = pValue;
        return this;
    }

    /**
     * Get sheet view
     *
     * @return PHPExcel_Worksheet_SheetView
     */
    public function getSheetView()
    {
        return this->_sheetView;
    }

    /**
     * Set sheet view
     *
     * @param PHPExcel_Worksheet_SheetView    pValue
     * @return PHPExcel_Worksheet
     */
    public function setSheetView(<ZExcel\Worksheet\SheetView> pValue)
    {
        let this->_sheetView = pValue;
        return this;
    }

    /**
     * Get Protection
     *
     * @return PHPExcel_Worksheet_Protection
     */
    public function getProtection()
    {
        return this->_protection;
    }

    /**
     * Set Protection
     *
     * @param PHPExcel_Worksheet_Protection    pValue
     * @return PHPExcel_Worksheet
     */
    public function setProtection(<\ZExcel\Worksheet\Protection> pValue)
    {
        let this->_protection = pValue;
        let this->_dirty = true;

        return this;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     row        Return the data highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return string Highest column name
     */
    public function getHighestColumn(string row = null)
    {
        if (row == null) {
            return this->_cachedHighestColumn;
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
    public function getHighestDataColumn(string row = null)
    {
        return this->_cellCollection->getHighestColumn(row);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     column     Return the highest data row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return int Highest row number
     */
    public function getHighestRow(string column = null)
    {
        if (column == null) {
            return this->_cachedHighestRow;
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
    public function getHighestDataRow(string column = null)
    {
        return this->_cellCollection->getHighestRow(column);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        return this->_cellCollection->getHighestRowAndColumn();
    }

    /**
     * Set a cell value
     *
     * @param string pCoordinate Coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param bool returnCell   Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel_Worksheet|PHPExcel_Cell    Depending on the last parameter being specified
     */
    public function setCellValue(string pCoordinate = 'A1', var pValue = null, boolean returnCell = false)
    {
    	var cell;
    	
        let cell = this->getCell(strtoupper(pCoordinate))->setValue(pValue);
        
        return (returnCell) ? cell : this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell (A = 0)
     * @param string pRow Numeric row coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel_Worksheet|PHPExcel_Cell    Depending on the last parameter being specified
     */
    public function setCellValueByColumnAndRow(int pColumn = 0, int pRow = 1, var pValue = null, boolean returnCell = false)
    {
    	var cell;
    	
        let cell = this->getCellByColumnAndRow(pColumn, pRow)->setValue(pValue);
        
        return (returnCell) ? cell : this;
    }

    /**
     * Set a cell value
     *
     * @param string pCoordinate Coordinate of the cell
     * @param mixed  pValue Value of the cell
     * @param string pDataType Explicit data type
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel_Worksheet|PHPExcel_Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicit(string pCoordinate = 'A1', var pValue = null, var pDataType = \ZExcel\Cell\DataType::TYPE_STRING, boolean returnCell = false)
    {
    	var cell;
    	
        // Set value
        let cell = this->getCell(strtoupper(pCoordinate))->setValueExplicit(pValue, pDataType);
        
        return (returnCell) ? cell : this;
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @param mixed pValue Value of the cell
     * @param string pDataType Explicit data type
     * @param bool returnCell Return the worksheet (false, default) or the cell (true)
     * @return PHPExcel_Worksheet|PHPExcel_Cell    Depending on the last parameter being specified
     */
    public function setCellValueExplicitByColumnAndRow(int pColumn = 0, int pRow = 1, var pValue = null, pDataType = \ZExcel\Cell\DataType::TYPE_STRING, boolean returnCell = false)
    {
    	var cell;
    	
        let cell = this->getCellByColumnAndRow(pColumn, pRow)->setValueExplicit(pValue, pDataType);
        
        return (returnCell) ? cell : this;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param string pCoordinate    Coordinate of the cell
     * @throws PHPExcel_Exception
     * @return PHPExcel_Cell Cell that was found
     */
    public function getCell(string pCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get cell at a specific coordinate by using numeric cell coordinates
     *
     * @param  string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @return PHPExcel_Cell Cell that was found
     */
    public function getCellByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Create a new cell at the specified coordinate
     *
     * @param string pCoordinate    Coordinate of the cell
     * @return PHPExcel_Cell Cell that was created
     */
    private function _createNewCell(string pCoordinate)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Does the cell at a specific coordinate exist?
     *
     * @param string pCoordinate  Coordinate of the cell
     * @throws PHPExcel_Exception
     * @return boolean
     */
    public function cellExists(string pCoordinate = 'A1')
    {
       throw new \Exception("Not implemented yet!");
    }

    /**
     * Cell at a specific coordinate by using numeric cell coordinates exists?
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @param string pRow Numeric row coordinate of the cell
     * @return boolean
     */
    public function cellExistsByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->cellExists(PHPExcel_Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Get row dimension at a specific row
     *
     * @param int pRow Numeric index of the row
     * @return PHPExcel_Worksheet_RowDimension
     */
    public function getRowDimension(int pRow = 1, boolean create = true)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get column dimension at a specific column
     *
     * @param string pColumn String index of the column
     * @return PHPExcel_Worksheet_ColumnDimension
     */
    public function getColumnDimension(string pColumn = 'A', boolean create = true)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get column dimension at a specific column by using numeric cell coordinates
     *
     * @param string pColumn Numeric column coordinate of the cell
     * @return PHPExcel_Worksheet_ColumnDimension
     */
    public function getColumnDimensionByColumn(int pColumn = 0)
    {
        return this->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(pColumn));
    }

    /**
     * Get styles
     *
     * @return PHPExcel_Style[]
     */
    public function getStyles()
    {
        return this->_styles;
    }

    /**
     * Get default style of workbook.
     *
     * @deprecated
     * @return PHPExcel_Style
     * @throws PHPExcel_Exception
     */
    public function getDefaultStyle()
    {
        return this->_parent->getDefaultStyle();
    }

    /**
     * Set default style - should only be used by PHPExcel_IReader implementations!
     *
     * @deprecated
     * @param PHPExcel_Style pValue
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setDefaultStyle(<\ZExcel\Style> pValue)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get style for cell
     *
     * @param string pCellCoordinate Cell coordinate (or range) to get style for
     * @return PHPExcel_Style
     * @throws PHPExcel_Exception
     */
    public function getStyle(string pCellCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get conditional styles for a cell
     *
     * @param string pCoordinate
     * @return PHPExcel_Style_Conditional[]
     */
    public function getConditionalStyles(string pCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Do conditional styles exist for this cell?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function conditionalStylesExists(string pCoordinate = 'A1') -> boolean
    {
        if (isset(this->_conditionalStylesCollection[strtoupper(pCoordinate)])) {
            return true;
        }
        return false;
    }

    /**
     * Removes conditional styles for a cell
     *
     * @param string pCoordinate
     * @return PHPExcel_Worksheet
     */
    public function removeConditionalStyles(string pCoordinate = 'A1')
    {
        unset(this->_conditionalStylesCollection[strtoupper(pCoordinate)]);
        return this;
    }

    /**
     * Get collection of conditional styles
     *
     * @return array
     */
    public function getConditionalStylesCollection()
    {
        return this->_conditionalStylesCollection;
    }

    /**
     * Set conditional styles
     *
     * @param pCoordinate string E.g. 'A1'
     * @param pValue PHPExcel_Style_Conditional[]
     * @return PHPExcel_Worksheet
     */
    public function setConditionalStyles(string pCoordinate = 'A1', <\ZExcel\Style\Conditional> pValue)
    {
        let this->_conditionalStylesCollection[strtoupper(pCoordinate)] = pValue;
        return this;
    }

    /**
     * Get style for cell by using numeric cell coordinates
     *
     * @param int pColumn  Numeric column coordinate of the cell
     * @param int pRow Numeric row coordinate of the cell
     * @param int pColumn2 Numeric column coordinate of the range cell
     * @param int pRow2 Numeric row coordinate of the range cell
     * @return PHPExcel_Style
     */
    public function getStyleByColumnAndRow(int pColumn = 0, int pRow = 1, var pColumn2 = null, var pRow2 = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set shared cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @deprecated
     * @param PHPExcel_Style pSharedCellStyle Cell style to share
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setSharedStyle(<\ZExcel\Style> pSharedCellStyle = null, string pRange = '')
    {
        this->duplicateStyle(pSharedCellStyle, pRange);
        return this;
    }

    /**
     * Duplicate cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param PHPExcel_Style pCellStyle Cell style to duplicate
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function duplicateStyle(<\ZExcel\Style> pCellStyle = null, string pRange = '')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Duplicate conditional style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param    array of PHPExcel_Style_Conditional    pCellStyle    Cell style to duplicate
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function duplicateConditionalStyle(array pCellStyle = null, pRange = '')
    {
        throw new \Exception("Not implemented yet!");
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
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function duplicateStyleArray(array pStyles = null, string pRange = '', boolean pAdvanced = true)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set break on a cell
     *
     * @param string pCell Cell coordinate (e.g. A1)
     * @param int pBreak Break type (type of PHPExcel_Worksheet::BREAK_*)
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setBreak(string pCell = 'A1', pBreak = \ZExcel\Worksheet::BREAK_NONE)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set break on a cell by using numeric cell coordinates
     *
     * @param integer pColumn Numeric column coordinate of the cell
     * @param integer pRow Numeric row coordinate of the cell
     * @param  integer pBreak Break type (type of PHPExcel_Worksheet::BREAK_*)
     * @return PHPExcel_Worksheet
     */
    public function setBreakByColumnAndRow(int pColumn = 0, int pRow = 1, pBreak = \ZExcel\Worksheet::BREAK_NONE)
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
        return this->_breaks;
    }

    /**
     * Set merge on a cell range
     *
     * @param string pRange  Cell range (e.g. A1:E1)
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function mergeCells(string pRange = 'A1:A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set merge on a cell range by using numeric cell coordinates
     *
     * @param int pColumn1    Numeric column coordinate of the first cell
     * @param int pRow1        Numeric row coordinate of the first cell
     * @param int pColumn2    Numeric column coordinate of the last cell
     * @param int pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function mergeCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove merge on a cell range
     *
     * @param    string            pRange        Cell range (e.g. A1:E1)
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function unmergeCells(string pRange = 'A1:A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove merge on a cell range by using numeric cell coordinates
     *
     * @param int pColumn1    Numeric column coordinate of the first cell
     * @param int pRow1        Numeric row coordinate of the first cell
     * @param int pColumn2    Numeric column coordinate of the last cell
     * @param int pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function unmergeCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get merge cells array.
     *
     * @return array[]
     */
    public function getMergeCells()
    {
        return this->_mergeCells;
    }

    /**
     * Set merge cells array for the entire sheet. Use instead mergeCells() to merge
     * a single cell range.
     *
     * @param array
     */
    public function setMergeCells(pValue = [])
    {
        let this->_mergeCells = pValue;

        return this;
    }

    /**
     * Set protection on a cell range
     *
     * @param    string            pRange                Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @param    string            pPassword            Password to unlock the protection
     * @param    boolean        pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function protectCells(string pRange = 'A1', string pPassword = '', boolean pAlreadyHashed = false)
    {
        throw new \Exception("Not implemented yet!");
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
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function protectCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1, string pPassword = '', boolean pAlreadyHashed = false)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove protection on a cell range
     *
     * @param    string            pRange        Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function unprotectCells(string pRange = 'A1')
    {
        throw new \Exception("Not implemented yet!");
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
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function unprotectCellsByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1, string pPassword = '', boolean pAlreadyHashed = false)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get protected cells
     *
     * @return array[]
     */
    public function getProtectedCells()
    {
        return this->_protectedCells;
    }

    /**
     *    Get Autofilter
     *
     *    @return PHPExcel_Worksheet_AutoFilter
     */
    public function getAutoFilter()
    {
        return this->_autoFilter;
    }

    /**
     *    Set AutoFilter
     *
     *    @param    PHPExcel_Worksheet_AutoFilter|string   pValue
     *            A simple string containing a Cell range like 'A1:E10' is permitted for backward compatibility
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet
     */
    public function setAutoFilter(var pValue)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Set Autofilter Range by using numeric cell coordinates
     *
     *    @param  integer  pColumn1    Numeric column coordinate of the first cell
     *    @param  integer  pRow1       Numeric row coordinate of the first cell
     *    @param  integer  pColumn2    Numeric column coordinate of the second cell
     *    @param  integer  pRow2       Numeric row coordinate of the second cell
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet
     */
    public function setAutoFilterByColumnAndRow(int pColumn1 = 0, int pRow1 = 1, int pColumn2 = 0, int pRow2 = 1)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove autofilter
     *
     * @return PHPExcel_Worksheet
     */
    public function removeAutoFilter()
    {
        this->_autoFilter->setRange(NULL);
        return this;
    }

    /**
     * Get Freeze Pane
     *
     * @return string
     */
    public function getFreezePane()
    {
        return this->_freezePane;
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
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function freezePane(string pCell = '')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Freeze Pane by using numeric cell coordinates
     *
     * @param int pColumn    Numeric column coordinate of the cell
     * @param int pRow        Numeric row coordinate of the cell
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function freezePaneByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->freezePane(\ZExcel\Cell::stringFromColumnIndex(pColumn) . pRow);
    }

    /**
     * Unfreeze Pane
     *
     * @return PHPExcel_Worksheet
     */
    public function unfreezePane()
    {
        return this->freezePane('');
    }

    /**
     * Insert a new row, updating all possible related data
     *
     * @param int pBefore    Insert before this one
     * @param int pNumRows    Number of rows to insert
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function insertNewRowBefore(int pBefore = 1, int pNumRows = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int pBefore    Insert before this one
     * @param int pNumCols    Number of columns to insert
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function insertNewColumnBefore(string pBefore = 'A', int pNumCols = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int pBefore    Insert before this one (numeric column coordinate of the cell)
     * @param int pNumCols    Number of columns to insert
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function insertNewColumnBeforeByIndex(int pBefore = 0, int pNumCols = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Delete a row, updating all possible related data
     *
     * @param int pRow        Remove starting with this one
     * @param int pNumRows    Number of rows to remove
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function removeRow(int pRow = 1, int pNumRows = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param string    pColumn     Remove starting with this one
     * @param int       pNumCols    Number of columns to remove
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function removeColumn(string pColumn = 'A', int pNumCols = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int pColumn    Remove starting with this one (numeric column coordinate of the cell)
     * @param int pNumCols    Number of columns to remove
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function removeColumnByIndex(int pColumn = 0, int pNumCols = 1) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Show gridlines?
     *
     * @return boolean
     */
    public function getShowGridlines() {
        return this->_showGridlines;
    }

    /**
     * Set show gridlines
     *
     * @param boolean pValue    Show gridlines (true/false)
     * @return PHPExcel_Worksheet
     */
    public function setShowGridlines(boolean pValue = false) {
        let this->_showGridlines = pValue;
        return this;
    }

    /**
    * Print gridlines?
    *
    * @return boolean
    */
    public function getPrintGridlines() {
        return this->_printGridlines;
    }

    /**
    * Set print gridlines
    *
    * @param boolean pValue Print gridlines (true/false)
    * @return PHPExcel_Worksheet
    */
    public function setPrintGridlines(boolean pValue = false) {
        let this->_printGridlines = pValue;
        return this;
    }

    /**
    * Show row and column headers?
    *
    * @return boolean
    */
    public function getShowRowColHeaders() {
        return this->_showRowColHeaders;
    }

    /**
    * Set show row and column headers
    *
    * @param boolean pValue Show row and column headers (true/false)
    * @return PHPExcel_Worksheet
    */
    public function setShowRowColHeaders(boolean pValue = false) {
        let this->_showRowColHeaders = pValue;
        return this;
    }

    /**
     * Show summary below? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryBelow() {
        return this->_showSummaryBelow;
    }

    /**
     * Set show summary below
     *
     * @param boolean pValue    Show summary below (true/false)
     * @return PHPExcel_Worksheet
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
     * @return PHPExcel_Worksheet
     */
    public function setShowSummaryRight(boolean pValue = true) {
        let this->_showSummaryRight = pValue;
        return this;
    }

    /**
     * Get comments
     *
     * @return PHPExcel_Comment[]
     */
    public function getComments()
    {
        return this->_comments;
    }

    /**
     * Set comments array for the entire sheet.
     *
     * @param array of PHPExcel_Comment
     * @return PHPExcel_Worksheet
     */
    public function setComments(array pValue = [])
    {
        let this->_comments = pValue;

        return this;
    }

    /**
     * Get comment for cell
     *
     * @param string pCellCoordinate    Cell coordinate to get comment for
     * @return PHPExcel_Comment
     * @throws PHPExcel_Exception
     */
    public function getComment(string pCellCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get comment for cell by using numeric cell coordinates
     *
     * @param int pColumn    Numeric column coordinate of the cell
     * @param int pRow        Numeric row coordinate of the cell
     * @return PHPExcel_Comment
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
     * @return string Example: 'A1'
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
     * @return PHPExcel_Worksheet
     */
    public function setSelectedCell(string pCoordinate = 'A1')
    {
        return this->setSelectedCells(pCoordinate);
    }

    /**
     * Select a range of cells.
     *
     * @param    string        pCoordinate    Cell range, examples: 'A1', 'B2:G5', 'A:C', '3:6'
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setSelectedCells(string pCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Selected cell by using numeric cell coordinates
     *
     * @param int pColumn Numeric column coordinate of the cell
     * @param int pRow Numeric row coordinate of the cell
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setSelectedCellByColumnAndRow(int pColumn = 0, int pRow = 1)
    {
        return this->setSelectedCells(PHPExcel_Cell::stringFromColumnIndex(pColumn) . pRow);
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
     * @return PHPExcel_Worksheet
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
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function fromArray(var source = null, var nullValue = null, string startCell = 'A1', boolean strictNullComparison = false) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Create array from a range of cells
     *
     * @param string pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param mixed nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean calculateFormulas Should formulas be calculated?
     * @param boolean formatData Should formatting be applied to cell values?
     * @param boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function rangeToArray(string pRange = 'A1', var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false) {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Create array from a range of cells
     *
     * @param  string pNamedRange Name of the Named Range
     * @param  mixed  nullValue Value returned in the array entry if a cell doesn't exist
     * @param  boolean calculateFormulas  Should formulas be calculated?
     * @param  boolean formatData  Should formatting be applied to cell values?
     * @param  boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                                True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     * @throws PHPExcel_Exception
     */
    public function namedRangeToArray(string pNamedRange = '', var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false) {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Create array from worksheet
     *
     * @param mixed nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean calculateFormulas Should formulas be calculated?
     * @param boolean formatData  Should formatting be applied to cell values?
     * @param boolean returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function toArray(var nullValue = null, boolean calculateFormulas = true, boolean formatData = true, boolean returnCellRef = false) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get row iterator
     *
     * @param   integer   startRow   The row number at which to start iterating
     * @param   integer   endRow     The row number at which to stop iterating
     *
     * @return PHPExcel_Worksheet_RowIterator
     */
    public function getRowIterator(int startRow = 1, int endRow = null) {
        return new \ZExcel\Worksheet\RowIterator(this, startRow, endRow);
    }

    /**
     * Get column iterator
     *
     * @param   string   startColumn The column address at which to start iterating
     * @param   string   endColumn   The column address at which to stop iterating
     *
     * @return PHPExcel_Worksheet_ColumnIterator
     */
    public function getColumnIterator(string startColumn = 'A', string endColumn = null) {
        return new \ZExcel\Worksheet\ColumnIterator(this, startColumn, endColumn);
    }

    /**
     * Run PHPExcel garabage collector.
     *
     * @return PHPExcel_Worksheet
     */
    public function garbageCollect() {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Extract worksheet title from range.
     *
     * Example: extractSheetTitle("testSheet!A1") ==> 'A1'
     * Example: extractSheetTitle("'testSheet 1'!A1", true) ==> array('testSheet 1', 'A1');
     *
     * @param string pRange    Range to extract title from
     * @param bool returnRange    Return range? (see example)
     * @return mixed
     */
    public static function extractSheetTitle(string pRange, boolean returnRange = false) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get hyperlink
     *
     * @param string pCellCoordinate    Cell coordinate to get hyperlink for
     */
    public function getHyperlink(string pCellCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set hyperlnk
     *
     * @param string pCellCoordinate    Cell coordinate to insert hyperlink
     * @param    PHPExcel_Cell_Hyperlink    pHyperlink
     * @return PHPExcel_Worksheet
     */
    public function setHyperlink(string pCellCoordinate = 'A1', <\ZExcel\Cell\Hyperlink> pHyperlink = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Hyperlink at a specific coordinate exists?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function hyperlinkExists(string pCoordinate = 'A1')
    {
        return isset(this->_hyperlinkCollection[pCoordinate]);
    }

    /**
     * Get collection of hyperlinks
     *
     * @return PHPExcel_Cell_Hyperlink[]
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
    public function getDataValidation(string pCellCoordinate = 'A1')
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set data validation
     *
     * @param string pCellCoordinate    Cell coordinate to insert data validation
     * @param    PHPExcel_Cell_DataValidation    pDataValidation
     * @return PHPExcel_Worksheet
     */
    public function setDataValidation(string pCellCoordinate = 'A1', <\ZExcel\Cell\DataValidation> pDataValidation = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Data validation at a specific coordinate exists?
     *
     * @param string pCoordinate
     * @return boolean
     */
    public function dataValidationExists(string pCoordinate = 'A1')
    {
        return isset(this->_dataValidationCollection[pCoordinate]);
    }

    /**
     * Get collection of data validations
     *
     * @return PHPExcel_Cell_DataValidation[]
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
    public function shrinkRangeToFit(range)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get tab color
     *
     * @return PHPExcel_Style_Color
     */
    public function getTabColor() -> <\ZExcel\Style\Color>
    {
        if (this->_tabColor === NULL) {
            let this->_tabColor = new \ZExcel\Style\Color();
        }

        return this->_tabColor;
    }

    /**
     * Reset tab color
     *
     * @return PHPExcel_Worksheet
     */
    public function resetTabColor() -> <\ZExcel\Worksheet>
    {
        let this->_tabColor = null;
        unset(this->_tabColor);

        return this;
    }

    /**
     * Tab color set?
     *
     * @return boolean
     */
    public function isTabColorSet() -> boolean
    {
        return (this->_tabColor !== NULL);
    }

    /**
     * Copy worksheet (!= clone!)
     *
     * @return PHPExcel_Worksheet
     */
    public function copy() -> <\ZExcel\Worksheet>
    {
    	var copied;
    	
        let copied = clone this;

        return copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        throw new \Exception("Not implemented yet!");
    }
    
	/**
     * Define the code name of the sheet
     *
     * @param null|string Same rule as Title minus space not allowed (but, like Excel, change silently space to underscore)
     * @return objWorksheet
     * @throws PHPExcel_Exception
    */
    public function setCodeName(pValue=null)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Return the code name of the sheet
     *
     * @return null|string
    */
    
    public function getCodeName()
    {
        return this->_codeName;
    }
    
    /**
     * Sheet has a code name ?
     * @return boolean
    */
    public function hasCodeName() -> boolean
    {
        return !(is_null(this->_codeName));
    }
}
