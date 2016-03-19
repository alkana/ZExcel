namespace ZExcel\CachedObjectStorage;

abstract class CacheBase
{
    /**
     * Parent worksheet
     *
     * @var \ZExcel\Worksheet
     */
    protected parent;

    /**
     * The currently active Cell
     *
     * @var \ZExcel\Cell
     */
    protected currentObject = null;

    /**
     * Coordinate address of the currently active Cell
     *
     * @var string
     */
    protected currentObjectID = null;

    /**
     * Flag indicating whether the currently active Cell requires saving
     *
     * @var boolean
     */
    protected currentCellIsDirty = true;

    /**
     * An array of cells or cell pointers for the worksheet cells held in this cache,
     *        and indexed by their coordinate address within the worksheet
     *
     * @var array of mixed
     */
    protected cellCache = [];

    /**
     * Initialise this new cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The worksheet for this cell collection
     */
    public function __construct(<\ZExcel\Worksheet> parent)
    {
        //    Set our parent worksheet.
        //    This is maintained within the cache controller to facilitate re-attaching it to \ZExcel\Cell objects when
        //        they are woken from a serialized state
        let this->parent = parent;
    }

    /**
     * Return the parent worksheet for this cell collection
     *
     * @return    \ZExcel\Worksheet
     */
    public function getParent() -> <\ZExcel\Worksheet>
    {
        return this->parent;
    }

    /**
     * Is a value set in the current \ZExcel\CachedObjectStorage_ICache for an indexed cell?
     *
     * @param    string        pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet(string pCoord) -> boolean
    {
        if (pCoord === this->currentObjectID) {
            return true;
        }
        
        //    Check if the requested entry exists in the cache
        return isset(this->cellCache[pCoord]);
    }

    /**
     * Move a cell object from one address to another
     *
     * @param    string        fromAddress    Current address of the cell to move
     * @param    string        toAddress        Destination address of the cell to move
     * @return    boolean
     */
    public function moveCell(string fromAddress, string toAddress)
    {
        if (fromAddress === this->currentObjectID) {
            let this->currentObjectID = toAddress;
        }
        
        let this->currentCellIsDirty = true;
        
        if (isset(this->cellCache[fromAddress])) {
            let this->cellCache[toAddress] = this->cellCache[fromAddress];
            unset(this->cellCache[fromAddress]);
        }

        return true;
    }
    
    public function addCacheData(pCoord, <\ZExcel\Cell> cell)
    {
        throw new \Exception("Can't be implemented in the abstract class");
    }

    /**
     * Add or Update a cell in cache
     *
     * @param    \ZExcel\Cell    cell        Cell to update
     * @return    \ZExcel\Cell
     * @throws    \ZExcel\Exception
     */
    public function updateCacheData(<\ZExcel\Cell> cell)
    {
        return this->addCacheData(cell->getCoordinate(), cell);
    }

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to delete
     * @throws    \ZExcel\Exception
     */
    public function deleteCacheData(string pCoord)
    {
        if (pCoord === this->currentObjectID && !is_null(this->currentObject)) {
            this->currentObject->detach();
            let this->currentObjectID = null;
            let this->currentObject = null;
        }

        if (is_object(this->cellCache[pCoord])) {
            this->cellCache[pCoord]->detach();
            unset(this->cellCache[pCoord]);
        }
        
        let this->currentCellIsDirty = false;
    }

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    string[]
     */
    public function getCellList()
    {
        return array_keys(this->cellCache);
    }

    /**
     * Sort the list of all cell addresses currently held in cache by row and column
     *
     * @return    string[]
     */
    public function getSortedCellList()
    {
        var coord, column, row;
        array sortKeys = [];
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", column, row);
            let sortKeys[sprintf("%09d%3s", row, column)] = coord;
        }
        ksort(sortKeys);

        return array_values(sortKeys);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        var coord, highestRow, highestColumn, c, r;
        array col, row;
        
        // Lookup highest column and highest row
        let col = ["A": "1A"];
        let row = [1];
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", c, r);
            let row[r] = r;
            let col[c] = strlen(c).c;
        }
        
        if (!empty(row)) {
            // Determine highest column and row
            let highestRow = max(row);
            let highestColumn = substr(max(col), 1);
        }

        return [
            "row": highestRow,
            "column": highestColumn
        ];
    }

    /**
     * Return the cell address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentAddress()
    {
        return this->currentObjectID;
    }

    /**
     * Return the column address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentColumn()
    {
        var column, row;
        
        sscanf(this->currentObjectID, "%[A-Z]%d", column, row);
        
        return column;
    }

    /**
     * Return the row address of the currently active cell object
     *
     * @return    integer
     */
    public function getCurrentRow()
    {
        var column, row;
        
        sscanf(this->currentObjectID, "%[A-Z]%d", column, row);
        
        return (int) row;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     row        Return the highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return  string     Highest column name
     */
    public function getHighestColumn(string row = null)
    {
        var colRow, coord, c, r;
        array columnList;
    
        if (row == null) {
            let colRow = this->getHighestRowAndColumn();
            return colRow["column"];
        }

        let columnList = [1];
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", c, r);
            if (r != row) {
                continue;
            }
            let columnList[] = \ZExcel\Cell::columnIndexFromString(c);
        }
        return \ZExcel\Cell::stringFromColumnIndex(max(columnList) - 1);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     column     Return the highest row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return  int        Highest row number
     */
    public function getHighestRow(string column = null)
    {
        var colRow, coord, c, r;
        array rowList;
        
        if (column == null) {
            let colRow = this->getHighestRowAndColumn();
            return colRow["row"];
        }

        let rowList = [0];
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", c, r);
            
            if (c != column) {
                continue;
            }
            
            let rowList[] = r;
        }

        return max(rowList);
    }

    /**
     * Generate a unique ID for cache referencing
     *
     * @return string Unique Reference
     */
    protected function getUniqueID()
    {
        var baseUnique;
        
        if (function_exists("posix_getpid")) {
            let baseUnique = posix_getpid();
        } else {
            let baseUnique = mt_rand();
        }
        
        return uniqid(baseUnique, true);
    }
    
    protected function storeData()
    {
    }

    /**
     * Clone the cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent)
    {
        this->storeData();

        let this->parent = parent;
        
        if ((this->currentObject !== null) && (is_object(this->currentObject))) {
            this->currentObject->attach(this);
        }
    }    //    function copyCellCollection()

    /**
     * Remove a row, deleting all cells in that row
     *
     * @param string    row    Row number to remove
     * @return void
     */
    public function removeRow(row)
    {
        var coord, c, r;
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", c, r);
            if (r == row) {
                this->deleteCacheData(coord);
            }
        }
    }

    /**
     * Remove a column, deleting all cells in that column
     *
     * @param string    column    Column ID to remove
     * @return void
     */
    public function removeColumn(column)
    {
        var coord, c, r;
        
        for coord in this->getCellList() {
            sscanf(coord, "%[A-Z]%d", c, r);
            if (c == column) {
                this->deleteCacheData(coord);
            }
        }
    }

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable()
    {
        return true;
    }
}
