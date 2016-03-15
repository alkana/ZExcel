namespace ZExcel\CachedObjectStorage;

class SQLite extends CacheBase implements ICache
{
    /**
     * Database table name
     *
     * @var string
     */
    private TableName = null;

    /**
     * Database handle
     *
     * @var resource
     */
    private DBHandle = null;

    /**
     * Store cell data in cache for the current cell object if it"s "dirty",
     *     and the "nullify" the current cell object
     *
     * @return    void
     * @throws    \ZExcel\Exception
     */
    protected function storeData()
    {
        if (this->currentCellIsDirty && !empty(this->currentObjectID)) {
            this->currentObject->detach();

            if (!this->DBHandle->queryExec("INSERT OR REPLACE INTO kvp_" . this->TableName . " VALUES('" . this->currentObjectID . "','" . sqlite_escape_string(serialize(this->currentObject)) . "')")) {
                throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
            }
            let this->currentCellIsDirty = false;
        }
        let this->currentObjectID = null;
        let this->currentObject = null;
    }

    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to update
     * @param    \ZExcel\Cell    cell        Cell to update
     * @return    \ZExcel\Cell
     * @throws    \ZExcel\Exception
     */
    public function addCacheData(pCoord, <\ZExcel\Cell> cell)
    {
        if ((pCoord !== this->currentObjectID) && (this->currentObjectID !== null)) {
            this->storeData();
        }

        let this->currentObjectID = pCoord;
        let this->currentObject = cell;
        let this->currentCellIsDirty = true;

        return cell;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param     string             pCoord        Coordinate of the cell
     * @throws     \ZExcel\Exception
     * @return     \ZExcel\Cell     Cell that was found, or null if not found
     */
    public function getCacheData(pCoord)
    {
        var cellResult, cellResultSet;
        string query;
        
        if (pCoord === this->currentObjectID) {
            return this->currentObject;
        }
        this->storeData();

        let query = "SELECT value FROM kvp_" . this->TableName . " WHERE id='" . pCoord . "'";
        let cellResultSet = this->DBHandle->query(query, SQLITE_ASSOC);
        
        if (cellResultSet === false) {
            throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
        } elseif (cellResultSet->numRows() == 0) {
            //    Return null if requested entry doesn"t exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;

        let cellResult = cellResultSet->fetchSingle();
        let this->currentObject = unserialize(cellResult);
        //    Re-attach this as the cell"s parent
        this->currentObject->attach(this);

        //    Return requested entry
        return this->currentObject;
    }

    /**
     * Is a value set for an indexed cell?
     *
     * @param    string        pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet(pCoord)
    {
        var cellResultSet;
        string query;
        
        if (pCoord === this->currentObjectID) {
            return true;
        }

        //    Check if the requested entry exists in the cache
        let query = "SELECT id FROM kvp_" . this->TableName . " WHERE id='" . pCoord . "'";
        let cellResultSet = this->DBHandle->query(query, SQLITE_ASSOC);
        if (cellResultSet === false) {
            throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
        } elseif (cellResultSet->numRows() == 0) {
            //    Return null if requested entry doesn"t exist in cache
            return false;
        }
        return true;
    }

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to delete
     * @throws    \ZExcel\Exception
     */
    public function deleteCacheData(pCoord)
    {
        string query;
        
        if (pCoord === this->currentObjectID) {
            this->currentObject->detach();
            let this->currentObjectID = null;
            let this->currentObject = null;
        }

        //    Check if the requested entry exists in the cache
        let query = "DELETE FROM kvp_" . this->TableName . " WHERE id='" . pCoord . "'";
        if (!this->DBHandle->queryExec(query)) {
            throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
        }

        let this->currentCellIsDirty = false;
    }

    /**
     * Move a cell object from one address to another
     *
     * @param    string        fromAddress    Current address of the cell to move
     * @param    string        toAddress        Destination address of the cell to move
     * @return    boolean
     */
    public function moveCell(fromAddress, toAddress)
    {
        var result;
        string query;
        
        if (fromAddress === this->currentObjectID) {
            let this->currentObjectID = toAddress;
        }

        let query = "DELETE FROM kvp_" . this->TableName . " WHERE id='" . toAddress . "'";
        let result = this->DBHandle->exec(query);
        
        if (result === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }

        let query = "UPDATE kvp_" . this->TableName . " SET id='" . toAddress . " WHERE id='" . fromAddress . "'";
        let result = this->DBHandle->exec(query);
        if (result === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }

        return true;
    }

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    string[]
     */
    public function getCellList()
    {
        var cellIdsResult, row;
        array cellKeys;
        string query;
        
        if (this->currentObjectID !== null) {
            this->storeData();
        }

        let query = "SELECT id FROM kvp_" . this->TableName;
        let cellIdsResult = this->DBHandle->unbufferedQuery(query, SQLITE_ASSOC);
        if (cellIdsResult === false) {
            throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
        }

        let cellKeys = [];
        for row in cellIdsResult {
            let cellKeys[] = row["id"];
        }

        return cellKeys;
    }

    /**
     * Clone the cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent)
    {
        var tableName;
        
        this->storeData();

        //    Get a new id for the new table name
        let tableName = str_replace(".", "_", this->getUniqueID());
        if (!this->DBHandle->queryExec("CREATE TABLE kvp_" . tableName . " (id VARCHAR(12) PRIMARY KEY, value BLOB) AS SELECT * FROM kvp_" . this->TableName)
        ) {
            throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
        }

        //    Copy the existing cell cache file
        let this->TableName = tableName;
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
     * @return    void
     */
    public function unsetWorksheetCells()
    {
        if (!is_null(this->currentObject)) {
            this->currentObject->detach();
            let this->currentObject = null;
            let this->currentObjectID = null;
        }
        //    detach ourself from the worksheet, so that it can then delete this object successfully
        let this->parent = null;

        //    Close down the temporary cache file
        this->__destruct();
    }

    /**
     * Initialise this new cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The worksheet for this cell collection
     */
    public function __construct(<\ZExcel\Worksheet> parent)
    {
        string _DBName;
        
        parent::__construct(parent);
        if (is_null(this->DBHandle)) {
            let this->TableName = str_replace(".", "_", this->getUniqueID());
            let _DBName = ":memory:";

            let this->DBHandle = new \SQLiteDatabase(_DBName);
            if (this->DBHandle === false) {
                throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
            }
            if (!this->DBHandle->queryExec("CREATE TABLE kvp_".this->TableName." (id VARCHAR(12) PRIMARY KEY, value BLOB)")) {
                throw new \ZExcel\Exception(sqlite_error_string(this->DBHandle->lastError()));
            }
        }
    }

    /**
     * Destroy this cell collection
     */
    public function __destruct()
    {
        if (!is_null(this->DBHandle)) {
            this->DBHandle->queryExec("DROP TABLE kvp_".this->TableName);
        }
        let this->DBHandle = null;
    }

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable()
    {
        if (!function_exists("sqlite_open")) {
            return false;
        }

        return true;
    }
}
