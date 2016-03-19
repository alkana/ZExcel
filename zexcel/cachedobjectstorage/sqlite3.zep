namespace ZExcel\CachedObjectStorage;

class SQLite3 extends CacheBase implements ICache
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
     * Prepared statement for a SQLite3 select query
     *
     * @var SQLite3Stmt
     */
    private selectQuery;

    /**
     * Prepared statement for a SQLite3 insert query
     *
     * @var SQLite3Stmt
     */
    private insertQuery;

    /**
     * Prepared statement for a SQLite3 update query
     *
     * @var SQLite3Stmt
     */
    private updateQuery;

    /**
     * Prepared statement for a SQLite3 delete query
     *
     * @var SQLite3Stmt
     */
    private deleteQuery;

    /**
     * Store cell data in cache for the current cell object if it's "dirty",
     *     and the 'nullify' the current cell object
     *
     * @return    void
     * @throws    \ZExcel\Exception
     */
    protected function storeData()
    {
        var result;
        
        if (this->currentCellIsDirty && !empty(this->currentObjectID)) {
            this->currentObject->detach();

            this->insertQuery->bindValue("id", this->currentObjectID, SQLITE3_TEXT);
            this->insertQuery->bindValue("data", serialize(this->currentObject), SQLITE3_BLOB);
            
            let result = this->insertQuery->execute();
            
            if (result === false) {
                throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
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
        var cellResult, cellData;
        
        if (pCoord === this->currentObjectID) {
            return this->currentObject;
        }
        
        this->storeData();

        this->selectQuery->bindValue("id", pCoord, SQLITE3_TEXT);
        
        let cellResult = this->selectQuery->execute();
        
        if (cellResult === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }
        
        let cellData = cellResult->fetchArray(SQLITE3_ASSOC);
        
        if (cellData === false) {
            //    Return null if requested entry doesn't exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;

        let this->currentObject = unserialize(cellData["value"]);
        //    Re-attach this as the cell's parent
        this->currentObject->attach(this);

        //    Return requested entry
        return this->currentObject;
    }

    /**
     *    Is a value set for an indexed cell?
     *
     * @param    string        pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet(pCoord)
    {
        var cellResult, cellData;
        
        if (pCoord === this->currentObjectID) {
            return true;
        }

        //    Check if the requested entry exists in the cache
        this->selectQuery->bindValue("id", pCoord, SQLITE3_TEXT);
        
        let cellResult = this->selectQuery->execute();
        
        if (cellResult === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }
        
        let cellData = cellResult->fetchArray(SQLITE3_ASSOC);

        return (cellData === false) ? false : true;
    }

    /**
     *    Delete a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to delete
     * @throws    \ZExcel\Exception
     */
    public function deleteCacheData(pCoord)
    {
        var result;
        
        if (pCoord === this->currentObjectID) {
            this->currentObject->detach();
            let this->currentObjectID = null;
            let this->currentObject = null;
        }

        //    Check if the requested entry exists in the cache
        this->deleteQuery->bindValue("id", pCoord, SQLITE3_TEXT);
        
        let result = this->deleteQuery->execute();
        
        if (result === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
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
        
        if (fromAddress === this->currentObjectID) {
            let this->currentObjectID = toAddress;
        }

        this->deleteQuery->bindValue("id", toAddress, SQLITE3_TEXT);
        
        let result = this->deleteQuery->execute();
        
        if (result === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }

        this->updateQuery->bindValue("toid", toAddress, SQLITE3_TEXT);
        this->updateQuery->bindValue("fromid", fromAddress, SQLITE3_TEXT);
        
        let result = this->updateQuery->execute();
        
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

        let query = "SELECT id FROM kvp_".this->TableName;
        let cellIdsResult = this->DBHandle->query(query);
        if (cellIdsResult === false) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
        }

        let cellKeys = [];
        let row = cellIdsResult->fetchArray(SQLITE3_ASSOC);
        while (row) {
            let cellKeys[] = row["id"];
            let row = cellIdsResult->fetchArray(SQLITE3_ASSOC);
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
        if (!this->DBHandle->exec("CREATE TABLE kvp_".tableName." (id VARCHAR(12) PRIMARY KEY, value BLOB) AS SELECT * FROM kvp_".this->TableName)
        ) {
            throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
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

            let this->DBHandle = new \SQLite3(_DBName);
            
            if (this->DBHandle === false) {
                throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
            }
            
            if (!this->DBHandle->exec("CREATE TABLE kvp_".this->TableName." (id VARCHAR(12) PRIMARY KEY, value BLOB)")) {
                throw new \ZExcel\Exception(this->DBHandle->lastErrorMsg());
            }
        }

        let this->selectQuery = this->DBHandle->prepare("SELECT value FROM kvp_".this->TableName." WHERE id = :id");
        let this->insertQuery = this->DBHandle->prepare("INSERT OR REPLACE INTO kvp_".this->TableName." VALUES(:id,:data)");
        let this->updateQuery = this->DBHandle->prepare("UPDATE kvp_".this->TableName." SET id=:toId WHERE id=:fromId");
        let this->deleteQuery = this->DBHandle->prepare("DELETE FROM kvp_".this->TableName." WHERE id = :id");
    }

    /**
     * Destroy this cell collection
     */
    public function __destruct()
    {
        if (!is_null(this->DBHandle)) {
            this->DBHandle->exec("DROP TABLE kvp_".this->TableName);
            this->DBHandle->close();
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
        if (!class_exists("SQLite3", false)) {
            return false;
        }

        return true;
    }
}
