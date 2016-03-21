namespace ZExcel\CachedObjectStorage;

class Apc extends CacheBase implements ICache
{/**
     * Prefix used to uniquely identify cache data for this worksheet
     *
     * @access    private
     * @var string
     */
    private cachePrefix = null;

    /**
     * Cache timeout
     *
     * @access    private
     * @var integer
     */
    private cacheTime = 600;

    /**
     * Store cell data in cache for the current cell object if it"s "dirty",
     *     and the "nullify" the current cell object
     *
     * @access  private
     * @return  void
     * @throws  \ZExcel\Exception
     */
    protected function storeData()
    {
        if (this->currentCellIsDirty && !empty(this->currentObjectID)) {
            this->currentObject->detach();

            if (!apc_store(
                this->cachePrefix . this->currentObjectID . ".cache",
                serialize(this->currentObject),
                this->cacheTime
            )) {
                this->__destruct();
                throw new \ZExcel\Exception("Failed to store cell " . this->currentObjectID . " in APC");
            }
            let this->currentCellIsDirty = false;
        }
        
        let this->currentObjectID = null;
        let this->currentObject = null;
    }

    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @access  public
     * @param   string         pCoord  Coordinate address of the cell to update
     * @param   \ZExcel\Cell  cell    Cell to update
     * @return  \ZExcel\Cell
     * @throws  \ZExcel\Exception
     */
    public function addCacheData(pCoord, <\ZExcel\Cell> cell)
    {
        if ((pCoord !== this->currentObjectID) && (this->currentObjectID !== null)) {
            this->storeData();
        }
        
        let this->cellCache[pCoord] = true;

        let this->currentObjectID = pCoord;
        let this->currentObject = cell;
        let this->currentCellIsDirty = true;

        return cell;
    }

    /**
     * Is a value set in the current \ZExcel\CachedObjectStorage\ICache for an indexed cell?
     *
     * @access  public
     * @param   string  pCoord  Coordinate address of the cell to check
     * @throws  \ZExcel\Exception
     * @return  boolean
     */
    public function isDataSet(pCoord)
    {
        var success;
        
        // Check if the requested entry is the current object, or exists in the cache
        if (parent::isDataSet(pCoord)) {
            if (this->currentObjectID == pCoord) {
                return true;
            }
            
            // Check if the requested entry still exists in apc
            let success = apc_fetch(this->cachePrefix . pCoord . ".cache");
            
            if (success === false) {
                // Entry no longer exists in APC, so clear it from the cache array
                parent::deleteCacheData(pCoord);
                throw new \ZExcel\Exception("Cell entry ".pCoord." no longer exists in APC cache");
            }
            
            return true;
        }
        
        return false;
    }

    /**
     * Get cell at a specific coordinate
     *
     * @access  public
     * @param   string         pCoord  Coordinate of the cell
     * @throws  \ZExcel\Exception
     * @return  \ZExcel\Cell  Cell that was found, or null if not found
     */
    public function getCacheData(pCoord)
    {
        var obj;
        
        if (pCoord === this->currentObjectID) {
            return this->currentObject;
        }
        
        this->storeData();

        //    Check if the entry that has been requested actually exists
        if (parent::isDataSet(pCoord)) {
            let obj = apc_fetch(this->cachePrefix . pCoord . ".cache");
            if (obj === false) {
                //    Entry no longer exists in APC, so clear it from the cache array
                parent::deleteCacheData(pCoord);
                throw new \ZExcel\Exception("Cell entry ".pCoord." no longer exists in APC cache");
            }
        } else {
            //    Return null if requested entry doesn"t exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;
        let this->currentObject = unserialize(obj);
        //    Re-attach this as the cell"s parent
        this->currentObject->attach(this);

        //    Return requested entry
        return this->currentObject;
    }

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return  string[]
     */
    public function getCellList()
    {
        if (this->currentObjectID !== null) {
            this->storeData();
        }

        return parent::getCellList();
    }

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @access  public
     * @param   string  pCoord  Coordinate address of the cell to delete
     * @throws  \ZExcel\Exception
     */
    public function deleteCacheData(pCoord)
    {
        //    Delete the entry from APC
        apc_delete(this->cachePrefix.pCoord.".cache");

        //    Delete the entry from our cell address array
        parent::deleteCacheData(pCoord);
    }

    /**
     * Clone the cell collection
     *
     * @access  public
     * @param   \ZExcel\Worksheet  parent  The new worksheet
     * @throws  \ZExcel\Exception
     * @return  void
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent)
    {
        var baseUnique, newCachePrefix, cacheList, cellID, obj;
        
        parent::copyCellCollection(parent);
        //    Get a new id for the new file name
        let baseUnique = this->getUniqueID();
        let newCachePrefix = substr(md5(baseUnique), 0, 8) . ".";
        let cacheList = this->getCellList();
        for cellID in cacheList {
            if (cellID != this->currentObjectID) {
                let obj = apc_fetch(this->cachePrefix . cellID . ".cache");
                if (obj === false) {
                    //    Entry no longer exists in APC, so clear it from the cache array
                    parent::deleteCacheData(cellID);
                    throw new \ZExcel\Exception("Cell entry " . cellID . " no longer exists in APC");
                }
                if (!apc_store(newCachePrefix . cellID . ".cache", obj, this->cacheTime)) {
                    this->__destruct();
                    throw new \ZExcel\Exception("Failed to store cell " . cellID . " in APC");
                }
            }
        }
        
        let this->cachePrefix = newCachePrefix;
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
     * @return  void
     */
    public function unsetWorksheetCells()
    {
        if (this->currentObject !== null) {
            this->currentObject->detach();
            let this->currentObject = null;
            let this->currentObjectID = null;
        }

        //    Flush the APC cache
        this->__destruct();

        let this->cellCache = [];

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        let this->parent = null;
    }

    /**
     * Initialise this new cell collection
     *
     * @param  \ZExcel\Worksheet  parent     The worksheet for this cell collection
     * @param  array of mixed      arguments  Additional initialisation arguments
     */
    public function __construct(<\ZExcel\Worksheet> parent, array arguments = [])
    {
        var cacheTime, baseUnique;
        
        let cacheTime = (isset(arguments["cacheTime"])) ? arguments["cacheTime"] : 600;

        if (this->cachePrefix === null) {
            let baseUnique = this->getUniqueID();
            let this->cachePrefix = substr(md5(baseUnique), 0, 8) . ".";
            let this->cacheTime = cacheTime;

            parent::__construct(parent);
        }
    }

    /**
     * Destroy this cell collection
     */
    public function __destruct()
    {
        var cacheList, cellID;
        
        let cacheList = this->getCellList();
        for cellID in cacheList {
            apc_delete(this->cachePrefix . cellID . ".cache");
        }
    }

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return  boolean
     */
    public static function cacheMethodIsAvailable()
    {
        if (!function_exists("apc_store")) {
            return false;
        }
        if (apc_sma_info() === false) {
            return false;
        }

        return true;
    }
}
