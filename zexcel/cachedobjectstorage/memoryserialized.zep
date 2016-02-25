namespace ZExcel\CachedObjectStorage;

class MemorySerialized extends CacheBase implements ICache
{
    /**
     * Store cell data in cache for the current cell object if it's "dirty",
     *     and the 'nullify' the current cell object
     *
     * @return    void
     * @throws    \ZExcel\Exception
     */
    protected function storeData()
    {
        if (this->currentCellIsDirty && !empty(this->currentObjectID)) {
            this->currentObject->detach();

            let this->cellCache[this->currentObjectID] = serialize(this->currentObject);
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
        if (pCoord === this->currentObjectID) {
            return this->currentObject;
        }
        this->storeData();

        //    Check if the entry that has been requested actually exists
        if (!isset(this->cellCache[pCoord])) {
            //    Return null if requested entry doesn't exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;
        let this->currentObject = unserialize(this->cellCache[pCoord]);
        //    Re-attach this as the cell's parent
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
        let this->cellCache = [];

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        let this->parent = null;
    }
}
