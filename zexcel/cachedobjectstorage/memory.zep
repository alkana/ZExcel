namespace ZExcel\CachedObjectStorage;

class Memory extends CacheBase implements ICache
{
    /**
     * Dummy method callable from CacheBase, but unused by Memory cache
     *
     * @return    void
     */
    protected function storeData()
    {
    }

    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to update
     * @param    \ZExcel\Cell    cell        Cell to update
     * @return    \ZExcel\Cell
     * @throws    \ZExcel\Exception
     */
    public function addCacheData(string pCoord, <\ZExcel\Cell> cell)
    {
        if (strlen(pCoord) === 0) {
            throw new \Exception("Cache Key must be a string");
        }
        
        let this->cellCache[pCoord] = cell;

        //    Set current entry to the new/updated entry
        let this->currentObjectID = pCoord;

        return cell;
    }


    /**
     * Get cell at a specific coordinate
     *
     * @param     string             pCoord        Coordinate of the cell
     * @throws     \ZExcel\Exception
     * @return     \ZExcel\Cell     Cell that was found, or null if not found
     */
    public function getCacheData(string pCoord)
    {
        //    Check if the entry that has been requested actually exists
        if (!isset(this->cellCache[pCoord])) {
            let this->currentObjectID = null;
            //    Return null if requested entry doesn't exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;

        //    Return requested entry
        return this->cellCache[pCoord];
    }


    /**
     * Clone the cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The new worksheet
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent)
    {
        var k, cell;
        array newCollection;
        
        parent::copyCellCollection(parent);

        let newCollection = [];
        for k, cell in this->cellCache {
            let newCollection[k] = clone cell;
            newCollection[k]->attach(this);
        }

        let this->cellCache = newCollection;
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
     */
    public function unsetWorksheetCells()
    {
        var k;
        // Because cells are all stored as intact objects in memory, we need to detach each one from the parent
        for k, _ in this->cellCache {
            this->cellCache[k]->detach();
            let this->cellCache[k] = null;
        }

        let this->cellCache = [];

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        let this->parent = null;
    }
}
