namespace ZExcel\CachedObjectStorage;

class DiscISAM extends CacheBase implements ICache
{
    /**
     * Name of the file for this cache
     *
     * @var string
     */
    private fileName = null;

    /**
     * File handle for this cache file
     *
     * @var resource
     */
    private fileHandle = null;

    /**
     * Directory/Folder where the cache file is located
     *
     * @var string
     */
    private cacheDirectory = null;

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

            fseek(this->fileHandle, 0, SEEK_END);

            let this->cellCache[this->currentObjectID] = [
                "ptr": ftell(this->fileHandle),
                "sz": fwrite(this->fileHandle, serialize(this->currentObject))
            ];
            
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
            //    Return null if requested entry doesn"t exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        let this->currentObjectID = pCoord;
        fseek(this->fileHandle, this->cellCache[pCoord]["ptr"]);
        let this->currentObject = unserialize(fread(this->fileHandle, this->cellCache[pCoord]["sz"]));
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
     * Clone the cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The new worksheet
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent)
    {
        var baseUnique, newFileName;
        
        parent::copyCellCollection(parent);
        //    Get a new id for the new file name
        let baseUnique = this->getUniqueID();
        let newFileName = this->cacheDirectory."/ZExcel.".baseUnique.".cache";
        //    Copy the existing cell cache file
        copy(this->fileName, newFileName);
        let this->fileName = newFileName;
        //    Open the copied cell cache file
        let this->fileHandle = fopen(this->fileName, "a+");
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
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

        //    Close down the temporary cache file
        this->__destruct();
    }

    /**
     * Initialise this new cell collection
     *
     * @param    \ZExcel\Worksheet    parent        The worksheet for this cell collection
     * @param    array of mixed        arguments    Additional initialisation arguments
     */
    public function __construct(<\ZExcel\Worksheet> parent, arguments)
    {
        var baseUnique;
        
        let this->cacheDirectory = ((isset(arguments["dir"])) && (arguments["dir"] !== null)) ? arguments["dir"] : \ZExcel\Shared\File::sys_get_temp_dir();

        parent::__construct(parent);
        
        if (is_null(this->fileHandle)) {
            let baseUnique = this->getUniqueID();
            let this->fileName = this->cacheDirectory."/PHPExcel.".baseUnique.".cache";
            let this->fileHandle = fopen(this->fileName, "a+");
        }
    }

    /**
     * Destroy this cell collection
     */
    public function __destruct()
    {
        if (!is_null(this->fileHandle)) {
            fclose(this->fileHandle);
            unlink(this->fileName);
        }
        let this->fileHandle = null;
    }
}
