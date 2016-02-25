namespace ZExcel\CachedObjectStorage;

interface ICache
{
    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to update
     * @param    PHPExcel_Cell    cell        Cell to update
     * @return    PHPExcel_Cell
     * @throws    PHPExcel_Exception
     */
    public function addCacheData(pCoord, <\ZExcel\Cell> cell);

    /**
     * Add or Update a cell in cache
     *
     * @param    PHPExcel_Cell    cell        Cell to update
     * @return    PHPExcel_Cell
     * @throws    PHPExcel_Exception
     */
    public function updateCacheData(<\ZExcel\Cell> cell);

    /**
     * Fetch a cell from cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to retrieve
     * @return PHPExcel_Cell     Cell that was found, or null if not found
     * @throws    PHPExcel_Exception
     */
    public function getCacheData(pCoord);

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            pCoord        Coordinate address of the cell to delete
     * @throws    PHPExcel_Exception
     */
    public function deleteCacheData(pCoord);

    /**
     * Is a value set in the current PHPExcel_CachedObjectStorage_ICache for an indexed cell?
     *
     * @param    string        pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet(pCoord);

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    string[]
     */
    public function getCellList();

    /**
     * Get the list of all cell addresses currently held in cache sorted by column and row
     *
     * @return    string[]
     */
    public function getSortedCellList();

    /**
     * Clone the cell collection
     *
     * @param    PHPExcel_Worksheet    parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(<\ZExcel\Worksheet> parent);

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable();
}
