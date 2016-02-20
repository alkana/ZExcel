namespace ZExcel\Worksheet;

class Row
{
    /**
     * \ZExcel\Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    private parent;

    /**
     * Row index
     *
     * @var int
     */
    private rowIndex = 0;

    /**
     * Create a new row
     *
     * @param \ZExcel\Worksheet parent
     * @param int               rowIndex
     */
    public function __construct(<\ZExcel\Worksheet> parent = null, int rowIndex = 1)
    {
        // Set parent and row index
        let this->parent   = parent;
        let this->rowIndex = rowIndex;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->parent);
    }

    /**
     * Get row index
     *
     * @return int
     */
    public function getRowIndex() -> int
    {
        return this->rowIndex;
    }

    /**
     * Get cell iterator
     *
     * @param    string startColumn The column address at which to start iterating
     * @param    string endColumn   Optionally, the column address at which to stop iterating
     * @return \ZExcel\Worksheet\CellIterator
     */
    public function getCellIterator(string startColumn = "A", string endColumn = null) -> <\ZExcel\Worksheet\CellIterator>
    {
        return new \ZExcel\Worksheet\RowCellIterator(this->parent, this->rowIndex, startColumn, endColumn);
    }
}
