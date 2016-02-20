namespace ZExcel\Worksheet;

class Column
{
    /**
     * \ZExcel\Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    private parent;

    /**
     * Column index
     *
     * @var string
     */
    private columnIndex;

    /**
     * Create a new column
     *
     * @param \ZExcel\Worksheet     parent
     * @param string                columnIndex
     */
    public function __construct(<\ZExcel\Worksheet> parent = null, string columnIndex = "A")
    {
        // Set parent and column index
        let this->parent      = parent;
        let this->columnIndex = columnIndex;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->parent);
    }

    /**
     * Get column index
     *
     * @return int
     */
    public function getColumnIndex() -> int
    {
        return this->columnIndex;
    }

    /**
     * Get cell iterator
     *
     * @param    integer                $startRow        The row number at which to start iterating
     * @param    integer                $endRow            Optionally, the row number at which to stop iterating
     * @return \ZExcel\Worksheet\CellIterator
     */
    public function getCellIterator(int startRow = 1, endRow = null) -> <\ZExcel\Worksheet\CellIterator>
    {
        return new \ZExcel\Worksheet\ColumnCellIterator(this->parent, this->columnIndex, startRow, endRow);
    }
}
