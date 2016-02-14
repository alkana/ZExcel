namespace ZExcel\Worksheet;

class RowCellIterator extends CellIterator implements \Iterator
{
    /**
     * Row index
     *
     * @var int
     */
    protected rowIndex;

    /**
     * Start position
     *
     * @var int
     */
    protected startColumn = 0;

    /**
     * End position
     *
     * @var int
     */
    protected endColumn = 0;

    /**
     * Create a new column iterator
     *
     * @param    \ZExcel\Worksheet    subject        The worksheet to iterate over
     * @param   integer             rowIndex       The row that we want to iterate
     * @param    string                startColumn    The column address at which to start iterating
     * @param    string                endColumn        Optionally, the column address at which to stop iterating
     */
    public function __construct(<\ZExcel\Worksheet> subject = null, int rowIndex = 1, string startColumn = "A", string endColumn = null)
    {
        // Set subject and row index
        let this->subject = subject;
        let this->rowIndex = rowIndex;
        this->resetEnd(endColumn);
        this->resetStart(startColumn);
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->subject);
    }

    /**
     * (Re)Set the start column and the current column pointer
     *
     * @param integer    startColumn    The column address at which to start iterating
     * @return \ZExcel\Worksheet\RowCellIterator
     * @throws \ZExcel\Exception
     */
    public function resetStart(startColumn = "A") -> <\ZExcel\Worksheet\RowCellIterator>
    {
        var startColumnIndex;
        
        let startColumnIndex = \ZExcel\Cell::columnIndexFromString(startColumn) - 1;
        let this->startColumn = startColumnIndex;
        
        this->adjustForExistingOnlyRange();
        this->seek(\ZExcel\Cell::stringFromColumnIndex(this->startColumn));

        return this;
    }

    /**
     * (Re)Set the end column
     *
     * @param string    endColumn    The column address at which to stop iterating
     * @return \ZExcel\Worksheet\RowCellIterator
     * @throws \ZExcel\Exception
     */
    public function resetEnd(endColumn = null) -> <\ZExcel\Worksheet\RowCellIterator>
    {
        let endColumn = (endColumn !== null) ? endColumn : this->subject->getHighestColumn();
        let this->endColumn = \ZExcel\Cell::columnIndexFromString(endColumn) - 1;
        
        this->adjustForExistingOnlyRange();

        return this;
    }

    /**
     * Set the column pointer to the selected column
     *
     * @param string    column    The column address to set the current pointer at
     * @return \ZExcel\Worksheet\RowCellIterator
     * @throws \ZExcel\Exception
     */
    public function seek(column = "A")
    {
        let column = \ZExcel\Cell::columnIndexFromString(column) - 1;
        
        if ((column < this->startColumn) || (column > this->endColumn)) {
            throw new \ZExcel\Exception("Column column is out of range (" . this->startColumn . " - " . this->endColumn . ")");
        } elseif (this->onlyExistingCells && !(this->subject->cellExistsByColumnAndRow(column, this->rowIndex))) {
            throw new \ZExcel\Exception("In \"IterateOnlyExistingCells\" mode and Cell does not exist");
        }
        
        let this->position = column;

        return this;
    }

    /**
     * Rewind the iterator to the starting column
     */
    public function rewind()
    {
        let this->position = this->startColumn;
    }

    /**
     * Return the current cell in this worksheet row
     *
     * @return \ZExcel\Cell
     */
    public function current() -> <\ZExcel\Cell>
    {
        return this->subject->getCellByColumnAndRow(this->position, this->rowIndex);
    }

    /**
     * Return the current iterator key
     *
     * @return string
     */
    public function key() -> string
    {
        return \ZExcel\Cell::stringFromColumnIndex(this->position);
    }

    /**
     * Set the iterator to its next value
     */
    public function next()
    {
        do {
            let this->position = this->position + 1;
        } while ((this->onlyExistingCells) &&
            (!this->subject->cellExistsByColumnAndRow(this->position, this->rowIndex)) &&
            (this->position <= this->endColumn));
    }

    /**
     * Set the iterator to its previous value
     *
     * @throws \ZExcel\Exception
     */
    public function prev()
    {
        if (this->position <= this->startColumn) {
            throw new \ZExcel\Exception(
                "Column is already at the beginning of range (" .
                \ZExcel\Cell::stringFromColumnIndex(this->endColumn) . " - " .
                \ZExcel\Cell::stringFromColumnIndex(this->endColumn) . ")"
            );
        }

        do {
            let this->position = this->position - 1;
        } while ((this->onlyExistingCells) &&
            (!this->subject->cellExistsByColumnAndRow(this->position, this->rowIndex)) &&
            (this->position >= this->startColumn));
    }

    /**
     * Indicate if more columns exist in the worksheet range of columns that we're iterating
     *
     * @return boolean
     */
    public function valid() -> boolean
    {
        return this->position <= this->endColumn;
    }

    /**
     * Validate start/end values for "IterateOnlyExistingCells" mode, and adjust if necessary
     *
     * @throws \ZExcel\Exception
     */
    protected function adjustForExistingOnlyRange()
    {
        if (this->onlyExistingCells) {
            while ((!this->subject->cellExistsByColumnAndRow(this->startColumn, this->rowIndex)) &&
                (this->startColumn <= this->endColumn)) {
                let this->startColumn = this->startColumn + 1;
            }
            if (this->startColumn > this->endColumn) {
                throw new \ZExcel\Exception("No cells exist within the specified range");
            }
            while ((!this->subject->cellExistsByColumnAndRow(this->endColumn, this->rowIndex)) &&
                (this->endColumn >= this->startColumn)) {
                let this->endColumn = this->endColumn - 1;
            }
            if (this->endColumn < this->startColumn) {
                throw new \ZExcel\Exception("No cells exist within the specified range");
            }
        }
    }
}
