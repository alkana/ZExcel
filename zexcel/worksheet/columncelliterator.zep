namespace ZExcel\Worksheet;

class ColumnCellIterator extends CellIterator implements \Iterator
{
    /**
     * Column index
     *
     * @var string
     */
    protected columnIndex;

    /**
     * Start position
     *
     * @var int
     */
    protected startRow = 1;

    /**
     * End position
     *
     * @var int
     */
    protected endRow = 1;

    /**
     * Create a new row iterator
     *
     * @param    \ZExcel\Worksheet    $subject        The worksheet to iterate over
     * @param   string              $columnIndex    The column that we want to iterate
     * @param    integer                $startRow        The row number at which to start iterating
     * @param    integer                $endRow            Optionally, the row number at which to stop iterating
     */
    public function __construct(<\ZExcel\Worksheet> subject = null, string columnIndex = "A", int startRow = 1, int endRow = null)
    {
        // Set subject
        let this->subject = subject;
        let this->columnIndex = \ZExcel\Cell::columnIndexFromString(columnIndex) - 1;
        
        this->resetEnd(endRow);
        this->resetStart(startRow);
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->subject);
    }

    /**
     * (Re)Set the start row and the current row pointer
     *
     * @param integer    $startRow    The row number at which to start iterating
     * @return \ZExcel\Worksheet\ColumnCellIterator
     * @throws \ZExcel\Exception
     */
    public function resetStart(int startRow = 1) -> <\ZExcel\Worksheet\ColumnCellIterator>
    {
        let this->startRow = startRow;
        
        this->adjustForExistingOnlyRange();
        this->seek(startRow);

        return this;
    }

    /**
     * (Re)Set the end row
     *
     * @param integer    $endRow    The row number at which to stop iterating
     * @return \ZExcel\Worksheet\ColumnCellIterator
     * @throws \ZExcel\Exception
     */
    public function resetEnd(endRow = null) -> <\ZExcel\Worksheet\ColumnCellIterator>
    {
        let this->endRow = ((endRow !== null) ? (int) endRow : this->subject->getHighestRow());
        this->adjustForExistingOnlyRange();

        return this;
    }

    /**
     * Set the row pointer to the selected row
     *
     * @param integer    $row    The row number to set the current pointer at
     * @return \ZExcel\Worksheet\ColumnCellIterator
     * @throws \ZExcel\Exception
     */
    public function seek(int row = 1)
    {
        if ((row < this->startRow) || (row > this->endRow)) {
            throw new \ZExcel\Exception("Row " . row . " is out of range (" . this->startRow . " - " . this->endRow . ")");
        } else {
            if (this->onlyExistingCells && !(this->subject->cellExistsByColumnAndRow(this->columnIndex, row))) {
                throw new \ZExcel\Exception("In \"IterateOnlyExistingCells\" mode and Cell does not exist");
            }
        }
        
        let this->position = row;

        return this;
    }

    /**
     * Rewind the iterator to the starting row
     */
    public function rewind()
    {
        let this->position = this->startRow;
    }

    /**
     * Return the current cell in this worksheet column
     *
     * @return \ZExcel\Worksheet\Row
     */
    public function current() -> <\ZExcel\Worksheet\Row>
    {
        return this->subject->getCellByColumnAndRow(this->columnIndex, this->position);
    }

    /**
     * Return the current iterator key
     *
     * @return int
     */
    public function key() -> int
    {
        return this->position;
    }

    /**
     * Set the iterator to its next value
     */
    public function next()
    {
        do {
            let this->position = this->position + 1;
        } while ((this->onlyExistingCells) &&
            (!this->subject->cellExistsByColumnAndRow(this->columnIndex, this->position)) &&
            (this->position <= this->endRow));
    }

    /**
     * Set the iterator to its previous value
     */
    public function prev()
    {
        if (this->position <= this->startRow) {
            throw new \ZExcel\Exception("Row is already at the beginning of range (" . $this->startRow . " - " . this->endRow . ")");
        }

        do {
            let this->position = this->position - 1;
        } while ((this->onlyExistingCells) &&
            (!this->subject->cellExistsByColumnAndRow(this->columnIndex, this->position)) &&
            (this->position >= this->startRow));
    }

    /**
     * Indicate if more rows exist in the worksheet range of rows that we're iterating
     *
     * @return boolean
     */
    public function valid() -> boolean
    {
        return this->position <= this->endRow;
    }

    /**
     * Validate start/end values for "IterateOnlyExistingCells" mode, and adjust if necessary
     *
     * @throws \ZExcel\Exception
     */
    protected function adjustForExistingOnlyRange()
    {
        if (this->onlyExistingCells) {
            while ((!this->subject->cellExistsByColumnAndRow(this->columnIndex, this->startRow)) && (this->startRow <= this->endRow)) {
                let this->startRow = this->startRow + 1;
            }
            if (this->startRow > this->endRow) {
                throw new \ZExcel\Exception("No cells exist within the specified range");
            }
            while ((!this->subject->cellExistsByColumnAndRow(this->columnIndex, this->endRow)) && (this->endRow >= this->startRow)) {
                let this->endRow = this->endRow - 1;
            }
            if (this->endRow < this->startRow) {
                throw new \ZExcel\Exception("No cells exist within the specified range");
            }
        }
    }
}
