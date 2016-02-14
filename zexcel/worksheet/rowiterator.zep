namespace ZExcel\Worksheet;

class RowIterator implements \Iterator
{
    /**
     * \ZExcel\Worksheet to iterate
     *
     * @var \ZExcel\Worksheet
     */
    private subject;

    /**
     * Current iterator position
     *
     * @var int
     */
    private position = 1;

    /**
     * Start position
     *
     * @var int
     */
    private startRow = 1;


    /**
     * End position
     *
     * @var int
     */
    private endRow = 1;


    /**
     * Create a new row iterator
     *
     * @param    \ZExcel\Worksheet    subject    The worksheet to iterate over
     * @param    integer                startRow    The row number at which to start iterating
     * @param    integer                endRow        Optionally, the row number at which to stop iterating
     */
    public function __construct(<\ZExcel\Worksheet> subject = null, int startRow = 1, int endRow = null)
    {
        // Set subject
        let this->subject = subject;
        
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
     * @param integer    startRow    The row number at which to start iterating
     * @return \ZExcel\Worksheet\RowIterator
     */
    public function resetStart(int startRow = 1) -> <\ZExcel\Worksheet\RowIterator>
    {
        let this->startRow = startRow;
        this->seek(startRow);

        return this;
    }

    /**
     * (Re)Set the end row
     *
     * @param integer    endRow    The row number at which to stop iterating
     * @return \ZExcel\Worksheet\RowIterator
     */
    public function resetEnd(endRow = null) -> <\ZExcel\Worksheet\RowIterator>
    {
        let this->endRow = (endRow !== null) ? endRow : this->subject->getHighestRow();

        return this;
    }

    /**
     * Set the row pointer to the selected row
     *
     * @param integer    row    The row number to set the current pointer at
     * @return \ZExcel\Worksheet\RowIterator
     * @throws \ZExcel\Exception
     */
    public function seek(int row = 1) -> <\ZExcel\Worksheet\RowIterator>
    {
        if ((row < this->startRow) || (row > this->endRow)) {
            throw new \ZExcel\Exception("Row row is out of range (" . this->startRow . " - " . this->endRow . ")");
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
     * Return the current row in this worksheet
     *
     * @return \ZExcel\Worksheet\Row
     */
    public function current() -> <\ZExcel\Worksheet\Row>
    {
        return new \ZExcel\Worksheet\Row(this->subject, this->position);
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
        let this->position = this->position + 1;
    }

    /**
     * Set the iterator to its previous value
     */
    public function prev()
    {
        if (this->position <= this->startRow) {
            throw new \ZExcel\Exception("Row is already at the beginning of range (" . this->startRow . " - " . this->endRow . ")");
        }

        let this->position = this->position - 1;
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
}
