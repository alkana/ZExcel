namespace ZExcel\Worksheet;

abstract class CellIterator
{
    /**
     * \ZExcel\Worksheet to iterate
     *
     * @var \ZExcel\Worksheet
     */
    protected subject;

    /**
     * Current iterator position
     *
     * @var mixed
     */
    protected position = null;

    /**
     * Iterate only existing cells
     *
     * @var boolean
     */
    protected onlyExistingCells = false;

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->subject);
    }

    /**
     * Get loop only existing cells
     *
     * @return boolean
     */
    public function getIterateOnlyExistingCells() -> boolean
    {
        return this->onlyExistingCells;
    }

    /**
     * Validate start/end values for "IterateOnlyExistingCells" mode, and adjust if necessary
     *
     * @throws \ZExcel\Exception
     */
    abstract protected function adjustForExistingOnlyRange();

    /**
     * Set the iterator to loop only existing cells
     *
     * @param    boolean        $value
     * @throws \ZExcel\Exception
     */
    public function setIterateOnlyExistingCells(boolean value = true)
    {
        let this->onlyExistingCells = value;

        this->adjustForExistingOnlyRange();
    }
}
