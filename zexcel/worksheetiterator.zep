namespace ZExcel;

class WorksheetIterator implements \Iterator
{
    /**
     * Spreadsheet to iterate
     *
     * @var PHPExcel
     */
    private subject;

    /**
     * Current iterator position
     *
     * @var int
     */
    private position = 0;

    /**
     * Create a new worksheet iterator
     *
     * @param PHPExcel         $subject
     */
    public function __construct(<\ZExcel> subject = null)
    {
        // Set subject
        let this->subject = subject;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset(this->subject);
    }
    
    public function rewind()
    {
        let this->position = 0;
    }
    
    public function current()
    {
       return this->subject->getSheet(this->position);
    }
    
    public function key()
    {
       return this->position;
    }
    
    public function next()
    {
       let this->position = this->position + 1;
    }
    
    public function valid()
    {
        return this->position < this->subject->getSheetCount();
    }
}
