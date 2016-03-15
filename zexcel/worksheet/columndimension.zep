namespace ZExcel\Worksheet;

class ColumnDimension extends \ZExcel\Worksheet\Dimension
{
    /**
     * Column index
     *
     * @var int
     */
    private columnIndex;

    /**
     * Column width
     *
     * When this is set to a negative value, the column width should be ignored by IWriter
     *
     * @var double
     */
    private width = -1;

    /**
     * Auto size?
     *
     * @var bool
     */
    private autoSize = false;

    /**
     * Create a new \ZExcel\Worksheet\ColumnDimension
     *
     * @param string pIndex Character column index
     */
    public function __construct(pIndex = "A")
    {
        // Initialise values
        let this->columnIndex = pIndex;

        // set dimension as unformatted by default
        parent::__construct(0);
    }

    /**
     * Get ColumnIndex
     *
     * @return string
     */
    public function getColumnIndex()
    {
        return this->columnIndex;
    }

    /**
     * Set ColumnIndex
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function setColumnIndex(pValue) -> <\ZExcel\Worksheet\ColumnDimension>
    {
        let this->columnIndex = pValue;
        
        return this;
    }

    /**
     * Get Width
     *
     * @return double
     */
    public function getWidth()
    {
        return this->width;
    }

    /**
     * Set Width
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function setWidth(double pValue = -1) -> <\ZExcel\Worksheet\ColumnDimension>
    {
        let this->width = pValue;
        
        return this;
    }

    /**
     * Get Auto Size
     *
     * @return bool
     */
    public function getAutoSize()
    {
        return this->autoSize;
    }

    /**
     * Set Auto Size
     *
     * @param bool pValue
     * @return \ZExcel\Worksheet\ColumnDimension
     */
    public function setAutoSize(boolean pValue = false) -> <\ZExcel\Worksheet\ColumnDimension>
    {
        let this->autoSize = pValue;
        
        return this;
    }
}
