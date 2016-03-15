namespace ZExcel\Worksheet;

class RowDimension extends \ZExcel\Worksheet\Dimension
{
    /**
     * Row index
     *
     * @var int
     */
    private rowIndex;

    /**
     * Row height (in pt)
     *
     * When this is set to a negative value, the row height should be ignored by IWriter
     *
     * @var double
     */
    private height = -1;

     /**
     * ZeroHeight for Row?
     *
     * @var bool
     */
    private zeroHeight = false;

    /**
     * Create a new \ZExcel\Worksheet\RowDimension
     *
     * @param int pIndex Numeric row index
     */
    public function __construct(int pIndex = 0)
    {
        // Initialise values
        let this->rowIndex = pIndex;

        // set dimension as unformatted by default
        parent::__construct(null);
    }

    /**
     * Get Row Index
     *
     * @return int
     */
    public function getRowIndex()
    {
        return this->rowIndex;
    }

    /**
     * Set Row Index
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\RowDimension
     */
    public function setRowIndex(int pValue) -> <\ZExcel\Worksheet\RowDimension>
    {
        let this->rowIndex = pValue;
        
        return this;
    }

    /**
     * Get Row Height
     *
     * @return double
     */
    public function getRowHeight()
    {
        return this->height;
    }

    /**
     * Set Row Height
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\RowDimension
     */
    public function setRowHeight(double pValue = -1) -> <\ZExcel\Worksheet\RowDimension>
    {
        let this->height = pValue;
        
        return this;
    }

    /**
     * Get ZeroHeight
     *
     * @return bool
     */
    public function getZeroHeight()
    {
        return this->zeroHeight;
    }

    /**
     * Set ZeroHeight
     *
     * @param bool pValue
     * @return \ZExcel\Worksheet\RowDimension
     */
    public function setZeroHeight(boolean pValue = false) -> <\ZExcel\Worksheet\RowDimension>
    {
        let this->zeroHeight = pValue;
        
        return this;
    }
}
