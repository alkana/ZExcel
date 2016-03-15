namespace ZExcel\Worksheet;

abstract class Dimension
{
    /**
     * Visible?
     *
     * @var bool
     */
    private visible = true;

    /**
     * Outline level
     *
     * @var int
     */
    private outlineLevel = 0;

    /**
     * Collapsed
     *
     * @var bool
     */
    private collapsed = false;

    /**
     * Index to cellXf. Null value means row has no explicit cellXf format.
     *
     * @var int|null
     */
    private xfIndex;

    /**
     * Create a new \ZExcel\Worksheet\Dimension
     *
     * @param int pIndex Numeric row index
     */
    public function __construct(int initialValue = null)
    {
        // set dimension as unformatted by default
        let this->xfIndex = initialValue;
    }

    /**
     * Get Visible
     *
     * @return bool
     */
    public function getVisible()
    {
        return this->visible;
    }

    /**
     * Set Visible
     *
     * @param bool pValue
     * @return \ZExcel\Worksheet\Dimension
     */
    public function setVisible(boolean pValue = true) -> <\ZExcel\Worksheet\Dimension>
    {
        let this->visible = pValue;
        
        return this;
    }

    /**
     * Get Outline Level
     *
     * @return int
     */
    public function getOutlineLevel()
    {
        return this->outlineLevel;
    }

    /**
     * Set Outline Level
     *
     * Value must be between 0 and 7
     *
     * @param int pValue
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet\Dimension
     */
    public function setOutlineLevel(int pValue) -> <\ZExcel\Worksheet\Dimension>
    {
        if (pValue < 0 || pValue > 7) {
            throw new \ZExcel\Exception("Outline level must range between 0 and 7.");
        }

        let this->outlineLevel = pValue;
        
        return this;
    }

    /**
     * Get Collapsed
     *
     * @return bool
     */
    public function getCollapsed()
    {
        return this->collapsed;
    }

    /**
     * Set Collapsed
     *
     * @param bool pValue
     * @return \ZExcel\Worksheet\Dimension
     */
    public function setCollapsed(boolean pValue = true) -> <\ZExcel\Worksheet\Dimension>
    {
        let this->collapsed = pValue;
        
        return this;
    }

    /**
     * Get index to cellXf
     *
     * @return int
     */
    public function getXfIndex()
    {
        return this->xfIndex;
    }

    /**
     * Set index to cellXf
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Dimension
     */
    public function setXfIndex(int pValue = 0) -> <\ZExcel\Worksheet\Dimension>
    {
        let this->xfIndex = pValue;
        
        return this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        var vars, key, value;
        
        let vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                let this->{key} = clone value;
            } else {
                let this->{key} = value;
            }
        }
    }
}
