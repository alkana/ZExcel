namespace ZExcel\Worksheet;

class PageMargins
{
    /**
     * Left
     *
     * @var double
     */
    private left = 0.7;

    /**
     * Right
     *
     * @var double
     */
    private right = 0.7;

    /**
     * Top
     *
     * @var double
     */
    private top = 0.75;

    /**
     * Bottom
     *
     * @var double
     */
    private bottom = 0.75;

    /**
     * Header
     *
     * @var double
     */
    private header = 0.3;

    /**
     * Footer
     *
     * @var double
     */
    private footer = 0.3;

    /**
     * Create a new \ZExcel\Worksheet\PageMargins
     */
    public function __construct()
    {
    }

    /**
     * Get Left
     *
     * @return double
     */
    public function getLeft()
    {
        return this->left;
    }

    /**
     * Set Left
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setLeft(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->left = pValue;
        
        return this;
    }

    /**
     * Get Right
     *
     * @return double
     */
    public function getRight()
    {
        return this->right;
    }

    /**
     * Set Right
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setRight(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->right = pValue;
        
        return this;
    }

    /**
     * Get Top
     *
     * @return double
     */
    public function getTop()
    {
        return this->top;
    }

    /**
     * Set Top
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setTop(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->top = pValue;
        
        return this;
    }

    /**
     * Get Bottom
     *
     * @return double
     */
    public function getBottom()
    {
        return this->bottom;
    }

    /**
     * Set Bottom
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setBottom(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->bottom = pValue;
        
        return this;
    }

    /**
     * Get Header
     *
     * @return double
     */
    public function getHeader()
    {
        return this->header;
    }

    /**
     * Set Header
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setHeader(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->header = pValue;
        
        return this;
    }

    /**
     * Get Footer
     *
     * @return double
     */
    public function getFooter()
    {
        return this->footer;
    }

    /**
     * Set Footer
     *
     * @param double pValue
     * @return \ZExcel\Worksheet\PageMargins
     */
    public function setFooter(double pValue) -> <\ZExcel\Worksheet\PageMargins>
    {
        let this->footer = pValue;
        
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
