namespace ZExcel;

class Comment implements IComparable
{
    /**
     * Author
     *
     * @var string
     */
    private _author;

    /**
     * Rich text comment
     *
     * @var \ZExcel\RichText
     */
    private _text;

    /**
     * Comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    private _width = "96pt";

    /**
     * Left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    private _marginLeft = "59.25pt";

    /**
     * Top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    private _marginTop = "1.5pt";

    /**
     * Visible
     *
     * @var boolean
     */
    private _visible = false;

    /**
     * Comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @var string
     */
    private _height = "55.5pt";

    /**
     * Comment fill color
     *
     * @var \ZExcel\Style\Color
     */
    private _fillColor;

    /**
     * Alignment
     *
     * @var string
     */
    private _alignment;

    /**
     * Create a new \ZExcel\Comment
     *
     * @throws \ZExcel\Exception
     */
    public function __construct()
    {
        // Initialise variables
        let this->_author    = "Author";
        let this->_text      = new \ZExcel\RichText();
        let this->_fillColor = new \ZExcel\Style\Color("FFFFFFE1");
        let this->_alignment = \ZExcel\Style\Alignment::HORIZONTAL_GENERAL;
    }

    /**
     * Get Author
     *
     * @return string
     */
    public function getAuthor()
    {
        return this->_author;
    }

    /**
     * Set Author
     *
     * @param string pValue
     * @return \ZExcel\Comment
     */
    public function setAuthor(var pValue = "") -> <\ZExcel\Comment>
    {
        let this->_author = pValue;
        
        return this;
    }

    /**
     * Get Rich text comment
     *
     * @return \ZExcel\RichText
     */
    public function getText()
    {
        return this->_text;
    }

    /**
     * Set Rich text comment
     *
     * @param \ZExcel\RichText pValue
     * @return \ZExcel\Comment
     */
    public function setText(<\ZExcel\RichText> pValue) -> <\ZExcel\Comment>
    {
        let this->_text = pValue;
        
        return this;
    }

    /**
     * Get comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getWidth()
    {
        return this->_width;
    }

    /**
     * Set comment width (CSS style, i.e. XXpx or YYpt)
     *
     * @param string value
     * @return \ZExcel\Comment
     */
    public function setWidth(value = "96pt") -> <\ZExcel\Comment>
    {
        let this->_width = value;
        
        return this;
    }

    /**
     * Get comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getHeight()
    {
        return this->_height;
    }

    /**
     * Set comment height (CSS style, i.e. XXpx or YYpt)
     *
     * @param string value
     * @return \ZExcel\Comment
     */
    public function setHeight(value = "55.5pt") -> <\ZExcel\Comment>
    {
        let this->_height = value;
        
        return this;
    }

    /**
     * Get left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getMarginLeft()
    {
        return this->_marginLeft;
    }

    /**
     * Set left margin (CSS style, i.e. XXpx or YYpt)
     *
     * @param string value
     * @return \ZExcel\Comment
     */
    public function setMarginLeft(var value = "59.25pt") -> <\ZExcel\Comment>
    {
        let this->_marginLeft = value;
        
        return this;
    }

    /**
     * Get top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @return string
     */
    public function getMarginTop()
    {
        return this->_marginTop;
    }

    /**
     * Set top margin (CSS style, i.e. XXpx or YYpt)
     *
     * @param string value
     * @return \ZExcel\Comment
     */
    public function setMarginTop(var value = "1.5pt") -> <\ZExcel\Comment>
    {
        let this->_marginTop = value;
        
        return this;
    }

    /**
     * Is the comment visible by default?
     *
     * @return boolean
     */
    public function getVisible()
    {
        return this->_visible;
    }

    /**
     * Set comment default visibility
     *
     * @param boolean value
     * @return \ZExcel\Comment
     */
    public function setVisible(var value = false) -> <\ZExcel\Comment>
    {
        let this->_visible = value;
        
        return this;
    }

    /**
     * Get fill color
     *
     * @return \ZExcel\Style\Color
     */
    public function getFillColor()
    {
        return this->_fillColor;
    }

    /**
     * Set Alignment
     *
     * @param string pValue
     * @return \ZExcel\Comment
     */
    public function setAlignment(var pValue = \ZExcel\Style\Alignment::HORIZONTAL_GENERAL) -> <\ZExcel\Comment>
    {
        let this->_alignment = pValue;
        
        return this;
    }

    /**
     * Get Alignment
     *
     * @return string
     */
    public function getAlignment()
    {
        return this->_alignment;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
              this->_author
            . this->_text->getHashCode()
            . this->_width
            . this->_height
            . this->_marginLeft
            . this->_marginTop
            . (this->_visible ? 1 : 0)
            . this->_fillColor->getHashCode()
            . this->_alignment
            . get_class(this)
        );
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

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return this->_text->getPlainText();
    }
}
