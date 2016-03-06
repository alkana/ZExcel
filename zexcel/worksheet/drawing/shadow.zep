namespace ZExcel\Worksheet\Drawing;

use ZExcel\IComparable as ZIComparable;

class Shadow implements ZIComparable
{
    /* Shadow alignment */
    const SHADOW_BOTTOM       = "b";
    const SHADOW_BOTTOM_LEFT  = "bl";
    const SHADOW_BOTTOM_RIGHT = "br";
    const SHADOW_CENTER       = "ctr";
    const SHADOW_LEFT         = "l";
    const SHADOW_TOP          = "t";
    const SHADOW_TOP_LEFT     = "tl";
    const SHADOW_TOP_RIGHT    = "tr";

    /**
     * Visible
     *
     * @var boolean
     */
    private _visible;

    /**
     * Blur radius
     *
     * Defaults to 6
     *
     * @var int
     */
    private _blurRadius;

    /**
     * Shadow distance
     *
     * Defaults to 2
     *
     * @var int
     */
    private _distance;

    /**
     * Shadow direction (in degrees)
     *
     * @var int
     */
    private _direction;

    /**
     * Shadow alignment
     *
     * @var int
     */
    private _alignment;

    /**
     * Color
     *
     * @var \ZExcel\Style\Color
     */
    private _color;

    /**
     * Alpha
     *
     * @var int
     */
    private _alpha;

    /**
     * Create a new \ZExcel\Worksheet\Drawing\Shadow
     */
    public function __construct()
    {
        // Initialise values
        let this->_visible    = false;
        let this->_blurRadius = 6;
        let this->_distance   = 2;
        let this->_direction  = 0;
        let this->_alignment  = \ZExcel\Worksheet\Drawing\Shadow::SHADOW_BOTTOM_RIGHT;
        let this->_color      = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_BLACK);
        let this->_alpha      = 50;
    }

    /**
     * Get Visible
     *
     * @return boolean
     */
    public function getVisible()
    {
        return this->_visible;
    }

    /**
     * Set Visible
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setVisible(pValue = false) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_visible = pValue;
        
        return this;
    }

    /**
     * Get Blur radius
     *
     * @return int
     */
    public function getBlurRadius()
    {
        return this->_blurRadius;
    }

    /**
     * Set Blur radius
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setBlurRadius(pValue = 6) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_blurRadius = pValue;
        
        return this;
    }

    /**
     * Get Shadow distance
     *
     * @return int
     */
    public function getDistance()
    {
        return this->_distance;
    }

    /**
     * Set Shadow distance
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setDistance(pValue = 2) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_distance = pValue;
        
        return this;
    }

    /**
     * Get Shadow direction (in degrees)
     *
     * @return int
     */
    public function getDirection()
    {
        return this->_direction;
    }

    /**
     * Set Shadow direction (in degrees)
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setDirection(pValue = 0) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_direction = pValue;
        
        return this;
    }

   /**
     * Get Shadow alignment
     *
     * @return int
     */
    public function getAlignment()
    {
        return this->_alignment;
    }

    /**
     * Set Shadow alignment
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setAlignment(pValue = 0) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_alignment = pValue;
        
        return this;
    }

   /**
     * Get Color
     *
     * @return \ZExcel\Style\Color
     */
    public function getColor()
    {
        return this->_color;
    }

    /**
     * Set Color
     *
     * @param     \ZExcel\Style\Color pValue
     * @throws     \ZExcel\Exception
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setColor(<\ZExcel\Style\Color> pValue = null) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
       let this->_color = pValue;
       
       return this;
    }

   /**
     * Get Alpha
     *
     * @return int
     */
    public function getAlpha()
    {
        return this->_alpha;
    }

    /**
     * Set Alpha
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function setAlpha(pValue = 0) -> <\ZExcel\Worksheet\Drawing\Shadow>
    {
        let this->_alpha = pValue;
        
        return this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
              (this->_visible ? "t" : "f")
            . this->_blurRadius
            . this->_distance
            . this->_direction
            . this->_alignment
            . this->_color->getHashCode()
            . this->_alpha
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
}
