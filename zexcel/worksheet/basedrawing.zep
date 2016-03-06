namespace ZExcel\Worksheet;

use ZExcel\IComparable as ZIComparable;

class BaseDrawing implements ZIComparable
{
    /**
     * Image counter
     *
     * @var int
     */
    private static _imageCounter = 0;

    /**
     * Image index
     *
     * @var int
     */
    private _imageIndex = 0;

    /**
     * Name
     *
     * @var string
     */
    protected _name;

    /**
     * Description
     *
     * @var string
     */
    protected _description;

    /**
     * Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    protected _worksheet;

    /**
     * Coordinates
     *
     * @var string
     */
    protected _coordinates;

    /**
     * Offset X
     *
     * @var int
     */
    protected _offsetX;

    /**
     * Offset Y
     *
     * @var int
     */
    protected _offsetY;

    /**
     * Width
     *
     * @var int
     */
    protected _width;

    /**
     * Height
     *
     * @var int
     */
    protected _height;

    /**
     * Proportional resize
     *
     * @var boolean
     */
    protected _resizeProportional;

    /**
     * Rotation
     *
     * @var int
     */
    protected _rotation;

    /**
     * Shadow
     *
     * @var \ZExcel\Worksheet\Drawing\Shadow
     */
    protected _shadow;

    /**
     * Create a new \ZExcel\Worksheet\BaseDrawing
     */
    public function __construct()
    {
        // Initialise values
        let this->_name               = "";
        let this->_description        = "";
        let this->_worksheet          = null;
        let this->_coordinates        = "A1";
        let this->_offsetX            = 0;
        let this->_offsetY            = 0;
        let this->_width              = 0;
        let this->_height             = 0;
        let this->_resizeProportional = true;
        let this->_rotation           = 0;
        let this->_shadow             = new \ZExcel\Worksheet\Drawing\Shadow();

        // Set image index
        let self::_imageCounter = self::_imageCounter + 1;
        let this->_imageIndex   = self::_imageCounter;
    }

    /**
     * Get image index
     *
     * @return int
     */
    public function getImageIndex()
    {
        return this->_imageIndex;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        return this->_name;
    }

    /**
     * Set Name
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setName(var pValue = "") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_name = pValue;
        
        return this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription()
    {
        return this->_description;
    }

    /**
     * Set Description
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setDescription(var pValue = "") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_description = pValue;
        
        return this;
    }

    /**
     * Get Worksheet
     *
     * @return \ZExcel\Worksheet
     */
    public function getWorksheet()
    {
        return this->_worksheet;
    }

    /**
     * Set Worksheet
     *
     * @param     \ZExcel\Worksheet     pValue
     * @param     bool                pOverrideOld    If a Worksheet has already been assigned, overwrite it and remove image from old Worksheet?
     * @throws     \ZExcel\Exception
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setWorksheet(<\ZExcel\Worksheet> pValue = null, var pOverrideOld = false) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        var iteratorr;
        
        if (is_null(this->_worksheet)) {
            // Add drawing to \ZExcel\Worksheet
            let this->_worksheet = pValue;
            
            this->_worksheet->getCell(this->_coordinates);
            this->_worksheet->getDrawingCollection()->append(this);
        } else {
            if (pOverrideOld) {
                // Remove drawing from old \ZExcel\Worksheet
                let iteratorr = this->_worksheet->getDrawingCollection()->getIterator();

                while (iteratorr->valid()) {
                    if (iteratorr->current()->getHashCode() == this->getHashCode()) {
                        this->_worksheet->getDrawingCollection()->offsetUnset( iteratorr->key() );
                        let this->_worksheet = null;
                        break;
                    }
                }

                // Set new \ZExcel\Worksheet
                this->setWorksheet(pValue);
            } else {
                throw new \ZExcel\Exception("A \ZExcel\Worksheet has already been assigned. Drawings can only exist on one \ZExcel\Worksheet.");
            }
        }
        
        return this;
    }

    /**
     * Get Coordinates
     *
     * @return string
     */
    public function getCoordinates()
    {
        return this->_coordinates;
    }

    /**
     * Set Coordinates
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setCoordinates(var pValue = "A1") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_coordinates = pValue;
        
        return this;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX()
    {
        return this->_offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setOffsetX(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_offsetX = pValue;
        
        return this;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY()
    {
        return this->_offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setOffsetY(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_offsetY = pValue;
        
        return this;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth()
    {
        return this->_width;
    }

    /**
     * Set Width
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setWidth(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        var ratio;
        
        // Resize proportional?
        if (this->_resizeProportional && pValue != 0) {
            if (this->_width != 0) {
                let ratio = this->_height / this->_width;
            } else {
                let ratio = this->_height;
            }
            
            let this->_height = round(ratio * pValue);
        }

        // Set width
        let this->_width = pValue;

        return this;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight()
    {
        return this->_height;
    }

    /**
     * Set Height
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setHeight(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        var ratio;
        
        // Resize proportional?
        if (this->_resizeProportional && pValue != 0) {
            if (this->_height != 0) {
                let ratio = this->_width / this->_height;
            } else {
                let ratio = this->_width;
            }
            
            let this->_width = round(ratio * pValue);
        }

        // Set height
        let this->_height = pValue;

        return this;
    }

    /**
     * Set width and height with proportional resize
     * Example:
     * <code>
     * objDrawing->setResizeProportional(true);
     * objDrawing->setWidthAndHeight(160,120);
     * </code>
     *
     * @author Vincent@luo MSN:kele_100@hotmail.com
     * @param int width
     * @param int height
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setWidthAndHeight(var width = 0, var height = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        var xratio, yratio;
        
        if (this->_width != 0) {
            let xratio = width / this->_width;
         } else {
             let xratio = width;
         }
         
         if (this->_height != 0) {
             let yratio = height / this->_height;
         } else {
             let yratio = height;
         }
        
        if (this->_resizeProportional && !(width == 0 || height == 0)) {
            if ((xratio * this->_height) < height) {
                let this->_height = ceil(xratio * this->_height);
                let this->_width  = width;
            } else {
                let this->_width    = ceil(yratio * this->_width);
                let this->_height    = height;
            }
        } else {
            let this->_width = width;
            let this->_height = height;
        }

        return this;
    }

    /**
     * Get ResizeProportional
     *
     * @return boolean
     */
    public function getResizeProportional()
    {
        return this->_resizeProportional;
    }

    /**
     * Set ResizeProportional
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setResizeProportional(pValue = true) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_resizeProportional = pValue;
        
        return this;
    }

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation()
    {
        return this->_rotation;
    }

    /**
     * Set Rotation
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setRotation(pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->_rotation = pValue;
        
        return this;
    }

    /**
     * Get Shadow
     *
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function getShadow()
    {
        return this->_shadow;
    }

    /**
     * Set Shadow
     *
     * @param     \ZExcel\Worksheet\Drawing\Shadow pValue
     * @throws     \ZExcel\Exception
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setShadow(<\ZExcel\Worksheet\Drawing\Shadow> pValue = null) -> <\ZExcel\Worksheet\BaseDrawing>
    {
           let this->_shadow = pValue;
           
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
              this->_name
            . this->_description
            . this->_worksheet->getHashCode()
            . this->_coordinates
            . this->_offsetX
            . this->_offsetY
            . this->_width
            . this->_height
            . this->_rotation
            . this->_shadow->getHashCode()
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
