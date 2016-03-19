namespace ZExcel\Worksheet;

use ZExcel\IComparable as ZIComparable;

class BaseDrawing implements ZIComparable
{
    /**
     * Image counter
     *
     * @var int
     */
    private static imageCounter = 0;

    /**
     * Image index
     *
     * @var int
     */
    private imageIndex = 0;

    /**
     * Name
     *
     * @var string
     */
    protected name;

    /**
     * Description
     *
     * @var string
     */
    protected description;

    /**
     * Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    protected worksheet;

    /**
     * Coordinates
     *
     * @var string
     */
    protected coordinates;

    /**
     * Offset X
     *
     * @var int
     */
    protected offsetX;

    /**
     * Offset Y
     *
     * @var int
     */
    protected offsetY;

    /**
     * Width
     *
     * @var int
     */
    protected width;

    /**
     * Height
     *
     * @var int
     */
    protected height;

    /**
     * Proportional resize
     *
     * @var boolean
     */
    protected resizeProportional;

    /**
     * Rotation
     *
     * @var int
     */
    protected rotation;

    /**
     * Shadow
     *
     * @var \ZExcel\Worksheet\Drawing\Shadow
     */
    protected shadow;

    /**
     * Create a new \ZExcel\Worksheet\BaseDrawing
     */
    public function __construct()
    {
        // Initialise values
        let this->name               = "";
        let this->description        = "";
        let this->worksheet          = null;
        let this->coordinates        = "A1";
        let this->offsetX            = 0;
        let this->offsetY            = 0;
        let this->width              = 0;
        let this->height             = 0;
        let this->resizeProportional = true;
        let this->rotation           = 0;
        let this->shadow             = new \ZExcel\Worksheet\Drawing\Shadow();

        // Set image index
        let self::imageCounter = self::imageCounter + 1;
        let this->imageIndex   = self::imageCounter;
    }

    /**
     * Get image index
     *
     * @return int
     */
    public function getImageIndex()
    {
        return this->imageIndex;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        return this->name;
    }

    /**
     * Set Name
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setName(var pValue = "") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->name = pValue;
        
        return this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription()
    {
        return this->description;
    }

    /**
     * Set Description
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setDescription(var pValue = "") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->description = pValue;
        
        return this;
    }

    /**
     * Get Worksheet
     *
     * @return \ZExcel\Worksheet
     */
    public function getWorksheet()
    {
        return this->worksheet;
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
        
        if (is_null(this->worksheet)) {
            // Add drawing to \ZExcel\Worksheet
            let this->worksheet = pValue;
            
            this->worksheet->getCell(this->coordinates);
            this->worksheet->getDrawingCollection()->append(this);
        } else {
            if (pOverrideOld) {
                // Remove drawing from old \ZExcel\Worksheet
                let iteratorr = this->worksheet->getDrawingCollection()->getIterator();

                while (iteratorr->valid()) {
                    if (iteratorr->current()->getHashCode() == this->getHashCode()) {
                        this->worksheet->getDrawingCollection()->offsetUnset( iteratorr->key() );
                        let this->worksheet = null;
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
        return this->coordinates;
    }

    /**
     * Set Coordinates
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setCoordinates(var pValue = "A1") -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->coordinates = pValue;
        
        return this;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX()
    {
        return this->offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setOffsetX(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->offsetX = pValue;
        
        return this;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY()
    {
        return this->offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setOffsetY(var pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->offsetY = pValue;
        
        return this;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth()
    {
        return this->width;
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
        if (this->resizeProportional && pValue != 0) {
            if (this->width != 0) {
                let ratio = this->height / this->width;
            } else {
                let ratio = this->height;
            }
            
            let this->height = round(ratio * pValue);
        }

        // Set width
        let this->width = pValue;

        return this;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight()
    {
        return this->height;
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
        if (this->resizeProportional && pValue != 0) {
            if (this->height != 0) {
                let ratio = this->width / this->height;
            } else {
                let ratio = this->width;
            }
            
            let this->width = round(ratio * pValue);
        }

        // Set height
        let this->height = pValue;

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
        
        if (this->width != 0) {
            let xratio = width / this->width;
         } else {
             let xratio = width;
         }
         
         if (this->height != 0) {
             let yratio = height / this->height;
         } else {
             let yratio = height;
         }
        
        if (this->resizeProportional && !(width == 0 || height == 0)) {
            if ((xratio * this->height) < height) {
                let this->height = ceil(xratio * this->height);
                let this->width  = width;
            } else {
                let this->width  = ceil(yratio * this->width);
                let this->height = height;
            }
        } else {
            let this->width = width;
            let this->height = height;
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
        return this->resizeProportional;
    }

    /**
     * Set ResizeProportional
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setResizeProportional(pValue = true) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->resizeProportional = pValue;
        
        return this;
    }

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation()
    {
        return this->rotation;
    }

    /**
     * Set Rotation
     *
     * @param int pValue
     * @return \ZExcel\Worksheet\BaseDrawing
     */
    public function setRotation(pValue = 0) -> <\ZExcel\Worksheet\BaseDrawing>
    {
        let this->rotation = pValue;
        
        return this;
    }

    /**
     * Get Shadow
     *
     * @return \ZExcel\Worksheet\Drawing\Shadow
     */
    public function getShadow()
    {
        return this->shadow;
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
           let this->shadow = pValue;
           
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
              this->name
            . this->description
            . this->worksheet->getHashCode()
            . this->coordinates
            . this->offsetX
            . this->offsetY
            . this->width
            . this->height
            . this->rotation
            . this->shadow->getHashCode()
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
