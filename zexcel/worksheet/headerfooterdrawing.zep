namespace ZExcel\Worksheet;

class HeaderFooterDrawing extends Drawing
{
    /**
     * Path
     *
     * @var string
     */
    private path;

    /**
     * Name
     *
     * @var string
     */
    protected name;

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
     * Create a new \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function __construct()
    {
        // Initialise values
        let this->path               = "";
        let this->name               = "";
        let this->offsetX            = 0;
        let this->offsetY            = 0;
        let this->width              = 0;
        let this->height             = 0;
        let this->resizeProportional = true;
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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setName(string pValue = "") -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        let this->name = pValue;
        
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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setOffsetX(int pValue = 0) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setOffsetY(int pValue = 0) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setWidth(int pValue = 0) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        var ratio;
        
        // Resize proportional?
        if (this->resizeProportional && pValue != 0) {
            let ratio = this->width / this->height;
            let this->height = (int) round(ratio * pValue);
        }

        // Set width
        let this->width = (int) pValue;

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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setHeight(int pValue = 0) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        var ratio;
        
        // Resize proportional?
        if (this->resizeProportional && pValue != 0) {
            let ratio = this->width / this->height;
            let this->width = (int) round(ratio * pValue);
        }

        // Set height
        let this->height = (int) pValue;

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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setWidthAndHeight(double width = 0, double height = 0) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        double xratio, yratio;
        
        let xratio = width / this->width;
        let yratio = height / this->height;
        
        if (this->resizeProportional && !(width == 0 || height == 0)) {
            if ((xratio * this->height) < height) {
                let this->width  = (int) width;
                let this->height = (int) ceil(xratio * this->height);
            } else {
                let this->width  = (int) ceil(yratio * this->width);
                let this->height = (int) height;
            }
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
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setResizeProportional(boolean pValue = true) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        let this->resizeProportional = pValue;
        
        return this;
    }

    /**
     * Get Filename
     *
     * @return string
     */
    public function getFilename()
    {
        return basename(this->path);
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        var parts;
        
        let parts = explode(".", basename(this->path));
        
        return end(parts);
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return this->path;
    }

    /**
     * Set Path
     *
     * @param     string         pValue            File path
     * @param     boolean        pVerifyFile    Verify file
     * @throws     \ZExcel\Exception
     * @return \ZExcel\Worksheet\HeaderFooterDrawing
     */
    public function setPath(string pValue = "", boolean pVerifyFile = true) -> <\ZExcel\Worksheet\HeaderFooterDrawing>
    {
        var tmp;
        
        if (pVerifyFile) {
            if (file_exists(pValue)) {
                let this->path = pValue;

                if (this->width == 0 && this->height == 0) {
                    let tmp = getimagesize(pValue);
                    // Get width/height
                    let this->width = tmp[0];
                    let this->height = tmp[1];
                }
            } else {
                throw new \ZExcel\Exception("File pValue not found!");
            }
        } else {
            let this->path = pValue;
        }
        
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
            this->path .
            this->name .
            this->offsetX .
            this->offsetY .
            this->width .
            this->height .
            get_class(this)
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
