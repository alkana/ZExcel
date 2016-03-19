namespace ZExcel\Worksheet;

class MemoryDrawing extends Drawing
{
    /* Rendering functions */
    const RENDERING_DEFAULT = "imagepng";
    const RENDERING_PNG     = "imagepng";
    const RENDERING_GIF     = "imagegif";
    const RENDERING_JPEG    = "imagejpeg";

    /* MIME types */
    const MIMETYPE_DEFAULT  = "image/png";
    const MIMETYPE_PNG      = "image/png";
    const MIMETYPE_GIF      = "image/gif";
    const MIMETYPE_JPEG     = "image/jpeg";

    /**
     * Image resource
     *
     * @var resource
     */
    private imageResource;

    /**
     * Rendering function
     *
     * @var string
     */
    private renderingFunction;

    /**
     * Mime type
     *
     * @var string
     */
    private mimeType;

    /**
     * Unique name
     *
     * @var string
     */
    private uniqueName;

    /**
     * Create a new \ZExcel\Worksheet\MemoryDrawing
     */
    public function __construct()
    {
        // Initialise values
        let this->imageResource     = null;
        let this->renderingFunction = self::RENDERING_DEFAULT;
        let this->mimeType          = self::MIMETYPE_DEFAULT;
        let this->uniqueName        = md5(rand(0, 9999). time() . rand(0, 9999));

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get image resource
     *
     * @return resource
     */
    public function getImageResource()
    {
        return this->imageResource;
    }

    /**
     * Set image resource
     *
     * @param    value resource
     * @return \ZExcel\Worksheet\MemoryDrawing
     */
    public function setImageResource(var value = null) -> <\Zexcel\Worksheet\MemoryDrawing>
    {
        let this->imageResource = value;

        if (!is_null(this->imageResource)) {
            // Get width/height
            let this->width  = imagesx(this->imageResource);
            let this->height = imagesy(this->imageResource);
        }
        return this;
    }

    /**
     * Get rendering function
     *
     * @return string
     */
    public function getRenderingFunction()
    {
        return this->renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param string value
     * @return \ZExcel\Worksheet\MemoryDrawing
     */
    public function setRenderingFunction(string value = \ZExcel\Worksheet\MemoryDrawing::RENDERING_DEFAULT) -> <\Zexcel\Worksheet\MemoryDrawing>
    {
        let this->renderingFunction = value;
        
        return this;
    }

    /**
     * Get mime type
     *
     * @return string
     */
    public function getMimeType()
    {
        return this->mimeType;
    }

    /**
     * Set mime type
     *
     * @param string value
     * @return \ZExcel\Worksheet\MemoryDrawing
     */
    public function setMimeType(string value = \ZExcel\Worksheet\MemoryDrawing::MIMETYPE_DEFAULT) -> <\Zexcel\Worksheet\MemoryDrawing>
    {
        let this->mimeType = value;
        
        return this;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        var extension;
        
        let extension = strtolower(this->getMimeType());
        let extension = explode("/", extension);
        let extension = extension[1];

        return this->uniqueName . this->getImageIndex() . "." . extension;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            this->renderingFunction .
            this->mimeType .
            this->uniqueName .
            parent::getHashCode() .
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
