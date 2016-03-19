namespace ZExcel\Worksheet;

class Drawing extends BaseDrawing
{
    /**
     * Path
     *
     * @var string
     */
    private path;

    /**
     * Create a new \ZExcel\Worksheet\Drawing
     */
    public function __construct()
    {
        // Initialise values
        let this->path = "";

        // Initialize parent
        parent::__construct();
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
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        var fileName;
        
        let fileName = this->getFilename();
        let fileName = str_replace(" ", "_", fileName);
        
        return str_replace("." . this->getExtension(), "", fileName) . this->getImageIndex() . "." . this->getExtension();
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        var exploded;
        
        let exploded = explode(".", basename(this->path));
        
        return exploded[count(exploded) - 1];
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
     * @return \ZExcel\Worksheet\Drawing
     */
    public function setPath(string pValue = "", boolean pVerifyFile = true)
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
