namespace ZExcel\Worksheet;

class HeaderFooter
{
    /* Header/footer image location */
    const IMAGE_HEADER_LEFT   = "LH";
    const IMAGE_HEADER_CENTER = "CH";
    const IMAGE_HEADER_RIGHT  = "RH";
    const IMAGE_FOOTER_LEFT   = "LF";
    const IMAGE_FOOTER_CENTER = "CF";
    const IMAGE_FOOTER_RIGHT  = "RF";

    /**
     * OddHeader
     *
     * @var string
     */
    private oddHeader = "";

    /**
     * OddFooter
     *
     * @var string
     */
    private oddFooter = "";

    /**
     * EvenHeader
     *
     * @var string
     */
    private evenHeader = "";

    /**
     * EvenFooter
     *
     * @var string
     */
    private evenFooter = "";

    /**
     * FirstHeader
     *
     * @var string
     */
    private firstHeader = "";

    /**
     * FirstFooter
     *
     * @var string
     */
    private firstFooter = "";

    /**
     * Different header for Odd/Even, defaults to false
     *
     * @var boolean
     */
    private differentOddEven = false;

    /**
     * Different header for first page, defaults to false
     *
     * @var boolean
     */
    private differentFirst = false;

    /**
     * Scale with document, defaults to true
     *
     * @var boolean
     */
    private scaleWithDocument = true;

    /**
     * Align with margins, defaults to true
     *
     * @var boolean
     */
    private alignWithMargins = true;

    /**
     * Header/footer images
     *
     * @var \ZExcel\Worksheet\HeaderFooterDrawing[]
     */
    private headerFooterImages = [];

    /**
     * Create a new \ZExcel\Worksheet\HeaderFooter
     */
    public function __construct()
    {
    }

    /**
     * Get OddHeader
     *
     * @return string
     */
    public function getOddHeader()
    {
        return this->oddHeader;
    }

    /**
     * Set OddHeader
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setOddHeader(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->oddHeader = pValue;
        
        return this;
    }

    /**
     * Get OddFooter
     *
     * @return string
     */
    public function getOddFooter()
    {
        return this->oddFooter;
    }

    /**
     * Set OddFooter
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setOddFooter(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->oddFooter = pValue;
        
        return this;
    }

    /**
     * Get EvenHeader
     *
     * @return string
     */
    public function getEvenHeader()
    {
        return this->evenHeader;
    }

    /**
     * Set EvenHeader
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setEvenHeader(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->evenHeader = pValue;
        
        return this;
    }

    /**
     * Get EvenFooter
     *
     * @return string
     */
    public function getEvenFooter()
    {
        return this->evenFooter;
    }

    /**
     * Set EvenFooter
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setEvenFooter(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->evenFooter = pValue;
        
        return this;
    }

    /**
     * Get FirstHeader
     *
     * @return string
     */
    public function getFirstHeader()
    {
        return this->firstHeader;
    }

    /**
     * Set FirstHeader
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setFirstHeader(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->firstHeader = pValue;
        
        return this;
    }

    /**
     * Get FirstFooter
     *
     * @return string
     */
    public function getFirstFooter()
    {
        return this->firstFooter;
    }

    /**
     * Set FirstFooter
     *
     * @param string pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setFirstFooter(string pValue) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->firstFooter = pValue;
        
        return this;
    }

    /**
     * Get DifferentOddEven
     *
     * @return boolean
     */
    public function getDifferentOddEven()
    {
        return this->differentOddEven;
    }

    /**
     * Set DifferentOddEven
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setDifferentOddEven(boolean pValue = false) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->differentOddEven = pValue;
        
        return this;
    }

    /**
     * Get DifferentFirst
     *
     * @return boolean
     */
    public function getDifferentFirst()
    {
        return this->differentFirst;
    }

    /**
     * Set DifferentFirst
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setDifferentFirst(boolean pValue = false) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->differentFirst = pValue;
        
        return this;
    }

    /**
     * Get ScaleWithDocument
     *
     * @return boolean
     */
    public function getScaleWithDocument()
    {
        return this->scaleWithDocument;
    }

    /**
     * Set ScaleWithDocument
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setScaleWithDocument(boolean pValue = true) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->scaleWithDocument = pValue;
        
        return this;
    }

    /**
     * Get AlignWithMargins
     *
     * @return boolean
     */
    public function getAlignWithMargins()
    {
        return this->alignWithMargins;
    }

    /**
     * Set AlignWithMargins
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setAlignWithMargins(boolean pValue = true) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->alignWithMargins = pValue;
        
        return this;
    }

    /**
     * Add header/footer image
     *
     * @param \ZExcel\Worksheet\HeaderFooterDrawing image
     * @param string location
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function addImage(<\ZExcel\Worksheet\HeaderFooterDrawing> image = null, string location = self::IMAGE_HEADER_LEFT) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        let this->headerFooterImages[location] = image;
        
        return this;
    }

    /**
     * Remove header/footer image
     *
     * @param string location
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function removeImage(string location = self::IMAGE_HEADER_LEFT) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        if (isset(this->headerFooterImages[location])) {
            unset(this->headerFooterImages[location]);
        }
        
        return this;
    }

    /**
     * Set header/footer images
     *
     * @param \ZExcel\Worksheet\HeaderFooterDrawing[] images
     * @throws \ZExcel\Exception
     * @return \ZExcel\Worksheet\HeaderFooter
     */
    public function setImages(images) -> <\ZExcel\Worksheet\HeaderFooter>
    {
        if (!is_array(images)) {
            throw new \ZExcel\Exception("Invalid parameter!");
        }

        let this->headerFooterImages = images;
        
        return this;
    }

    /**
     * Get header/footer images
     *
     * @return \ZExcel\Worksheet\HeaderFooterDrawing[]
     */
    public function getImages() -> array
    {
        // Sort array
        array images = [];
        
        if (isset(this->headerFooterImages[self::IMAGE_HEADER_LEFT])) {
            let images[self::IMAGE_HEADER_LEFT] = this->headerFooterImages[self::IMAGE_HEADER_LEFT];
        }
        
        if (isset(this->headerFooterImages[self::IMAGE_HEADER_CENTER])) {
            let images[self::IMAGE_HEADER_CENTER] = this->headerFooterImages[self::IMAGE_HEADER_CENTER];
        }
        
        if (isset(this->headerFooterImages[self::IMAGE_HEADER_RIGHT])) {
            let images[self::IMAGE_HEADER_RIGHT] = this->headerFooterImages[self::IMAGE_HEADER_RIGHT];
        }
        
        if (isset(this->headerFooterImages[self::IMAGE_FOOTER_LEFT])) {
            let images[self::IMAGE_FOOTER_LEFT] = this->headerFooterImages[self::IMAGE_FOOTER_LEFT];
        }
        
        if (isset(this->headerFooterImages[self::IMAGE_FOOTER_CENTER])) {
            let images[self::IMAGE_FOOTER_CENTER] = this->headerFooterImages[self::IMAGE_FOOTER_CENTER];
        }
        
        if (isset(this->headerFooterImages[self::IMAGE_FOOTER_RIGHT])) {
            let images[self::IMAGE_FOOTER_RIGHT] = this->headerFooterImages[self::IMAGE_FOOTER_RIGHT];
        }
        
        let this->headerFooterImages = images;

        return this->headerFooterImages;
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
