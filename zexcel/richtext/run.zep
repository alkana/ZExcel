namespace ZExcel\RichText;

class Run extends TextElement implements ITextElement
{
    /**
     * Font
     *
     * @var PHPExcel_Style_Font
     */
    private font;

    /**
     * Create a new PHPExcel_RichText_Run instance
     *
     * @param     string        $pText        Text
     */
    public function __construct(string pText = "")
    {
        // Initialise variables
        this->setText(pText);
        let this->font = new \ZExcel\Style\Font();
    }

    /**
     * Get font
     *
     * @return PHPExcel_Style_Font
     */
    public function getFont()
    {
        return this->font;
    }

    /**
     * Set font
     *
     * @param    PHPExcel_Style_Font        $pFont        Font
     * @throws     PHPExcel_Exception
     * @return PHPExcel_RichText_ITextElement
     */
    public function setFont(<\ZExcel\Style\Font> pFont = null)
    {
        let this->font = pFont;
        return this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(this->getText() . this->font->getHashCode() . __CLASS__);
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        var key, value, vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                let this->{key} = clone value;
            } else {
                let this->{key} = value;
            }
        }
    }
}
