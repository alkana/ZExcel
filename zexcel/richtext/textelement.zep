namespace ZExcel\RichText;

class TextElement implements ITextElement
{
    /**
     * Text
     *
     * @var string
     */
    private text;

    /**
     * Create a new PHPExcel_RichText_TextElement instance
     *
     * @param     string        $pText        Text
     */
    public function __construct(string pText = "")
    {
        // Initialise variables
        let this->text = pText;
    }

    /**
     * Get text
     *
     * @return string    Text
     */
    public function getText()
    {
        return this->text;
    }

    /**
     * Set text
     *
     * @param     $pText string    Text
     * @return PHPExcel_RichText_ITextElement
     */
    public function setText(string pText = "")
    {
        let this->text = pText;
        return this;
    }

    /**
     * Get font
     *
     * @return PHPExcel_Style_Font
     */
    public function getFont()
    {
        return null;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(this->text . __CLASS__);
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
