namespace ZExcel;

class RichText implements IComparable
{
    /**
     * Rich text elements
     *
     * @var \ZExcel\RichText\ITextElement[]
     */
    private richTextElements;

    /**
     * Create a new \ZExcel\RichText instance
     *
     * @param \ZExcel\Cell pCell
     * @throws \ZExcel\Exception
     */
    public function __construct(<\ZExcel\Cell> pCell = null)
    {
        var objRun;
    
        // Initialise variables
        let this->richTextElements = [];

        // Rich-Text string attached to cell?
        if (pCell !== null) {
            // Add cell text and style
            if (pCell->getValue() != "") {
                let objRun = new \ZExcel\RichText\Run(pCell->getValue());
                objRun->setFont(clone pCell->getParent()->getStyle(pCell->getCoordinate())->getFont());
                this->addText(objRun);
            }

            // Set parent value
            pCell->setValueExplicit(this, \ZExcel\Cell\DataType::TYPE_STRING);
        }
    }

    /**
     * Add text
     *
     * @param \ZExcel\RichText\ITextElement pText Rich text element
     * @throws \ZExcel\Exception
     * @return \ZExcel\RichText
     */
    public function addText(<\ZExcel\RichText\ITextElement> pText = null)
    {
        let this->richTextElements[] = pText;
        return this;
    }

    /**
     * Create text
     *
     * @param string pText Text
     * @return \ZExcel\RichText\TextElement
     * @throws \ZExcel\Exception
     */
    public function createText(string pText = "")
    {
        var objText;
    
        let objText = new \ZExcel\RichText\TextElement(pText);
        this->addText(objText);
        return objText;
    }

    /**
     * Create text run
     *
     * @param string pText Text
     * @return \ZExcel\RichText\Run
     * @throws \ZExcel\Exception
     */
    public function createTextRun(string pText = "")
    {
        var objText;
        
        let objText = new \ZExcel\RichText\Run(pText);
        this->addText(objText);
        return objText;
    }

    /**
     * Get plain text
     *
     * @return string
     */
    public function getPlainText()
    {
        var text, returnValue = "";

        // Loop through all \ZExcel\RichText\ITextElement
        for text in this->richTextElements {
            let returnValue = returnValue + text->getText();
        }

        // Return
        return returnValue;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return this->getPlainText();
    }

    /**
     * Get Rich Text elements
     *
     * @return \ZExcel\RichText\ITextElement[]
     */
    public function getRichTextElements()
    {
        return this->richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param \ZExcel\RichText\ITextElement[] pElements Array of elements
     * @throws \ZExcel\Exception
     * @return \ZExcel\RichText
     */
    public function setRichTextElements(array pElements = null)
    {
        if (is_array(pElements)) {
            let this->richTextElements = pElements;
        } else {
            throw new \ZExcel\Exception("Invalid \\ZExcel\\RichText\\ITextElement[] array passed.");
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
        var element, hashElements = "";
        
        for element in this->richTextElements {
            let hashElements = hashElements + element->getHashCode();
        }

        return md5(hashElements . get_class(this));
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
