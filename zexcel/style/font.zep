namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Font extends Supervisor implements ZIComparable
{
	/* Underline types */
    const UNDERLINE_NONE             = "none";
    const UNDERLINE_DOUBLE           = "double";
    const UNDERLINE_DOUBLEACCOUNTING = "doubleAccounting";
    const UNDERLINE_SINGLE           = "single";
    const UNDERLINE_SINGLEACCOUNTING = "singleAccounting";

    /**
     * Font Name
     *
     * @var string
     */
    protected name = "Calibri";

    /**
     * Font Size
     *
     * @var float
     */
    protected size = 11;

    /**
     * Bold
     *
     * @var boolean
     */
    protected bold = false;

    /**
     * Italic
     *
     * @var boolean
     */
    protected italic = false;

    /**
     * Superscript
     *
     * @var boolean
     */
    protected superScript = false;

    /**
     * Subscript
     *
     * @var boolean
     */
    protected subScript = false;

    /**
     * Underline
     *
     * @var string
     */
    protected underline = self::UNDERLINE_NONE;

    /**
     * Strikethrough
     *
     * @var boolean
     */
    protected strikethrough = false;

    /**
     * Foreground color
     *
     * @var \ZExcel\Style\Color
     */
    protected color;

    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        // Initialise values
        if (isConditional) {
            let this->name = null;
            let this->size = null;
            let this->bold = null;
            let this->italic = null;
            let this->superScript = null;
            let this->subScript = null;
            let this->underline = null;
            let this->strikethrough = null;
            let this->color = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_BLACK, isSupervisor, isConditional);
        } else {
            let this->color = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_BLACK, isSupervisor);
        }
        // bind parent if we are a supervisor
        if (isSupervisor) {
            this->color->bindParent(this, "color");
        }
    }

    public function getSharedComponent()
    {
        return this->parent->getSharedComponent()->getFont();
    }

    public function getStyleArray(array arry)
    {
        return ["font": arry];
    }

    public function applyFromArray(array pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (array_key_exists("name", pStyles)) {
                    this->setName(pStyles["name"]);
                }
                if (array_key_exists("bold", pStyles)) {
                    this->setBold(pStyles["bold"]);
                }
                if (array_key_exists("italic", pStyles)) {
                    this->setItalic(pStyles["italic"]);
                }
                if (array_key_exists("superScript", pStyles)) {
                    this->setSuperScript(pStyles["superScript"]);
                }
                if (array_key_exists("subScript", pStyles)) {
                    this->setSubScript(pStyles["subScript"]);
                }
                if (array_key_exists("underline", pStyles)) {
                    this->setUnderline(pStyles["underline"]);
                }
                if (array_key_exists("strike", pStyles)) {
                    this->setStrikethrough(pStyles["strike"]);
                }
                if (array_key_exists("color", pStyles)) {
                    this->getColor()->applyFromArray(pStyles["color"]);
                }
                if (array_key_exists("size", pStyles)) {
                    this->setSize(pStyles["size"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    public function getName()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getName();
        }
        
        return this->name;
    }

    public function setName(string pValue = "Calibri")
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = "Calibri";
        }
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["name": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->name = pValue;
        }
        
        return this;
    }

    public function getSize()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getSize();
        }
        
        return this->size;
    }

    public function setSize(pValue = 10)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = 10;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["size": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->size = pValue;
        }
        
        return this;
    }

    public function getBold()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getBold();
        }
        
        return this->bold;
    }

    public function setBold(pValue = false)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["bold": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->bold = pValue;
        }
        
        return this;
    }

    public function getItalic()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getItalic();
        }
        
        return this->italic;
    }

    public function setItalic(var pValue = false)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["italic": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->italic = pValue;
        }
        
        return this;
    }

    public function getSuperScript()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getSuperScript();
        }
        
        return this->superScript;
    }

    public function setSuperScript(var pValue = false)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["superScript": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->superScript = pValue;
            let this->subScript = !pValue;
        }
        
        return this;
    }

    public function getSubScript()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getSubScript();
        }
        
        return this->subScript;
    }

    public function setSubScript(var pValue = false)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["subScript": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->subScript = pValue;
            let this->superScript = !pValue;
        }
        
        return this;
    }

    public function getUnderline()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getUnderline();
        }
        
        return this->underline;
    }

    public function setUnderline(pValue = self::UNDERLINE_NONE)
    {
        var styleArray;
        
        if (is_bool(pValue)) {
            let pValue = (pValue) ? self::UNDERLINE_SINGLE : self::UNDERLINE_NONE;
        } elseif (pValue == "") {
            let pValue = self::UNDERLINE_NONE;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["underline": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->underline = pValue;
        }
        
        return this;
    }

    public function getStrikethrough()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getStrikethrough();
        }
        
        return this->strikethrough;
    }

    public function setStrikethrough(var pValue = false)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["strike": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->strikethrough = pValue;
        }
        
        return this;
    }

    public function getColor()
    {
        return this->color;
    }

    public function setColor(<\ZExcel\Style\Color> pValue = null)
    {
        var color, styleArray;
    	
    	let color = pValue;
    	
        // make sure parameter is a real color and not a supervisor
        if (pValue->getIsSupervisor()) {
        	let color = pValue->getSharedComponent();
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getColor()->getStyleArray(["argb": color->getARGB()]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->color = color;
        }
        
        return this;
    }

    public function getHashCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->name .
            this->size .
            (this->bold ? "t" : "f") .
            (this->italic ? "t" : "f") .
            (this->superScript ? "t" : "f") .
            (this->subScript ? "t" : "f") .
            this->underline .
            (this->strikethrough ? "t" : "f") .
            this->color->getHashCode() .
            get_class(this)
        );
    }
}
