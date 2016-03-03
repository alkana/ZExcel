namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Alignment extends Supervisor implements ZIComparable
{
	/* Horizontal alignment styles */
    const HORIZONTAL_GENERAL           = "general";
    const HORIZONTAL_LEFT              = "left";
    const HORIZONTAL_RIGHT             = "right";
    const HORIZONTAL_CENTER            = "center";
    const HORIZONTAL_CENTER_CONTINUOUS = "centerContinuous";
    const HORIZONTAL_JUSTIFY           = "justify";
    const HORIZONTAL_FILL              = "fill";
    const HORIZONTAL_DISTRIBUTED       = "distributed";        // Excel2007 only

    /* Vertical alignment styles */
    const VERTICAL_BOTTOM      = "bottom";
    const VERTICAL_TOP         = "top";
    const VERTICAL_CENTER      = "center";
    const VERTICAL_JUSTIFY     = "justify";
    const VERTICAL_DISTRIBUTED = "distributed";        // Excel2007 only

    /* Read order */
    const READORDER_CONTEXT = 0;
    const READORDER_LTR     = 1;
    const READORDER_RTL     = 2;

    /**
     * Horizontal alignment
     *
     * @var string
     */
    protected horizontal = \ZExcel\Style\Alignment::HORIZONTAL_GENERAL;

    /**
     * Vertical alignment
     *
     * @var string
     */
    protected vertical = \ZExcel\Style\Alignment::VERTICAL_BOTTOM;

    /**
     * Text rotation
     *
     * @var integer
     */
    protected textRotation = 0;

    /**
     * Wrap text
     *
     * @var boolean
     */
    protected wrapText = false;

    /**
     * Shrink to fit
     *
     * @var boolean
     */
    protected shrinkToFit = false;

    /**
     * Indent - only possible with horizontal alignment left and right
     *
     * @var integer
     */
    protected indent = 0;

    /**
     * Read order
     *
     * @var integer
     */
    protected readorder = 0;
    
    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        if (isConditional == true) {
            let this->horizontal   = null;
            let this->vertical     = null;
            let this->textRotation = null;
        }
    }
    
    public function getSharedComponent() -> <\ZExcel\Style\Alignment>
    {
        return this->parent->getSharedComponent()->getAlignment();
    }
    
    public function getStyleArray(array arry) -> array
    {
        return ["alignment": arry];
    }
    
    public function applyFromArray(array pStyles = null) -> <\ZExcel\Style\Alignment>
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor == true) {
                this->getActiveSheet()
                	->getStyle(this->getSelectedCells())
                    ->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (isset(pStyles["horizontal"])) {
                    this->setHorizontal(pStyles["horizontal"]);
                }
                if (isset(pStyles["vertical"])) {
                    this->setVertical(pStyles["vertical"]);
                }
                if (isset(pStyles["rotation"])) {
                    this->setTextRotation(pStyles["rotation"]);
                }
                if (isset(pStyles["wrap"])) {
                    this->setWrapText(pStyles["wrap"]);
                }
                if (isset(pStyles["shrinkToFit"])) {
                    this->setShrinkToFit(pStyles["shrinkToFit"]);
                }
                if (isset(pStyles["indent"])) {
                    this->setIndent(pStyles["indent"]);
                }
                if (isset(pStyles["readorder"])) {
                    this->setReadorder(pStyles["readorder"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }
    
    public function getHorizontal() -> string
    {
        if (this->isSupervisor == true) {
            return this->getSharedComponent()->getHorizontal();
        }
        
        return this->horizontal;
    }
    
    public function setHorizontal(var pValue = \ZExcel\Style\Alignment::HORIZONTAL_GENERAL) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = \ZExcel\Style\Alignment::HORIZONTAL_GENERAL;
        }

        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["horizontal": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->horizontal = pValue;
        }
        
        return this;
    }
    
    public function getVertical() -> string
    {
        if (this->isSupervisor == true) {
            return this->getSharedComponent()->getVertical();
        }
        return this->vertical;
    }
    
    public function setVertical(var pValue = \ZExcel\Style\Alignment::VERTICAL_BOTTOM) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = \ZExcel\Style\Alignment::VERTICAL_BOTTOM;
        }

        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["vertical": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->vertical = pValue;
        }
        
        return this;
    }
    
    public function getTextRotation() -> int
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getTextRotation();
        }
        return this->textRotation;
    }
    
    public function setTextRotation(int pValue = 0) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        // Excel2007 value 255 => PHPExcel value -165
        if (pValue == 255) {
            let pValue = -165;
        }

        // Set rotation
        if ((pValue >= -90 && pValue <= 90) || pValue == -165) {
            if (this->isSupervisor == true) {
                let styleArray = this->getStyleArray(["rotation": pValue]);
                this->getActiveSheet()
                	->getStyle(this->getSelectedCells())
                	->applyFromArray(styleArray);
            } else {
                let this->textRotation = pValue;
            }
        } else {
            throw new \ZExcel\Exception("Text rotation should be a value between -90 and 90.");
        }

        return this;
    }
    
    public function getWrapText() -> boolean
    {
        if (this->isSupervisor == true) {
            return this->getSharedComponent()->getWrapText();
        }
        return this->wrapText;
    }
    
    public function setWrapText(var pValue = false) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["wrap": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->wrapText = pValue;
        }
        
        return this;
    }
    
    public function getShrinkToFit() -> boolean
    {
        if (this->isSupervisor == null) {
            return this->getSharedComponent()->getShrinkToFit();
        }
        
        return this->shrinkToFit;
    }
    
    public function setShrinkToFit(var pValue = false) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = false;
        }
        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["shrinkToFit": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->shrinkToFit = pValue;
        }
        
        return this;
    }
    
    public function getIndent() -> int
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getIndent();
        }
        return this->indent;
    }
    
    public function setIndent(int pValue = 0) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue > 0) {
            if (this->getHorizontal() != self::HORIZONTAL_GENERAL &&
                this->getHorizontal() != self::HORIZONTAL_LEFT &&
                this->getHorizontal() != self::HORIZONTAL_RIGHT) {
                let pValue = 0; // indent not supported
            }
        }
        
        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["indent": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->indent = pValue;
        }
        
        return this;
    }
    
    public function getReadorder() -> int
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getReadorder();
        }
        return this->readorder;
    }
    
    public function setReadorder(int pValue = 0) -> <\ZExcel\Style\Alignment>
    {
        var styleArray;
        
        if (pValue < 0 || pValue > 2) {
            let pValue = 0;
        }
        
        if (this->isSupervisor == true) {
            let styleArray = this->getStyleArray(["readorder": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->readorder = pValue;
        }
        
        return this;
    }
    
    public function getHashCode() -> string
    {
        if (this->isSupervisor == true) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->horizontal .
            this->vertical .
            this->textRotation .
            (this->wrapText ? 't' : 'f') .
            (this->shrinkToFit ? 't' : 'f') .
            this->indent .
            this->readorder .
            get_class(this)
        );
    }
}
