namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Border extends Supervisor implements ZIComparable
{
	/* Border style */
    const BORDER_NONE             = "none";
    const BORDER_DASHDOT          = "dashDot";
    const BORDER_DASHDOTDOT       = "dashDotDot";
    const BORDER_DASHED           = "dashed";
    const BORDER_DOTTED           = "dotted";
    const BORDER_DOUBLE           = "double";
    const BORDER_HAIR             = "hair";
    const BORDER_MEDIUM           = "medium";
    const BORDER_MEDIUMDASHDOT    = "mediumDashDot";
    const BORDER_MEDIUMDASHDOTDOT = "mediumDashDotDot";
    const BORDER_MEDIUMDASHED     = "mediumDashed";
    const BORDER_SLANTDASHDOT     = "slantDashDot";
    const BORDER_THICK            = "thick";
    const BORDER_THIN             = "thin";
    
    /**
     * Border style
     *
     * @var string
     */
    protected borderStyle = \ZExcel\Style\Border::BORDER_NONE;

    /**
     * Border color
     *
     * @var \ZExcel\Style\Color
     */
    protected color;

    /**
     * Parent property name
     *
     * @var string
     */
    protected parentPropertyName;
    
    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        // Initialise values
        let this->color = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_BLACK, isSupervisor);

        // bind parent if we are a supervisor
        if (isSupervisor) {
            this->color->bindParent(this, "color");
        }
    }
    
    public function bindParent(<\ZExcel\Style\Border> parent, string parentPropertyName = null)
    {
        let this->parent = parent;
        let this->parentPropertyName = parentPropertyName;
        
        return this;
    }

    public function getSharedComponent()
    {
        switch (this->parentPropertyName) {
            case "allBorders":
            case "horizontal":
            case "inside":
            case "outline":
            case "vertical":
                throw new \ZExcel\Exception("Cannot get shared component for a pseudo-border.");
                break;
            case "bottom":
                return this->parent->getSharedComponent()->getBottom();
            case "diagonal":
                return this->parent->getSharedComponent()->getDiagonal();
            case "left":
                return this->parent->getSharedComponent()->getLeft();
            case "right":
                return this->parent->getSharedComponent()->getRight();
            case "top":
                return this->parent->getSharedComponent()->getTop();
        }
    }

    public function getStyleArray(array arry)
    {
    	array output;
    
        switch (this->parentPropertyName) {
            case "allBorders":
            case "bottom":
            case "diagonal":
            case "horizontal":
            case "inside":
            case "left":
            case "outline":
            case "right":
            case "top":
            case "vertical":
                let key = strtolower("vertical");
                break;
        }
        
        let output[key] = arry;
        
        return this->parent->getStyleArray(output);
    }

    public function applyFromArray(array pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()
                	->getStyle(this->getSelectedCells())
                	->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (isset(pStyles["style"])) {
                    this->setBorderStyle(pStyles["style"]);
                }
                if (isset(pStyles["color"])) {
                    this->getColor()->applyFromArray(pStyles["color"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    public function getBorderStyle()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getBorderStyle();
        }
        return this->borderStyle;
    }

    public function setBorderStyle(var pValue = \ZExcel\Style\Border::BORDER_NONE)
    {

        if (empty(pValue)) {
            let pValue = \ZExcel\Style\Border::BORDER_NONE;
        } elseif (is_bool(pValue) && pValue) {
            let pValue = \ZExcel\Style\Border::BORDER_MEDIUM;
        }
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["style": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->borderStyle = pValue;
        }
        return this;
    }

    public function getColor()
    {
        return this->color;
    }

    public function setColor(<\ZExcel\Style\Color> pValue = null)
    {
    	var color;
    	
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
            this->borderStyle .
            this->color->getHashCode() .
            __CLASS__
        );
    }
}
