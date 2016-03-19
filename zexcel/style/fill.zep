namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Fill extends Supervisor implements ZIComparable
{
    /* Fill types */
    const FILL_NONE                    = "none";
    const FILL_SOLID                   = "solid";
    const FILL_GRADIENT_LINEAR         = "linear";
    const FILL_GRADIENT_PATH           = "path";
    const FILL_PATTERN_DARKDOWN        = "darkDown";
    const FILL_PATTERN_DARKGRAY        = "darkGray";
    const FILL_PATTERN_DARKGRID        = "darkGrid";
    const FILL_PATTERN_DARKHORIZONTAL  = "darkHorizontal";
    const FILL_PATTERN_DARKTRELLIS     = "darkTrellis";
    const FILL_PATTERN_DARKUP          = "darkUp";
    const FILL_PATTERN_DARKVERTICAL    = "darkVertical";
    const FILL_PATTERN_GRAY0625        = "gray0625";
    const FILL_PATTERN_GRAY125         = "gray125";
    const FILL_PATTERN_LIGHTDOWN       = "lightDown";
    const FILL_PATTERN_LIGHTGRAY       = "lightGray";
    const FILL_PATTERN_LIGHTGRID       = "lightGrid";
    const FILL_PATTERN_LIGHTHORIZONTAL = "lightHorizontal";
    const FILL_PATTERN_LIGHTTRELLIS    = "lightTrellis";
    const FILL_PATTERN_LIGHTUP         = "lightUp";
    const FILL_PATTERN_LIGHTVERTICAL   = "lightVertical";
    const FILL_PATTERN_MEDIUMGRAY      = "mediumGray";

    /**
     * Fill type
     *
     * @var string
     */
    protected fillType = \ZExcel\Style\Fill::FILL_NONE;

    /**
     * Rotation
     *
     * @var double
     */
    protected rotation = 0;

    /**
     * Start color
     *
     * @var \ZExcel\Style\Color
     */
    protected startColor;

    /**
     * End color
     *
     * @var \ZExcel\Style\Color
     */
    protected endColor;

    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        // Initialise values
        if (isConditional) {
            let this->fillType = null;
        }
        let this->startColor = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_WHITE, isSupervisor, isConditional);
        let this->endColor = new \ZExcel\Style\Color(\ZExcel\Style\Color::COLOR_BLACK, isSupervisor, isConditional);

        // bind parent if we are a supervisor
        if (isSupervisor) {
            this->startColor->bindParent(this, "startColor");
            this->endColor->bindParent(this, "endColor");
        }
    }

    public function getSharedComponent()
    {
        return this->parent->getSharedComponent()->getFill();
    }

    public function getStyleArray(array arry)
    {
        return ["fill": arry];
    }

    public function applyFromArray(array pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()
                    ->getStyle(this->getSelectedCells())
                    ->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (array_key_exists("type", pStyles)) {
                    this->setFillType(pStyles["type"]);
                }
                if (array_key_exists("rotation", pStyles)) {
                    this->setRotation(pStyles["rotation"]);
                }
                if (array_key_exists("startcolor", pStyles)) {
                    this->getStartColor()->applyFromArray(pStyles["startcolor"]);
                }
                if (array_key_exists("endcolor", pStyles)) {
                    this->getEndColor()->applyFromArray(pStyles["endcolor"]);
                }
                if (array_key_exists("color", pStyles)) {
                    this->getStartColor()->applyFromArray(pStyles["color"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    public function getFillType()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getFillType();
        }
        
        return this->fillType;
    }

    public function setFillType(pValue = \ZExcel\Style\Fill::FILL_NONE)
    {
        var styleArray;
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["type": pValue]);
            this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(styleArray);
        } else {
            let this->fillType = pValue;
        }
        
        return this;
    }

    public function getRotation()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getRotation();
        }
        
        return this->rotation;
    }

    public function setRotation(int pValue = 0)
    {
        var styleArray;
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["rotation": pValue]);
            this->getActiveSheet()
                ->getStyle(this->getSelectedCells())
                ->applyFromArray(styleArray);
        } else {
            let this->rotation = pValue;
        }
        
        return this;
    }

    public function getStartColor()
    {
        return this->startColor;
    }

    public function setStartColor(<\ZExcel\Style\Color> pValue = null)
    {
        var color, styleArray;
        
        let color = pValue;
        
        // make sure parameter is a real color and not a supervisor
        if (pValue->getIsSupervisor()) {
            let color = pValue->getSharedComponent();
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStartColor()->getStyleArray(["argb": color->getARGB()]);
            this->getActiveSheet()
                ->getStyle(this->getSelectedCells())
                ->applyFromArray(styleArray);
        } else {
            let this->startColor = color;
        }
        
        return this;
    }

    public function getEndColor()
    {
        return this->endColor;
    }

    public function setEndColor(<\ZExcel\Style\Color> pValue = null)
    {
        var color, styleArray;
        
        let color = pValue;
        
        // make sure parameter is a real color and not a supervisor
        if (pValue->getIsSupervisor()) {
            let color = pValue->getSharedComponent();
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getEndColor()->getStyleArray(["argb": color->getARGB()]);
            this->getActiveSheet()
                ->getStyle(this->getSelectedCells())
                ->applyFromArray(styleArray);
        } else {
            let this->endColor = color;
        }
        return this;
    }

    public function getHashCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->getFillType() .
            this->getRotation() .
            this->getStartColor()->getHashCode() .
            this->getEndColor()->getHashCode() .
            get_class(this)
        );
    }
}
