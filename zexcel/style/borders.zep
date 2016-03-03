namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Borders extends Supervisor implements ZIComparable
{
	/* Diagonal directions */
    const DIAGONAL_NONE = 0;
    const DIAGONAL_UP   = 1;
    const DIAGONAL_DOWN = 2;
    const DIAGONAL_BOTH = 3;

    /**
     * Left
     *
     * @var \ZExcel\Style\Border
     */
    protected left;

    /**
     * Right
     *
     * @var \ZExcel\Style\Border
     */
    protected right;

    /**
     * Top
     *
     * @var \ZExcel\Style\Border
     */
    protected top;

    /**
     * Bottom
     *
     * @var \ZExcel\Style\Border
     */
    protected bottom;

    /**
     * Diagonal
     *
     * @var \ZExcel\Style\Border
     */
    protected diagonal;

    /**
     * DiagonalDirection
     *
     * @var int
     */
    protected diagonalDirection;

    /**
     * All borders psedo-border. Only applies to supervisor.
     *
     * @var \ZExcel\Style\Border
     */
    protected allBorders;

    /**
     * Outline psedo-border. Only applies to supervisor.
     *
     * @var \ZExcel\Style\Border
     */
    protected outline;

    /**
     * Inside psedo-border. Only applies to supervisor.
     *
     * @var \ZExcel\Style\Border
     */
    protected inside;

    /**
     * Vertical pseudo-border. Only applies to supervisor.
     *
     * @var \ZExcel\Style\Border
     */
    protected vertical;

    /**
     * Horizontal pseudo-border. Only applies to supervisor.
     *
     * @var \ZExcel\Style\Border
     */
    protected horizontal;

    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        // Initialise values
        let this->left = new \ZExcel\Style\Border(isSupervisor, isConditional);
        let this->right = new \ZExcel\Style\Border(isSupervisor, isConditional);
        let this->top = new \ZExcel\Style\Border(isSupervisor, isConditional);
        let this->bottom = new \ZExcel\Style\Border(isSupervisor, isConditional);
        let this->diagonal = new \ZExcel\Style\Border(isSupervisor, isConditional);
        let this->diagonalDirection = \ZExcel\Style\Borders::DIAGONAL_NONE;

        // Specially for supervisor
        if (isSupervisor) {
            // Initialize pseudo-borders
            let this->allBorders = new \ZExcel\Style\Border(true);
            let this->outline = new \ZExcel\Style\Border(true);
            let this->inside = new \ZExcel\Style\Border(true);
            let this->vertical = new \ZExcel\Style\Border(true);
            let this->horizontal = new \ZExcel\Style\Border(true);

            // bind parent if we are a supervisor
            this->left->bindParent(this, "left");
            this->right->bindParent(this, "right");
            this->top->bindParent(this, "top");
            this->bottom->bindParent(this, "bottom");
            this->diagonal->bindParent(this, "diagonal");
            this->allBorders->bindParent(this, "allBorders");
            this->outline->bindParent(this, "outline");
            this->inside->bindParent(this, "inside");
            this->vertical->bindParent(this, "vertical");
            this->horizontal->bindParent(this, "horizontal");
        }
    }

    public function getSharedComponent()
    {
        return this->parent->getSharedComponent()->getBorders();
    }

    public function getStyleArray(arry)
    {
        return ["borders": arry];
    }

    public function applyFromArray(array pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()
                	->getStyle(this->getSelectedCells())
                	->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (array_key_exists("left", pStyles)) {
                    this->getLeft()->applyFromArray(pStyles["left"]);
                }
                if (array_key_exists("right", pStyles)) {
                    this->getRight()->applyFromArray(pStyles["right"]);
                }
                if (array_key_exists("top", pStyles)) {
                    this->getTop()->applyFromArray(pStyles["top"]);
                }
                if (array_key_exists("bottom", pStyles)) {
                    this->getBottom()->applyFromArray(pStyles["bottom"]);
                }
                if (array_key_exists("diagonal", pStyles)) {
                    this->getDiagonal()->applyFromArray(pStyles["diagonal"]);
                }
                if (array_key_exists("diagonaldirection", pStyles)) {
                    this->setDiagonalDirection(pStyles["diagonaldirection"]);
                }
                if (array_key_exists("allborders", pStyles)) {
                    this->getLeft()->applyFromArray(pStyles["allborders"]);
                    this->getRight()->applyFromArray(pStyles["allborders"]);
                    this->getTop()->applyFromArray(pStyles["allborders"]);
                    this->getBottom()->applyFromArray(pStyles["allborders"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        return this;
    }

    public function getLeft()
    {
        return this->left;
    }

    public function getRight()
    {
        return this->right;
    }

    public function getTop()
    {
        return this->top;
    }

    public function getBottom()
    {
        return this->bottom;
    }

    public function getDiagonal()
    {
        return this->diagonal;
    }

    public function getAllBorders()
    {
        if (!this->isSupervisor) {
            throw new \ZExcel\Exception("Can only get pseudo-border for supervisor.");
        }
        return this->allBorders;
    }

    public function getOutline()
    {
        if (!this->isSupervisor) {
            throw new \ZExcel\Exception("Can only get pseudo-border for supervisor.");
        }
        return this->outline;
    }

    public function getInside()
    {
        if (!this->isSupervisor) {
            throw new \ZExcel\Exception("Can only get pseudo-border for supervisor.");
        }
        return this->inside;
    }

    public function getVertical()
    {
        if (!this->isSupervisor) {
            throw new \ZExcel\Exception("Can only get pseudo-border for supervisor.");
        }
        return this->vertical;
    }

    public function getHorizontal()
    {
        if (!this->isSupervisor) {
            throw new \ZExcel\Exception("Can only get pseudo-border for supervisor.");
        }
        return this->horizontal;
    }

    public function getDiagonalDirection()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getDiagonalDirection();
        }
        return this->diagonalDirection;
    }

    public function setDiagonalDirection(pValue = \ZExcel\Style\Borders::DIAGONAL_NONE)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = \ZExcel\Style\Borders::DIAGONAL_NONE;
        }
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["diagonaldirection": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->diagonalDirection = pValue;
        }
        
        return this;
    }

    public function getHashCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashcode();
        }
        
        return md5(
            this->getLeft()->getHashCode() .
            this->getRight()->getHashCode() .
            this->getTop()->getHashCode() .
            this->getBottom()->getHashCode() .
            this->getDiagonal()->getHashCode() .
            this->getDiagonalDirection() .
            get_class(this)
        );
    }
}
