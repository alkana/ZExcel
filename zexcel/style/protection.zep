namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Protection extends Supervisor implements ZIComparable
{
    /** Protection styles */
    const PROTECTION_INHERIT      = "inherit";
    const PROTECTION_PROTECTED    = "protected";
    const PROTECTION_UNPROTECTED  = "unprotected";

    /**
     * Locked
     *
     * @var string
     */
    protected locked;

    /**
     * Hidden
     *
     * @var string
     */
    protected hidden;

    /**
     * Create a new \ZExcel\Style\Protection
     *
     * @param    boolean    isSupervisor    Flag indicating if this is a supervisor or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     * @param    boolean    isConditional    Flag indicating if this is a conditional style or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     */
    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        parent::__construct(isSupervisor);

        // Initialise values
        if (!isConditional) {
            let this->locked = self::PROTECTION_INHERIT;
            let this->hidden = self::PROTECTION_INHERIT;
        }
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return \ZExcel\Style\Protection
     */
    public function getSharedComponent()
    {
        return this->parent->getSharedComponent()->getProtection();
    }

    /**
     * Build style array from subcomponents
     *
     * @param array array
     * @return array
     */
    public function getStyleArray(array arry)
    {
        return ["protection": arry];
    }

    /**
     * Apply styles from array
     *
     * <code>
     * objPHPExcel->getActiveSheet()->getStyle("B2")->getLocked()->applyFromArray(
     *        [
     *            "locked": true,
     *            "hidden": false
     *        ]
     * );
     * </code>
     *
     * @param    array    pStyles    Array containing style information
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Style\Protection
     */
    public function applyFromArray(pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (isset(pStyles["locked"])) {
                    this->setLocked(pStyles["locked"]);
                }
                if (isset(pStyles["hidden"])) {
                    this->setHidden(pStyles["hidden"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    /**
     * Get locked
     *
     * @return string
     */
    public function getLocked()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getLocked();
        }
        return this->locked;
    }

    /**
     * Set locked
     *
     * @param string pValue
     * @return \ZExcel\Style\Protection
     */
    public function setLocked(pValue = self::PROTECTION_INHERIT)
    {
        var styleArray;
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["locked": pValue]);
            this->getActiveSheet()
                ->getStyle(this->getSelectedCells())
                ->applyFromArray(styleArray);
        } else {
            let this->locked = pValue;
        }
        return this;
    }

    /**
     * Get hidden
     *
     * @return string
     */
    public function getHidden()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHidden();
        }
        return this->hidden;
    }

    /**
     * Set hidden
     *
     * @param string pValue
     * @return \ZExcel\Style\Protection
     */
    public function setHidden(pValue = self::PROTECTION_INHERIT)
    {
        var styleArray;
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["hidden": pValue]);
            this->getActiveSheet()
                ->getStyle(this->getSelectedCells())
                ->applyFromArray(styleArray);
        } else {
            let this->hidden = pValue;
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
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->locked .
            this->hidden .
            get_class(this)
        );
    }
}
