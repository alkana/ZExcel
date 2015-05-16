namespace ZExcel;

class Style extends Style\Supervisor implements IComparable
{
	/**
     * Font
     *
     * @var PHPExcel_Style_Font
     */
    protected font;

    /**
     * Fill
     *
     * @var PHPExcel_Style_Fill
     */
    protected fill;

    /**
     * Borders
     *
     * @var PHPExcel_Style_Borders
     */
    protected borders;

    /**
     * Alignment
     *
     * @var PHPExcel_Style_Alignment
     */
    protected alignment;

    /**
     * Number Format
     *
     * @var PHPExcel_Style_NumberFormat
     */
    protected numberFormat;

    /**
     * Conditional styles
     *
     * @var PHPExcel_Style_Conditional[]
     */
    protected conditionalStyles;

    /**
     * Protection
     *
     * @var PHPExcel_Style_Protection
     */
    protected protection;

    /**
     * Index of style in collection. Only used for real style.
     *
     * @var int
     */
    protected index;

    /**
     * Use Quote Prefix when displaying in cell editor. Only used for real style.
     *
     * @var boolean
     */
    protected quotePrefix = false;
    
    protected isSupervisor = false;

    /**
     * Create a new PHPExcel_Style
     *
     * @param boolean isSupervisor Flag indicating if this is a supervisor or not
     *         Leave this value at default unless you understand exactly what
     *    its ramifications are
     * @param boolean isConditional Flag indicating if this is a conditional style or not
     *       Leave this value at default unless you understand exactly what
     *    its ramifications are
     */
    public function __construct(boolean isSupervisor = false, boolean isConditional = false)
    {
        // Supervisor?
        let this->isSupervisor = isSupervisor;

        // Initialise values
        let this->conditionalStyles = [];
        let this->font         = new \ZExcel\Style\Font(isSupervisor, isConditional);
        let this->fill         = new \ZExcel\Style\Fill(isSupervisor, isConditional);
        let this->borders      = new \ZExcel\Style\Borders(isSupervisor, isConditional);
        let this->alignment    = new \ZExcel\Style\Alignment(isSupervisor, isConditional);
        let this->numberFormat = new \ZExcel\Style\NumberFormat(isSupervisor, isConditional);
        let this->protection   = new \ZExcel\Style\Protection(isSupervisor, isConditional);

        // bind parent if we are a supervisor
        if (isSupervisor == true) {
            this->font->bindParent(this);
            this->fill->bindParent(this);
            this->borders->bindParent(this);
            this->alignment->bindParent(this);
            this->numberFormat->bindParent(this);
            this->protection->bindParent(this);
        }
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return PHPExcel_Style
     */
    public function getSharedComponent()
    {
    	var activeSheet, selectedCell, xfIndex = 0;
    	
        let activeSheet = this->getActiveSheet();
        let selectedCell = this->getActiveCell(); // e.g. "A1"

        if (activeSheet->cellExists(selectedCell)) {
            let xfIndex = activeSheet->getCell(selectedCell)->getXfIndex();
        }

        return this->parent->getCellXfByIndex(xfIndex);
    }

    /**
     * Get parent. Only used for style supervisor
     *
     * @return PHPExcel
     */
    public function getParent()
    {
        return this->parent;
    }

    /**
     * Build style array from subcomponents
     *
     * @param array array
     * @return array
     */
    public function getStyleArray(array arry)
    {
        return ["quotePrefix": arry];
    }

    /**
     * Apply styles from array
     *
     * <code>
     * objPHPExcel->getActiveSheet()->getStyle("B2")->applyFromArray(
     *         array(
     *             "font"    => array(
     *                 "name"      => "Arial",
     *                 "bold"      => true,
     *                 "italic"    => false,
     *                 "underline" => PHPExcel_Style_Font::UNDERLINE_DOUBLE,
     *                 "strike"    => false,
     *                 "color"     => array(
     *                     "rgb" => "808080"
     *                 )
     *             ),
     *             "borders" => array(
     *                 "bottom"     => array(
     *                     "style" => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     "color" => array(
     *                         "rgb" => "808080"
     *                     )
     *                 ),
     *                 "top"     => array(
     *                     "style" => PHPExcel_Style_Border::BORDER_DASHDOT,
     *                     "color" => array(
     *                         "rgb" => "808080"
     *                     )
     *                 )
     *             ),
     *             "quotePrefix"    => true
     *         )
     * );
     * </code>
     *
     * @param    array    pStyles    Array containing style information
     * @param     boolean        pAdvanced    Advanced mode for setting borders.
     * @throws    PHPExcel_Exception
     * @return PHPExcel_Style
     */
    public function applyFromArray(array pStyles = null, boolean pAdvanced = true)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Fill
     *
     * @return PHPExcel_Style_Fill
     */
    public function getFill()
    {
        return this->fill;
    }

    /**
     * Get Font
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
     * @param PHPExcel_Style_Font font
     * @return PHPExcel_Style
     */
    public function setFont(<\ZExcel\Style\Font> font)
    {
        let this->font = font;
        return this;
    }

    /**
     * Get Borders
     *
     * @return PHPExcel_Style_Borders
     */
    public function getBorders()
    {
        return this->borders;
    }

    /**
     * Get Alignment
     *
     * @return PHPExcel_Style_Alignment
     */
    public function getAlignment()
    {
        return this->alignment;
    }

    /**
     * Get Number Format
     *
     * @return PHPExcel_Style_NumberFormat
     */
    public function getNumberFormat()
    {
        return this->numberFormat;
    }

    /**
     * Get Conditional Styles. Only used on supervisor.
     *
     * @return PHPExcel_Style_Conditional[]
     */
    public function getConditionalStyles()
    {
        return this->getActiveSheet()->getConditionalStyles(this->getActiveCell());
    }

    /**
     * Set Conditional Styles. Only used on supervisor.
     *
     * @param PHPExcel_Style_Conditional[] pValue Array of condtional styles
     * @return PHPExcel_Style
     */
    public function setConditionalStyles(array pValue = null)
    {
        if (is_array(pValue)) {
            this->getActiveSheet()->setConditionalStyles(this->getSelectedCells(), pValue);
        }
        
        return this;
    }

    /**
     * Get Protection
     *
     * @return PHPExcel_Style_Protection
     */
    public function getProtection()
    {
        return this->protection;
    }

    /**
     * Get quote prefix
     *
     * @return boolean
     */
    public function getQuotePrefix()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getQuotePrefix();
        }
        return this->quotePrefix;
    }

    /**
     * Set quote prefix
     *
     * @param boolean pValue
     */
    public function setQuotePrefix(var pValue)
    {
    	array styleArray = [];
    	
        if (pValue == "") {
            let pValue = false;
        }
        
        if (this->isSupervisor) {
            let styleArray = ["quotePrefix": pValue];
            this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(styleArray);
        } else {
            let this->quotePrefix = pValue;
        }
        
        return this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
    	var conditional;
        string hashConditionals = "";
        
        for conditional in this->conditionalStyles {
            let hashConditionals .= conditional->getHashCode();
        }

        return md5(
            this->fill->getHashCode() .
            this->font->getHashCode() .
            this->borders->getHashCode() .
            this->alignment->getHashCode() .
            this->numberFormat->getHashCode() .
            hashConditionals .
            this->protection->getHashCode() .
            (this->quotePrefix  ? "t" : "f") .
            __CLASS__
        );
    }

    /**
     * Get own index in style collection
     *
     * @return int
     */
    public function getIndex() -> int
    {
        return this->index;
    }

    /**
     * Set own index in style collection
     *
     * @param int pValue
     */
    public function setIndex(int pValue)
    {
        let this->index = pValue;
    }
}
