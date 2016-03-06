namespace ZExcel;

class Style extends Style\Supervisor implements IComparable
{
	/**
     * Font
     *
     * @var \ZExcel\Style_Font
     */
    protected font;

    /**
     * Fill
     *
     * @var \ZExcel\Style_Fill
     */
    protected fill;

    /**
     * Borders
     *
     * @var \ZExcel\Style_Borders
     */
    protected borders;

    /**
     * Alignment
     *
     * @var \ZExcel\Style_Alignment
     */
    protected alignment;

    /**
     * Number Format
     *
     * @var \ZExcel\Style_NumberFormat
     */
    protected numberFormat;

    /**
     * Conditional styles
     *
     * @var \ZExcel\Style_Conditional[]
     */
    protected conditionalStyles;

    /**
     * Protection
     *
     * @var \ZExcel\Style_Protection
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
     * Create a new \ZExcel\Style
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
     * @return \ZExcel\Style
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
     *                 "underline" => \ZExcel\Style_Font::UNDERLINE_DOUBLE,
     *                 "strike"    => false,
     *                 "color"     => array(
     *                     "rgb" => "808080"
     *                 )
     *             ),
     *             "borders" => array(
     *                 "bottom"     => array(
     *                     "style" => \ZExcel\Style_Border::BORDER_DASHDOT,
     *                     "color" => array(
     *                         "rgb" => "808080"
     *                     )
     *                 ),
     *                 "top"     => array(
     *                     "style" => \ZExcel\Style_Border::BORDER_DASHDOT,
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
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Style
     */
    public function applyFromArray(array pStyles = null, boolean pAdvanced = true)
    {
        var pRange, rangeA, rangeB, rangeStart, rangeEnd, tmp, component,
            x, y, xMax, yMax, colStart, colEnd, edges, rangee,
            rowStart, rowEnd, regionStyles, innerEdges, innerEdge,
            selectionType, oldXfIndexes, oldXfIndex, col, row, workbook,
            style, newStyle, existingStyle, newXfIndexes,
            columnDimension, rowDimension, cell;
        
        if (is_array(pStyles)) {
            if (this->isSupervisor) {

                let pRange = this->getSelectedCells();

                // Uppercase coordinate
                let pRange = strtoupper(pRange);

                // Is it a cell range or a single cell?
                if (strpos(pRange, ":") === false) {
                    let rangeA = pRange;
                    let rangeB = pRange;
                } else {
                    let tmp = explode(":", pRange);
                    let rangeA = tmp[0];
                    let rangeB = tmp[1];
                }

                // Calculate range outer borders
                let rangeStart = \ZExcel\Cell::coordinateFromString(rangeA);
                let rangeEnd   = \ZExcel\Cell::coordinateFromString(rangeB);

                // Translate column into index
                let rangeStart[0]    = \ZExcel\Cell::columnIndexFromString(rangeStart[0]) - 1;
                let rangeEnd[0]    = \ZExcel\Cell::columnIndexFromString(rangeEnd[0]) - 1;

                // Make sure we can loop upwards on rows and columns
                if (rangeStart[0] > rangeEnd[0] && rangeStart[1] > rangeEnd[1]) {
                    let tmp = rangeStart;
                    let rangeStart = rangeEnd;
                    let rangeEnd = tmp;
                }

                // ADVANCED MODE:
                if (pAdvanced && isset(pStyles["borders"])) {
                    // "allborders" is a shorthand property for "outline" and "inside" and
                    //        it applies to components that have not been set explicitly
                    if (isset(pStyles["borders"]["allborders"])) {
                        for component in ["outline", "inside"] {
                            if (!isset(pStyles["borders"][component])) {
                                let pStyles["borders"][component] = pStyles["borders"]["allborders"];
                            }
                        }
                        unset(pStyles["borders"]["allborders"]); // not needed any more
                    }
                    
                    // "outline" is a shorthand property for "top", "right", "bottom", "left"
                    //        it applies to components that have not been set explicitly
                    if (isset(pStyles["borders"]["outline"])) {
                        for component in ["top", "right", "bottom", "left"] {
                            if (!isset(pStyles["borders"][component])) {
                                let pStyles["borders"][component] = pStyles["borders"]["outline"];
                            }
                        }
                        unset(pStyles["borders"]["outline"]); // not needed any more
                    }
                    
                    // "inside" is a shorthand property for "vertical" and "horizontal"
                    //        it applies to components that have not been set explicitly
                    if (isset(pStyles["borders"]["inside"])) {
                        for component in ["vertical", "horizontal"] {
                            if (!isset(pStyles["borders"][component])) {
                                let pStyles["borders"][component] = pStyles["borders"]["inside"];
                            }
                        }
                        unset(pStyles["borders"]["inside"]); // not needed any more
                    }
                    
                    // width and height characteristics of selection, 1, 2, or 3 (for 3 or more)
                    let xMax = min(rangeEnd[0] - rangeStart[0] + 1, 3);
                    let yMax = min(rangeEnd[1] - rangeStart[1] + 1, 3);

                    // loop through up to 3 x 3 = 9 regions
                    for x in range(1, xMax) {
                        // start column index for region
                        let colStart = (x == 3) ?
                            \ZExcel\Cell::stringFromColumnIndex(rangeEnd[0])
                                : \ZExcel\Cell::stringFromColumnIndex(rangeStart[0] + x - 1);
                        // end column index for region
                        let colEnd = (x == 1) ?
                            \ZExcel\Cell::stringFromColumnIndex(rangeStart[0])
                                : \ZExcel\Cell::stringFromColumnIndex(rangeEnd[0] - xMax + x);

                        for y in range(1, yMax) {
                            // which edges are touching the region
                            let edges = [];
                            
                            if (x == 1) {
                                // are we at left edge
                                let edges[] = "left";
                            }
                            
                            if (x == xMax) {
                                // are we at right edge
                                let edges[] = "right";
                            }
                            
                            if (y == 1) {
                                // are we at top edge?
                                let edges[] = "top";
                            }
                            
                            if (y == yMax) {
                                // are we at bottom edge?
                                let edges[] = "bottom";
                            }

                            // start row index for region
                            let rowStart = (y == 3) ?
                                rangeEnd[1] : rangeStart[1] + y - 1;

                            // end row index for region
                            let rowEnd = (y == 1) ?
                                rangeStart[1] : rangeEnd[1] - yMax + y;

                            // build range for region
                            let rangee = colStart . rowStart . ":" . colEnd . rowEnd;

                            // retrieve relevant style array for region
                            let regionStyles = pStyles;
                            unset(regionStyles["borders"]["inside"]);

                            // what are the inner edges of the region when looking at the selection
                            let innerEdges = array_diff(["top", "right", "bottom", "left"], edges );

                            // inner edges that are not touching the region should take the "inside" border properties if they have been set
                            for innerEdge in innerEdges {
                                switch (innerEdge) {
                                    case "top":
                                    case "bottom":
                                        // should pick up "horizontal" border property if set
                                        if (isset(pStyles["borders"]["horizontal"])) {
                                            let regionStyles["borders"][innerEdge] = pStyles["borders"]["horizontal"];
                                        } else {
                                            unset(regionStyles["borders"][innerEdge]);
                                        }
                                        break;
                                    case "left":
                                    case "right":
                                        // should pick up "vertical" border property if set
                                        if (isset(pStyles["borders"]["vertical"])) {
                                            let regionStyles["borders"][innerEdge] = pStyles["borders"]["vertical"];
                                        } else {
                                            unset(regionStyles["borders"][innerEdge]);
                                        }
                                        break;
                                }
                            }

                            // apply region style to region by calling applyFromArray() in simple mode
                            this->getActiveSheet()->getStyle(rangee)->applyFromArray(regionStyles, false);
                        }
                    }
                    
                    return this;
                }

                // SIMPLE MODE:
                // Selection type, inspect
                if (preg_match("/^[A-Z]+1:[A-Z]+1048576$/", pRange)) {
                    let selectionType = "COLUMN";
                } elseif (preg_match("/^A[0-9]+:XFD[0-9]+$/", pRange)) {
                    let selectionType = "ROW";
                } else {
                    let selectionType = "CELL";
                }

                // First loop through columns, rows, or cells to find out which styles are affected by this operation
                switch (selectionType) {
                    case "COLUMN":
                        let oldXfIndexes = [];
                        
                        for col in range(rangeStart[0], rangeEnd[0]) {
                            let oldXfIndexes[this->getActiveSheet()->getColumnDimensionByColumn(col)->getXfIndex()] = true;
                        }
                        break;
                    case "ROW":
                        let oldXfIndexes = [];
                        
                        for row in range(rangeStart[1], rangeEnd[1]) {
                            if (this->getActiveSheet()->getRowDimension(row)->getXfIndex() == null) {
                                let oldXfIndexes[0] = true; // row without explicit style should be formatted based on default style
                            } else {
                                let oldXfIndexes[this->getActiveSheet()->getRowDimension(row)->getXfIndex()] = true;
                            }
                        }
                        break;
                    case "CELL":
                        let oldXfIndexes = [];
                        
                        for col in range(rangeStart[0], rangeEnd[0]) {
                            for row in range(rangeStart[1], rangeEnd[1]) {
                                let oldXfIndexes[this->getActiveSheet()->getCellByColumnAndRow(col, row)->getXfIndex()] = true;
                            }
                        }
                        break;
                }

                // clone each of the affected styles, apply the style array, and add the new styles to the workbook
                let workbook = this->getActiveSheet()->getParent();
                
                for oldXfIndex, _ in oldXfIndexes {
                    let style = workbook->getCellXfByIndex(oldXfIndex);
                    let newStyle = clone style;
                    
                    newStyle->applyFromArray(pStyles);

                    let existingStyle = workbook->getCellXfByHashCode(newStyle->getHashCode());

                    if (existingStyle) {
                        // there is already such cell Xf in our collection
                        let newXfIndexes[oldXfIndex] = existingStyle->getIndex();
                    } else {
                        // we don"t have such a cell Xf, need to add
                        workbook->addCellXf(newStyle);
                        let newXfIndexes[oldXfIndex] = newStyle->getIndex();
                    }
                }

                // Loop through columns, rows, or cells again and update the XF index
                switch (selectionType) {
                    case "COLUMN":
                        for col in range(rangeStart[0], rangeEnd[0]) {
                            let columnDimension = this->getActiveSheet()->getColumnDimensionByColumn(col);
                            let oldXfIndex = columnDimension->getXfIndex();
                            
                            columnDimension->setXfIndex(newXfIndexes[oldXfIndex]);
                        }
                        break;

                    case "ROW":
                        for row in range(rangeStart[1], rangeEnd[1]) {
                            let rowDimension = this->getActiveSheet()->getRowDimension(row);
                            let oldXfIndex = rowDimension->getXfIndex() === null ? 0 : rowDimension->getXfIndex(); // row without explicit style should be formatted based on default style
                            
                            rowDimension->setXfIndex(newXfIndexes[oldXfIndex]);
                        }
                        break;

                    case "CELL":
                        for col in range(rangeStart[0], rangeEnd[0]) {
                            for row in range(rangeStart[1], rangeEnd[1]) {
                                let cell = this->getActiveSheet()->getCellByColumnAndRow(col, row);
                                let oldXfIndex = cell->getXfIndex();
                                
                                cell->setXfIndex(newXfIndexes[oldXfIndex]);
                            }
                        }
                        break;
                }

            } else {
                // not a supervisor, just apply the style array directly on style object
                if (array_key_exists("fill", pStyles)) {
                    this->getFill()->applyFromArray(pStyles["fill"]);
                }
                if (array_key_exists("font", pStyles)) {
                    this->getFont()->applyFromArray(pStyles["font"]);
                }
                if (array_key_exists("borders", pStyles)) {
                    this->getBorders()->applyFromArray(pStyles["borders"]);
                }
                if (array_key_exists("alignment", pStyles)) {
                    this->getAlignment()->applyFromArray(pStyles["alignment"]);
                }
                if (array_key_exists("numberformat", pStyles)) {
                    this->getNumberFormat()->applyFromArray(pStyles["numberformat"]);
                }
                if (array_key_exists("protection", pStyles)) {
                    this->getProtection()->applyFromArray(pStyles["protection"]);
                }
                if (array_key_exists("quotePrefix", pStyles)) {
                    let this->quotePrefix = pStyles["quotePrefix"];
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    /**
     * Get Fill
     *
     * @return \ZExcel\Style_Fill
     */
    public function getFill()
    {
        return this->fill;
    }

    /**
     * Get Font
     *
     * @return \ZExcel\Style_Font
     */
    public function getFont()
    {
        return this->font;
    }

    /**
     * Set font
     *
     * @param \ZExcel\Style_Font font
     * @return \ZExcel\Style
     */
    public function setFont(<\ZExcel\Style\Font> font)
    {
        let this->font = font;
        return this;
    }

    /**
     * Get Borders
     *
     * @return \ZExcel\Style_Borders
     */
    public function getBorders()
    {
        return this->borders;
    }

    /**
     * Get Alignment
     *
     * @return \ZExcel\Style_Alignment
     */
    public function getAlignment()
    {
        return this->alignment;
    }

    /**
     * Get Number Format
     *
     * @return \ZExcel\Style_NumberFormat
     */
    public function getNumberFormat()
    {
        return this->numberFormat;
    }

    /**
     * Get Conditional Styles. Only used on supervisor.
     *
     * @return \ZExcel\Style_Conditional[]
     */
    public function getConditionalStyles()
    {
        return this->getActiveSheet()->getConditionalStyles(this->getActiveCell());
    }

    /**
     * Set Conditional Styles. Only used on supervisor.
     *
     * @param \ZExcel\Style_Conditional[] pValue Array of condtional styles
     * @return \ZExcel\Style
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
     * @return \ZExcel\Style_Protection
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
            get_class(this)
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
