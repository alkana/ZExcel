namespace ZExcel\Worksheet;

class SheetView
{
    /* Sheet View types */
    const SHEETVIEW_NORMAL             = "normal";
    const SHEETVIEW_PAGE_LAYOUT        = "pageLayout";
    const SHEETVIEW_PAGE_BREAK_PREVIEW = "pageBreakPreview";

    private _sheetViewTypes = [
        self::SHEETVIEW_NORMAL,
        self::SHEETVIEW_PAGE_LAYOUT,
        self::SHEETVIEW_PAGE_BREAK_PREVIEW
    ];

    /**
     * ZoomScale
     *
     * Valid values range from 10 to 400.
     *
     * @var int
     */
    private _zoomScale     = 100;

    /**
     * ZoomScaleNormal
     *
     * Valid values range from 10 to 400.
     *
     * @var int
     */
    private _zoomScaleNormal = 100;

    /**
     * View
     *
     * Valid values range from 10 to 400.
     *
     * @var string
     */
    private _sheetviewType = self::SHEETVIEW_NORMAL;

    /**
     * Create a new PHPExcel_Worksheet_SheetView
     */
    public function __construct()
    {
    }

    /**
     * Get ZoomScale
     *
     * @return int
     */
    public function getZoomScale() {
        return this->_zoomScale;
    }

    /**
     * Set ZoomScale
     *
     * Valid values range from 10 to 400.
     *
     * @param   int     pValue
     * @throws  \ZExcel\Exception
     * @return PHPExcel_Worksheet_SheetView
     */
    public function setZoomScale(pValue = 100) {
        // Microsoft Office Excel 2007 only allows setting a scale between 10 and 400 via the user interface,
        // but it is apparently still able to handle any scale >= 1
        if ((pValue >= 1) || is_null(pValue)) {
            let this->_zoomScale = pValue;
        } else {
            throw new \ZExcel\Exception("Scale must be greater than or equal to 1.");
        }
        return this;
    }

    /**
     * Get ZoomScaleNormal
     *
     * @return int
     */
    public function getZoomScaleNormal() {
        return this->_zoomScaleNormal;
    }

    /**
     * Set ZoomScale
     *
     * Valid values range from 10 to 400.
     *
     * @param   int     pValue
     * @throws  \ZExcel\Exception
     * @return PHPExcel_Worksheet_SheetView
     */
    public function setZoomScaleNormal(pValue = 100) {
        if ((pValue >= 1) || is_null(pValue)) {
            let this->_zoomScaleNormal = pValue;
        } else {
            throw new \ZExcel\Exception("Scale must be greater than or equal to 1.");
        }
        return this;
    }

    /**
     * Get View
     *
     * @return string
     */
    public function getView() {
        return this->_sheetviewType;
    }

    /**
     * Set View
     *
     * Valid values are
     *      "normal"            self::SHEETVIEW_NORMAL
     *      "pageLayout"        self::SHEETVIEW_PAGE_LAYOUT
     *      "pageBreakPreview"  self::SHEETVIEW_PAGE_BREAK_PREVIEW
     *
     * @param   string  pValue
     * @throws  \ZExcel\Exception
     * @return PHPExcel_Worksheet_SheetView
     */
    public function setView(pValue = null) {
        //  MS Excel 2007 allows setting the view to "normal", "pageLayout" or "pageBreakPreview"
        //      via the user interface
        if (pValue === null) {
            let pValue = self::SHEETVIEW_NORMAL;
        }
        
        if (in_array(pValue, this->_sheetViewTypes)) {
            let this->_sheetviewType = pValue;
        } else {
            throw new \ZExcel\Exception("Invalid sheetview layout type.");
        }

        return this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
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
