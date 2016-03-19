namespace ZExcel\Calculation;

class Functionn
{
    /* Function categories */
    const CATEGORY_CUBE                 = "Cube";
    const CATEGORY_DATABASE             = "Database";
    const CATEGORY_DATE_AND_TIME        = "Date and Time";
    const CATEGORY_ENGINEERING          = "Engineering";
    const CATEGORY_FINANCIAL            = "Financial";
    const CATEGORY_INFORMATION          = "Information";
    const CATEGORY_LOGICAL              = "Logical";
    const CATEGORY_LOOKUP_AND_REFERENCE = "Lookup and Reference";
    const CATEGORY_MATH_AND_TRIG        = "Math and Trig";
    const CATEGORY_STATISTICAL          = "Statistical";
    const CATEGORY_TEXT_AND_DATA        = "Text and Data";

    /**
     * Category (represented by CATEGORY_*)
     *
     * @var string
     */
    private category;

    /**
     * Excel name
     *
     * @var string
     */
    private excelName;

    /**
     * PHPExcel name
     *
     * @var string
     */
    private phpExcelName;

    /**
     * Create a new PHPExcel_Calculation_Function
     *
     * @param     string        pCategory         Category (represented by CATEGORY_*)
     * @param     string        pExcelName        Excel function name
     * @param     string        pPHPExcelName    PHPExcel function mapping
     * @throws     PHPExcel_Calculation_Exception
     */
    public function __construct(string pCategory = null, string pExcelName = null, string pPHPExcelName = null)
    {
        if ((pCategory !== null) && (pExcelName !== null) && (pPHPExcelName !== null)) {
            // Initialise values
            let this->category     = pCategory;
            let this->excelName    = pExcelName;
            let this->phpExcelName = pPHPExcelName;
        } else {
            throw new \ZExcel\Calculation\Exception("Invalid parameters passed.");
        }
    }

    /**
     * Get Category (represented by CATEGORY_*)
     *
     * @return string
     */
    public function getCategory()
    {
        return this->category;
    }

    /**
     * Set Category (represented by CATEGORY_*)
     *
     * @param     string        value
     * @throws     PHPExcel_Calculation_Exception
     */
    public function setCategory(value = null)
    {
        if (!is_null(value)) {
            let this->category = value;
        } else {
            throw new \ZExcel\Calculation\Exception("Invalid parameter passed.");
        }
    }

    /**
     * Get Excel name
     *
     * @return string
     */
    public function getExcelName()
    {
        return this->excelName;
    }

    /**
     * Set Excel name
     *
     * @param string    value
     */
    public function setExcelName(value)
    {
        let this->excelName = value;
    }

    /**
     * Get PHPExcel name
     *
     * @return string
     */
    public function getPHPExcelName()
    {
        return this->phpExcelName;
    }

    /**
     * Set PHPExcel name
     *
     * @param string    value
     */
    public function setPHPExcelName(value)
    {
        let this->phpExcelName = value;
    }
}
