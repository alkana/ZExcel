namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class NumberFormat extends Supervisor implements ZIComparable
{
    /* Pre-defined formats */
    const FORMAT_GENERAL                 = "General";

    const FORMAT_TEXT                    = "@";

    const FORMAT_NUMBER                  = "0";
    const FORMAT_NUMBER_00               = "0.00";
    const FORMAT_NUMBER_COMMA_SEPARATED1 = "#,##0.00";
    const FORMAT_NUMBER_COMMA_SEPARATED2 = "#,##0.00_-";

    const FORMAT_PERCENTAGE              = "0%";
    const FORMAT_PERCENTAGE_00           = "0.00%";

    const FORMAT_DATE_YYYYMMDD2          = "yyyy-mm-dd";
    const FORMAT_DATE_YYYYMMDD           = "yy-mm-dd";
    const FORMAT_DATE_DDMMYYYY           = "dd/mm/yy";
    const FORMAT_DATE_DMYSLASH           = "d/m/y";
    const FORMAT_DATE_DMYMINUS           = "d-m-y";
    const FORMAT_DATE_DMMINUS            = "d-m";
    const FORMAT_DATE_MYMINUS            = "m-y";
    const FORMAT_DATE_XLSX14             = "mm-dd-yy";
    const FORMAT_DATE_XLSX15             = "d-mmm-yy";
    const FORMAT_DATE_XLSX16             = "d-mmm";
    const FORMAT_DATE_XLSX17             = "mmm-yy";
    const FORMAT_DATE_XLSX22             = "m/d/yy h:mm";
    const FORMAT_DATE_DATETIME           = "d/m/y h:mm";
    const FORMAT_DATE_TIME1              = "h:mm AM/PM";
    const FORMAT_DATE_TIME2              = "h:mm:ss AM/PM";
    const FORMAT_DATE_TIME3              = "h:mm";
    const FORMAT_DATE_TIME4              = "h:mm:ss";
    const FORMAT_DATE_TIME5              = "mm:ss";
    const FORMAT_DATE_TIME6              = "h:mm:ss";
    const FORMAT_DATE_TIME7              = "i:s.S";
    const FORMAT_DATE_TIME8              = "h:mm:ss;@";
    const FORMAT_DATE_YYYYMMDDSLASH      = "yy/mm/dd;@";

    const FORMAT_CURRENCY_USD_SIMPLE     = '""#,##0.00_-';
    const FORMAT_CURRENCY_USD            = "#,##0_-";
    const FORMAT_CURRENCY_EUR_SIMPLE     = "[EUR ]#,##0.00_-";

    /**
     * Excel built-in number formats
     *
     * @var array
     */
    protected static builtInFormats;

    /**
     * Excel built-in number formats (flipped, for faster lookups)
     *
     * @var array
     */
    protected static flippedBuiltInFormats;
    
    /**
     * Search/replace values to convert Excel date/time format masks to PHP format masks
     *
     * @var array
     */
    private static dateFormatReplacements;
    
    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (24 hr clock)
     *
     * @var array
     */
    private static dateFormatReplacements24;
    
    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (12 hr clock)
     *
     * @var array
     */
    private static dateFormatReplacements12;

    /**
     * Format Code
     *
     * @var string
     */
    protected formatCode = \ZExcel\Style\NumberFormat::FORMAT_GENERAL;

    /**
     * Built-in format Code
     *
     * @var string
     */
    protected builtInFormatCode    = 0;

    /**
     * Create a new \ZExcel\Style\NumberFormat
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

        if (isConditional) {
            let this->formatCode = null;
            let this->builtInFormatCode = false;
        }
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return \ZExcel\Style\NumberFormat
     */
    public function getSharedComponent()
    {
        return this->parent->getSharedComponent()->getNumberFormat();
    }

    /**
     * Build style array from subcomponents
     *
     * @param array array
     * @return array
     */
    public function getStyleArray(arry)
    {
        return ["numberformat": arry];
    }

    /**
     * Apply styles from array
     *
     * <code>
     * objPHPExcel->getActiveSheet()->getStyle("B2")->getNumberFormat()->applyFromArray(
     *        array(
     *            "code" => \ZExcel\Style\NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE
     *        )
     * );
     * </code>
     *
     * @param    array    pStyles    Array containing style information
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Style\NumberFormat
     */
    public function applyFromArray(array pStyles = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Format Code
     *
     * @return string
     */
    public function getFormatCode()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Format Code
     *
     * @param string pValue
     * @return \ZExcel\Style\NumberFormat
     */
    public function setFormatCode(pValue = \ZExcel\Style\NumberFormat::FORMAT_GENERAL)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get Built-In Format Code
     *
     * @return int
     */
    public function getBuiltInFormatCode()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set Built-In Format Code
     *
     * @param int pValue
     * @return \ZExcel\Style\NumberFormat
     */
    public function setBuiltInFormatCode(pValue = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Fill built-in format codes
     */
    private static function fillBuiltInFormatCodes()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get built-in format code
     *
     * @param    int        pIndex
     * @return    string
     */
    public static function builtInFormatCode(pIndex)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get built-in format code index
     *
     * @param    string        formatCode
     * @return    int|boolean
     */
    public static function builtInFormatCodeIndex(formatCode)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        throw new \Exception("Not implemented yet!");
    }
    
    protected static function getDateFormatReplacements()
    {
        throw new \Exception("Not implemented yet!");
    }
        
    private static function getDateFormatReplacements24()
    {
        throw new \Exception("Not implemented yet!");
    }
    
    private static function getDateFormatReplacements12()
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function formatAsDate(value, format)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function formatAsPercentage(value, format)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function formatAsFraction(value, format)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function complexNumberFormatMask(number, mask, level = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Convert a value in a pre-defined format to a PHP string
     *
     * @param mixed    value        Value to format
     * @param string    format        Format code
     * @param array        callBack    Callback function for additional formatting of string
     * @return string    Formatted string
     */
    public static function toFormattedString(var value = "0", var format = \ZExcel\Style\NumberFormat::FORMAT_GENERAL, var callback = null)
    {
        throw new \Exception("Not implemented yet!");
    }
}
