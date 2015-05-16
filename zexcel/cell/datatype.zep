namespace ZExcel\Cell;

class DataType
{
    /* Data types */
    const TYPE_STRING2  = "str";
    const TYPE_STRING   = "s";
    const TYPE_FORMULA  = "f";
    const TYPE_NUMERIC  = "n";
    const TYPE_BOOL     = "b";
    const TYPE_NULL     = "null";
    const TYPE_INLINE   = "inlineStr";
    const TYPE_ERROR    = "e";

    /**
     * List of error codes
     *
     * @var array
     */
    private static errorCodes;

    /**
     * Get list of error codes
     *
     * @return array
     */
    public static function getErrorCodes()
    {
        if (self::errorCodes == null) {
            let self::errorCodes = [
                "#NULL!": 0,
                "#DIV/0!": 1,
                "#VALUE!": 2,
                "#REF!": 3,
                "#NAME?": 4,
                "#NUM!": 5,
                "#N/A": 6
            ];
        }
    
        return self::errorCodes;
    }

    /**
     * DataType for value
     *
     * @deprecated  Replaced by \ZExcel\Cell\IValueBinder infrastructure, will be removed in version 1.8.0
     * @param       mixed  pValue
     * @return      string
     */
    public static function dataTypeForValue(pValue = null)
    {
        return \ZExcel\Cell\DefaultValueBinder::dataTypeForValue(pValue);
    }

    /**
     * Check a string that it satisfies Excel requirements
     *
     * @param  mixed  Value to sanitize to an Excel string
     * @return mixed  Sanitized value
     */
    public static function checkString(pValue = null)
    {
        if (pValue instanceof \ZExcel\RichText) {
            // TODO: Sanitize Rich-Text string (max. character count is 32,767)
            return pValue;
        }

        // string must never be longer than 32,767 characters, truncate if necessary
        let pValue = \ZExcel\Shared\Strin::Substring(pValue, 0, 32767);

        // we require that newline is represented as "\n" in core, not as "\r\n" or "\r"
        let pValue = str_replace(["\r\n", "\r"], "\n", pValue);

        return pValue;
    }

    /**
     * Check a value that it is a valid error code
     *
     * @param  mixed   Value to sanitize to an Excel error code
     * @return string  Sanitized value
     */
    public static function checkErrorCode(pValue = null)
    {
        let pValue = (string) pValue;

        if (!array_key_exists(pValue, self::errorCodes)) {
            let pValue = "#NULL!";
        }

        return pValue;
    }
}
