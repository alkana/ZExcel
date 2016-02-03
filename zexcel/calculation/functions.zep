namespace ZExcel\Calculation;

class Functions
{
    /** constants */
    const COMPATIBILITY_EXCEL      = "Excel";
    const COMPATIBILITY_GNUMERIC   = "Gnumeric";
    const COMPATIBILITY_OPENOFFICE = "OpenOfficeCalc";

    const RETURNDATE_PHP_NUMERIC = "P";
    const RETURNDATE_PHP_OBJECT  = "O";
    const RETURNDATE_EXCEL       = "E";


    /**
     * Compatibility mode to use for error checking and responses
     *
     * @access    private
     * @var string
     */
    protected static compatibilityMode = self::COMPATIBILITY_EXCEL;

    /**
     * Data Type to use when returning date values
     *
     * @access    private
     * @var string
     */
    protected static returnDateType = self::RETURNDATE_EXCEL;

    /**
     * List of error codes
     *
     * @access    private
     * @var array
     */
    protected static errorCodes = [
        "null": "#NULL!",
        "divisionbyzero": "#DIV/0!",
        "value": "#VALUE!",
        "reference": "#REF!",
        "name": "#NAME?",
        "num": "#NUM!",
        "na": "#N/A",
        "gettingdata": "#GETTING_DATA"
    ];
    
    public static function setCompatibilityMode(string compatibilityMode) -> boolean
    {
        if ((compatibilityMode == self::COMPATIBILITY_EXCEL)
                || (compatibilityMode == self::COMPATIBILITY_GNUMERIC)
                || (compatibilityMode == self::COMPATIBILITY_OPENOFFICE)) {
            let self::compatibilityMode = compatibilityMode;
            
            return true;
        }
        return false;
    }
    
    public static function getCompatibilityMode() -> string
    {
        return self::compatibilityMode;
    }
    
    public static function setReturnDateType(string returnDateType) -> boolean
    {
        if ((returnDateType == self::RETURNDATE_PHP_NUMERIC)
                || (returnDateType == self::RETURNDATE_PHP_OBJECT)
                || (returnDateType == self::RETURNDATE_EXCEL)) {
            let self::returnDateType = returnDateType;

            return true;
        }
        return false;
    }
    
    public static function getReturnDateType() -> string
    {
        return self::returnDateType;
    }

    /**
     * DUMMY
     *
     * @access    public
     * @category Error Returns
     * @return    string    #Not Yet Implemented
     */
    public static function dummy()
    {
        return "#Not Yet Implemented";
    }


    /**
     * DIV0
     *
     * @access    public
     * @category Error Returns
     * @return    string    #Not Yet Implemented
     */
    public static function div0() -> string
    {
        return self::errorCodes["divisionbyzero"];
    }
    
    public static function na()
    {
        return self::errorCodes["na"];
    }
    
    public static function nan() -> string
    {
        return self::errorCodes["num"];
    }
    
    public static function name() -> string
    {
        return self::errorCodes["name"];
    }

    public static function ref() -> string
    {
        return self::errorCodes["reference"];
    }
    
    public static function nulll() -> string
    {
        return self::errorCodes["null"];
    }

    public static function value() -> string
    {
        return self::errorCodes["value"];
    }

    public static function isMatrixValue(int idx) -> boolean
    {
        return ((substr_count(idx, ".") <= 1) || (preg_match("/\.[A-Z]/", idx) > 0));
    }


    public static function isValue(int idx) -> boolean
    {
        return (substr_count(idx, ".") == 0);
    }


    public static function isCellValue(int idx) -> boolean
    {
        return (substr_count(idx, ".") > 1);
    }


    public static function ifCondition(condition)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * ERROR_TYPE
     *
     * @param    mixed    value    Value to check
     * @return    boolean
     */
    public static function error_type(value = "")
    {
        var i, errorCode;
        
        let value = self::flattenSingleValue(value);

        let i = 1;
        for errorCode in self::errorCodes {
            if (value === errorCode) {
                return i;
            }
            let i = i + 1;
        }
        return self::na();
    }

    /**
     * IS_BLANK
     *
     * @param    mixed    value    Value to check
     * @return    boolean
     */
    public static function is_blank(value = null) -> boolean
    {
        if (!is_null(value)) {
            let value = self::flattenSingleValue(value);
        }

        return is_null(value);
    }

    /**
     * IS_ERR
     *
     * @param    mixed    value    Value to check
     * @return    boolean
     */
    public static function is_err(value = "") -> boolean
    {
        let value = self::flattenSingleValue(value);

        return self::is_error(value) && (!self::is_na(value));
    }

    /**
     * IS_ERROR
     *
     * @param    mixed    value    Value to check
     * @return    boolean
     */
    public static function is_error(value = "") -> boolean
    {
        let value = self::flattenSingleValue(value);

        if (!is_string(value)) {
            return false;
        }
        return in_array(value, array_values(self::errorCodes));
    }

    /**
     * IS_NA
     *
     * @param    mixed    $value    Value to check
     * @return    boolean
     */
    public static function is_na(value = "") -> boolean
    {
        let value = self::flattenSingleValue(value);

        return (value === self::na());
    }
    
    public static function flattenArray(var arry)
    {
        var value, val, v;
        
        if (!is_array(arry)) {
            return [arry];
        }

        var arrayValues = [];
        for value in arry {
            if (is_array(value)) {
                for val in value {
                    if (is_array(val)) {
                        for v in val {
                            let arrayValues[] = v;
                        }
                    } else {
                        let arrayValues[] = val;
                    }
                }
            } else {
                let arrayValues[] = value;
            }
        }

        return arrayValues;
    }
    
    public static function flattenSingleValue(var value = "")
    {
        while (is_array(value)) {
            let value = array_pop(value);
        }

        return value;
    }
}
