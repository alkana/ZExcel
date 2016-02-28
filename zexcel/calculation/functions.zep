namespace ZExcel\Calculation;

class Functions
{
    /** MAX_VALUE */
    const MAX_VALUE= "1.2e308";

/** 2 / PI */
    const M_2DIVPI = 0.63661977236758134307553505349006;

/** MAX_ITERATIONS */
    const MAX_ITERATIONS = 256;

/** PRECISION */
    const PRECISION = 0.000000000000000888;

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


    public static function ifCondition(var condition)
    {
        var operator, operand, returnValue = null;
        array matches = [];
        
        let condition = \ZExcel\Calculation\Functions::flattenSingleValue(condition);
        
        if (strlen(condition) == 0) {
            let condition = "=\"\"";
        }
        
        if (!in_array(substr(condition, 0, 1), [">", "<", "="])) {
            if (!is_numeric(condition)) {
                let condition = \ZExcel\Calculation::_wrapResult(strtoupper(condition));
            }
            let returnValue = "=" . condition;
        } else {
            preg_match("/([<>=]+)(.*)/", condition, matches);
            
            let operator = matches[1];
            let operand = matches[2];
            
            if (!is_numeric(operand)) {
                let operand = str_replace("\"", "\"\"", operand);
                let operand = \ZExcel\Calculation::_wrapResult(strtoupper(operand));
            }

            let returnValue = operator . operand;
        }
        
        return returnValue;
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
    
    /**
     * IS_EVEN
     *
     * @param    mixed    $value    Value to check
     * @return   boolean
     */
    public static function is_even(var value = null) -> boolean
    {
        let value = self::flattenSingleValue(value);

        if (value === null) {
            return self::NaME();
        } elseif ((is_bool(value)) || ((is_string(value)) && (!is_numeric(value)))) {
            return self::VaLUE();
        }

        return (value % 2 == 0);
    }
    
    /**
     * IS_ODD
     *
     * @param    mixed    $value    Value to check
     * @return   boolean
     */
    public static function is_odd(var value = null) -> boolean
    {
        let value = self::flattenSingleValue(value);

        if (value === null) {
            return self::NaME();
        } elseif ((is_bool(value)) || ((is_string(value)) && (!is_numeric(value)))) {
            return self::VaLUE();
        }

        return (abs(value) % 2 == 1);
    }
    
    /**
     * IS_NUMBER
     *
     * @param    mixed    $value        Value to check
     * @return   boolean
     */
    public static function is_number(var value = null) -> boolean
    {
        let value = self::flattenSingleValue(value);

        if (is_string(value)) {
            return false;
        }
        
        return is_numeric(value);
    }
    
    /**
     * IS_LOGICAL
     *
     * @param    mixed    $value        Value to check
     * @return   boolean
     */
    public static function is_logical(var value = null) -> boolean
    {
        let value = self::flattenSingleValue(value);

        return is_bool(value);
    }
    
    /**
     * IS_TEXT
     *
     * @param    mixed    $value        Value to check
     * @return   boolean
     */
    public static function is_text(var value = null) -> boolean
    {
        let value = self::flattenSingleValue(value);

        return (is_string(value) && !self::iS_ERROR(value));
    }
    
    /**
     * IS_NONTEXT
     *
     * @param    mixed    $value        Value to check
     * @return   boolean
     */
    public static function is_nontext(var value = null) -> boolean
    {
        return !self::is_text(value);
    }
    
    /**
     * VERSION
     *
     * @return    string    Version information
     */
    public static function version()
    {
        return "\ZExcel ##VERSION##, ##DATE##";
    }
    
    /**
     * N
     *
     * Returns a value converted to a number
     *
     * @param    value        The value you want converted
     * @return    number        N converts values listed in the following table
     *        If value is or refers to N returns
     *        A number            That number
     *        A date              The serial number of that date
     *        TRUE                1
     *        FALSE               0
     *        An error value      The error value
     *        Anything else       0
     */
    public static function n(var value = null)
    {
        while (is_array(value)) {
            let value = array_shift(value);
        }

        switch (gettype(value)) {
            case "double":
            case "float":
            case "integer":
                return value;
            case "boolean":
                return (int) value;
            case "string":
                //    Errors
                if ((strlen(value) > 0) && (substr(value, 0, 1) == "#")) {
                    return value;
                }
                break;
        }
        return 0;
    }
    
    /**
     * TYPE
     *
     * Returns a number that identifies the type of a value
     *
     * @param    value        The value you want tested
     * @return    number        N converts values listed in the following table
     *        If value is or refers to N returns
     *        A number            1
     *        Text                2
     *        Logical Value       4
     *        An error value      16
     *        Array or Matrix     64
     */
    public static function type(var value = null)
    {
        var a;
        
        let value = self::flattenArrayIndexed(value);
        
        if (is_array(value) && (count(value) > 1)) {
            end(value);
            let a = key(value);
            //    Range of cells is an error
            if (self::isCellValue(a)) {
                return 16;
            //    Test for Matrix
            } elseif (self::isMatrixValue(a)) {
                return 64;
            }
        } elseif (empty(value)) {
            //    Empty Cell
            return 1;
        }
        
        let value = self::flattenSingleValue(value);

        if ((value === null) || (is_float(value)) || (is_int(value))) {
                return 1;
        } elseif (is_bool(value)) {
                return 4;
        } elseif (is_array(value)) {
                return 64;
        } elseif (is_string(value)) {
            //    Errors
            if ((strlen(value) > 0) && (substr(value, 0, 1) == "#")) {
                return 16;
            }
            return 2;
        }
        
        return 0;
    }
    
    /**
     * Convert a multi-dimensional array to a simple 1-dimensional array
     *
     * @param    array    $array    Array to be flattened
     * @return   array    Flattened array
     */
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
    
    /**
     * Convert a multi-dimensional array to a simple 1-dimensional array, but retain an element of indexing
     *
     * @param    array    $array    Array to be flattened
     * @return   array    Flattened array
     */
    public static function flattenArrayIndexed(var arr) -> array
    {
        var k1, k2, k3, val1, val2, val3, arrayValues;
    
        if (!is_array(arr)) {
            return (array) arr;
        }

        let arrayValues = [];
        
        for k1, val1 in arr {
            if (is_array(val1)) {
                for k2, val2 in val1 {
                    if (is_array(val2)) {
                        for k3, val3 in val2 {
                            let arrayValues[k1.".".k2.".".k3] = val3;
                        }
                    } else {
                        let arrayValues[k1.".".k2] = val2;
                    }
                }
            } else {
                let arrayValues[k1] = val1;
            }
        }

        return arrayValues;
    }
    
    /**
     * Convert an array to a single scalar value by extracting the first element
     *
     * @param    mixed $value Array or scalar value
     * @return   mixed
     */
    public static function flattenSingleValue(var value = "")
    {
        while (is_array(value)) {
            let value = array_pop(value);
        }

        return value;
    }
}
