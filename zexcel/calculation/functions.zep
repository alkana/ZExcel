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
    protected static errorCodes;
    
    protected static function getErrorCodes()
    {
        if (self::errorCodes == null) {
            let self::errorCodes = [
                "null": "#NULL!",
                "divisionbyzero": "#DIV/0!",
                "value": "#VALUE!",
                "reference": "#REF!",
                "name": "#NAME?",
                "num": "#NUM!",
                "na": "#N/A",
                "gettingdata": "#GETTING_DATA"
            ];
        }
    
        return self::errorCodes;
    }
    
    public static function VaLUE()
    {
        var errorCodes;
        
        let errorCodes = self::getErrorCodes();
        
        return errorCodes["value"];
    }
    
    public static function NaN()
    {
        var errorCodes;
        let errorCodes = self::getErrorCodes();
        
        return errorCodes["num"];
    }
    
    public static function NaME()
    {
        var errorCodes;
        
        let errorCodes = self::getErrorCodes();
        
        return errorCodes["name"];
    }

    public static function ReF()
    {
        var errorCodes;
        
        let errorCodes = self::getErrorCodes();
        
        return errorCodes["reference"];
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
