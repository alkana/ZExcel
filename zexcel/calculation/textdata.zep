namespace ZExcel\Calculation;

class TextData
{
    private static invalidChars = null;

    private static function unicodeToOrd(c)
    {
        var c0, c1, c2, c3, c4, c5;
        
        let c0 = substr(c, 0, 1);
        let c1 = substr(c, 1, 1);
        let c2 = substr(c, 2, 1);
        let c3 = substr(c, 3, 1);
        let c4 = substr(c, 4, 1);
        let c5 = substr(c, 5, 1);
        
        if (ord(c0) >= 0 && ord(c0) <= 127) {
            return ord(c0);
        } else {
            if (ord(c0) >= 192 && ord(c0) <= 223) {
                return (ord(c0) - 192) * 64 + (ord(c1) - 128);
            } else {
                if (ord(c0) >= 224 && ord(c0) <= 239) {
                    return (ord(c0) - 224) * 4096 + (ord(c1) - 128) * 64 + (ord(c2) - 128);
                } else {
                    if (ord(c0) >= 240 && ord(c0) <= 247) {
                        return (ord(c0) - 240) * 262144 + (ord(c1) - 128) * 4096 + (ord(c2) - 128) * 64 + (ord(c3) - 128);
                    } else {
                        if (ord(c0) >= 248 && ord(c0) <= 251) {
                            return (ord(c0) - 248) * 16777216 + (ord(c1) - 128) * 262144 + (ord(c2) - 128) * 4096 + (ord(c3) - 128) * 64 + (ord(c4) - 128);
                        } else {
                            if (ord(c0) >= 252 && ord(c0) <= 253) {
                                return (ord(c0) - 252) * 1073741824 + (ord(c1) - 128) * 16777216 + (ord(c2) - 128) * 262144 + (ord(c3) - 128) * 4096 + (ord(c4) - 128) * 64 + (ord(c5) - 128);
                            } else {
                                if (ord(c0) >= 254 && ord(c0) <= 255) {
                                    // error
                                    return \ZExcel\Calculation\Functions::VaLUE();
                                }
                            }
                        }
                    }
                }
            }
        }
        
        return 0;
    }

    /**
     * CHARACTER
     *
     * @param    string    character    Value
     * @return    int
     */
    public static function character(var character)
    {
        let character = \ZExcel\Calculation\Functions::flattenSingleValue(character);

        if ((!is_numeric(character)) || (character < 0)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (function_exists("mb_convert_encoding")) {
            return mb_convert_encoding("&#" . intval(character) . ";", "UTF-8", "HTML-ENTITIES");
        } else {
            return chr(intval(character));
        }
    }


    /**
     * TRIMNONPRINTABLE
     *
     * @param    mixed    stringValue    Value to check
     * @return    string
     */
    public static function trimNonPrintable(var stringValue = "")
    {
        string trimer = chr(0) . ".." . chr(31);
        
        let stringValue = \ZExcel\Calculation\Functions::flattenSingleValue(stringValue);

        if (is_bool(stringValue)) {
            return (stringValue) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        if (!is_array(self::invalidChars)) {
            let self::invalidChars = range(chr(0), chr(31));
        }

        if (is_string(stringValue) || is_numeric(stringValue)) {
            return str_replace(self::invalidChars, "", trim(stringValue, trimer));
        }
        
        return null;
    }


    /**
     * TRIMSPACES
     *
     * @param    mixed    stringValue    Value to check
     * @return    string
     */
    public static function trimspaces(var stringValue = "")
    {
        let stringValue = \ZExcel\Calculation\Functions::flattenSingleValue(stringValue);
        
        if (is_bool(stringValue)) {
            return (stringValue) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        if (is_string(stringValue) || is_numeric(stringValue)) {
            return trim(preg_replace("/ +/", " ", trim(stringValue, " ")), " ");
        }
        
        return null;
    }


    /**
     * ASCIICODE
     *
     * @param    string    characters        Value
     * @return    int
     */
    public static function asciiCode(var characters)
    {
        var character;
        
        if ((characters === null) || (characters === "")) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let characters = \ZExcel\Calculation\Functions::flattenSingleValue(characters);
        
        if (is_bool(characters)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let characters = (int) characters;
            } else {
                if (characters === true) {
                    let characters = \ZExcel\Calculation::getTRUE();
                } else {
                    let characters = \ZExcel\Calculation::getFALSE();
                }
            }
        }

        let character = characters;
        
        if ((function_exists("mb_strlen")) && (function_exists("mb_substr"))) {
            if (mb_strlen(characters, "UTF-8") > 1) {
                let character = mb_substr(characters, 0, 1, "UTF-8");
            }
            
            return self::unicodeToOrd(character);
        } else {
            if (strlen(characters) > 0) {
                let character = substr(characters, 0, 1);
            }
            
            return ord(character);
        }
    }


    /**
     * CONCATENATE
     *
     * @return    string
     */
    public static function concatenate() -> string
    {
        var aArgs, arg;
        string returnValue = "";

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        for arg in aArgs {
            if (is_bool(arg)) {
                if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                    let arg = (int) arg;
                } else {
                    if (arg) {
                        let arg = \ZExcel\Calculation::getTRUE();
                    } else {
                        let arg = \ZExcel\Calculation::getFALSE();
                    }
                }
            }
            let returnValue = returnValue . (string) arg;
        }

        return returnValue;
    }


    /**
     * DOLLAR
     *
     * This function converts a number to text using currency format, with the decimals rounded to the specified place.
     * The format used is #,##0.00_);(#,##0.00)..
     *
     * @param    float    value            The value to format
     * @param    int        decimals        The number of digits to display to the right of the decimal point.
     *                                    If decimals is negative, number is rounded to the left of the decimal point.
     *                                    If you omit decimals, it is assumed to be 2
     * @return    string
     */
    public static function dollar(var value = 0, var decimals = 2)
    {
        var mask, round;
        
        let value    = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let decimals = is_null(decimals) ? 0 : \ZExcel\Calculation\Functions::flattenSingleValue(decimals);

        // Validate parameters
        if (!is_numeric(value) || !is_numeric(decimals)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        let decimals = floor(decimals);

        let mask = "$#,##0";
        if (decimals > 0) {
            let mask = mask . "." . str_repeat(strval(0), decimals);
        } else {
            let round = pow(10, abs(decimals));
            if (value < 0) {
                let round = 0 - round;
            }
            let value = \ZExcel\Calculation\MathTrig::MRoUND(value, round);
        }

        return \ZExcel\Style\NumberFormat::toFormattedString(value, mask);

    }


    /**
     * SEARCHSENSITIVE
     *
     * @param    string    needle        The string to look for
     * @param    string    haystack    The string in which to look
     * @param    int        offset        Offset within haystack
     * @return    string
     */
    public static function searchSensitive(var needle, var haystack, int offset = 1)
    {
        var pos;
        
        let needle   = \ZExcel\Calculation\Functions::flattenSingleValue(needle);
        let haystack = \ZExcel\Calculation\Functions::flattenSingleValue(haystack);
        let offset   = \ZExcel\Calculation\Functions::flattenSingleValue(offset);

        if (!is_bool(needle)) {
            if (is_bool(haystack)) {
                let haystack = (haystack) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
            }

            if ((offset > 0) && (\ZExcel\Shared\Stringg::CountCharacters(haystack) > offset)) {
                if (\ZExcel\Shared\Stringg::CountCharacters(needle) == 0) {
                    return offset;
                }
                
                let offset = offset - 1;
                
                if (function_exists("mb_strpos")) {
                    let pos = mb_strpos(haystack, needle, offset, "UTF-8");
                } else {
                    let pos = strpos(haystack, needle, offset);
                }
                
                if (pos !== false) {
                    return pos + 1;
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SEARCHINSENSITIVE
     *
     * @param    string    needle        The string to look for
     * @param    string    haystack    The string in which to look
     * @param    int        offset        Offset within haystack
     * @return    string
     */
    public static function searchInsensitive(var needle, var haystack, int offset = 1)
    {
        var pos;
        
        let needle   = \ZExcel\Calculation\Functions::flattenSingleValue(needle);
        let haystack = \ZExcel\Calculation\Functions::flattenSingleValue(haystack);
        let offset   = \ZExcel\Calculation\Functions::flattenSingleValue(offset);

        if (!is_bool(needle)) {
            if (is_bool(haystack)) {
                let haystack = (haystack) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
            }

            if ((offset > 0) && (\ZExcel\Shared\Stringg::CountCharacters(haystack) > offset)) {
                if (\ZExcel\Shared\Stringg::CountCharacters(needle) == 0) {
                    return offset;
                }
                
                let offset = offset - 1;
                
                if (function_exists("mb_stripos")) {
                    let pos = mb_stripos(haystack, needle, offset, "UTF-8");
                } else {
                    let pos = stripos(haystack, needle, offset);
                }
                
                if (pos !== false) {
                    return pos + 1;
                }
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FIXEDFORMAT
     *
     * @param    mixed        value    Value to check
     * @param    integer        decimals
     * @param    boolean        no_commas
     * @return    boolean
     */
    public static function fixedFormat(var value, var decimals = 2, var no_commas = false)
    {
        var valueResult;
        
        let value     = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let decimals  = \ZExcel\Calculation\Functions::flattenSingleValue(decimals);
        let no_commas = \ZExcel\Calculation\Functions::flattenSingleValue(no_commas);

        // Validate parameters
        if (!is_numeric(value) || !is_numeric(decimals)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        let decimals = floor(decimals);

        let valueResult = round(value, decimals);
        
        if (decimals < 0) {
            let decimals = 0;
        }
        if (!no_commas) {
            let valueResult = number_format(valueResult, decimals);
        }

        return (string) valueResult;
    }


    /**
     * LEFT
     *
     * @param    string    value    Value
     * @param    int        chars    Number of characters
     * @return    string
     */
    public static function left(var value = "", int chars = 1)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let chars = \ZExcel\Calculation\Functions::flattenSingleValue(chars);

        if (chars < 0) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (is_bool(value)) {
            let value = (value) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        if (function_exists("mb_substr")) {
            return mb_substr(value, 0, chars, "UTF-8");
        } else {
            return substr(value, 0, chars);
        }
    }


    /**
     * MID
     *
     * @param    string    value    Value
     * @param    int        start    Start character
     * @param    int        chars    Number of characters
     * @return    string
     */
    public static function mid(var value = "", var start = 1, var chars = null)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let start = \ZExcel\Calculation\Functions::flattenSingleValue(start);
        let chars = \ZExcel\Calculation\Functions::flattenSingleValue(chars);

        if ((start < 1) || (chars < 0)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (is_bool(value)) {
            if value {
                let value = \ZExcel\Calculation::getTRUE();
            } else {
                let value = \ZExcel\Calculation::getFALSE();
            }
        }

        let start = start - 1;
        
        if (function_exists("mb_substr")) {
            return mb_substr(value, start, chars, "UTF-8");
        } else {
            return substr(value, start, chars);
        }
    }


    /**
     * RIGHT
     *
     * @param    string    value    Value
     * @param    int        chars    Number of characters
     * @return    string
     */
    public static function right(var value = "", int chars = 1)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let chars = \ZExcel\Calculation\Functions::flattenSingleValue(chars);

        if (chars < 0) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (is_bool(value)) {
            let value = (value) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        if ((function_exists("mb_substr")) && (function_exists("mb_strlen"))) {
            return mb_substr(value, mb_strlen(value, "UTF-8") - chars, chars, "UTF-8");
        } else {
            return substr(value, strlen(value) - chars);
        }
    }


    /**
     * STRINGLENGTH
     *
     * @param    string    value    Value
     * @return    string
     */
    public static function stringlength(var value = "")
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        if (is_bool(value)) {
            let value = (value) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        if (function_exists("mb_strlen")) {
            return mb_strlen(value, "UTF-8");
        } else {
            return strlen(value);
        }
    }


    /**
     * LOWERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        mixedCaseString
     * @return    string
     */
    public static function lowercase(var mixedCaseString) -> string
    {
        let mixedCaseString = \ZExcel\Calculation\Functions::flattenSingleValue(mixedCaseString);

        if (is_bool(mixedCaseString)) {
            let mixedCaseString = (mixedCaseString) ? \ZExcel\Calculation::getTRUE() : \ZExcel\Calculation::getFALSE();
        }

        return \ZExcel\Shared\Stringg::StrToLower(mixedCaseString);
    }


    /**
     * UPPERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        mixedCaseString
     * @return    string
     */
    public static function uppercase(var mixedCaseString) -> string
    {
        let mixedCaseString = \ZExcel\Calculation\Functions::flattenSingleValue(mixedCaseString);

        if (is_bool(mixedCaseString)) {
            if (mixedCaseString) {
                let mixedCaseString =  \ZExcel\Calculation::getTRUE();
            } else {
                let mixedCaseString =  \ZExcel\Calculation::getFALSE();
            }
        }

        return \ZExcel\Shared\Stringg::StrToUpper(mixedCaseString);
    }


    /**
     * PROPERCASE
     *
     * Converts a string value to upper case.
     *
     * @param    string        mixedCaseString
     * @return    string
     */
    public static function propercase(var mixedCaseString) -> string
    {
        let mixedCaseString = \ZExcel\Calculation\Functions::flattenSingleValue(mixedCaseString);

        if (is_bool(mixedCaseString)) {
            if (mixedCaseString) {
                let mixedCaseString =  \ZExcel\Calculation::getTRUE();
            } else {
                let mixedCaseString =  \ZExcel\Calculation::getFALSE();
            }
        }

        return \ZExcel\Shared\Stringg::strToTitle(mixedCaseString);
    }


    /**
     * REPLACE
     *
     * @param    string    oldText    String to modify
     * @param    int        start        Start character
     * @param    int        chars        Number of characters
     * @param    string    newText    String to replace in defined position
     * @return    string
     */
    public static function replace(var oldText = "", int start = 1, var chars = null, var newText)
    {
        var left, right;
        
        let oldText = \ZExcel\Calculation\Functions::flattenSingleValue(oldText);
        let start   = \ZExcel\Calculation\Functions::flattenSingleValue(start);
        let chars   = \ZExcel\Calculation\Functions::flattenSingleValue(chars);
        let newText = \ZExcel\Calculation\Functions::flattenSingleValue(newText);

        let left = self::LeFT(oldText, start - 1);
        let right = self::RiGHT(oldText, self::STRiNGLENGTH(oldText) - (start + chars) + 1);

        return left . newText . right;
    }


    /**
     * SUBSTITUTE
     *
     * @param    string    text        Value
     * @param    string    fromText    From Value
     * @param    string    toText        To Value
     * @param    integer    instance    Instance Number
     * @return    string
     */
    public static function substitute(var text = "", var fromText = "", var toText = "", int instance = 0)
    {
        var pos;
        
        let text     = \ZExcel\Calculation\Functions::flattenSingleValue(text);
        let fromText = \ZExcel\Calculation\Functions::flattenSingleValue(fromText);
        let toText   = \ZExcel\Calculation\Functions::flattenSingleValue(toText);
        let instance = floor(\ZExcel\Calculation\Functions::flattenSingleValue(instance));

        if (instance == 0) {
            // if (function_exists("mb_str_replace")) {
            //    return mb_str_replace(fromText, toText, text);
            // } else {
                return str_replace(fromText, toText, text);
            // }
        } else {
            let pos = -1;
            while (instance > 0) {
                if (function_exists("mb_strpos")) {
                    let pos = mb_strpos(text, fromText, pos+1, "UTF-8");
                } else {
                    let pos = strpos(text, fromText, pos+1);
                }
                
                if (pos === false) {
                    break;
                }
                
                let instance = instance - 1;
            }
            if (pos !== false) {
                let pos = pos + 1;
                
                if (function_exists("mb_strlen")) {
                    return self::RePLACE(text, pos, mb_strlen(fromText, "UTF-8"), toText);
                } else {
                    return self::RePLACE(text, pos, strlen(fromText), toText);
                }
            }
        }

        return text;
    }


    /**
     * RETURNSTRING
     *
     * @param    mixed    testValue    Value to check
     * @return    boolean
     */
    public static function returnString(var testValue = "")
    {
        let testValue = \ZExcel\Calculation\Functions::flattenSingleValue(testValue);

        if (is_string(testValue)) {
            return testValue;
        }
        return null;
    }


    /**
     * TEXTFORMAT
     *
     * @param    mixed    value    Value to check
     * @param    string    format    Format mask to use
     * @return    boolean
     */
    public static function textFormat(var value, var format)
    {
        let value  = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let format = \ZExcel\Calculation\Functions::flattenSingleValue(format);

        if ((is_string(value)) && (!is_numeric(value)) && \ZExcel\Shared\Date::isDateTimeFormatCode(format)) {
            let value = \ZExcel\Calculation\DateTime::DaTEVALUE(value);
        }

        return (string) \ZExcel\Style\NumberFormat::toFormattedString(value, format);
    }

    /**
     * VALUE
     *
     * @param    mixed    value    Value to check
     * @return    boolean
     */
    public static function value(var value = "")
    {
        var dateSetting, numberValue, dateValue, timeValue;
        
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        if (!is_numeric(value)) {
            let numberValue = str_replace(
                \ZExcel\Shared\Stringg::getThousandsSeparator(),
                "",
                trim(value, chr(0) . chr(9) . chr(10) . chr(11) . chr(32) . chr(104) . \ZExcel\Shared\Stringg::getCurrencyCode())
            );
            
            if (is_numeric(numberValue)) {
                return (float) numberValue;
            }

            let dateSetting = \ZExcel\Calculation\Functions::getReturnDateType();
            \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);

            if (strpos(value, ":") !== false) {
                let timeValue = \ZExcel\Calculation\DateTime::TiMEVALUE(value);
                if (timeValue !== \ZExcel\Calculation\Functions::VaLUE()) {
                    \ZExcel\Calculation\Functions::setReturnDateType(dateSetting);
                    return timeValue;
                }
            }
            
            let dateValue = \ZExcel\Calculation\DateTime::DaTEVALUE(value);
            
            if (dateValue !== \ZExcel\Calculation\Functions::VaLUE()) {
                \ZExcel\Calculation\Functions::setReturnDateType(dateSetting);
                return dateValue;
            }
            
            \ZExcel\Calculation\Functions::setReturnDateType(dateSetting);

            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        return (float) value;
    }
}
