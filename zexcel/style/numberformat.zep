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

    const FORMAT_CURRENCY_USD_SIMPLE     = "\"\"#,##0.00_-";
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
    private static dateFormatReplacements = [
            // first remove escapes related to non-format characters
            "\\": "",
            //    12-hour suffix
            "am/pm": "A",
            //    4-digit year
            "e": "Y",
            "yyyy": "Y",
            //    2-digit year
            "yy": "y",
            //    first letter of month - no php equivalent
            "mmmmm": "M",
            //    full month name
            "mmmm": "F",
            //    short month name
            "mmm": "M",
            //    mm is minutes if time, but can also be month w/leading zero
            //    so we try to identify times be the inclusion of a : separator in the mask
            //    It isn"t perfect, but the best way I know how
            ":mm": ":i",
            "mm:": "i:",
            //    month leading zero
            "mm": "m",
            //    month no leading zero
            "m": "n",
            //    full day of week name
            "dddd": "l",
            //    short day of week name
            "ddd" : "D",
            //    days leading zero
            "dd": "d",
            //    days no leading zero
            "d": "j",
            //    seconds
            "ss": "s",
            //    fractional seconds - no php equivalent
            ".s": ""
    ];
    
    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (24 hr clock)
     *
     * @var array
     */
    private static dateFormatReplacements24 = [
        "hh": "H",
        "h":  "G"
    ];
    
    /**
     * Search/replace values to convert Excel date/time format masks hours to PHP format masks (12 hr clock)
     *
     * @var array
     */
    private static dateFormatReplacements12 = [
        "hh": "h",
        "h":  "g"
    ];

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
    protected builtInFormatCode = 0;

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
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (array_key_exists("code", pStyles)) {
                    this->setFormatCode(pStyles["code"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        
        return this;
    }

    /**
     * Get Format Code
     *
     * @return string
     */
    public function getFormatCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getFormatCode();
        }
        
        if (this->builtInFormatCode !== false) {
            return self::builtInFormatCode(this->builtInFormatCode);
        }
        
        return this->formatCode;
    }

    /**
     * Set Format Code
     *
     * @param string pValue
     * @return \ZExcel\Style\NumberFormat
     */
    public function setFormatCode(var pValue = \ZExcel\Style\NumberFormat::FORMAT_GENERAL)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = \ZExcel\Style\NumberFormat::FORMAT_GENERAL;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["code": pValue]);
            this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(styleArray);
        } else {
            let this->formatCode = pValue;
            let this->builtInFormatCode = self::builtInFormatCodeIndex(pValue);
        }
        
        return this;
    }

    /**
     * Get Built-In Format Code
     *
     * @return int
     */
    public function getBuiltInFormatCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getBuiltInFormatCode();
        }
        
        return this->builtInFormatCode;
    }

    /**
     * Set Built-In Format Code
     *
     * @param int pValue
     * @return \ZExcel\Style\NumberFormat
     */
    public function setBuiltInFormatCode(var pValue = 0)
    {
        var styleArray;
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["code": self::builtInFormatCode(pValue)]);
            this->getActiveSheet()->getStyle(this->getSelectedCells())->applyFromArray(styleArray);
        } else {
            let this->builtInFormatCode = pValue;
            let this->formatCode = self::builtInFormatCode(pValue);
        }
        
        return this;
    }

    /**
     * Fill built-in format codes
     */
    private static function fillBuiltInFormatCodes()
    {
        //  [MS-OI29500: Microsoft Office Implementation Information for ISO/IEC-29500 Standard Compliance]
        //  18.8.30. numFmt (Number Format)
        //
        //  The ECMA standard defines built-in format IDs
        //      14: "mm-dd-yy"
        //      22: "m/d/yy h:mm"
        //      37: "#,##0 ;(#,##0)"
        //      38: "#,##0 ;[Red](#,##0)"
        //      39: "#,##0.00;(#,##0.00)"
        //      40: "#,##0.00;[Red](#,##0.00)"
        //      47: "mmss.0"
        //      KOR fmt 55: "yyyy-mm-dd"
        //  Excel defines built-in format IDs
        //      14: "m/d/yyyy"
        //      22: "m/d/yyyy h:mm"
        //      37: "#,##0_);(#,##0)"
        //      38: "#,##0_);[Red](#,##0)"
        //      39: "#,##0.00_);(#,##0.00)"
        //      40: "#,##0.00_);[Red](#,##0.00)"
        //      47: "mm:ss.0"
        //      KOR fmt 55: "yyyy/mm/dd"
 
        // Built-in format codes
        if (is_null(self::builtInFormats)) {
            let self::builtInFormats = [];

            // General
            let self::builtInFormats[0] = \ZExcel\Style\NumberFormat::FORMAT_GENERAL;
            let self::builtInFormats[1] = strval(0);
            let self::builtInFormats[2] = "0.00";
            let self::builtInFormats[3] = "#,##0";
            let self::builtInFormats[4] = "#,##0.00";

            let self::builtInFormats[9] = "0%";
            let self::builtInFormats[10] = "0.00%";
            let self::builtInFormats[11] = "0.00E+00";
            let self::builtInFormats[12] = "# ?/?";
            let self::builtInFormats[13] = "# \?\?/\?\?";
            let self::builtInFormats[14] = "m/d/yyyy";                     // Despite ECMA 'mm-dd-yy';
            let self::builtInFormats[15] = "d-mmm-yy";
            let self::builtInFormats[16] = "d-mmm";
            let self::builtInFormats[17] = "mmm-yy";
            let self::builtInFormats[18] = "h:mm AM/PM";
            let self::builtInFormats[19] = "h:mm:ss AM/PM";
            let self::builtInFormats[20] = "h:mm";
            let self::builtInFormats[21] = "h:mm:ss";
            let self::builtInFormats[22] = "m/d/yyyy h:mm";                // Despite ECMA 'm/d/yy h:mm';

            let self::builtInFormats[37] = "#,##0_);(#,##0)";              //  Despite ECMA '#,##0 ;(#,##0)';
            let self::builtInFormats[38] = "#,##0_);[Red](#,##0)";         //  Despite ECMA '#,##0 ;[Red](#,##0)';
            let self::builtInFormats[39] = "#,##0.00_);(#,##0.00)";        //  Despite ECMA '#,##0.00;(#,##0.00)';
            let self::builtInFormats[40] = "#,##0.00_);[Red](#,##0.00)";   //  Despite ECMA '#,##0.00;[Red](#,##0.00)';

            // @FIXME zephir try to translate string
//             let self::builtInFormats[44] = "_(\"$\"* #,##0.00_);_(\"$\"* \(#,##0.00\);_(\"$\"* "-"??_);_(@_)";
            let self::builtInFormats[45] = "mm:ss";
            let self::builtInFormats[46] = "[h]:mm:ss";
            let self::builtInFormats[47] = "mm:ss.0";                      //  Despite ECMA 'mmss.0';
            let self::builtInFormats[48] = "##0.0E+0";
            let self::builtInFormats[49] = "@";

            // CHT
            let self::builtInFormats[27] = "[-404]e/m/d";
            let self::builtInFormats[30] = "m/d/yy";
            let self::builtInFormats[36] = "[$-404]e/m/d";
            let self::builtInFormats[50] = "[$-404]e/m/d";
            let self::builtInFormats[57] = "[$-404]e/m/d";

            // THA
            let self::builtInFormats[59] = "t0";
            let self::builtInFormats[60] = "t0.00";
            let self::builtInFormats[61] = "t#,##0";
            let self::builtInFormats[62] = "t#,##0.00";
            let self::builtInFormats[67] = "t0%";
            let self::builtInFormats[68] = "t0.00%";
            let self::builtInFormats[69] = "t# ?/?";
            let self::builtInFormats[70] = "t# \?\?/\?\?";

            // Flip array (for faster lookups)
            let self::flippedBuiltInFormats = array_flip(self::builtInFormats);
        }
    }

    /**
     * Get built-in format code
     *
     * @param    int        pIndex
     * @return    string
     */
    public static function builtInFormatCode(var pIndex)
    {
        // Clean parameter
        let pIndex = intval(pIndex);

        // Ensure built-in format codes are available
        self::fillBuiltInFormatCodes();
        
        // Lookup format code
        if (isset(self::builtInFormats[pIndex])) {
            return self::builtInFormats[pIndex];
        }

        return "";
    }

    /**
     * Get built-in format code index
     *
     * @param    string        formatCode
     * @return    int|boolean
     */
    public static function builtInFormatCodeIndex(var formatCode)
    {
        // Ensure built-in format codes are available
        self::fillBuiltInFormatCodes();

        // Lookup format code
        if (isset(self::flippedBuiltInFormats[formatCode])) {
            return self::flippedBuiltInFormats[formatCode];
        }

        return false;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->formatCode .
            this->builtInFormatCode .
            get_class(this)
        );
    }
    
    public static function setLowercaseCallback(matches) {
        return mb_strtolower(matches[0]);
    }
    
    public static function escapeQuotesCallback(matches) {
        return "\\" . implode("\\", str_split(matches[1]));
    }

    private static function formatAsDate(value, format)
    {
        var key, blocks, dateObj;
        
        let format = preg_replace("/^(\[\$[A-Z]*-[0-9A-F]*\])/i", "", format);

        // OpenOffice.org uses upper-case number formats, e.g. "YYYY", convert to lower-case;
        //    but we don"t want to change any quoted strings
        let format = preg_replace_callback("/(?:^|\")([^\"]*)(?:$|\")/", ["\\ZExcel\\Style\\NumberFormat", "setLowercaseCallback"], format);

        // Only process the non-quoted blocks for date format characters
        let blocks = explode("\"", format);
        
        for key, _ in blocks {
            if (key % 2 == 0) {
                let blocks[key] = strtr(blocks[key], self::dateFormatReplacements);
                
                if (!strpos(blocks[key], "A")) {
                    // 24-hour time format
                    let blocks[key] = strtr(blocks[key], self::dateFormatReplacements24);
                } else {
                    // 12-hour time format
                    let blocks[key] = strtr(blocks[key], self::dateFormatReplacements12);
                }
            }
        }
        
        let format = implode("\"", blocks);

        // escape any quoted characters so that DateTime format() will render them correctly
        let format = preg_replace_callback("/\"(.*)\"/U", ["\\ZExcel\\Style\\NumberFormat", "escapeQuotesCallback"], format);

        let dateObj = \ZExcel\Shared\Date::ExcelToPHPObject(value);
        let value = dateObj->format(format);
        
        return [value, format];
    }

    private static function formatAsPercentage(var value, var format)
    {
        var s;
        array m = [];
        
        if (format === self::FORMAT_PERCENTAGE) {
            let value = round((100 * value), 0) . "%";
        } else {
            if (preg_match("/\.[#0]+/i", format, m)) {
                let s = substr(m[0], 0, 1) . (strlen(m[0]) - 1);
                let format = str_replace(m[0], s, format);
            }
            
            if (preg_match("/^[#0]+/", format, m)) {
                let format = str_replace(m[0], strlen(m[0]), format);
            }
            
            let format = "%" . str_replace("%", "f%%", format);

            let value = sprintf(format, 100 * value);
        }
        
        return [value, format];
    }

    private static function formatAsFraction(var value, var format)
    {
        var sign, integerPart, decimalPart, decimalLength, decimalDivisor, adjustedDecimalPart, adjustedDecimalDivisor, gcd;
        
        let sign = (floatval(value) < 0) ? "-" : "";

        let integerPart = floor(abs(value));
        let decimalPart = trim(fmod(abs(value), 1), "0.");
        let decimalLength = strlen(decimalPart);
        let decimalDivisor = pow(10, decimalLength);

        let gcd = call_user_func("\ZExcel\Calculation\MathTrig::GCD", decimalPart, decimalDivisor);

        let adjustedDecimalPart = decimalPart / gcd;
        let adjustedDecimalDivisor = decimalDivisor / gcd;

        if ((strpos(format, strval(0)) !== false) || (strpos(format, "#") !== false) || (substr(format, 0, 3) == "? ?")) {
            if (integerPart == 0) {
                let integerPart = "";
            }
            
            let value = sprintf("%s%s %s/%s", sign, integerPart, adjustedDecimalPart, adjustedDecimalDivisor);
        } else {
            let adjustedDecimalPart = adjustedDecimalPart + ((double) integerPart * adjustedDecimalDivisor);
            let value = sprintf("%s%s/%s", sign, adjustedDecimalPart, adjustedDecimalDivisor);
        }
        
        return [value, format];
    }

    /**
     * @FIXME segment fault + zend_mm_heap corrupted
     */
    private static function complexNumberFormatMask(number, mask, level = 0)
    {
        var sign, numbers, mask, masks, result, result1, result2,
            r, block, blockValue, divisor, size, offset;
        
        let sign = (number < 0.0);
        let number = abs(number);
        
        if (strpos(mask, ".") !== false) {
            let numbers = explode(".", number . ".0");
            let masks = explode(".", mask . ".0");
            let result1 = self::complexNumberFormatMask(numbers[0], masks[0], 1);
            let result2 = strrev(self::complexNumberFormatMask(strrev(numbers[1]), strrev(masks[1]), 1));
            
            let result = result1 . "." . result2;
            
            if (sign == true) {
                let result = "-" . result;
            }
            
            return result;
        }
        
        let result = [];
        let r = preg_match_all("/0+/", mask, result, PREG_OFFSET_CAPTURE);
        
        if (r > 1) {
            let result = array_reverse(result[0]);

            for block in result {
                let divisor = "1" . block[0];
                let size = strlen(block[0]);
                let offset = block[1];

                let blockValue = sprintf(
                    "%0" . size . "d",
                    fmod(number, divisor)
                );
                
                let number = floor(number / (double) divisor);
                let mask = substr_replace(mask, blockValue, offset, size);
            }
            
            if (number > 0) {
                let mask = substr_replace(mask, number, offset, 0);
            }
            
            let result = mask;
        } else {
            let result = number;
        }

        if (sign == true) {
            let result = "-" . result;
        }
        
        return result;
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
        var sections, formatColor, color_regex, useThousands, n, m, number_regex, currencyFormat, currencyCode,
            scale, left, dec, right, minWidth, sprintf_pattern, writerInstance, functionn;
        array matches;
        
        // For now we do not treat strings although section 4 of a format code affects strings
        if (!is_numeric(value)) {
            return value;
        }

        // For "General" format code, we just pass the value although this is not entirely the way Excel does it,
        // it seems to round numbers to a total of 10 digits.
        if ((format === \ZExcel\Style\NumberFormat::FORMAT_GENERAL) || (format === \ZExcel\Style\NumberFormat::FORMAT_TEXT)) {
            return value;
        }

        // Convert any other escaped characters to quoted strings, e.g. (\T to "T")
        let format = preg_replace("/(\\\(.))(?=(?:[^\"]|\"[^\"]*\")*$)/u", "\"${2}\"", format);

        // Get the sections, there can be up to four sections, separated with a semi-colon (but only if not a quoted literal)
        let sections = preg_split("/(;)(?=(?:[^\"]|\"[^\"]*\")*$)/u", format);

        // Extract the relevant section depending on whether number is positive, negative, or zero?
        // Text not supported yet.
        // Here is how the sections apply to various values in Excel:
        //   1 section:   [POSITIVE/NEGATIVE/ZERO/TEXT]
        //   2 sections:  [POSITIVE/ZERO/TEXT] [NEGATIVE]
        //   3 sections:  [POSITIVE/TEXT] [NEGATIVE] [ZERO]
        //   4 sections:  [POSITIVE] [NEGATIVE] [ZERO] [TEXT]
        switch (count(sections)) {
            case 1:
                let format = sections[0];
                break;
            case 2:
                let format = (value >= 0) ? sections[0] : sections[1];
                let value = abs(value); // Use the absolute value
                break;
            case 3:
                if (value > 0) {
                    let format = sections[0];
                } else {
                    if (value < 0) {
                        let format = sections[1];
                    } else {
                        let format = sections[2];
                    }
                }
                
                let value = abs(value); // Use the absolute value
                break;
            case 4:
                if (value > 0) {
                    let format = sections[0];
                } else {
                    if (value < 0) {
                        let format = sections[1];
                    } else {
                        let format = sections[2];
                    }
                }
                
                let value = abs(value); // Use the absolute value
                break;
            default:
                // something is wrong, just use first section
                let format = sections[0];
                break;
        }
        
        // In Excel formats, "_" is used to add spacing,
        //    The following character indicates the size of the spacing, which we can"t do in HTML, so we just use a standard space
        let format = preg_replace("/_./", " ", format);

        // Save format with color information for later use below
        let formatColor = format;

        // Strip color information
        let color_regex = "/^\\[[a-zA-Z]+\\]/";
        let format = preg_replace(color_regex, "", format);

        // Let"s begin inspecting the format and converting the value to a formatted string

        //  Check for date/time characters (not inside quotes)
        let matches = [];
        if (preg_match("/(\\[\\$[A-Z]*-[0-9A-F]*\\])*[hmsdy](?=(?:[^\"]|\"[^\"]*\")*$)/miu", format, matches)) {
            // datetime format
            let sections = self::formatAsDate(value, format);
            let value = sections[0];
            let format = sections[1];
        } else {
            if (preg_match("/%$/", format)) {
                // % number format
                var sections = self::formatAsPercentage(value, format);
                let value = sections[0];
                let format = sections[1];
            } else {
                if (format === self::FORMAT_CURRENCY_EUR_SIMPLE) {
                    let value = "EUR " . sprintf("%1.2f", value);
                } else {
                    // Some non-number strings are quoted, so we"ll get rid of the quotes, likewise any positional * symbols
                    let format = str_replace(["\"", "*"], "", format);
    
                    // Find out if we need thousands separator
                    // This is indicated by a comma enclosed by a digit placeholder:
                    //        #,#   or   0,0
                    let useThousands = preg_match("/(#,#|0,0)/", format);
                    
                    if (useThousands) {
                        let format = preg_replace("/0,0/", "00", format);
                        let format = preg_replace("/#,#/", "##", format);
                    }
    
                    // Scale thousands, millions,...
                    // This is indicated by a number of commas after a digit placeholder:
                    //        #,   or    0.0,,
                    let scale = 1; // same as no scale
                    let matches = [];
                    
                    if (preg_match("/(#|0)(,+)/", format, matches)) {
                        let scale = pow(1000, strlen(matches[2]));
    
                        // strip the commas
                        let format = preg_replace("/0,+/", strval(0), format);
                        let format = preg_replace("/#,+/", "#", format);
                    }
                    
                    let m = [];
                    
                    if (preg_match("/#?.*\\?\/\\?/", format, m)) {
                        if (value != (int) value) {
                            let sections = self::formatAsFraction(value, format);
                            let value  = sections[0];
                            let format = sections[1];
                        }
                    } else {
                        // Handle the number itself
    
                        // scale number
                        let value = value / scale;
    
                        // Strip #
                        let format = preg_replace("/\\#/", strval(0), format);
    
                        let n = "/\\[[^\\]]+\\]/";
                        let m = preg_replace(n, "", format);
                        let number_regex = "/(0+)(\\.?)(0*)/";
                        let matches = [];
                        
                        if (preg_match(number_regex, m, matches)) {
                            let left = matches[1];
                            let dec = matches[2];
                            let right = matches[3];
    
                            // minimun width of formatted number (including dot)
                            let minWidth = strlen(left) + strlen(dec) + strlen(right);
                            if (useThousands) {
                                let value = number_format(
                                    value,
                                    strlen(right),
                                    \ZExcel\Shared\Stringg::getDecimalSeparator(),
                                    \ZExcel\Shared\Stringg::getThousandsSeparator()
                                );
                                let value = preg_replace(number_regex, value, format);
                            } else {
                                if (preg_match("/[0#]E[+-]0/i", format)) {
                                    //    Scientific format
                                    let value = sprintf("%5.2E", value);
                                } else {
                                    if (preg_match("/0([^\\d\\.]+)0/", format)) {
                                        let value = self::complexNumberFormatMask(value, format);
                                    } else {
                                        let sprintf_pattern = "%0" . minWidth . "." . strlen(right) . "f";
                                        let value = sprintf(sprintf_pattern, value);
                                        let value = preg_replace(number_regex, value, format);
                                    }
                                }
                            }
                        }
                    }
                
                    if (preg_match("/\\[\\$(.*)\\]/u", format, m)) {
                        //  Currency or Accounting
                        let currencyFormat = m[0];
                        let currencyCode = m[1];
                        let currencyCode = current(explode("-", currencyCode));
                        
                        if (currencyCode == "") {
                            let currencyCode = \ZExcel\Shared\Stringg::getCurrencyCode();
                        }
                        
                        let value = preg_replace("/\\[\\$([^\\]]*)\\]/u", currencyCode, value);
                    }
                }
            }
        }
        
        // Escape any escaped slashes to a single slash
        let format = preg_replace("/\\\\/u", "\\", format);

        // Additional formatting provided by callback function
        if (callback !== null) {
            let writerInstance = callback[0];
            let functionn = callback[1];
            let value = writerInstance->{functionn}(value, formatColor);
        }
        
        return value;
    }
}
