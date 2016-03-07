namespace ZExcel\Shared;

class Stringg
{
    const STRING_REGEXP_FRACTION = "(-?)(\d+)\s+(\d+\/\d+)";


    /**
     * Control characters array
     *
     * @var string[]
     */
    private static _controlCharacters;

    /**
     * SYLK Characters array
     *
     * @var array
     */
    private static _SYLKCharacters;

    /**
     * Decimal separator
     *
     * @var string
     */
    private static decimalSeparator;

    /**
     * Thousands separator
     *
     * @var string
     */
    private static thousandsSeparator;

    /**
     * Currency code
     *
     * @var string
     */
    private static currencyCode;

    /**
     * Is mbstring extension avalable?
     *
     * @var boolean
     */
    private static _isMbstringEnabled;

    /**
     * Is iconv extension avalable?
     *
     * @var boolean
     */
    private static _isIconvEnabled;

    /**
     * Build control characters array
     */
    private static function _buildControlCharacters()
    {
        var i, find, replace;
        
        for i in range(0, 31) {
            if (i != 9 && i != 10 && i != 13) {
                let find = "_x" . sprintf("%04s" , strtoupper(dechex(i))) . "_";
                let replace = chr(i);
                let self::_controlCharacters[find] = replace;
            }
        }
    }

    /**
     * Build SYLK characters array
     */
    private static function _buildSYLKCharacters()
    {
        let self::_SYLKCharacters = [
            "\x1B 0": chr(0),
            "\x1B 1": chr(1),
            "\x1B 2": chr(2),
            "\x1B 3": chr(3),
            "\x1B 4": chr(4),
            "\x1B 5": chr(5),
            "\x1B 6": chr(6),
            "\x1B 7": chr(7),
            "\x1B 8": chr(8),
            "\x1B 9": chr(9),
            "\x1B :": chr(10),
            "\x1B ;": chr(11),
            "\x1B <": chr(12),
            "\x1B :": chr(13),
            "\x1B >": chr(14),
            "\x1B ?": chr(15),
            "\x1B!0": chr(16),
            "\x1B!1": chr(17),
            "\x1B!2": chr(18),
            "\x1B!3": chr(19),
            "\x1B!4": chr(20),
            "\x1B!5": chr(21),
            "\x1B!6": chr(22),
            "\x1B!7": chr(23),
            "\x1B!8": chr(24),
            "\x1B!9": chr(25),
            "\x1B!:": chr(26),
            "\x1B!;": chr(27),
            "\x1B!<": chr(28),
            "\x1B!=": chr(29),
            "\x1B!>": chr(30),
            "\x1B!?": chr(31),
            "\x1B'?": chr(127),
            "\x1B(0": "€", // 128 in CP1252
            "\x1B(2": "‚", // 130 in CP1252
            "\x1B(3": "ƒ", // 131 in CP1252
            "\x1B(4": "„", // 132 in CP1252
            "\x1B(5": "…", // 133 in CP1252
            "\x1B(6": "†", // 134 in CP1252
            "\x1B(7": "‡", // 135 in CP1252
            "\x1B(8": "ˆ", // 136 in CP1252
            "\x1B(9": "‰", // 137 in CP1252
            "\x1B(:": "Š", // 138 in CP1252
            "\x1B(;": "‹", // 139 in CP1252
            "\x1BNj": "Œ", // 140 in CP1252
            "\x1B(>": "Ž", // 142 in CP1252
            "\x1B)1": "‘", // 145 in CP1252
            "\x1B)2": "’", // 146 in CP1252
            "\x1B)3": "“", // 147 in CP1252
            "\x1B)4": "”", // 148 in CP1252
            "\x1B)5": "•", // 149 in CP1252
            "\x1B)6": "–", // 150 in CP1252
            "\x1B)7": "—", // 151 in CP1252
            "\x1B)8": "˜", // 152 in CP1252
            "\x1B)9": "™", // 153 in CP1252
            "\x1B):": "š", // 154 in CP1252
            "\x1B);": "›", // 155 in CP1252
            "\x1BNz": "œ", // 156 in CP1252
            "\x1B)>": "ž", // 158 in CP1252
            "\x1B)?": "Ÿ", // 159 in CP1252
            "\x1B*0": " ", // 160 in CP1252
            "\x1BN!": "¡", // 161 in CP1252
            "\x1BN\"": "¢", // 162 in CP1252
            "\x1BN#": "£", // 163 in CP1252
            "\x1BN(": "¤", // 164 in CP1252
            "\x1BN%": "¥", // 165 in CP1252
            "\x1B*6": "¦", // 166 in CP1252
            "\x1BN\"": "§", // 167 in CP1252
            "\x1BNH ": "¨", // 168 in CP1252
            "\x1BNS": "©", // 169 in CP1252
            "\x1BNc": "ª", // 170 in CP1252
            "\x1BN+": "«", // 171 in CP1252
            "\x1B*<": "¬", // 172 in CP1252
            "\x1B*=": "­", // 173 in CP1252
            "\x1BNR": "®", // 174 in CP1252
            "\x1B*?": "¯", // 175 in CP1252
            "\x1BN0": "°", // 176 in CP1252
            "\x1BN1": "±", // 177 in CP1252
            "\x1BN2": "²", // 178 in CP1252
            "\x1BN3": "³", // 179 in CP1252
            "\x1BNB ": "´", // 180 in CP1252
            "\x1BN5": "µ", // 181 in CP1252
            "\x1BN6": "¶", // 182 in CP1252
            "\x1BN7": "·", // 183 in CP1252
            "\x1B+8": "¸", // 184 in CP1252
            "\x1BNQ": "¹", // 185 in CP1252
            "\x1BNk": "º", // 186 in CP1252
            "\x1BN;": "»", // 187 in CP1252
            "\x1BN<": "¼", // 188 in CP1252
            "\x1BN=": "½", // 189 in CP1252
            "\x1BN>": "¾", // 190 in CP1252
            "\x1BN?": "¿", // 191 in CP1252
            "\x1BNAA": "À", // 192 in CP1252
            "\x1BNBA": "Á", // 193 in CP1252
            "\x1BNCA": "Â", // 194 in CP1252
            "\x1BNDA": "Ã", // 195 in CP1252
            "\x1BNHA": "Ä", // 196 in CP1252
            "\x1BNJA": "Å", // 197 in CP1252
            "\x1BNa": "Æ", // 198 in CP1252
            "\x1BNKC": "Ç", // 199 in CP1252
            "\x1BNAE": "È", // 200 in CP1252
            "\x1BNBE": "É", // 201 in CP1252
            "\x1BNCE": "Ê", // 202 in CP1252
            "\x1BNHE": "Ë", // 203 in CP1252
            "\x1BNAI": "Ì", // 204 in CP1252
            "\x1BNBI": "Í", // 205 in CP1252
            "\x1BNCI": "Î", // 206 in CP1252
            "\x1BNHI": "Ï", // 207 in CP1252
            "\x1BNb": "Ð", // 208 in CP1252
            "\x1BNDN": "Ñ", // 209 in CP1252
            "\x1BNAO": "Ò", // 210 in CP1252
            "\x1BNBO": "Ó", // 211 in CP1252
            "\x1BNCO": "Ô", // 212 in CP1252
            "\x1BNDO": "Õ", // 213 in CP1252
            "\x1BNHO": "Ö", // 214 in CP1252
            "\x1B-7": "×", // 215 in CP1252
            "\x1BNi": "Ø", // 216 in CP1252
            "\x1BNAU": "Ù", // 217 in CP1252
            "\x1BNBU": "Ú", // 218 in CP1252
            "\x1BNCU": "Û", // 219 in CP1252
            "\x1BNHU": "Ü", // 220 in CP1252
            "\x1B-=": "Ý", // 221 in CP1252
            "\x1BNl": "Þ", // 222 in CP1252
            "\x1BN{": "ß", // 223 in CP1252
            "\x1BNAa": "à", // 224 in CP1252
            "\x1BNBa": "á", // 225 in CP1252
            "\x1BNCa": "â", // 226 in CP1252
            "\x1BNDa": "ã", // 227 in CP1252
            "\x1BNHa": "ä", // 228 in CP1252
            "\x1BNJa": "å", // 229 in CP1252
            "\x1BNq": "æ", // 230 in CP1252
            "\x1BNKc": "ç", // 231 in CP1252
            "\x1BNAe": "è", // 232 in CP1252
            "\x1BNBe": "é", // 233 in CP1252
            "\x1BNCe": "ê", // 234 in CP1252
            "\x1BNHe": "ë", // 235 in CP1252
            "\x1BNAi": "ì", // 236 in CP1252
            "\x1BNBi": "í", // 237 in CP1252
            "\x1BNCi": "î", // 238 in CP1252
            "\x1BNHi": "ï", // 239 in CP1252
            "\x1BNs": "ð", // 240 in CP1252
            "\x1BNDn": "ñ", // 241 in CP1252
            "\x1BNAo": "ò", // 242 in CP1252
            "\x1BNBo": "ó", // 243 in CP1252
            "\x1BNCo": "ô", // 244 in CP1252
            "\x1BNDo": "õ", // 245 in CP1252
            "\x1BNHo": "ö", // 246 in CP1252
            "\x1B/7": "÷", // 247 in CP1252
            "\x1BNy": "ø", // 248 in CP1252
            "\x1BNAu": "ù", // 249 in CP1252
            "\x1BNBu": "ú", // 250 in CP1252
            "\x1BNCu": "û", // 251 in CP1252
            "\x1BNHu": "ü", // 252 in CP1252
            "\x1B/=": "ý", // 253 in CP1252
            "\x1BN|": "þ", // 254 in CP1252
            "\x1BNHy": "ÿ" // 255 in CP1252
        ];
    }

    /**
     * Get whether mbstring extension is available
     *
     * @return boolean
     */
    public static function getIsMbstringEnabled()
    {
        if (isset(self::_isMbstringEnabled)) {
            return self::_isMbstringEnabled;
        }

        let self::_isMbstringEnabled = function_exists("mb_convert_encoding") ? true : false;

        return self::_isMbstringEnabled;
    }

    /**
     * Get whether iconv extension is available
     *
     * @return boolean
     */
    public static function getIsIconvEnabled()
    {
        if (isset(self::_isIconvEnabled)) {
            return self::_isIconvEnabled;
        }
        
        // Fail if iconv doesn't exist
        if (!function_exists("iconv")) {
            let self::_isIconvEnabled = false;
            return false;
        }
        
        // Sometimes iconv is not working, and e.g. iconv('UTF-8', 'UTF-16LE', 'x') just returns false,
        if (!iconv("UTF-8", "UTF-16LE", "x")) {
            let self::_isIconvEnabled = false;
            return false;
        }
        
        // Sometimes iconv_substr('A', 0, 1, 'UTF-8') just returns false in PHP 5.2.0
        // we cannot use iconv in that case either (http://bugs.php.net/bug.php?id=37773)
        if (!iconv_substr("A", 0, 1, "UTF-8")) {
            let self::_isIconvEnabled = false;
            return false;
        }
        
        // CUSTOM: IBM AIX iconv() does not work
        if ( defined("PHP_OS") && stristr(PHP_OS, "AIX")
                && defined("ICONV_IMPL") && (strcasecmp(ICONV_IMPL, "unknown") == 0)
                && defined("ICONV_VERSION") && (strcasecmp(ICONV_VERSION, "unknown") == 0)) {
            let self::_isIconvEnabled = false;
            return false;
        }
        
        // If we reach here no problems were detected with iconv
        let self::_isIconvEnabled = true;
        return true;
    }

    public static function buildCharacterSets()
    {
        if(empty(self::_controlCharacters)) {
            self::_buildControlCharacters();
        }
        
        if(empty(self::_SYLKCharacters)) {
            self::_buildSYLKCharacters();
        }
    }

    /**
     * Convert from OpenXML escaped control character to PHP control character
     *
     * Excel 2007 team:
     * ----------------
     * That's correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param     string    value    Value to unescape
     * @return     string
     */
    public static function controlCharacterOOXML2PHP(var value = "") -> string
    {
        return str_replace(
            array_keys(self::_controlCharacters),
            array_values(self::_controlCharacters),
            value
        );
    }

    /**
     * Convert from PHP control character to OpenXML escaped control character
     *
     * Excel 2007 team:
     * ----------------
     * That"s correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param     string    value    Value to escape
     * @return     string
     */
    public static function controlCharacterPHP2OOXML(var value = "") -> string
    {
        return str_replace(
            array_values(self::_controlCharacters),
            array_keys(self::_controlCharacters),
            value
        );
    }

    /**
     * Try to sanitize UTF8, stripping invalid byte sequences. Not perfect. Does not surrogate characters.
     *
     * @param string value
     * @return string
     */
    public static function sanitizeUTF8(var value)
    {
        if (self::getIsIconvEnabled()) {
            let value = iconv("UTF-8", "UTF-8", value);
            
            return value;
        }

        if (self::getIsMbstringEnabled()) {
            let value = mb_convert_encoding(value, "UTF-8", "UTF-8");
            
            return value;
        }

        // else, no conversion
        return value;
    }

    /**
     * Check if a string contains UTF8 data
     *
     * @param string value
     * @return boolean
     */
    public static function isUTF8(var value = "") -> boolean
    {
        return (value === "" || preg_match("'/^./su", value) === 1);
    }

    /**
     * Formats a numeric value as a string for output in various output writers forcing
     * point as decimal separator in case locale is other than English.
     *
     * @param mixed value
     * @return string
     */
    public static function formatNumber(var value) -> string
    {
        if (is_float(value)) {
            return str_replace(",", ".", value);
        }
        
        return (string) value;
    }

    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (8-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3
     *
     * @param string  value    UTF-8 encoded string
     * @param mixed[] arrcRuns Details of rich text runs in value
     * @return string
     */
    public static function utf8toBIFF8UnicodeShort(var value, var arrcRuns = [])
    {
        var ln, opt, data, cRun;
        
        // character count
        let ln = self::CountCharacters(value, "UTF-8");
        // option flags
        if (empty(arrcRuns)) {
            let opt = (self::getIsIconvEnabled() || self::getIsMbstringEnabled()) ?
                0x0001 : 0x0000;
            let data = pack("CC", ln, opt);
            // characters
            let data = data . self::ConvertEncoding(value, "UTF-16LE", "UTF-8");
        } else {
            let data = pack("vC", ln, 0x09);
            let data = data . pack("v", count(arrcRuns));
            // characters
            let data = data . self::ConvertEncoding(value, "UTF-16LE", "UTF-8");
            for cRun in arrcRuns {
                let data = data . pack("v", cRun["strlen"]);
                let data = data . pack("v", cRun["fontidx"]);
            }
        }
        
        return data;
    }

    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (16-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3
     *
     * @param string value UTF-8 encoded string
     * @return string
     */
    public static function utf8toBIFF8UnicodeLong(var value)
    {
        var ln, opt, chars, data;
        
        // character count
        let ln = self::CountCharacters(value, "UTF-8");
        
        // option flags
        if (self::getIsIconvEnabled() || self::getIsMbstringEnabled()) {
            let opt = 0x0001;
        } else {
            let opt = 0x0000;
        }
        
        // characters
        let chars = self::ConvertEncoding(value, "UTF-16LE", "UTF-8");

        let data = pack("vC", ln, opt) . chars;
        
        return data;
    }

    /**
     * Convert string from one encoding to another. First try mbstring, then iconv, finally strlen
     *
     * @param string value
     * @param string to Encoding to convert to, e.g. 'UTF-8'
     * @param string from Encoding to convert from, e.g. 'UTF-16LE'
     * @return string
     */
    public static function convertEncoding(var value, var to, var from)
    {
        if (self::getIsIconvEnabled()) {
            return iconv(from, to, value);
        }

        if (self::getIsMbstringEnabled()) {
            return mb_convert_encoding(value, to, from);
        }

        if (from == "UTF-16LE"){
            return self::utf16_decode(value, false);
        } elseif (from == "UTF-16BE"){
            return self::utf16_decode(value);
        }
        
        // else, no conversion
        return value;
    }

    /**
     * Decode UTF-16 encoded strings.
     *
     * Can handle both BOM'ed data and un-BOM'ed data.
     * Assumes Big-Endian byte order if no BOM is available.
     * This function was taken from http://php.net/manual/en/function.utf8-decode.php
     * and bom_be parameter added.
     *
     * @param   string  str  UTF-16 encoded data to decode.
     * @return  string  UTF-8 / ISO encoded data.
     * @access  public
     * @version 0.2 / 2010-05-13
     * @author  Rasmus Andersson {@link http://rasmusandersson.se/}
     * @author vadik56
     */
    public static function utf16_decode(var str, var bom_be = true)
    {
        var c0, c1, len, newstr, val, i;
        
        if( strlen(str) < 2 ) {
            return str;
        }
        
        let c0 = ord(substr(str, 0, 1));
        let c1 = ord(substr(str, 1, 1));
        
        if( c0 == 0xfe && c1 == 0xff ) {
            let str = substr(str,2);
        } elseif( c0 == 0xff && c1 == 0xfe ) {
            let str = substr(str,2);
            let bom_be = false;
        }
        
        let len = strlen(str);
        let newstr = "";
        
        for i in range(0, len - 1) {
            if( bom_be ) {
                let val = ord(substr(str, i, 1)) << 4;
                let val = val + ord(substr(str, i + 1, 1));
            } else {
                let val = ord(substr(str, i + 1, 1)) << 4;
                let val = val + ord(substr(str, i, 1));
            }
            
            if (val == 0x228) {
                let newstr = newstr . "\n";
            } else {
                let newstr = newstr . chr(val);
            }
            
            let i = i + 1; // equal with for step i + 2
        }
        
        return newstr;
    }

    /**
     * Get character count. First try mbstring, then iconv, finally strlen
     *
     * @param string value
     * @param string enc Encoding
     * @return int Character count
     */
    public static function countCharacters(value, enc = "UTF-8")
    {
        if (self::getIsMbstringEnabled()) {
            return mb_strlen(value, enc);
        }

        if (self::getIsIconvEnabled()) {
            return iconv_strlen(value, enc);
        }

        // else strlen
        return strlen(value);
    }

    /**
     * Get a substring of a UTF-8 encoded string. First try mbstring, then iconv, finally strlen
     *
     * @param string pValue UTF-8 encoded string
     * @param int pStart Start offset
     * @param int pLength Maximum number of characters in substring
     * @return string
     */
    public static function substring(pValue = "", pStart = 0, pLength = 0)
    {
        if (self::getIsMbstringEnabled()) {
            return mb_substr(pValue, pStart, pLength, "UTF-8");
        }

        if (self::getIsIconvEnabled()) {
            return iconv_substr(pValue, pStart, pLength, "UTF-8");
        }

        // else substr
        return substr(pValue, pStart, pLength);
    }

    /**
     * Convert a UTF-8 encoded string to upper case
     *
     * @param string pValue UTF-8 encoded string
     * @return string
     */
    public static function strToUpper(string pValue = "") -> string
    {
        if (function_exists("mb_convert_case")) {
            return mb_convert_case(pValue, MB_CASE_UPPER, "UTF-8");
        }
        
        return strtoupper(pValue);
    }

    /**
     * Convert a UTF-8 encoded string to lower case
     *
     * @param string pValue UTF-8 encoded string
     * @return string
     */
    public static function strToLower(string pValue = "") -> string
    {
        if (function_exists("mb_convert_case")) {
            return mb_convert_case(pValue, MB_CASE_LOWER, "UTF-8");
        }
        
        return strtolower(pValue);
    }

    /**
     * Convert a UTF-8 encoded string to title/proper case
     *    (uppercase every first character in each word, lower case all other characters)
     *
     * @param string pValue UTF-8 encoded string
     * @return string
     */
    public static function strToTitle(string pValue = "") -> string
    {
        if (function_exists("mb_convert_case")) {
            return mb_convert_case(pValue, MB_CASE_TITLE, "UTF-8");
        }
        return ucwords(pValue);
    }

    public static function mb_is_upper(chr)
    {
        return mb_strtolower(chr, "UTF-8") != chr;
    }

    public static function mb_str_split(string stringg)
    {
        return preg_split("/(?<!^)(?!$)/u", stringg);
    }

    /**
     * Reverse the case of a string, so that all uppercase characters become lowercase
     *    and all lowercase characters become uppercase
     *
     * @param string pValue UTF-8 encoded string
     * @return string
     */
    public static function StrCaseReverse(pValue = "")
    {
        var characters, character, k;
        
        if (self::getIsMbstringEnabled()) {
            let characters = self::mb_str_split(pValue);
            for k, character in characters {
                if(self::mb_is_upper(character)) {
                    let characters[k] = mb_strtolower(character, "UTF-8");
                } else {
                    let character = mb_strtoupper(character, "UTF-8");
                }
            }
            return implode("", characters);
        }
        return strtolower(pValue) ^ strtoupper(pValue) ^ pValue;
    }
    
    /**
     * Identify whether a string contains a fractional numeric value,
     *    and convert it to a numeric if it is
     *
     * @param string &operand string value to test
     * @return boolean
     */
    public static function convertToNumberIfFraction(var operand)
    {
        var sign, fractionFormula;
        array match = [];
        
        if (preg_match("/^" . self::STRING_REGEXP_FRACTION . "$/i", operand, match)) {
            let sign = (match[1] == "-") ? "-" : "+";
            let fractionFormula = "=" . sign . match[2] . sign . match[3];
            let operand = \ZExcel\Calculation::getInstance()->_calculateFormulaValue(fractionFormula);
            
            return operand;
        }
        
        return false;
    }

    /**
     * Get the decimal separator. If it has not yet been set explicitly, try to obtain number
     * formatting information from locale.
     *
     * @return string
     */
    public static function getDecimalSeparator() -> string
    {
        var localeconv;
        
        if (!isset(self::decimalSeparator)) {
            let localeconv = localeconv();
            
            if (localeconv["decimal_point"] != "") {
                let self::decimalSeparator = localeconv["decimal_point"];
            } else {
                let self::decimalSeparator = localeconv["mon_decimal_point"];
            }

            if (self::decimalSeparator == "") {
                // Default to .
                let self::decimalSeparator = ".";
            }
        }
        return self::decimalSeparator;
    }

    /**
     * Set the decimal separator. Only used by \ZExcel\Style_NumberFormat::toFormattedString()
     * to format output by \ZExcel\Writer_HTML and \ZExcel\Writer_PDF
     *
     * @param string pValue Character for decimal separator
     */
    public static function setDecimalSeparator(string pValue = ".")
    {
        let self::decimalSeparator = pValue;
    }

    /**
     * Get the thousands separator. If it has not yet been set explicitly, try to obtain number
     * formatting information from locale.
     *
     * @return string
     */
    public static function getThousandsSeparator() -> string
    {
        var localeconv;
        
        if (!isset(self::thousandsSeparator)) {
            let localeconv = localeconv();
            
            if (localeconv["thousands_sep"] != "") {
                let self::thousandsSeparator = localeconv["thousands_sep"];
            } else {
                let self::thousandsSeparator = localeconv["mon_thousands_sep"];
            }

            if (self::thousandsSeparator == "") {
                // Default to .
                let self::thousandsSeparator = ",";
            }
        }
        
        return self::thousandsSeparator;
    }

    /**
     * Set the thousands separator. Only used by \ZExcel\Style_NumberFormat::toFormattedString()
     * to format output by \ZExcel\Writer_HTML and \ZExcel\Writer_PDF
     *
     * @param string pValue Character for thousands separator
     */
    public static function setThousandsSeparator(string pValue = ",")
    {
        let self::thousandsSeparator = pValue;
    }

    /**
     *    Get the currency code. If it has not yet been set explicitly, try to obtain the
     *        symbol information from locale.
     *
     * @return string
     */
    public static function getCurrencyCode() -> string
    {
        var localeconv;
        
        if (!isset(self::currencyCode)) {
            let localeconv = localeconv();
            
            if (localeconv["currency_symbol"] != "") {
                let self::currencyCode = localeconv["currency_symbol"];
            } else {
                let self::currencyCode = localeconv["int_curr_symbol"];
            }

            if (self::currencyCode == "") {
                // Default to 
                let self::currencyCode = "$";
            }
        }
        
        return self::currencyCode;
    }

    /**
     * Set the currency code. Only used by \ZExcel\Style_NumberFormat::toFormattedString()
     *        to format output by \ZExcel\Writer_HTML and \ZExcel\Writer_PDF
     *
     * @param string pValue Character for currency code
     */
    public static function setCurrencyCode(string pValue = "$")
    {
        let self::currencyCode = pValue;
    }

    /**
     * Convert SYLK encoded string to UTF-8
     *
     * @param string pValue
     * @return string UTF-8 encoded string
     */
    public static function SYLKtoUTF8(var pValue = "")
    {
        var k, v;
        
        // If there is no escape character in the string there is nothing to do
        if (strpos(pValue, "") === false) {
            return pValue;
        }

        for k, v in self::_SYLKCharacters {
            let pValue = str_replace(k, v, pValue);
        }

        return pValue;
    }

    /**
     * Retrieve any leading numeric part of a string, or return the full string if no leading numeric
     *    (handles basic integer or float, but not exponent or non decimal)
     *
     * @param    string    value
     * @return    mixed    string or only the leading numeric part of the string
     */
    public static function testStringAsNumeric(var value)
    {
        var v;
    
        if (is_numeric(value)) {
            return value;
        }
        
        let v = floatval(value);
        
        return ((is_numeric(substr(value, 0, strlen(v)))) ? v : value);
    }
}
