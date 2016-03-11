namespace ZExcel\Shared;

class Font
{
    /* Methods for resolving autosize value */
    const AUTOSIZE_METHOD_APPROX    = "approx";
    const AUTOSIZE_METHOD_EXACT        = "exact";

    /** Character set codes used by BIFF5-8 in Font records */
    const CHARSET_ANSI_LATIN                = 0x00;
    const CHARSET_SYSTEM_DEFAULT            = 0x01;
    const CHARSET_SYMBOL                    = 0x02;
    const CHARSET_APPLE_ROMAN               = 0x4D;
    const CHARSET_ANSI_JAPANESE_SHIFTJIS    = 0x80;
    const CHARSET_ANSI_KOREAN_HANGUL        = 0x81;
    const CHARSET_ANSI_KOREAN_JOHAB         = 0x82;
    const CHARSET_ANSI_CHINESE_SIMIPLIFIED  = 0x86;        //    gb2312
    const CHARSET_ANSI_CHINESE_TRADITIONAL  = 0x88;        //    big5
    const CHARSET_ANSI_GREEK                = 0xA1;
    const CHARSET_ANSI_TURKISH              = 0xA2;
    const CHARSET_ANSI_VIETNAMESE           = 0xA3;
    const CHARSET_ANSI_HEBREW               = 0xB1;
    const CHARSET_ANSI_ARABIC               = 0xB2;
    const CHARSET_ANSI_BALTIC               = 0xBA;
    const CHARSET_ANSI_CYRILLIC             = 0xCC;
    const CHARSET_ANSI_THAI                 = 0xDD;
    const CHARSET_ANSI_LATIN_II             = 0xEE;
    const CHARSET_OEM_LATIN_I               = 0xFF;

    //  XXX: Constants created!
    /** Font filenames */
    const ARIAL                                = "arial.ttf";
    const ARIAL_BOLD                        = "arialbd.ttf";
    const ARIAL_ITALIC                        = "ariali.ttf";
    const ARIAL_BOLD_ITALIC                    = "arialbi.ttf";

    const CALIBRI                            = "CALIBRI.TTF";
    const CALIBRI_BOLD                        = "CALIBRIB.TTF";
    const CALIBRI_ITALIC                    = "CALIBRII.TTF";
    const CALIBRI_BOLD_ITALIC                = "CALIBRIZ.TTF";

    const COMIC_SANS_MS                        = "comic.ttf";
    const COMIC_SANS_MS_BOLD                = "comicbd.ttf";

    const COURIER_NEW                        = "cour.ttf";
    const COURIER_NEW_BOLD                    = "courbd.ttf";
    const COURIER_NEW_ITALIC                = "couri.ttf";
    const COURIER_NEW_BOLD_ITALIC            = "courbi.ttf";

    const GEORGIA                            = "georgia.ttf";
    const GEORGIA_BOLD                        = "georgiab.ttf";
    const GEORGIA_ITALIC                    = "georgiai.ttf";
    const GEORGIA_BOLD_ITALIC                = "georgiaz.ttf";

    const IMPACT                            = "impact.ttf";

    const LIBERATION_SANS                    = "LiberationSans-Regular.ttf";
    const LIBERATION_SANS_BOLD                = "LiberationSans-Bold.ttf";
    const LIBERATION_SANS_ITALIC            = "LiberationSans-Italic.ttf";
    const LIBERATION_SANS_BOLD_ITALIC        = "LiberationSans-BoldItalic.ttf";

    const LUCIDA_CONSOLE                    = "lucon.ttf";
    const LUCIDA_SANS_UNICODE                = "l_10646.ttf";

    const MICROSOFT_SANS_SERIF                = "micross.ttf";

    const PALATINO_LINOTYPE                    = "pala.ttf";
    const PALATINO_LINOTYPE_BOLD            = "palab.ttf";
    const PALATINO_LINOTYPE_ITALIC            = "palai.ttf";
    const PALATINO_LINOTYPE_BOLD_ITALIC        = "palabi.ttf";

    const SYMBOL                            = "symbol.ttf";

    const TAHOMA                            = "tahoma.ttf";
    const TAHOMA_BOLD                        = "tahomabd.ttf";

    const TIMES_NEW_ROMAN                    = "times.ttf";
    const TIMES_NEW_ROMAN_BOLD                = "timesbd.ttf";
    const TIMES_NEW_ROMAN_ITALIC            = "timesi.ttf";
    const TIMES_NEW_ROMAN_BOLD_ITALIC        = "timesbi.ttf";

    const TREBUCHET_MS                        = "trebuc.ttf";
    const TREBUCHET_MS_BOLD                    = "trebucbd.ttf";
    const TREBUCHET_MS_ITALIC                = "trebucit.ttf";
    const TREBUCHET_MS_BOLD_ITALIC            = "trebucbi.ttf";

    const VERDANA                            = "verdana.ttf";
    const VERDANA_BOLD                        = "verdanab.ttf";
    const VERDANA_ITALIC                    = "verdanai.ttf";
    const VERDANA_BOLD_ITALIC                = "verdanaz.ttf";

    /**
     * AutoSize method
     *
     * @var string
     */
    private static autoSizeMethod = self::AUTOSIZE_METHOD_APPROX;

    /**
     * Path to folder containing TrueType font .ttf files
     *
     * @var string
     */
    private static trueTypeFontPath = null;

    private static _autoSizeMethods = [
        self::AUTOSIZE_METHOD_APPROX,
        self::AUTOSIZE_METHOD_EXACT
    ];
    
    /**
     * How wide is a default column for a given default font and size?
     * Empirical data found by inspecting real Excel files and reading off the pixel width
     * in Microsoft Office Excel 2007.
     *
     * @var array
     */
    public static defaultColumnWidths = [
        "Arial": [
             1: ["px": 24, "width": 12.00000000],
             2: ["px": 24, "width": 12.00000000],
             3: ["px": 32, "width": 10.66406250],
             4: ["px": 32, "width": 10.66406250],
             5: ["px": 40, "width": 10.00000000],
             6: ["px": 48, "width":  9.59765625],
             7: ["px": 48, "width":  9.59765625],
             8: ["px": 56, "width":  9.33203125],
             9: ["px": 64, "width":  9.14062500],
            10: ["px": 64, "width":  9.14062500]
        ],
        "Calibri": [
             1: ["px": 24, "width": 12.00000000],
             2: ["px": 24, "width": 12.00000000],
             3: ["px": 32, "width": 10.66406250],
             4: ["px": 32, "width": 10.66406250],
             5: ["px": 40, "width": 10.00000000],
             6: ["px": 48, "width":  9.59765625],
             7: ["px": 48, "width":  9.59765625],
             8: ["px": 56, "width":  9.33203125],
             9: ["px": 56, "width":  9.33203125],
            10: ["px": 64, "width":  9.14062500],
            11: ["px": 64, "width":  9.14062500]
        ],
        "Verdana": [
             1: ["px": 24, "width": 12.00000000],
             2: ["px": 24, "width": 12.00000000],
             3: ["px": 32, "width": 10.66406250],
             4: ["px": 32, "width": 10.66406250],
             5: ["px": 40, "width": 10.00000000],
             6: ["px": 48, "width":  9.59765625],
             7: ["px": 48, "width":  9.59765625],
             8: ["px": 64, "width":  9.14062500],
             9: ["px": 72, "width":  9.00000000],
            10: ["px": 72, "width":  9.00000000]
        ]
    ];

    /**
     * Set autoSize method
     *
     * @param string pValue
     * @return     boolean                    Success or failure
     */
    public static function setAutoSizeMethod(string pValue = self::AUTOSIZE_METHOD_APPROX)
    {
        if (!in_array(pValue, self::_autoSizeMethods)) {
            return false;
        }
        let self::autoSizeMethod = pValue;

        return true;
    }

    /**
     * Get autoSize method
     *
     * @return string
     */
    public static function getAutoSizeMethod()
    {
        return self::autoSizeMethod;
    }

    /**
     * Set the path to the folder containing .ttf files. There should be a trailing slash.
     * Typical locations on variout some platforms:
     *    <ul>
     *        <li>C:/Windows/Fonts/</li>
     *        <li>/usr/share/fonts/truetype/</li>
     *        <li>~/.fonts/</li>
     *    </ul>
     *
     * @param string pValue
     */
    public static function setTrueTypeFontPath(string pValue = "")
    {
        let self::trueTypeFontPath = pValue;
    }

    /**
     * Get the path to the folder containing .ttf files.
     *
     * @return string
     */
    public static function getTrueTypeFontPath()
    {
        return self::trueTypeFontPath;
    }

    /**
     * Calculate an (approximate) OpenXML column width, based on font size and text contained
     *
     * @param     \ZExcel\Style_Font            font            Font object
     * @param     \ZExcel\RichText|string    cellText        Text to calculate width
     * @param     integer                        rotation        Rotation angle
     * @param     \ZExcel\Style_Font|NULL    defaultFont    Font object
     * @return     integer        Column width
     */
    public static function calculateColumnWidth(<\ZExcel\Style\Font> font, var cellText = "", int rotation = 0, <\ZExcel\Style\Font> defaultFont = null)
    {
        var lineTexts, lineText, approximate, columnWidthAdjust, columnWidth, e;
        array lineWidths = []; 
        
        // If it is rich text, use plain text
        if (cellText instanceof \ZExcel\RichText) {
            let cellText = cellText->getPlainText();
        }

        // Special case if there are one or more newline characters ("\n")
        if (strpos(cellText, "\n") !== false) {
            let lineTexts = explode("\n", cellText);
            let lineWidths = [];
            
            for lineText in lineTexts {
                let lineWidths[] = self::calculateColumnWidth(font, lineText, rotation, defaultFont);
            }
            return max(lineWidths); // width of longest line in cell
        }

        // Try to get the exact text width in pixels
        let approximate = (self::autoSizeMethod == self::AUTOSIZE_METHOD_APPROX);
        
        if (!approximate) {
            let columnWidthAdjust = ceil(self::getTextWidthPixelsExact("n", font, 0) * 1.07);
            try {
                // Width of text in pixels excl. padding
                // and addition because Excel adds some padding, just use approx width of "n" glyph
                let columnWidth = self::getTextWidthPixelsExact(cellText, font, rotation) + columnWidthAdjust;
            } catch \ZExcel\Exception, e {
                let approximate = true;
            }
        }

        if (approximate) {
            let columnWidthAdjust = self::getTextWidthPixelsApprox("n", font, 0);
            // Width of text in pixels excl. padding, approximation
            // and addition because Excel adds some padding, just use approx width of "n" glyph
            let columnWidth = self::getTextWidthPixelsApprox(cellText, font, rotation) + columnWidthAdjust;
        }

        // Convert from pixel width to column width
        let columnWidth = \ZExcel\Shared\Drawing::pixelsToCellDimension(columnWidth, defaultFont);

        // Return
        return round(columnWidth, 6);
    }

    /**
     * Get GD text width in pixels for a string of text in a certain font at a certain rotation angle
     *
     * @param string text
     * @param \ZExcel\Style_Font
     * @param int rotation
     * @return int
     * @throws \ZExcel\Exception
     */
    public static function getTextWidthPixelsExact(string text, <\ZExcel\Style\Font> font, int rotation = 0)
    {
        var fontFile, textBox, textWidth, lowerLeftCornerX, lowerRightCornerX, upperRightCornerX, upperLeftCornerX;
        
        if (!function_exists("imagettfbbox")) {
            throw new \ZExcel\Exception("GD library needs to be enabled");
        }

        // font size should really be supplied in pixels in GD2,
        // but since GD2 seems to assume 72dpi, pixels and points are the same
        let fontFile = self::getTrueTypeFontFileFromFont(font);
        let textBox = imagettfbbox(font->getSize(), rotation, fontFile, text);

        // Get corners positions
        let lowerLeftCornerX  = textBox[0];
        let lowerRightCornerX = textBox[2];
        let upperRightCornerX = textBox[4];
        let upperLeftCornerX  = textBox[6];

        // Consider the rotation when calculating the width
        let textWidth = max(lowerRightCornerX - upperLeftCornerX, upperRightCornerX - lowerLeftCornerX);

        return textWidth;
    }

    /**
     * Get approximate width in pixels for a string of text in a certain font at a certain rotation angle
     *
     * @param string columnText
     * @param \ZExcel\Style_Font font
     * @param int rotation
     * @return int Text width in pixels (no padding added)
     */
    public static function getTextWidthPixelsApprox(string columnText, <\ZExcel\Style\Font> font = null, int rotation = 0)
    {
        var fontName, fontSize, columnWidth;
        
        let fontName = font->getName();
        let fontSize = font->getSize();

        // Calculate column width in pixels. We assume fixed glyph width. Result varies with font name and size.
        switch (fontName) {
            case "Calibri":
                // value 8.26 was found via interpolation by inspecting real Excel files with Calibri 11 font.
                let columnWidth = (int) (8.26 * \ZExcel\Shared\Stringg::CountCharacters(columnText));
                let columnWidth = columnWidth * fontSize / 11; // extrapolate from font size
                break;
            case "Arial":
                // value 7 was found via interpolation by inspecting real Excel files with Arial 10 font.
//                columnWidth = (int) (7 * \ZExcel\Shared\Stringg::CountCharacters(columnText));
                // value 8 was set because of experience in different exports at Arial 10 font.
                let columnWidth = (int) (8 * \ZExcel\Shared\Stringg::CountCharacters(columnText));
                let columnWidth = columnWidth * fontSize / 10; // extrapolate from font size
                break;
            case "Verdana":
                // value 8 was found via interpolation by inspecting real Excel files with Verdana 10 font.
                let columnWidth = (int) (8 * \ZExcel\Shared\Stringg::CountCharacters(columnText));
                let columnWidth = columnWidth * fontSize / 10; // extrapolate from font size
                break;
            default:
                // just assume Calibri
                let columnWidth = (int) (8.26 * \ZExcel\Shared\Stringg::CountCharacters(columnText));
                let columnWidth = columnWidth * fontSize / 11; // extrapolate from font size
                break;
        }

        // Calculate approximate rotated column width
        if (rotation !== 0) {
            if (rotation == -165) {
                // stacked text
                let columnWidth = 4; // approximation
            } else {
                // rotated text
                let columnWidth = columnWidth * cos(deg2rad(rotation)) + fontSize * abs(sin(deg2rad(rotation))) / 5; // approximation
            }
        }

        // pixel width is an integer
        return (int) columnWidth;
    }

    /**
     * Calculate an (approximate) pixel size, based on a font points size
     *
     * @param     int        fontSizeInPoints    Font size (in points)
     * @return     int        Font size (in pixels)
     */
    public static function fontSizeToPixels(float fontSizeInPoints = 11) -> int
    {
        return (int) ((4 / 3) * fontSizeInPoints);
    }

    /**
     * Calculate an (approximate) pixel size, based on inch size
     *
     * @param     int        sizeInInch    Font size (in inch)
     * @return     int        Size (in pixels)
     */
    public static function inchSizeToPixels(float sizeInInch = 1) -> float
    {
        return (sizeInInch * 96);
    }

    /**
     * Calculate an (approximate) pixel size, based on centimeter size
     *
     * @param     int        sizeInCm    Font size (in centimeters)
     * @return     int        Size (in pixels)
     */
    public static function centimeterSizeToPixels(float sizeInCm = 1) -> float
    {
        return (sizeInCm * 37.795275591);
    }

    /**
     * Returns the font path given the font
     *
     * @param \ZExcel\Style_Font
     * @return string Path to TrueType font file
     */
    public static function getTrueTypeFontFileFromFont(<\ZExcel\Style\Font> font)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Returns the associated charset for the font name.
     *
     * @param string name Font name
     * @return int Character set code
     */
    public static function getCharsetFromFontName(string name)
    {
        switch (name) {
            // Add more cases. Check FONT records in real Excel files.
            case "EucrosiaUPC":
                return self::CHARSET_ANSI_THAI;
            case "Wingdings":
                return self::CHARSET_SYMBOL;
            case "Wingdings 2":
                return self::CHARSET_SYMBOL;
            case "Wingdings 3":
                return self::CHARSET_SYMBOL;
            default:
                return self::CHARSET_ANSI_LATIN;
        }
    }

    /**
     * Get the effective column width for columns without a column dimension or column with width -1
     * For example, for Calibri 11 this is 9.140625 (64 px)
     *
     * @param \ZExcel\Style_Font font The workbooks default font
     * @param boolean pPixels true = return column width in pixels, false = return in OOXML units
     * @return mixed Column width
     */
    public static function getDefaultColumnWidthByFont(<\ZExcel\Style\Font> font, boolean pPixels = false)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the effective row height for rows without a row dimension or rows with height -1
     * For example, for Calibri 11 this is 15 points
     *
     * @param \ZExcel\Style_Font font The workbooks default font
     * @return float Row height in points
     */
    public static function getDefaultRowHeightByFont(<\ZExcel\Style\Font> font)
    {
        throw new \Exception("Not implemented yet!");
    }
}
