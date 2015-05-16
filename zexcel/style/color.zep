namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Color extends Supervisor implements ZIComparable
{
	/* Colors */
    const COLOR_BLACK      = "FF000000";
    const COLOR_WHITE      = "FFFFFFFF";
    const COLOR_RED        = "FFFF0000";
    const COLOR_DARKRED    = "FF800000";
    const COLOR_BLUE       = "FF0000FF";
    const COLOR_DARKBLUE   = "FF000080";
    const COLOR_GREEN      = "FF00FF00";
    const COLOR_DARKGREEN  = "FF008000";
    const COLOR_YELLOW     = "FFFFFF00";
    const COLOR_DARKYELLOW = "FF808000";

    /**
     * Indexed colors array
     *
     * @var array
     */
    protected static indexedColors;

    /**
     * ARGB - Alpha RGB
     *
     * @var string
     */
    protected argb = null;

    /**
     * Parent property name
     *
     * @var string
     */
    protected parentPropertyName;

    public function __construct(string pARGB = \ZExcel\Style\Color::COLOR_BLACK, boolean isSupervisor = false, boolean isConditional = false)
    {
        //    Supervisor?
        parent::__construct(isSupervisor);

        //    Initialise values
        if (!isConditional) {
            let this->argb = pARGB;
        }
    }

    public function bindParent(var parent, string parentPropertyName = null)
    {
        let this->parent = parent;
        let this->parentPropertyName = parentPropertyName;
        
        return this;
    }

    public function getSharedComponent()
    {
        switch (this->parentPropertyName) {
            case "endColor":
                return this->parent->getSharedComponent()->getEndColor();
            case "color":
                return this->parent->getSharedComponent()->getColor();
            case "startColor":
                return this->parent->getSharedComponent()->getStartColor();
        }
    }

    public function getStyleArray(array arry)
    {
        var key;
        
    	array output = [];
    	
        switch (this->parentPropertyName) {
            case "endColor":
                let key = "endcolor";
                break;
            case "color":
                let key = "color";
                break;
            case "startColor":
                let key = "startcolor";
                break;

        }
        
        let output[key] = arry; 
        
        return this->parent->getStyleArray(output);
    }

    public function applyFromArray(array pStyles = null)
    {
        if (is_array(pStyles)) {
            if (this->isSupervisor) {
                this->getActiveSheet()
                	->getStyle(this->getSelectedCells())
                	->applyFromArray(this->getStyleArray(pStyles));
            } else {
                if (array_key_exists("rgb", pStyles)) {
                    this->setRGB(pStyles["rgb"]);
                }
                if (array_key_exists("argb", pStyles)) {
                    this->setARGB(pStyles["argb"]);
                }
            }
        } else {
            throw new \ZExcel\Exception("Invalid style array passed.");
        }
        return this;
    }
    
    public function getARGB()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getARGB();
        }
        
        return this->argb;
    }

    public function setARGB(pValue = \ZExcel\Style\Color::COLOR_BLACK)
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = \ZExcel\Style\Color::COLOR_BLACK;
        }
        
        if (this->isSupervisor) {
            let styleArray = this->getStyleArray(["argb": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->argb = pValue;
        }
        
        return this;
    }

    public function getRGB()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getRGB();
        }
        return substr(this->argb, 2);
    }

    public function setRGB(var pValue = "000000")
    {
        var styleArray;
        
        if (pValue == "") {
            let pValue = "000000";
        }
        
        if (this->isSupervisor) {
        	let pValue = "FF" . pValue;
            let styleArray = this->getStyleArray(["argb": pValue]);
            this->getActiveSheet()
            	->getStyle(this->getSelectedCells())
            	->applyFromArray(styleArray);
        } else {
            let this->argb = "FF" . pValue;
        }
        
        return this;
    }

    private static function getColourComponent(string Rgb, int offset, boolean hex = true)
    {
        var color;
        
        let color = substr(Rgb, offset, 2);
        
        if (!hex) {
            let color = hexdec(color);
        }
        
        return color;
    }

    public static function getRed(string Rgb, boolean hex = true)
    {
        return self::getColourComponent(Rgb, strlen(Rgb) - 6, hex);
    }

    public static function getGreen(string Rgb, boolean hex = true)
    {
        return self::getColourComponent(Rgb, strlen(Rgb) - 4, hex);
    }

    public static function getBlue(string Rgb, boolean hex = true)
    {
        return self::getColourComponent(Rgb, strlen(Rgb) - 2, hex);
    }

    public static function changeBrightness(string hex, var adjustPercentage)
    {
        var rgba, red, green, blue, rgb;
        
        let rgba = (strlen(hex) == 8);

        let red = self::getRed(hex, false);
        let green = self::getGreen(hex, false);
        let blue = self::getBlue(hex, false);
        
        if (adjustPercentage > 0) {
            let red += (255 - red) * adjustPercentage;
            let green += (255 - green) * adjustPercentage;
            let blue += (255 - blue) * adjustPercentage;
        } else {
            let red += red * adjustPercentage;
            let green += green * adjustPercentage;
            let blue += blue * adjustPercentage;
        }

        if (red < 0) {
            let red = 0;
        } elseif (red > 255) {
            let red = 255;
        }
        
        if (green < 0) {
            let green = 0;
        } elseif (green > 255) {
            let green = 255;
        }
        if (blue < 0) {
            let blue = 0;
        } elseif (blue > 255) {
            let blue = 255;
        }

        let rgb = strtoupper(
            str_pad(dechex(red), 2, "0", 0) .
            str_pad(dechex(green), 2, "0", 0) .
            str_pad(dechex(blue), 2, "0", 0)
        );
        
        return ((rgba) ? "FF" : "") . rgb;
    }

    public static function indexedColor(int pIndex, boolean background = false)
    {
        // Clean parameter
        let pIndex = intval(pIndex);

        // Indexed colors
        if (is_null(self::indexedColors)) {
            let self::indexedColors = [
                    1: "FF000000",    //    System Colour #1 - Black
                    2: "FFFFFFFF",    //    System Colour #2 - White
                    3: "FFFF0000",    //    System Colour #3 - Red
                    4: "FF00FF00",    //    System Colour #4 - Green
                    5: "FF0000FF",    //    System Colour #5 - Blue
                    6: "FFFFFF00",    //    System Colour #6 - Yellow
                    7: "FFFF00FF",    //    System Colour #7- Magenta
                    8: "FF00FFFF",    //    System Colour #8- Cyan
                    9: "FF800000",    //    Standard Colour #9
                    10: "FF008000",    //    Standard Colour #10
                    11: "FF000080",    //    Standard Colour #11
                    12: "FF808000",    //    Standard Colour #12
                    13: "FF800080",    //    Standard Colour #13
                    14: "FF008080",    //    Standard Colour #14
                    15: "FFC0C0C0",    //    Standard Colour #15
                    16: "FF808080",    //    Standard Colour #16
                    17: "FF9999FF",    //    Chart Fill Colour #17
                    18: "FF993366",    //    Chart Fill Colour #18
                    19: "FFFFFFCC",    //    Chart Fill Colour #19
                    20: "FFCCFFFF",    //    Chart Fill Colour #20
                    21: "FF660066",    //    Chart Fill Colour #21
                    22: "FFFF8080",    //    Chart Fill Colour #22
                    23: "FF0066CC",    //    Chart Fill Colour #23
                    24: "FFCCCCFF",    //    Chart Fill Colour #24
                    25: "FF000080",    //    Chart Line Colour #25
                    26: "FFFF00FF",    //    Chart Line Colour #26
                    27: "FFFFFF00",    //    Chart Line Colour #27
                    28: "FF00FFFF",    //    Chart Line Colour #28
                    29: "FF800080",    //    Chart Line Colour #29
                    30: "FF800000",    //    Chart Line Colour #30
                    31: "FF008080",    //    Chart Line Colour #31
                    32: "FF0000FF",    //    Chart Line Colour #32
                    33: "FF00CCFF",    //    Standard Colour #33
                    34: "FFCCFFFF",    //    Standard Colour #34
                    35: "FFCCFFCC",    //    Standard Colour #35
                    36: "FFFFFF99",    //    Standard Colour #36
                    37: "FF99CCFF",    //    Standard Colour #37
                    38: "FFFF99CC",    //    Standard Colour #38
                    39: "FFCC99FF",    //    Standard Colour #39
                    40: "FFFFCC99",    //    Standard Colour #40
                    41: "FF3366FF",    //    Standard Colour #41
                    42: "FF33CCCC",    //    Standard Colour #42
                    43: "FF99CC00",    //    Standard Colour #43
                    44: "FFFFCC00",    //    Standard Colour #44
                    45: "FFFF9900",    //    Standard Colour #45
                    46: "FFFF6600",    //    Standard Colour #46
                    47: "FF666699",    //    Standard Colour #47
                    48: "FF969696",    //    Standard Colour #48
                    49: "FF003366",    //    Standard Colour #49
                    50: "FF339966",    //    Standard Colour #50
                    51: "FF003300",    //    Standard Colour #51
                    52: "FF333300",    //    Standard Colour #52
                    53: "FF993300",    //    Standard Colour #53
                    54: "FF993366",    //    Standard Colour #54
                    55: "FF333399",    //    Standard Colour #55
                    56: "FF333333"    //    Standard Colour #56
                ];
        }

        if (array_key_exists(pIndex, self::indexedColors)) {
            return new \ZExcel\Style\Color(self::indexedColors[pIndex]);
        }

        if (background) {
            return new \ZExcel\Style\Color(self::COLOR_WHITE);
        }
        return new \ZExcel\Style\Color(self::COLOR_BLACK);
    }

    public function getHashCode()
    {
        if (this->isSupervisor) {
            return this->getSharedComponent()->getHashCode();
        }
        return md5(
            this->argb .
            __CLASS__
        );
    }
}
