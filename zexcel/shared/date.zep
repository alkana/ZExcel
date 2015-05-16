namespace ZExcel\Shared;

class Date
{
    /** constants */
    const CALENDAR_WINDOWS_1900 = 1900;        //    Base date of 1st Jan 1900 = 1.0
    const CALENDAR_MAC_1904 = 1904;            //    Base date of 2nd Jan 1904 = 1.0

    /*
     * Base calendar year to use for calculations
     *
     * @private
     * @var    int
     */
    protected static excelBaseDate = self::CALENDAR_WINDOWS_1900;
    
    private static possibleDateFormatCharacters = "eymdHs";
    
    /*
     * Names of the months of the year, indexed by shortname
     * Planned usage for locale settings
     *
     * @public
     * @var    string[]
     */
    public static monthNames;

    /*
     * Names of the months of the year, indexed by shortname
     * Planned usage for locale settings
     *
     * @public
     * @var    string[]
     */
    public static numberSuffixes;

    public static function initStaticArray()
    {
    	if (self::numberSuffixes == null) {
    		let self::numberSuffixes = [
	    		"st",
		        "nd",
		        "rd",
		        "th"
    		];
    	}
    	
    	if (self::monthNames == null) {
    		let self::monthNames = [
		        "Jan": "January",
		        "Feb": "February",
		        "Mar": "March",
		        "Apr": "April",
		        "May": "May",
		        "Jun": "June",
		        "Jul": "July",
		        "Aug": "August",
		        "Sep": "September",
		        "Oct": "October",
		        "Nov": "November",
		        "Dec": "December"
    		];
    	}
    }

    /**
     * Set the Excel calendar (Windows 1900 or Mac 1904)
     *
     * @param     integer    $baseDate            Excel base date (1900 or 1904)
     * @return     boolean                        Success or failure
     */
    public static function setExcelCalendar(baseDate)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Return the Excel calendar (Windows 1900 or Mac 1904)
     *
     * @return     integer    Excel base date (1900 or 1904)
     */
    public static function getExcelCalendar()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    Convert a date from Excel to PHP
     *
     *    @param        long        $dateValue            Excel date/time value
     *    @param        boolean        $adjustToTimezone    Flag indicating whether $dateValue should be treated as
     *                                                    a UST timestamp, or adjusted to UST
     *    @param        string         $timezone            The timezone for finding the adjustment from UST
     *    @return        long        PHP serialized date/time
     */
    public static function ExcelToPHP(dateValue = 0, adjustToTimezone = false, timezone = null)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Convert a date from Excel to a PHP Date/Time object
     *
     * @param    integer        $dateValue        Excel date/time value
     * @return    DateTime                    PHP date/time object
     */
    public static function ExcelToPHPObject($dateValue = 0)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    Convert a date from PHP to Excel
     *
     *    @param    mixed        $dateValue            PHP serialized date/time or date object
     *    @param    boolean        $adjustToTimezone    Flag indicating whether $dateValue should be treated as
     *                                                    a UST timestamp, or adjusted to UST
     *    @param    string         $timezone            The timezone for finding the adjustment from UST
     *    @return    mixed        Excel date/time value
     *                            or boolean FALSE on failure
     */
    public static function PHPToExcel(dateValue = 0, adjustToTimezone = false, timezone = null)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * FormattedPHPToExcel
     *
     * @param    long    $year
     * @param    long    $month
     * @param    long    $day
     * @param    long    $hours
     * @param    long    $minutes
     * @param    long    $seconds
     * @return  long                Excel date/time value
     */
    public static function FormattedPHPToExcel(year, month, day, hours = 0, minutes = 0, seconds = 0)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Is a given cell a date/time?
     *
     * @param     \ZExcel\CellCell    $pCell
     * @return     boolean
     */
    public static function isDateTime(<\ZExcel\CellCell> pCell)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Is a given number format a date/time?
     *
     * @param     \ZExcel\CellStyle_NumberFormat    $pFormat
     * @return     boolean
     */
    public static function isDateTimeFormat(<\ZExcel\CellStyle\NumberFormat> pFormat)
    {
        throw new \Exception("Not implemented yet!");
    }



    /**
     * Is a given number format code a date/time?
     *
     * @param     string    $pFormatCode
     * @return     boolean
     */
    public static function isDateTimeFormatCode(pFormatCode = "")
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Convert a date/time string to Excel time
     *
     * @param    string    $dateValue        Examples: '2009-12-31', '2009-12-31 15:59', '2009-12-31 15:59:10'
     * @return    float|FALSE        Excel date/time serial value
     */
    public static function stringToExcel(dateValue = "")
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function monthStringToNumber(month)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function dayStringToNumber(day)
    {
        throw new \Exception("Not implemented yet!");
    }
}
