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
    public static monthNames = [
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

    /*
     * Names of the months of the year, indexed by shortname
     * Planned usage for locale settings
     *
     * @public
     * @var    string[]
     */
    public static numberSuffixes = [
        "st",
        "nd",
        "rd",
        "th"
    ];

    /**
     * Set the Excel calendar (Windows 1900 or Mac 1904)
     *
     * @param     integer    $baseDate            Excel base date (1900 or 1904)
     * @return     boolean                        Success or failure
     */
    public static function setExcelCalendar(baseDate)
    {
        if ((baseDate == self::CALENDAR_WINDOWS_1900) || (baseDate == self::CALENDAR_MAC_1904)) {
            let self::excelBaseDate = baseDate;
            return true;
        }
        return false;
    }


    /**
     * Return the Excel calendar (Windows 1900 or Mac 1904)
     *
     * @return     integer    Excel base date (1900 or 1904)
     */
    public static function getExcelCalendar()
    {
        return self::excelBaseDate;
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
    public static function ExcelToPHP(var dateValue = 0, boolean adjustToTimezone = false, var timezone = null)
    {
        var myexcelBaseDate, utcDays, returnValue, hours, mins, secs, timezoneAdjustment;
        
        if (self::excelBaseDate == self::CALENDAR_WINDOWS_1900) {
            let myexcelBaseDate = 25569;
            //    Adjust for the spurious 29-Feb-1900 (Day 60)
            if (dateValue < 60) {
                let myexcelBaseDate = myexcelBaseDate - 1;
            }
        } else {
            let myexcelBaseDate = 24107;
        }

        // Perform conversion
        if (dateValue >= 1) {
            let utcDays = dateValue - myexcelBaseDate;
            let returnValue = round(utcDays * 86400);
            if ((returnValue <= PHP_INT_MAX) && (returnValue >= -PHP_INT_MAX)) {
                let returnValue = (int) returnValue;
            }
        } else {
            let hours = round(dateValue * 24);
            let mins = round(dateValue * 1440) - round(hours * 60);
            let secs = round(dateValue * 86400) - round(hours * 3600) - round(mins * 60);
            let returnValue = (int) gmmktime(hours, mins, secs);
        }

        let timezoneAdjustment = (adjustToTimezone) ? \ZExcel\Shared\TimeZone::getTimezoneAdjustment(timezone, returnValue) : 0;

        return returnValue + timezoneAdjustment;
    }


    /**
     * Convert a date from Excel to a PHP Date/Time object
     *
     * @param    integer  dateValue Excel date/time value
     * @return   DateTime           PHP date/time object
     */
    public static function ExcelToPHPObject(int dateValue = 0) -> <\DateTime>
    {
        var dateTime, days, time, hours, minutes, seconds, dateObj;
    
        let dateTime = self::ExcelToPHP(dateValue);
        let days = floor(dateTime / 86400);
        let time = round(((dateTime / 86400) - days) * 86400);
        let hours = round(time / 3600);
        let minutes = round(time / 60) - (hours * 60);
        let seconds = round(time) - (hours * 3600) - (minutes * 60);

        let dateObj = date_create("1-Jan-1970+" . days . " days");
        dateObj->setTime(hours,minutes,seconds);

        return dateObj;
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
     * @param  long year
     * @param  long month
     * @param  long day
     * @param  long hours
     * @param  long minutes
     * @param  long seconds
     * @return long Excel date/time value
     */
    public static function FormattedPHPToExcel(var year, var month, var day, var hours = 0, var minutes = 0, var seconds = 0) -> double
    {
        var excel1900isLeapYear, myexcelBaseDate, century, decade, excelDate, excelTime;
    
        if (self::excelBaseDate == self::CALENDAR_WINDOWS_1900) {
            //
            //    Fudge factor for the erroneous fact that the year 1900 is treated as a Leap Year in MS Excel
            //    This affects every date following 28th February 1900
            //
            let excel1900isLeapYear = true;
            if ((year == 1900) && (month <= 2)) {
                let excel1900isLeapYear = false;
            }
            let myexcelBaseDate = 2415020;
        } else {
            let myexcelBaseDate = 2416481;
            let excel1900isLeapYear = false;
        }

        //    Julian base date Adjustment
        if (month > 2) {
            let month = month - 3;
        } else {
            let month = month + 9;
            let year = year - 1;
        }

        //    Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        let century = substr(year, 0, 2);
        let decade = substr(year, 2, 2);
        let excelDate = (float) (floor((146097 * century) / 4) + floor((1461 * decade) / 4) + floor((153 * month + 2) / 5) + day + 1721119 - myexcelBaseDate + excel1900isLeapYear);

        let excelTime = (float) (((hours * 3600) + (minutes * 60) + seconds) / 86400);

        return excelDate + excelTime;
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
    public static function isDateTimeFormatCode(string pFormatCode = "") -> boolean
    {
        var segMatcher, subVal;
        
        if (strtolower(pFormatCode) === strtolower(\ZExcel\Style\NumberFormat::FORMAT_GENERAL)) {
            //    "General" contains an epoch letter 'e', so we trap for it explicitly here (case-insensitive check)
            return false;
        }
        if (preg_match("/[0#]E[+-]0/i", pFormatCode)) {
            //    Scientific format
            return false;
        }

        // Switch on formatcode
        switch (pFormatCode) {
            //    Explicitly defined date formats
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_YYYYMMDD:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_YYYYMMDD2:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_DDMMYYYY:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_DMYSLASH:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_DMYMINUS:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_DMMINUS:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_MYMINUS:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_DATETIME:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME1:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME2:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME3:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME4:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME5:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME6:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME7:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_TIME8:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_YYYYMMDDSLASH:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_XLSX14:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_XLSX15:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_XLSX16:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_XLSX17:
            case \ZExcel\Style\NumberFormat::FORMAT_DATE_XLSX22:
                return true;
        }

        //    Typically number, currency or accounting (or occasionally fraction) formats
        if ((substr(pFormatCode, 0, 1) == "_") || (substr(pFormatCode, 0, 2) == "0 ")) {
            return false;
        }
        
        // Try checking for any of the date formatting characters that don"t appear within square braces
        if (preg_match("/(^|\])[^\[]*[" . self::possibleDateFormatCharacters . "]/i", pFormatCode)) {
            //    We might also have a format mask containing quoted strings...
            //        we don"t want to test for any of our characters within the quoted blocks
            if (strpos(pFormatCode, "\"") !== false) {
                let segMatcher = false;
                for subVal in explode("\"", pFormatCode) {
                    //    Only test in alternate array entries (the non-quoted blocks)
                    let segMatcher = !segMatcher;
                    if (segMatcher && (preg_match("/(^|\])[^\[]*[" . self::possibleDateFormatCharacters . "]/i", subVal))) {
                        return true;
                    }
                }
                return false;
            }
            return true;
        }

        // No date...
        return false;
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

    public static function monthStringToNumber(string month)
    {
        var shortMonthName, longMonthName, monthIndex = 1;
        
        for shortMonthName, longMonthName in self::monthNames {
            if ((month === longMonthName) || (month === shortMonthName)) {
                return monthIndex;
            }
            let monthIndex = monthIndex + 1;
        }
        
        return month;
    }

    public static function dayStringToNumber(string day)
    {
        var strippedDayValue;
        
        let strippedDayValue = (str_replace(self::numberSuffixes, "", day));
        
        if (is_numeric(strippedDayValue)) {
            return strippedDayValue;
        }
        
        return day;
    }
}
