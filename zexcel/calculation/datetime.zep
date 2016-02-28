namespace ZExcel\Calculation;

class DateTime
{
/**
     * Identify if a year is a leap year or not
     *
     * @param    integer    year    The year to test
     * @return    boolean            TRUE if the year is a leap year, otherwise FALSE
     */
    public static function isLeapYear(int year)
    {
        return (((year % 4) == 0) && ((year % 100) != 0) || ((year % 400) == 0));
    }


    /**
     * Return the number of days between two dates based on a 360 day calendar
     *
     * @param    integer    startDay        Day of month of the start date
     * @param    integer    startMonth        Month of the start date
     * @param    integer    startYear        Year of the start date
     * @param    integer    endDay            Day of month of the start date
     * @param    integer    endMonth        Month of the start date
     * @param    integer    endYear        Year of the start date
     * @param    boolean methodUS        Whether to use the US method or the European method of calculation
     * @return    integer    Number of days between the start date and the end date
     */
    private static function dateDiff360(int startDay, int startMonth, int startYear, int endDay, int endMonth, int endYear, boolean methodUS)
    {
        if (startDay == 31) {
            let startDay = startDay - 1;
        } elseif (methodUS && (startMonth == 2 && (startDay == 29 || (startDay == 28 && !self::isLeapYear(startYear))))) {
            let startDay = 30;
        }
        if (endDay == 31) {
            if (methodUS && startDay != 30) {
                let endDay = 1;
                if (endMonth == 12) {
                    let endYear = endYear + 1;
                    let endMonth = 1;
                } else {
                    let endMonth = endMonth + 1;
                }
            } else {
                let endDay = 30;
            }
        }

        return endDay + endMonth * 30 + endYear * 360 - startDay - startMonth * 30 - startYear * 360;
    }


    /**
     * getDateValue
     *
     * @param    string    dateValue
     * @return    mixed    Excel date/time serial value, or string if error
     */
    public static function getDateValue(var dateValue)
    {
        var saveReturnDateType;
        
        if (!is_numeric(dateValue)) {
            if ((is_string(dateValue)) && (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            
            if ((is_object(dateValue)) && (dateValue instanceof \DateTime)) {
                let dateValue = \ZExcel\Shared\Date::PHPToExcel(dateValue);
            } else {
                let saveReturnDateType = \ZExcel\Calculation\Functions::getReturnDateType();
                \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
                
                let dateValue = self::datevalue(dateValue);
                \ZExcel\Calculation\Functions::setReturnDateType(saveReturnDateType);
            }
        }
        
        return dateValue;
    }


    /**
     * getTimeValue
     *
     * @param    string    timeValue
     * @return    mixed    Excel date/time serial value, or string if error
     */
    private static function getTimeValue(string timeValue) -> string
    {
        var saveReturnDateType, timeValue;
        
        let saveReturnDateType = \ZExcel\Calculation\Functions::getReturnDateType();
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        
        let timeValue = self::timevalue(timeValue);
        \ZExcel\Calculation\Functions::setReturnDateType(saveReturnDateType);
        
        return timeValue;
    }


    private static function adjustDateByMonths(int dateValue = 0, int adjustmentMonths = 0) -> <\DateTime>
    {
        var oMonth, oYear, adjustmentMonthsString, monthDiff,
            nMonth, nYear, adjustDays, adjustDaysString, PHPDateObject;
        
        // Execute function
        let PHPDateObject = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);
        let oMonth = (int) PHPDateObject->format("m");
        let oYear = (int) PHPDateObject->format("Y");

        let adjustmentMonthsString = adjustmentMonths;
        if (adjustmentMonths > 0) {
            let adjustmentMonthsString = "+" . adjustmentMonthsString;
        }
        if (adjustmentMonths != 0) {
            PHPDateObject->modify(adjustmentMonthsString." months");
        }
        let nMonth = (int) PHPDateObject->format("m");
        let nYear = (int) PHPDateObject->format("Y");

        let monthDiff = (nMonth - oMonth) + ((nYear - oYear) * 12);
        if (monthDiff != adjustmentMonths) {
            let adjustDays = (int) PHPDateObject->format("d");
            let adjustDaysString = "-".adjustDays." days";
            PHPDateObject->modify(adjustDaysString);
        }
        
        return PHPDateObject;
    }


    /**
     * DATETIMENOW
     *
     * Returns the current date and time.
     * The NOW function is useful when you need to display the current date and time on a worksheet or
     * calculate a value based on the current date and time, and have that value updated each time you
     * open the worksheet.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the date
     * and time format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        NOW()
     *
     * @access    public
     * @category Date/Time Functions
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function datetimenow()
    {
        var saveTimeZone, retValue;
        
        let retValue = false;
        let saveTimeZone = date_default_timezone_get();
        
        date_default_timezone_set("UTC");
        
        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                let retValue = (float) \ZExcel\Shared\Date::PHPToExcel(time());
                break;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                let retValue = (int) time();
                break;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                let retValue = new DateTime();
                break;
        }
        
        date_default_timezone_set(saveTimeZone);

        return retValue;
    }


    /**
     * DATENOW
     *
     * Returns the current date.
     * The NOW function is useful when you need to display the current date and time on a worksheet or
     * calculate a value based on the current date and time, and have that value updated each time you
     * open the worksheet.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the date
     * and time format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        TODAY()
     *
     * @access    public
     * @category Date/Time Functions
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function datenow()
    {
        var saveTimeZone, retValue, excelDateTime;
        
        let retValue = false;
        let saveTimeZone = date_default_timezone_get();
        
        date_default_timezone_set("UTC");
        
        let excelDateTime = floor(\ZExcel\Shared\Date::PHPToExcel(time()));
        
        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                let retValue = (float) excelDateTime;
                break;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                let retValue = (int) \ZExcel\Shared\Date::ExcelToPHP(excelDateTime);
                break;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                let retValue = \ZExcel\Shared\Date::ExcelToPHPObject(excelDateTime);
                break;
        }
        
        date_default_timezone_set(saveTimeZone);

        return retValue;
    }


    /**
     * DATE
     *
     * The DATE function returns a value that represents a particular date.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the date
     * format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        DATE(year,month,day)
     *
     * PHPExcel is a lot more forgiving than MS Excel when passing non numeric values to this function.
     * A Month name or abbreviation (English only at this point) such as "January" or "Jan" will still be accepted,
     *     as will a day value with a suffix (e.g. "21st" rather than simply 21); again only English language.
     *
     * @access    public
     * @category Date/Time Functions
     * @param    integer        year    The value of the year argument can include one to four digits.
     *                                Excel interprets the year argument according to the configured
     *                                date system: 1900 or 1904.
     *                                If year is between 0 (zero) and 1899 (inclusive), Excel adds that
     *                                value to 1900 to calculate the year. For example, DATE(108,1,2)
     *                                returns January 2, 2008 (1900+108).
     *                                If year is between 1900 and 9999 (inclusive), Excel uses that
     *                                value as the year. For example, DATE(2008,1,2) returns January 2,
     *                                2008.
     *                                If year is less than 0 or is 10000 or greater, Excel returns the
     *                                #NUM! error value.
     * @param    integer        month    A positive or negative integer representing the month of the year
     *                                from 1 to 12 (January to December).
     *                                If month is greater than 12, month adds that number of months to
     *                                the first month in the year specified. For example, DATE(2008,14,2)
     *                                returns the serial number representing February 2, 2009.
     *                                If month is less than 1, month subtracts the magnitude of that
     *                                number of months, plus 1, from the first month in the year
     *                                specified. For example, DATE(2008,-3,2) returns the serial number
     *                                representing September 2, 2007.
     * @param    integer        day    A positive or negative integer representing the day of the month
     *                                from 1 to 31.
     *                                If day is greater than the number of days in the month specified,
     *                                day adds that number of days to the first day in the month. For
     *                                example, DATE(2008,1,35) returns the serial number representing
     *                                February 4, 2008.
     *                                If day is less than 1, day subtracts the magnitude that number of
     *                                days, plus one, from the first day of the month specified. For
     *                                example, DATE(2008,1,-15) returns the serial number representing
     *                                December 16, 2007.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function date(var year = 0, var month = 1, var day = 1)
    {
        var baseYear, excelDateValue;
    
        let year  = \ZExcel\Calculation\Functions::flattenSingleValue(year);
        let month = \ZExcel\Calculation\Functions::flattenSingleValue(month);
        let day   = \ZExcel\Calculation\Functions::flattenSingleValue(day);

        if ((month !== null) && (!is_numeric(month))) {
            let month = \ZExcel\Shared\Date::monthStringToNumber(month);
        }

        if ((day !== null) && (!is_numeric(day))) {
            let day = \ZExcel\Shared\Date::dayStringToNumber(day);
        }

        let year = (year !== null) ? \ZExcel\Shared\Stringg::testStringAsNumeric(year) : 0;
        let month = (month !== null) ? \ZExcel\Shared\Stringg::testStringAsNumeric(month) : 0;
        let day = (day !== null) ? \ZExcel\Shared\Stringg::testStringAsNumeric(day) : 0;
        if ((!is_numeric(year)) ||
            (!is_numeric(month)) ||
            (!is_numeric(day))) {
            return \ZExcel\Calculation\Functions::value();
        }
        let year  = (int) year;
        let month = (int) month;
        let day   = (int) day;

        let baseYear = \ZExcel\Shared\Date::getExcelCalendar();
        // Validate parameters
        if (year < (baseYear - 1900)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (((baseYear - 1900) != 0) && (year < baseYear) && (year >= 1900)) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        if ((year < baseYear) && (year >= (baseYear - 1900))) {
            let year += 1900;
        }

        if (month < 1) {
            //    Handle year/month adjustment if month < 1
            let month = month - 1;
            let year = year + ceil(month / 12) - 1;
            let month = 13 - abs(month % 12);
        } elseif (month > 12) {
            //    Handle year/month adjustment if month > 12
            let year = year + floor(month / 12);
            let month = (month % 12);
        }

        // Re-validate the year parameter after adjustments
        if ((year < baseYear) || (year >= 10000)) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Execute function
        let excelDateValue = \ZExcel\Shared\Date::FormattedPHPToExcel(year, month, day);
        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                return (float) excelDateValue;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                return (int) \ZExcel\Shared\Date::ExcelToPHP(excelDateValue);
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                return \ZExcel\Shared\Date::ExcelToPHPObject(excelDateValue);
        }
    }


    /**
     * TIME
     *
     * The TIME function returns a value that represents a particular time.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the time
     * format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        TIME(hour,minute,second)
     *
     * @access    public
     * @category Date/Time Functions
     * @param    integer        hour        A number from 0 (zero) to 32767 representing the hour.
     *                                    Any value greater than 23 will be divided by 24 and the remainder
     *                                    will be treated as the hour value. For example, TIME(27,0,0) =
     *                                    TIME(3,0,0) = .125 or 3:00 AM.
     * @param    integer        minute        A number from 0 to 32767 representing the minute.
     *                                    Any value greater than 59 will be converted to hours and minutes.
     *                                    For example, TIME(0,750,0) = TIME(12,30,0) = .520833 or 12:30 PM.
     * @param    integer        second        A number from 0 to 32767 representing the second.
     *                                    Any value greater than 59 will be converted to hours, minutes,
     *                                    and seconds. For example, TIME(0,0,2000) = TIME(0,33,22) = .023148
     *                                    or 12:33:20 AM
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function time(var hour = 0, var minute = 0, var second = 0)
    {
        var date, dayAdjust, calendar, phpDateObject;
        
        let hour = \ZExcel\Calculation\Functions::flattenSingleValue(hour);
        let minute = \ZExcel\Calculation\Functions::flattenSingleValue(minute);
        let second = \ZExcel\Calculation\Functions::flattenSingleValue(second);

        if (hour == "") {
            let hour = 0;
        }
        if (minute == "") {
            let minute = 0;
        }
        if (second == "") {
            let second = 0;
        }

        if ((!is_numeric(hour)) || (!is_numeric(minute)) || (!is_numeric(second))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let hour = (int) hour;
        let minute = (int) minute;
        let second = (int) second;

        if (second < 0) {
            let minute = minute + floor(second / 60);
            let second = 60 - abs(second % 60);
            
            if (second == 60) {
                let second = 0;
            }
        } elseif (second >= 60) {
            let minute = minute + floor(second / 60);
            let second = second % 60;
        }
        if (minute < 0) {
            let hour = hour + floor(minute / 60);
            let minute = 60 - abs(minute % 60);
            if (minute == 60) {
                let minute = 0;
            }
        } elseif (minute >= 60) {
            let hour = hour + floor(minute / 60);
            let minute = minute % 60;
        }

        if (hour > 23) {
            let hour = hour % 24;
        } elseif (hour < 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Execute function
        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                let date = 0;
                let calendar = \ZExcel\Shared\Date::getExcelCalendar();
                
                if (calendar != \ZExcel\Shared\Date::CALENDAR_WINDOWS_1900) {
                    let date = 1;
                }
                
                return (float) \ZExcel\Shared\Date::FormattedPHPToExcel(calendar, 1, date, hour, minute, second);
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                return (int) \ZExcel\Shared\Date::ExcelToPHP(\ZExcel\Shared\Date::FormattedPHPToExcel(1970, 1, 1, hour, minute, second));    // -2147468400; //    -2147472000 + 3600
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                let dayAdjust = 0;
                
                if (hour < 0) {
                    let dayAdjust = floor(hour / 24);
                    let hour = 24 - abs(hour % 24);
                    if (hour == 24) {
                        let hour = 0;
                    }
                } elseif (hour >= 24) {
                    let dayAdjust = floor(hour / 24);
                    let hour = hour % 24;
                }
                
                let phpDateObject = new \DateTime("1900-01-01 ".hour.":".minute.":".second);
                
                if (dayAdjust != 0) {
                    phpDateObject->modify(dayAdjust." days");
                }
                
                return phpDateObject;
        }
    }


    /**
     * DATEVALUE
     *
     * Returns a value that represents a particular date.
     * Use DATEVALUE to convert a date represented by a text string to an Excel or PHP date/time stamp
     * value.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the date
     * format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        DATEVALUE(dateValue)
     *
     * @access    public
     * @category Date/Time Functions
     * @param    string    dateValue        Text that represents a date in a Microsoft Excel date format.
     *                                    For example, "1/30/2008" or "30-Jan-2008" are text strings within
     *                                    quotation marks that represent dates. Using the default date
     *                                    system in Excel for Windows, date_text must represent a date from
     *                                    January 1, 1900, to December 31, 9999. Using the default date
     *                                    system in Excel for the Macintosh, date_text must represent a date
     *                                    from January 1, 1904, to December 31, 9999. DATEVALUE returns the
     *                                    #VALUE! error value if date_text is out of this range.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function datevalue(var dateValue = 1)
    {
        var t1, t, k, PHPDateArray, testVal1, testVal2, testVal3, excelDateValue;
        boolean yearFound = false;
    
        let dateValue = trim(\ZExcel\Calculation\Functions::flattenSingleValue(dateValue), "\"");
        //    Strip any ordinals because they're allowed in Excel (English only)
        let dateValue = preg_replace("/(\d)(st|nd|rd|th)([ -\/])/Ui", "$1$3", dateValue);
        //    Convert separators (/ . or space) to hyphens (should also handle dot used for ordinals in some countries, e.g. Denmark, Germany)
        let dateValue = str_replace(["/", ".", "-", "  "], [" ", " ", " ", " "], dateValue);

        let t1 = explode(" ", dateValue);
        
        for k, t in t1 {
            if ((is_numeric(t)) && (t > 31)) {
                if (yearFound) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                } else {
                    if (t < 100) {
                        let t = t + 1900;
                        let t1[k] = t;
                    }
                    let yearFound = true;
                }
            }
        }
        
        if ((count(t1) == 1) && (strpos(t, ":") != false)) {
            //    We've been fed a time value without any date
            return 0.0;
        } elseif (count(t1) == 2) {
            //    We only have two parts of the date: either day/month or month/year
            if (yearFound) {
                array_unshift(t1, 1);
            } else {
                array_push(t1, date("Y"));
            }
        }

        let dateValue = implode(" ", t1);
        
        let PHPDateArray = date_parse(dateValue);
        
        if ((PHPDateArray === false) || (PHPDateArray["error_count"] > 0)) {
            let testVal1 = call_user_func("strtok" , dateValue, "- ");
            
            if (testVal1 !== false) {
                let testVal2 = call_user_func("strtok", "- ");
                
                if (testVal2 !== false) {
                    let testVal3 = call_user_func("strtok", "- ");
                    
                    if (testVal3 === false) {
                        let testVal3 = strftime("%Y");
                    }
                } else {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            
            let PHPDateArray = date_parse(testVal1 . "-" . testVal2 . "-" . testVal3);
            
            if ((PHPDateArray === false) || (PHPDateArray["error_count"] > 0)) {
                let PHPDateArray = date_parse(testVal2 . "-" . testVal1 . "-" . testVal3);
                
                if ((PHPDateArray === false) || (PHPDateArray["error_count"] > 0)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            }
        }

        if ((PHPDateArray !== false) && (PHPDateArray["error_count"] == 0)) {
            // Execute function
            if (PHPDateArray["year"] == "") {
                let PHPDateArray["year"] = strftime("%Y");
            }
            if (PHPDateArray["year"] < 1900) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            if (PHPDateArray["month"] == "") {
                let PHPDateArray["month"] = strftime("%m");
            }
            if (PHPDateArray["day"] == "") {
                let PHPDateArray["day"] = strftime("%d");
            }
            
            let excelDateValue = floor(
                \ZExcel\Shared\Date::FormattedPHPToExcel(
                    PHPDateArray["year"],
                    PHPDateArray["month"],
                    PHPDateArray["day"],
                    PHPDateArray["hour"],
                    PHPDateArray["minute"],
                    PHPDateArray["second"]
                )
            );

            switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
                case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                    return (float) excelDateValue;
                case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                    return (int) \ZExcel\Shared\Date::ExcelToPHP(excelDateValue);
                case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                    return new \DateTime(PHPDateArray["year"]."-".PHPDateArray["month"]."-".PHPDateArray["day"]." 00:00:00");
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * TIMEVALUE
     *
     * Returns a value that represents a particular time.
     * Use TIMEVALUE to convert a time represented by a text string to an Excel or PHP date/time stamp
     * value.
     *
     * NOTE: When used in a Cell Formula, MS Excel changes the cell format so that it matches the time
     * format of your regional settings. PHPExcel does not change cell formatting in this way.
     *
     * Excel Function:
     *        timevalue(timeValue)
     *
     * @access    public
     * @category Date/Time Functions
     * @param    string    timeValue        A text string that represents a time in any one of the Microsoft
     *                                    Excel time formats; for example, "6:45 PM" and "18:45" text strings
     *                                    within quotation marks that represent time.
     *                                    Date information in time_text is ignored.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function timevalue(var timeValue)
    {
        var PHPDateArray, excelDateValue;
        
        let timeValue = trim(\ZExcel\Calculation\Functions::flattenSingleValue(timeValue), "\"");
        let timeValue = str_replace(["/", "."], ["-", "-"], timeValue);

        let PHPDateArray = date_parse(timeValue);
        
        if ((PHPDateArray !== false) && (PHPDateArray["error_count"] == 0)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let excelDateValue = \ZExcel\Shared\Date::FormattedPHPToExcel(
                    PHPDateArray["year"],
                    PHPDateArray["month"],
                    PHPDateArray["day"],
                    PHPDateArray["hour"],
                    PHPDateArray["minute"],
                    PHPDateArray["second"]
                );
            } else {
                let excelDateValue = (double) \ZExcel\Shared\Date::FormattedPHPToExcel(1900, 1, 1, PHPDateArray["hour"], PHPDateArray["minute"], PHPDateArray["second"]) - 1.0;
            }
            
            switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
                case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                    return (float) excelDateValue;
                case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                    return (int) (\ZExcel\Shared\Date::ExcelToPHP(excelDateValue + 25569.0) - 3600);
                case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                    return new \DateTime("1900-01-01 " . PHPDateArray["hour"] . ":" . PHPDateArray["minute"] . ":" . PHPDateArray["second"]);
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * DATEDIF
     *
     * @param    mixed    startDate        Excel date serial value, PHP date/time stamp, PHP DateTime object
     *                                    or a standard date string
     * @param    mixed    endDate        Excel date serial value, PHP date/time stamp, PHP DateTime object
     *                                    or a standard date string
     * @param    string    unit
     * @return    integer    Interval between the dates
     */
    public static function datedif(var startDate = 0, var endDate = 0, var unit = "D")
    {
        var difference, PHPStartDateObject, PHPEndDateObject, startDays, startMonths, startYears, endDays, endMonths, endYears, adjustDays, retVal;
        
        let startDate = \ZExcel\Calculation\Functions::flattenSingleValue(startDate);
        let endDate   = \ZExcel\Calculation\Functions::flattenSingleValue(endDate);
        let unit      = strtoupper(\ZExcel\Calculation\Functions::flattenSingleValue(unit));

        let startDate = self::getDateValue(startDate);
        let endDate = self::getDateValue(endDate);

        if (is_string(startDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if (is_string(endDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Validate parameters
        if (startDate >= endDate) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Execute function
        let difference = endDate - startDate;

        let PHPStartDateObject = \ZExcel\Shared\Date::ExcelToPHPObject(startDate);
        let startDays = PHPStartDateObject->format("j");
        let startMonths = PHPStartDateObject->format("n");
        let startYears = PHPStartDateObject->format("Y");

        let PHPEndDateObject = \ZExcel\Shared\Date::ExcelToPHPObject(endDate);
        let endDays = PHPEndDateObject->format("j");
        let endMonths = PHPEndDateObject->format("n");
        let endYears = PHPEndDateObject->format("Y");

        let retVal = \ZExcel\Calculation\Functions::NaN();
        switch (unit) {
            case "D":
                let retVal = intval(difference);
                break;
            case "M":
                let retVal = intval(endMonths - startMonths) + (intval(endYears - startYears) * 12);
                //    We"re only interested in full months
                if (endDays < startDays) {
                    let retVal = retVal - 1;
                }
                break;
            case "Y":
                let retVal = intval(endYears - startYears);
                //    We"re only interested in full months
                if (endMonths < startMonths) {
                    let retVal = retVal - 1;
                } elseif ((endMonths == startMonths) && (endDays < startDays)) {
                    let retVal = retVal - 1;
                }
                break;
            case "MD":
                if (endDays < startDays) {
                    let retVal = endDays;
                    PHPEndDateObject->modify("-".endDays." days");
                    let adjustDays = PHPEndDateObject->format("j");
                    if (adjustDays > startDays) {
                        let retVal = retVal + (adjustDays - startDays);
                    }
                } else {
                    let retVal = endDays - startDays;
                }
                break;
            case "YM":
                let retVal = intval(endMonths - startMonths);
                if (retVal < 0) {
                    let retVal = retVal + 12;
                }
                //    We"re only interested in full months
                if (endDays < startDays) {
                    let retVal = retVal - 1;
                }
                break;
            case "YD":
                let retVal = intval(difference);
                if (endYears > startYears) {
                    while (endYears > startYears) {
                        PHPEndDateObject->modify("-1 year");
                        let endYears = PHPEndDateObject->format("Y");
                    }
                    let retVal = PHPEndDateObject->format("z") - PHPStartDateObject->format("z");
                    if (retVal < 0) {
                        let retVal = retVal + 365;
                    }
                }
                break;
            default:
                let retVal = \ZExcel\Calculation\Functions::NaN();
        }
        return retVal;
    }


    /**
     * DAYS360
     *
     * Returns the number of days between two dates based on a 360-day year (twelve 30-day months),
     * which is used in some accounting calculations. Use this function to help compute payments if
     * your accounting system is based on twelve 30-day months.
     *
     * Excel Function:
     *        DAYS360(startDate,endDate[,method])
     *
     * @access    public
     * @category Date/Time Functions
     * @param    mixed        startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    mixed        endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    boolean        method            US or European Method
     *                                        FALSE or omitted: U.S. (NASD) method. If the starting date is
     *                                        the last day of a month, it becomes equal to the 30th of the
     *                                        same month. If the ending date is the last day of a month and
     *                                        the starting date is earlier than the 30th of a month, the
     *                                        ending date becomes equal to the 1st of the next month;
     *                                        otherwise the ending date becomes equal to the 30th of the
     *                                        same month.
     *                                        TRUE: European method. Starting dates and ending dates that
     *                                        occur on the 31st of a month become equal to the 30th of the
     *                                        same month.
     * @return    integer        Number of days between start date and end date
     */
    public static function days360(var startDate = 0, var endDate = 0, var method = false)
    {
        var PHPStartDateObject, startDay, startMonth, startYear, PHPEndDateObject, endDay, endMonth, endYear;
        
        let startDate  = \ZExcel\Calculation\Functions::flattenSingleValue(startDate);
        let endDate    = \ZExcel\Calculation\Functions::flattenSingleValue(endDate);

        let startDate = self::getDateValue(startDate);
        let endDate = self::getDateValue(endDate);

        if (is_string(startDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        if (is_string(endDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (!is_bool(method)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Execute function
        let PHPStartDateObject = \ZExcel\Shared\Date::ExcelToPHPObject(startDate);
        let startDay = PHPStartDateObject->format("j");
        let startMonth = PHPStartDateObject->format("n");
        let startYear = PHPStartDateObject->format("Y");

        let PHPEndDateObject = \ZExcel\Shared\Date::ExcelToPHPObject(endDate);
        let endDay = PHPEndDateObject->format("j");
        let endMonth = PHPEndDateObject->format("n");
        let endYear = PHPEndDateObject->format("Y");

        let method = !method;
        
        return self::dateDiff360(startDay, startMonth, startYear, endDay, endMonth, endYear, method);
    }


    /**
     * YEARFRAC
     *
     * Calculates the fraction of the year represented by the number of whole days between two dates
     * (the start_date and the end_date).
     * Use the YEARFRAC worksheet function to identify the proportion of a whole year"s benefits or
     * obligations to assign to a specific term.
     *
     * Excel Function:
     *        YEARFRAC(startDate,endDate[,method])
     *
     * @access    public
     * @category Date/Time Functions
     * @param    mixed    startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    mixed    endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    integer    method            Method used for the calculation
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float    fraction of the year
     */
    public static function yearfrac(var startDate = 0, var endDate = 0, var method = 0)
    {
        var startDay, startMonth, year, endDay, endMonth;
        int days, startYear, endYear, years;
        double leapDays;
        
        let startDate = \ZExcel\Calculation\Functions::flattenSingleValue(startDate);
        let endDate   = \ZExcel\Calculation\Functions::flattenSingleValue(endDate);
        let method    = \ZExcel\Calculation\Functions::flattenSingleValue(method);

        let startDate = self::getDateValue(startDate);
        let endDate   = self::getDateValue(endDate);
        
        if (is_string(startDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if (is_string(endDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (((is_numeric(method)) && (!is_string(method))) || (method == "")) {
            switch(method) {
                case 0:
                    return self::DaYS360(startDate, endDate) / 360;
                case 1:
                    let days = 0 + self::DaTEDIF(startDate, endDate);
                    let startYear = 0 + self::YeAR(startDate);
                    let endYear = 0 + self::YeAR(endDate);
                    let years = endYear - startYear + 1;
                    let leapDays = 0;
                    
                    if (years == 1) {
                        if (self::isLeapYear(endYear)) {
                            let startMonth = self::MoNTHOFYEAR(startDate);
                            let endMonth = self::MoNTHOFYEAR(endDate);
                            let endDay = self::DaYOFMONTH(endDate);
                            
                            if ((startMonth < 3) || ((endMonth * 100 + endDay) >= (2 * 100 + 29))) {
                                let leapDays = leapDays + 1;
                            }
                        }
                    } else {
                        for year in range(startYear, endYear) {
                            if (year == startYear) {
                                let startMonth = self::MoNTHOFYEAR(startDate);
                                let startDay = self::DaYOFMONTH(startDate);
                                
                                if (startMonth < 3 && self::isLeapYear(year)) {
                                    let leapDays = leapDays + 1;
                                }
                            } elseif (year == endYear) {
                                let endMonth = self::MoNTHOFYEAR(endDate);
                                let endDay = self::DaYOFMONTH(endDate);
                                
                                if ((endMonth * 100 + endDay) >= 229 && self::isLeapYear(year)) {
                                    let leapDays = leapDays + 1;
                                }
                            } elseif (self::isLeapYear(year)) {
                                let leapDays = leapDays + 1;
                            }
                        }
                        
                        if (years == 2) {
                            if ((leapDays == 0) && (self::isLeapYear(startYear)) && (days > 365)) {
                                let leapDays = 1;
                            } elseif (days < 366) {
                                let years = 1;
                            }
                        }

                        let leapDays = leapDays / years;
                    }
                    
                    return days / (365 + leapDays);
                case 2:
                    return self::DaTEDIF(startDate, endDate) / 360;
                case 3:
                    return self::DaTEDIF(startDate, endDate) / 365;
                case 4:
                    return self::DaYS360(startDate, endDate, true) / 360;
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * NETWORKDAYS
     *
     * Returns the number of whole working days between start_date and end_date. Working days
     * exclude weekends and any dates identified in holidays.
     * Use NETWORKDAYS to calculate employee benefits that accrue based on the number of days
     * worked during a specific term.
     *
     * Excel Function:
     *        NETWORKDAYS(startDate,endDate[,holidays[,holiday[,...]]])
     *
     * @access    public
     * @category Date/Time Functions
     * @param    mixed            startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                            PHP DateTime object, or a standard date string
     * @param    mixed            endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                            PHP DateTime object, or a standard date string
     * @param    mixed            holidays,...    Optional series of Excel date serial value (float), PHP date
     *                                            timestamp (int), PHP DateTime object, or a standard date
     *                                            strings that will be excluded from the working calendar, such
     *                                            as state and federal holidays and floating holidays.
     * @return    integer            Interval between the dates
     */
    public static function networkdays()
    {
        var startDate, endDate, dateArgs, sDate, eDate, holidayDate;
        double startDoW, endDoW, wholeWeekDays, partWeekDays;
        array holidayCountedArray = [];
        
        //    Flush the mandatory start and end date that are referenced in the function definition, and get the optional days
        let dateArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        //    Retrieve the mandatory start and end date that are referenced in the function definition
        let startDate = \ZExcel\Calculation\Functions::flattenSingleValue(array_shift(dateArgs));
        let endDate   = \ZExcel\Calculation\Functions::flattenSingleValue(array_shift(dateArgs));
        
        let startDate = self::getDateValue(startDate);
        let sDate = startDate;

        //    Validate the start and end dates
        if (is_string(startDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let startDate = (float) floor(startDate);
        let endDate = self::getDateValue(endDate);
        let eDate = endDate;
        
        if (is_string(endDate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let endDate = (float) floor(endDate);

        if (sDate > eDate) {
            let startDate = eDate;
            let endDate = sDate;
        }

        // Execute function
        let startDoW = 6 - self::DaYOFWEEK(startDate, 2);
        if (startDoW < 0) {
            let startDoW = 0;
        }
        
        let endDoW = 0 + self::DaYOFWEEK(endDate, 2);
        if (endDoW >= 6) {
            let endDoW = 0;
        }

        let wholeWeekDays = 5 * floor((endDate - startDate) / 7);
        let partWeekDays = endDoW + startDoW;
        if (partWeekDays > 5) {
            let partWeekDays = partWeekDays - 5;
        }

        //    Test any extra holiday parameters
        for holidayDate in dateArgs {
            let holidayDate = self::getDateValue(holidayDate);
        
            if (is_string(holidayDate)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            
            if ((holidayDate >= startDate) && (holidayDate <= endDate)) {
                if ((self::DaYOFWEEK(holidayDate, 2) < 6) && (!in_array(holidayDate, holidayCountedArray))) {
                    let partWeekDays = partWeekDays - 1;
                    let holidayCountedArray[] = holidayDate;
                }
            }
        }

        if (sDate > eDate) {
            return 0 - (wholeWeekDays + partWeekDays);
        }
        
        return wholeWeekDays + partWeekDays;
    }


    /**
     * WORKDAY
     *
     * Returns the date that is the indicated number of working days before or after a date (the
     * starting date). Working days exclude weekends and any dates identified as holidays.
     * Use WORKDAY to exclude weekends or holidays when you calculate invoice due dates, expected
     * delivery times, or the number of days of work performed.
     *
     * Excel Function:
     *        WORKDAY(startDate,endDays[,holidays[,holiday[,...]]])
     *
     * @access    public
     * @category Date/Time Functions
     * @param    mixed        startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    integer        endDays        The number of nonweekend and nonholiday days before or after
     *                                        startDate. A positive value for days yields a future date; a
     *                                        negative value yields a past date.
     * @param    mixed        holidays,...    Optional series of Excel date serial value (float), PHP date
     *                                        timestamp (int), PHP DateTime object, or a standard date
     *                                        strings that will be excluded from the working calendar, such
     *                                        as state and federal holidays and floating holidays.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function workday()
    {
        var startDate, endDays, dateArgs, holidayDate;
        array holidayCountedArray = [], holidayDates = [];
        double startDoW, endDoW, endDate;
        boolean decrementing;
        
        //    Flush the mandatory start date and days that are referenced in the function definition, and get the optional days
        let dateArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        //    Retrieve the mandatory start date and days that are referenced in the function definition
        let startDate = array_shift(dateArgs);
        let endDays = array_shift(dateArgs);
        
        let startDate = \ZExcel\Calculation\Functions::flattenSingleValue(startDate);
        let endDays   = \ZExcel\Calculation\Functions::flattenSingleValue(endDays);

        let startDate = self::getDateValue(startDate);

        if ((is_string(startDate)) || (!is_numeric(endDays))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let startDate = (float) floor(startDate);
        let endDays = (int) floor(endDays);
        
        //    If endDays is 0, we always return startDate
        if (endDays == 0) {
            return startDate;
        }

        let decrementing = false;
        if (endDays < 0) {
            let decrementing = true;
        }

        //    Adjust the start date if it falls over a weekend

        let startDoW = 0 + self::DaYOFWEEK(startDate, 3);
        
        if (startDoW >= 5) {
            let startDate = decrementing ? (startDate + 4 - startDoW) : (startDate + 7 - startDoW);
            let endDays = decrementing ? (endDays + 1) : (endDays - 1);
        }

        //    Add endDays
        let endDate = (float) startDate + (intval(endDays / 5) * 7) + (endDays % 5);

        //    Adjust the calculated end date if it falls over a weekend
        let endDoW = 0 + self::DaYOFWEEK(endDate, 3);
        if (endDoW >= 5) {
            let endDate = decrementing ? (endDate + 4 - endDoW) : (endDate + 7 - endDoW);
        }

        //    Test any extra holiday parameters
        if (!empty(dateArgs)) {
            for holidayDate in dateArgs {
                if ((holidayDate !== null) && (strlen(trim(holidayDate)) > 0)) {
                    let holidayDate = self::getDateValue(holidayDate);
                    
                    if (is_string(holidayDate)) {
                        return \ZExcel\Calculation\Functions::VaLUE();
                    }
                    
                    if (self::DaYOFWEEK(holidayDate, 3) < 5) {
                        let holidayDates[] = holidayDate;
                    }
                }
            }
            
            if (decrementing) {
                rsort(holidayDates, SORT_NUMERIC);
            } else {
                sort(holidayDates, SORT_NUMERIC);
            }
            
            for holidayDate in holidayDates {
                if (decrementing) {
                    if ((holidayDate <= startDate) && (holidayDate >= endDate)) {
                        if (!in_array(holidayDate, holidayCountedArray)) {
                            let endDate = endDate - 1;
                            let holidayCountedArray[] = holidayDate;
                        }
                    }
                } else {
                    if ((holidayDate >= startDate) && (holidayDate <= endDate)) {
                        if (!in_array(holidayDate, holidayCountedArray)) {
                            let endDate = endDate + 1;
                            let holidayCountedArray[] = holidayDate;
                        }
                    }
                }
                //    Adjust the calculated end date if it falls over a weekend
                let endDoW = 0 + self::DaYOFWEEK(endDate, 3);
                if (endDoW >= 5) {
                    if (decrementing) {
                        let endDate = endDate + (-1 * endDoW) + 4;
                    } else {
                        let endDate = endDate + 7 - endDoW;
                    }
                }
            }
        }

        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                return (float) endDate;
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                return (int) \ZExcel\Shared\Date::ExcelToPHP(endDate);
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                return \ZExcel\Shared\Date::ExcelToPHPObject(endDate);
        }
    }


    /**
     * DAYOFMONTH
     *
     * Returns the day of the month, for a specified date. The day is given as an integer
     * ranging from 1 to 31.
     *
     * Excel Function:
     *        DAY(dateValue)
     *
     * @param    mixed    dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Day of the month
     */
    public static function dayofmonth(var dateValue = 1)
    {
        let dateValue = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);

        if (dateValue === null) {
            let dateValue = 1;
        } else {
        
            let dateValue = self::getDateValue(dateValue);
            
            if (is_string(dateValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (dateValue == 0.0) {
                return 0;
            } elseif (dateValue < 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }

        // Execute function
        let dateValue = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);

        return (int) dateValue->format("j");
    }


    /**
     * DAYOFWEEK
     *
     * Returns the day of the week for a specified date. The day is given as an integer
     * ranging from 0 to 7 (dependent on the requested style).
     *
     * Excel Function:
     *        WEEKDAY(dateValue[,style])
     *
     * @param    mixed    dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    int        style            A number that determines the type of return value
     *                                        1 or omitted    Numbers 1 (Sunday) through 7 (Saturday).
     *                                        2                Numbers 1 (Monday) through 7 (Sunday).
     *                                        3                Numbers 0 (Monday) through 6 (Sunday).
     * @return    int        Day of the week value
     */
    public static function dayofweek(var dateValue = 1, var style = 1)
    {
        var DoW, firstDay;
        
        let dateValue = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);
        let style     = \ZExcel\Calculation\Functions::flattenSingleValue(style);

        if (!is_numeric(style)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        } elseif ((style < 1) || (style > 3)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let style = floor(style);

        if (dateValue === null) {
            let dateValue = 1;
        } else {
            let dateValue = self::getDateValue(dateValue);
    
            if (is_string(dateValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (dateValue < 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        
        // Execute function
        let dateValue = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);
        let DoW = dateValue->format("w");

        let firstDay = 1;
        switch (style) {
            case 1:
                let DoW = DoW + 1;
                break;
            case 2:
                if (DoW == 0) {
                    let DoW = 7;
                }
                break;
            case 3:
                if (DoW == 0) {
                    let DoW = 7;
                }
                let firstDay = 0;
                let DoW = DoW - 1;
                break;
        }
        if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_EXCEL) {
            //    Test for Excel's 1900 leap year, and introduce the error as required
            if ((dateValue->format("Y") == 1900) && (dateValue->format("n") <= 2)) {
                let DoW = DoW - 1;
                if (DoW < firstDay) {
                    let DoW = DoW + 7;
                }
            }
        }

        return (int) DoW;
    }


    /**
     * WEEKOFYEAR
     *
     * Returns the week of the year for a specified date.
     * The WEEKNUM function considers the week containing January 1 to be the first week of the year.
     * However, there is a European standard that defines the first week as the one with the majority
     * of days (four or more) falling in the new year. This means that for years in which there are
     * three days or less in the first week of January, the WEEKNUM function returns week numbers
     * that are incorrect according to the European standard.
     *
     * Excel Function:
     *        WEEKNUM(dateValue[,style])
     *
     * @param    mixed    dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    boolean    method            Week begins on Sunday or Monday
     *                                        1 or omitted    Week begins on Sunday.
     *                                        2                Week begins on Monday.
     * @return    int        Week Number
     */
    public static function weekofyear(var dateValue = 1, var method = 1)
    {
        var dayOfYear, dow;
        int daysInFirstWeek, weekOfYear;
        
        let dateValue = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);
        let method    = \ZExcel\Calculation\Functions::flattenSingleValue(method);

        if (!is_numeric(method)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        } elseif ((method < 1) || (method > 2)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let method = floor(method);

        if (dateValue === null) {
            let dateValue = 1;
        } else {
            let dateValue = self::getDateValue(dateValue);
    
            if (is_string(dateValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (dateValue < 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        
        // Execute function
        let dateValue = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);
        let dayOfYear = dateValue->format("z");
        let dow = dateValue->format("w");
        
        dateValue->modify("-" . dayOfYear . " days");
        
        let dow = dateValue->format("w");
        let daysInFirstWeek = 7 - ((dow + (2 - method)) % 7);
        let dayOfYear = dayOfYear - daysInFirstWeek;
        let weekOfYear = ceil(dayOfYear / 7) + 1;

        return (int) weekOfYear;
    }


    /**
     * MONTHOFYEAR
     *
     * Returns the month of a date represented by a serial number.
     * The month is given as an integer, ranging from 1 (January) to 12 (December).
     *
     * Excel Function:
     *        MONTH(dateValue)
     *
     * @param    mixed    dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Month of the year
     */
    public static function monthofyear(var dateValue = 1)
    {
        let dateValue = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);

        if (dateValue === null) {
            let dateValue = 1;
        } else {
            let dateValue = self::getDateValue(dateValue);
            
            if (is_string(dateValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (dateValue < 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        
        // Execute function
        let dateValue = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);

        return (int) dateValue->format("n");
    }


    /**
     * YEAR
     *
     * Returns the year corresponding to a date.
     * The year is returned as an integer in the range 1900-9999.
     *
     * Excel Function:
     *        YEAR(dateValue)
     *
     * @param    mixed    dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Year
     */
    public static function year(var dateValue = 1)
    {
        let dateValue = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);

        if (dateValue === null) {
            let dateValue = 1;
        } else {
            let dateValue = self::getDateValue(dateValue);
            
            if (is_string(dateValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (dateValue < 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }

        // Execute function
        let dateValue = \ZExcel\Shared\Date::ExcelToPHPObject(dateValue);

        return (int) dateValue->format("Y");
    }


    /**
     * HOUROFDAY
     *
     * Returns the hour of a time value.
     * The hour is given as an integer, ranging from 0 (12:00 A.M.) to 23 (11:00 P.M.).
     *
     * Excel Function:
     *        HOUR(timeValue)
     *
     * @param    mixed    timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Hour
     */
    public static function hourofday(var timeValue = 0)
    {
        var testVal;
        
        let timeValue = \ZExcel\Calculation\Functions::flattenSingleValue(timeValue);

        if (!is_numeric(timeValue)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                let testVal = call_user_func("strtok", timeValue, "/-: ");
                if (strlen(testVal) < strlen(timeValue)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            }
            
            let timeValue = self::getTimeValue(timeValue);
            
            if (is_string(timeValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        // Execute function
        if (timeValue >= 1) {
            let timeValue = fmod(timeValue, 1);
        } elseif (timeValue < 0.0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let timeValue = \ZExcel\Shared\Date::ExcelToPHP(timeValue);

        return (int) gmdate("G", timeValue);
    }


    /**
     * MINUTEOFHOUR
     *
     * Returns the minutes of a time value.
     * The minute is given as an integer, ranging from 0 to 59.
     *
     * Excel Function:
     *        MINUTE(timeValue)
     *
     * @param    mixed    timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Minute
     */
    public static function minuteofhour(var timeValue = 0)
    {
        var testVal;
        
        let timeValue = \ZExcel\Calculation\Functions::flattenSingleValue(timeValue);
        
        if (!is_numeric(timeValue)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                let testVal = call_user_func("strtok", timeValue, "/-: ");
                if (strlen(testVal) < strlen(timeValue)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            }
            
            let timeValue = self::getTimeValue(timeValue);
            
            if (is_string(timeValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        // Execute function
        if (timeValue >= 1) {
            let timeValue = fmod(timeValue, 1);
        } elseif (timeValue < 0.0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let timeValue = \ZExcel\Shared\Date::ExcelToPHP(timeValue);

        return (int) gmdate("i", timeValue);
    }


    /**
     * SECONDOFMINUTE
     *
     * Returns the seconds of a time value.
     * The second is given as an integer in the range 0 (zero) to 59.
     *
     * Excel Function:
     *        SECOND(timeValue)
     *
     * @param    mixed    timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Second
     */
    public static function secondofminute(var timeValue = 0)
    {
        var testVal;
        
        let timeValue = \ZExcel\Calculation\Functions::flattenSingleValue(timeValue);

        if (!is_numeric(timeValue)) {
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                let testVal = call_user_func("strtok", timeValue, "/-: ");
                if (strlen(testVal) < strlen(timeValue)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            }
            
            let timeValue = self::getTimeValue(timeValue);
            
            if (is_string(timeValue)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        // Execute function
        if (timeValue >= 1) {
            let timeValue = fmod(timeValue, 1);
        } elseif (timeValue < 0.0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let timeValue = \ZExcel\Shared\Date::ExcelToPHP(timeValue);

        return (int) gmdate("s", timeValue);
    }


    /**
     * EDATE
     *
     * Returns the serial number that represents the date that is the indicated number of months
     * before or after a specified date (the start_date).
     * Use EDATE to calculate maturity dates or due dates that fall on the same day of the month
     * as the date of issue.
     *
     * Excel Function:
     *        EDATE(dateValue,adjustmentMonths)
     *
     * @param    mixed    dateValue            Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    int        adjustmentMonths    The number of months before or after start_date.
     *                                        A positive value for months yields a future date;
     *                                        a negative value yields a past date.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function edate(var dateValue = 1, var adjustmentMonths = 0)
    {
        let dateValue        = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);
        let adjustmentMonths = \ZExcel\Calculation\Functions::flattenSingleValue(adjustmentMonths);

        if (!is_numeric(adjustmentMonths)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let adjustmentMonths = floor(adjustmentMonths);
        let dateValue = self::getDateValue(dateValue);

        if (is_string(dateValue)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Execute function
        let dateValue = self::adjustDateByMonths(dateValue, adjustmentMonths);

        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                return (float) \ZExcel\Shared\Date::PHPToExcel(dateValue);
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                return (int) \ZExcel\Shared\Date::ExcelToPHP(\ZExcel\Shared\Date::PHPToExcel(dateValue));
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                return dateValue;
        }
    }


    /**
     * EOMONTH
     *
     * Returns the date value for the last day of the month that is the indicated number of months
     * before or after start_date.
     * Use EOMONTH to calculate maturity dates or due dates that fall on the last day of the month.
     *
     * Excel Function:
     *        EOMONTH(dateValue,adjustmentMonths)
     *
     * @param    mixed    dateValue            Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    int        adjustmentMonths    The number of months before or after start_date.
     *                                        A positive value for months yields a future date;
     *                                        a negative value yields a past date.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function eomonth(var dateValue = 1, var adjustmentMonths = 0)
    {
        var adjustDays, adjustDaysString;
        
        let dateValue        = \ZExcel\Calculation\Functions::flattenSingleValue(dateValue);
        let adjustmentMonths = \ZExcel\Calculation\Functions::flattenSingleValue(adjustmentMonths);

        if (!is_numeric(adjustmentMonths)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let adjustmentMonths = floor(adjustmentMonths);
        let dateValue = self::getDateValue(dateValue);
        
        if (is_string(dateValue)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Execute function
        let dateValue = self::adjustDateByMonths(dateValue, adjustmentMonths + 1);
        let adjustDays = (int) dateValue->format("d");
        let adjustDaysString = "-" . adjustDays . " days";
        
        dateValue->modify(adjustDaysString);

        switch (\ZExcel\Calculation\Functions::getReturnDateType()) {
            case \ZExcel\Calculation\Functions::RETURNDATE_EXCEL:
                return (float) \ZExcel\Shared\Date::PHPToExcel(dateValue);
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC:
                return (int) \ZExcel\Shared\Date::ExcelToPHP(\ZExcel\Shared\Date::PHPToExcel(dateValue));
            case \ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT:
                return dateValue;
        }
    }
}
