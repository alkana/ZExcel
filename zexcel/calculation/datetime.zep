namespace ZExcel\Calculation;

class DateTime
{
/**
     * Identify if a year is a leap year or not
     *
     * @param    integer    $year    The year to test
     * @return    boolean            TRUE if the year is a leap year, otherwise FALSE
     */
    public static function isLeapYear(int year)
    {
        return (((year % 4) == 0) && ((year % 100) != 0) || ((year % 400) == 0));
    }


    /**
     * Return the number of days between two dates based on a 360 day calendar
     *
     * @param    integer    $startDay        Day of month of the start date
     * @param    integer    $startMonth        Month of the start date
     * @param    integer    $startYear        Year of the start date
     * @param    integer    $endDay            Day of month of the start date
     * @param    integer    $endMonth        Month of the start date
     * @param    integer    $endYear        Year of the start date
     * @param    boolean $methodUS        Whether to use the US method or the European method of calculation
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
     * @param    string    $dateValue
     * @return    mixed    Excel date/time serial value, or string if error
     */
    public static function getDateValue(string dateValue)
    {
        var saveReturnDateType;
        
        if (!is_numeric(dateValue)) {
            if ((is_string(dateValue)) &&
                (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            if ((is_object(dateValue)) && (typeof dateValue == "DateTime")) {
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
     * @param    string    $timeValue
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


    private static function adjustDateByMonths(int dateValue = 0, int adjustmentMonths = 0) -> <\ZExcel\Shared\Date>
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
        throw new \Exception("Not implemented yet!");
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
        throw new \Exception("Not implemented yet!");
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
     * @param    integer        $year    The value of the year argument can include one to four digits.
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
     * @param    integer        $month    A positive or negative integer representing the month of the year
     *                                from 1 to 12 (January to December).
     *                                If month is greater than 12, month adds that number of months to
     *                                the first month in the year specified. For example, DATE(2008,14,2)
     *                                returns the serial number representing February 2, 2009.
     *                                If month is less than 1, month subtracts the magnitude of that
     *                                number of months, plus 1, from the first month in the year
     *                                specified. For example, DATE(2008,-3,2) returns the serial number
     *                                representing September 2, 2007.
     * @param    integer        $day    A positive or negative integer representing the day of the month
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
            let year = year + floor($month / 12);
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
     * @param    integer        $hour        A number from 0 (zero) to 32767 representing the hour.
     *                                    Any value greater than 23 will be divided by 24 and the remainder
     *                                    will be treated as the hour value. For example, TIME(27,0,0) =
     *                                    TIME(3,0,0) = .125 or 3:00 AM.
     * @param    integer        $minute        A number from 0 to 32767 representing the minute.
     *                                    Any value greater than 59 will be converted to hours and minutes.
     *                                    For example, TIME(0,750,0) = TIME(12,30,0) = .520833 or 12:30 PM.
     * @param    integer        $second        A number from 0 to 32767 representing the second.
     *                                    Any value greater than 59 will be converted to hours, minutes,
     *                                    and seconds. For example, TIME(0,0,2000) = TIME(0,33,22) = .023148
     *                                    or 12:33:20 AM
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function time($hour = 0, $minute = 0, $second = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    string    $dateValue        Text that represents a date in a Microsoft Excel date format.
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
        var yearFound, t1, t, k, PHPDateArray, testVal, excelDateValue;
    
        let dateValue = trim(\ZExcel\Calculation\Functions::flattenSingleValue(dateValue), "\"");
        //    Strip any ordinals because they're allowed in Excel (English only)
        let dateValue = preg_replace("/(\d)(st|nd|rd|th)([ -\/])/Ui", "$1$3", dateValue);
        //    Convert separators (/ . or space) to hyphens (should also handle dot used for ordinals in some countries, e.g. Denmark, Germany)
        let dateValue = str_replace(["/", ".", "-", "  "], [" ", " ", " ", " "], dateValue);

        let yearFound = false;
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
            let testVal = explode(" ", dateValue);
            
            if (count(testVal) < 2) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (!isset(testVal[2])) {
                let testVal[2] = strftime("%Y");
            }
            
            let PHPDateArray = date_parse(implode("-", testVal));
            
            if ((PHPDateArray === false) || (PHPDateArray["error_count"] > 0)) {
                
                let PHPDateArray = date_parse(testVal[1] . "-" . testVal[0] . "-" . testVal[2]);
                
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
            
            // @FIXME
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
     * @param    string    $timeValue        A text string that represents a time in any one of the Microsoft
     *                                    Excel time formats; for example, "6:45 PM" and "18:45" text strings
     *                                    within quotation marks that represent time.
     *                                    Date information in time_text is ignored.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function timevalue($timeValue)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * DATEDIF
     *
     * @param    mixed    $startDate        Excel date serial value, PHP date/time stamp, PHP DateTime object
     *                                    or a standard date string
     * @param    mixed    $endDate        Excel date serial value, PHP date/time stamp, PHP DateTime object
     *                                    or a standard date string
     * @param    string    $unit
     * @return    integer    Interval between the dates
     */
    public static function datedif($startDate = 0, $endDate = 0, $unit = "D")
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed        $startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    mixed        $endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    boolean        $method            US or European Method
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
    public static function days360($startDate = 0, $endDate = 0, $method = false)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    mixed    $endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    integer    $method            Method used for the calculation
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float    fraction of the year
     */
    public static function yearfrac($startDate = 0, $endDate = 0, $method = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed            $startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                            PHP DateTime object, or a standard date string
     * @param    mixed            $endDate        Excel date serial value (float), PHP date timestamp (int),
     *                                            PHP DateTime object, or a standard date string
     * @param    mixed            $holidays,...    Optional series of Excel date serial value (float), PHP date
     *                                            timestamp (int), PHP DateTime object, or a standard date
     *                                            strings that will be excluded from the working calendar, such
     *                                            as state and federal holidays and floating holidays.
     * @return    integer            Interval between the dates
     */
    public static function networkdays($startDate, $endDate)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed        $startDate        Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    integer        $endDays        The number of nonweekend and nonholiday days before or after
     *                                        startDate. A positive value for days yields a future date; a
     *                                        negative value yields a past date.
     * @param    mixed        $holidays,...    Optional series of Excel date serial value (float), PHP date
     *                                        timestamp (int), PHP DateTime object, or a standard date
     *                                        strings that will be excluded from the working calendar, such
     *                                        as state and federal holidays and floating holidays.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function workday($startDate, $endDays)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Day of the month
     */
    public static function dayofmonth($dateValue = 1)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    int        $style            A number that determines the type of return value
     *                                        1 or omitted    Numbers 1 (Sunday) through 7 (Saturday).
     *                                        2                Numbers 1 (Monday) through 7 (Sunday).
     *                                        3                Numbers 0 (Monday) through 6 (Sunday).
     * @return    int        Day of the week value
     */
    public static function dayofweek($dateValue = 1, $style = 1)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @param    boolean    $method            Week begins on Sunday or Monday
     *                                        1 or omitted    Week begins on Sunday.
     *                                        2                Week begins on Monday.
     * @return    int        Week Number
     */
    public static function weekofyear($dateValue = 1, $method = 1)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Month of the year
     */
    public static function monthofyear($dateValue = 1)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard date string
     * @return    int        Year
     */
    public static function year($dateValue = 1)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Hour
     */
    public static function hourofday($timeValue = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Minute
     */
    public static function minuteofhour($timeValue = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $timeValue        Excel date serial value (float), PHP date timestamp (int),
     *                                    PHP DateTime object, or a standard time string
     * @return    int        Second
     */
    public static function secondofminute($timeValue = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue            Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    int        $adjustmentMonths    The number of months before or after start_date.
     *                                        A positive value for months yields a future date;
     *                                        a negative value yields a past date.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function edate($dateValue = 1, $adjustmentMonths = 0)
    {
        throw new \Exception("Not implemented yet!");
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
     * @param    mixed    $dateValue            Excel date serial value (float), PHP date timestamp (int),
     *                                        PHP DateTime object, or a standard date string
     * @param    int        $adjustmentMonths    The number of months before or after start_date.
     *                                        A positive value for months yields a future date;
     *                                        a negative value yields a past date.
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function eomonth($dateValue = 1, $adjustmentMonths = 0)
    {
        throw new \Exception("Not implemented yet!");
    }
}
