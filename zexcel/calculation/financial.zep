namespace ZExcel\Calculation;

class Financial
{
    /**
      * FINANCIAL_MAX_ITERATIONS
      */
    const FINANCIAL_MAX_ITERATIONS  = 128;
    
    /**
     * FINANCIAL_PRECISION
     */
    const FINANCIAL_PRECISION = 0.00000001;
    
    /**
     * isLastDayOfMonth
     *
     * Returns a boolean TRUE/FALSE indicating if this date is the last date of the month
     *
     * @param    DateTime    testDate    The date for testing
     * @return    boolean
     */
    private static function isLastDayOfMonth(<\DateTime> testDate) -> boolean
    {
        return (testDate->format("d") == testDate->format("t"));
    }

    /**
     * isFirstDayOfMonth
     *
     * Returns a boolean TRUE/FALSE indicating if this date is the first date of the month
     *
     * @param    DateTime    testDate    The date for testing
     * @return    boolean
     */
    private static function isFirstDayOfMonth(<\DateTime> testDate) -> boolean
    {
        return (testDate->format("d") == 1);
    }

    private static function couponFirstPeriodDate(int settlement, var maturity, var frequency, boolean next)
    {
        var result, eom, months;
        
        let months = strval(12 / frequency);

        let result = \ZExcel\Shared\Date::ExcelToPHPObject(maturity);
        let eom = self::isLastDayOfMonth(result);

        while (settlement < \ZExcel\Shared\Date::PHPToExcel(result)) {
            result->modify("-" . months . " months");
        }
        if (next) {
            result->modify("+" . months . " months");
        }

        if (eom) {
            result->modify("-1 day");
        }

        return \ZExcel\Shared\Date::PHPToExcel(result);
    }

    private static function isValidFrequency(int frequency) -> boolean
    {
        if ((frequency == 1) || (frequency == 2) || (frequency == 4)) {
            return true;
        }
        
        if ((\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) && ((frequency == 6) || (frequency == 12))) {
            return true;
        }
        
        return false;
    }

    /**
     * daysPerYear
     *
     * Returns the number of days in a specified year, as defined by the "basis" value
     *
     * @param    integer        year    The year against which we"re testing
     * @param   integer        basis    The type of day count:
     *                                    0 or omitted US (NASD)    360
     *                                    1                        Actual (365 or 366 in a leap year)
     *                                    2                        360
     *                                    3                        365
     *                                    4                        European 360
     * @return    integer
     */
    private static function daysPerYear(int year, int basis = 0)
    {
        int daysPerYear;
        
        switch (basis) {
            case 0:
            case 2:
            case 4:
                let daysPerYear = 360;
                break;
            case 3:
                let daysPerYear = 365;
                break;
            case 1:
                if (\ZExcel\Calculation\DateTime::isLeapYear(year)) {
                    let daysPerYear = 366;
                } else {
                    let daysPerYear = 365;
                }
                break;
            default:
                return \ZExcel\Calculation\Functions::NaN();
        }
        
        return daysPerYear;
    }

    private static function interestAndPrincipal(var rate = 0, var per = 0, var nper = 0, var pv = 0, var fv = 0, var type = 0) -> array
    {
        var pmt, capital, interest, principal;
        int i;
        
        let pmt = self::pmt(rate, nper, pv, fv, type);
        
        let capital = pv;
        
        for i in range(1, per) {
            let interest = (type && i == 1) ? 0 : -capital * rate;
            let principal = pmt - interest;
            let capital = capital + principal;
        }
        
        return [interest, principal];
    }


    /**
     * ACCRINT
     *
     * Returns the accrued interest for a security that pays periodic interest.
     *
     * Excel Function:
     *        ACCRINT(issue,firstinterest,settlement,rate,par,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    issue            The security"s issue date.
     * @param    mixed    firstinterest    The security"s first interest date.
     * @param    mixed    settlement        The security"s settlement date.
     *                                    The security settlement date is the date after the issue date
     *                                    when the security is traded to the buyer.
     * @param    float    rate            The security"s annual coupon rate.
     * @param    float    par            The security"s par value.
     *                                    If you omit par, ACCRINT uses 1,000.
     * @param    integer    frequency        the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer    basis            The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function accrint(var issue, var firstinterest, var settlement, var rate, var par = 1000, var frequency = 1, var basis = 0)
    {
        var daysBetweenIssueAndSettlement;
        
        let issue         = \ZExcel\Calculation\Functions::flattenSingleValue(issue);
        let firstinterest = \ZExcel\Calculation\Functions::flattenSingleValue(firstinterest);
        let settlement    = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let rate          = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let par           = (is_null(par))       ? 1000 :  \ZExcel\Calculation\Functions::flattenSingleValue(par);
        let frequency     = (is_null(frequency)) ? 1    :  \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis         = (is_null(basis))     ? 0    :  \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if ((is_numeric(rate)) && (is_numeric(par))) {
            let rate = (float) rate;
            let par  = (float) par;
            
            if ((rate <= 0) || (par <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenIssueAndSettlement = \ZExcel\Calculation\DateTime::YeARFRAC(issue, settlement, basis);
            
            if (!is_numeric(daysBetweenIssueAndSettlement)) {
                //    return date error
                return daysBetweenIssueAndSettlement;
            }

            return par * rate * daysBetweenIssueAndSettlement;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * ACCRINTM
     *
     * Returns the accrued interest for a security that pays interest at maturity.
     *
     * Excel Function:
     *        ACCRINTM(issue,settlement,rate[,par[,basis]])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    issue        The security"s issue date.
     * @param    mixed    settlement    The security"s settlement (or maturity) date.
     * @param    float    rate        The security"s annual coupon rate.
     * @param    float    par            The security"s par value.
     *                                    If you omit par, ACCRINT uses 1,000.
     * @param    integer    basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function accrintm(var issue, var settlement, var rate, var par = 1000, var basis = 0)
    {
        var daysBetweenIssueAndSettlement;
        
        let issue         = \ZExcel\Calculation\Functions::flattenSingleValue(issue);
        let settlement    = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let rate          = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let par           = (is_null(par))       ? 1000 :  \ZExcel\Calculation\Functions::flattenSingleValue(par);
        let basis         = (is_null(basis))     ? 0    :  \ZExcel\Calculation\Functions::flattenSingleValue(basis);
        
        if ((is_numeric(rate)) && (is_numeric(par))) {
            let rate = (float) rate;
            let par  = (float) par;
            
            if ((rate <= 0) || (par <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenIssueAndSettlement = \ZExcel\Calculation\DateTime::YeARFRAC(issue, settlement, basis);
            
            if (!is_numeric(daysBetweenIssueAndSettlement)) {
                //    return date error
                return daysBetweenIssueAndSettlement;
            }
            
            return par * rate * daysBetweenIssueAndSettlement;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * AMORDEGRC
     *
     * Returns the depreciation for each accounting period.
     * This function is provided for the French accounting system. If an asset is purchased in
     * the middle of the accounting period, the prorated depreciation is taken into account.
     * The function is similar to AMORLINC, except that a depreciation coefficient is applied in
     * the calculation depending on the life of the assets.
     * This function will return the depreciation until the last period of the life of the assets
     * or until the cumulated value of depreciation is greater than the cost of the assets minus
     * the salvage value.
     *
     * Excel Function:
     *        AMORDEGRC(cost,purchased,firstPeriod,salvage,period,rate[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    cost        The cost of the asset.
     * @param    mixed    purchased    Date of the purchase of the asset.
     * @param    mixed    firstPeriod    Date of the end of the first period.
     * @param    mixed    salvage        The salvage value at the end of the life of the asset.
     * @param    float    period        The period.
     * @param    float    rate        Rate of depreciation.
     * @param    integer    basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function amorDegRC(double cost, var purchased, var firstPeriod, var salvage, var period, var rate, var basis = 0)
    {
        double fUsePer, amortiseCoeff, fNRate, fRest;
        int n;
        
        let cost        = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let purchased   = \ZExcel\Calculation\Functions::flattenSingleValue(purchased);
        let firstPeriod = \ZExcel\Calculation\Functions::flattenSingleValue(firstPeriod);
        let salvage     = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let period      = floor(\ZExcel\Calculation\Functions::flattenSingleValue(period));
        let rate        = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let basis       = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    The depreciation coefficients are:
        //    Life of assets (1/rate)        Depreciation coefficient
        //    Less than 3 years            1
        //    Between 3 and 4 years        1.5
        //    Between 5 and 6 years        2
        //    More than 6 years            2.5
        let fUsePer = 1.0 / rate;
        
        if (fUsePer < 3.0) {
            let amortiseCoeff = 1.0;
        } else {
            if (fUsePer < 5.0) {
                let amortiseCoeff = 1.5;
            } else {
                if (fUsePer <= 6.0) {
                    let amortiseCoeff = 2.0;
                } else {
                    let amortiseCoeff = 2.5;
                }
            }
        }

        let rate = (double) rate * amortiseCoeff;
        let fNRate = 0.0 + round(\ZExcel\Calculation\DateTime::YeARFRAC(purchased, firstPeriod, basis) * (double) rate * cost, 0);
        let cost = cost - fNRate;
        let fRest = cost - salvage;

        for n in range(0, period - 1) {
            let fNRate = 0.0 + round((double) rate * cost, 0);
            let fRest  = fRest - fNRate;

            if (fRest < 0.0) {
                switch (period - n) {
                    case 0:
                    case 1:
                        return round(cost * 0.5, 0);
                    default:
                        return 0.0;
                }
            }
            
            let cost = cost - fNRate;
        }
        
        return fNRate;
    }


    /**
     * AMORLINC
     *
     * Returns the depreciation for each accounting period.
     * This function is provided for the French accounting system. If an asset is purchased in
     * the middle of the accounting period, the prorated depreciation is taken into account.
     *
     * Excel Function:
     *        AMORLINC(cost,purchased,firstPeriod,salvage,period,rate[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    cost        The cost of the asset.
     * @param    mixed    purchased    Date of the purchase of the asset.
     * @param    mixed    firstPeriod    Date of the end of the first period.
     * @param    mixed    salvage        The salvage value at the end of the life of the asset.
     * @param    float    period        The period.
     * @param    float    rate        Rate of depreciation.
     * @param    integer    basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function amorlinc(cost, purchased, firstPeriod, salvage, period, rate, basis = 0)
    {
        var purchasedYear, yearFrac;
        double fOneRate, fCostDelta, f0Rate;
        int nNumOfFullPeriods;
        
        let cost        = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let purchased   = \ZExcel\Calculation\Functions::flattenSingleValue(purchased);
        let firstPeriod = \ZExcel\Calculation\Functions::flattenSingleValue(firstPeriod);
        let salvage     = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let period      = floor(\ZExcel\Calculation\Functions::flattenSingleValue(period));
        let rate        = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let basis       = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);
        
        let fOneRate = cost * rate;
        let fCostDelta = cost - salvage;
        //    Note, quirky variation for leap years on the YEARFRAC for this function
        let purchasedYear = \ZExcel\Calculation\DateTime::YeAR(purchased);
        let yearFrac = \ZExcel\Calculation\DateTime::YeARFRAC(purchased, firstPeriod, basis);

        if ((basis == 1) && (yearFrac < 1) && (\ZExcel\Calculation\DateTime::isLeapYear(purchasedYear))) {
            let yearFrac = yearFrac * (365 / 366);
        }

        let f0Rate = yearFrac * rate * cost;
        let nNumOfFullPeriods = intval((cost - (double) salvage - f0Rate) / fOneRate);

        if (period == 0) {
            return f0Rate;
        } else {
            if (period <= nNumOfFullPeriods) {
                return fOneRate;
            } else {
                if (period == (nNumOfFullPeriods + 1)) {
                    return (fCostDelta - fOneRate * nNumOfFullPeriods - f0Rate);
                } else {
                    return 0.0;
                }
            }
        }
    }


    /**
     * COUPDAYBS
     *
     * Returns the number of days from the beginning of the coupon period to the settlement date.
     *
     * Excel Function:
     *        COUPDAYBS(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function coupDayBS(var settlement, var maturity, var frequency, var basis = 0)
    {
        var daysPerYear, prev;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(settlement), basis);
        let prev = self::couponFirstPeriodDate(settlement, maturity, frequency, false);

        return \ZExcel\Calculation\DateTime::YeARFRAC(prev, settlement, basis) * daysPerYear;
    }


    /**
     * COUPDAYS
     *
     * Returns the number of days in the coupon period that contains the settlement date.
     *
     * Excel Function:
     *        COUPDAYS(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function coupDays(var settlement, var maturity, var frequency, var basis = 0)
    {
        var daysPerYear, prev, next;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        switch (basis) {
            case 3:
                // Actual/365
                return 365 / frequency;
            case 1:
                // Actual/actual
                if (frequency == 1) {
                    let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(maturity), basis);
                    return (daysPerYear / frequency);
                }
                
                let prev = self::couponFirstPeriodDate(settlement, maturity, frequency, false);
                let next = self::couponFirstPeriodDate(settlement, maturity, frequency, true);
                
                return (next - prev);
            default:
                // US (NASD) 30/360, Actual/360 or European 30/360
                return 360 / frequency;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * COUPDAYSNC
     *
     * Returns the number of days from the settlement date to the next coupon date.
     *
     * Excel Function:
     *        COUPDAYSNC(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function coupDaySnc(var settlement, var maturity, var frequency, var basis = 0)
    {
        var daysPerYear, next;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(settlement), basis);
        let next = self::couponFirstPeriodDate(settlement, maturity, frequency, true);

        return \ZExcel\Calculation\DateTime::YeARFRAC(settlement, next, basis) * daysPerYear;
    }


    /**
     * COUPNCD
     *
     * Returns the next coupon date after the settlement date.
     *
     * Excel Function:
     *        COUPNCD(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function coupncd(var settlement, var maturity, var frequency, var basis = 0)
    {
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return self::couponFirstPeriodDate(settlement, maturity, frequency, true);
    }


    /**
     * COUPNUM
     *
     * Returns the number of coupons payable between the settlement date and maturity date,
     * rounded up to the nearest whole coupon.
     *
     * Excel Function:
     *        COUPNUM(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    integer
     */
    public static function coupnum(var settlement, var maturity, var frequency, var basis = 0)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let settlement = self::couponFirstPeriodDate(settlement, maturity, frequency, true);
        let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis) * 365;

        switch (frequency) {
            case 1: // annual payments
                return ceil(daysBetweenSettlementAndMaturity / 360);
            case 2: // half-yearly
                return ceil(daysBetweenSettlementAndMaturity / 180);
            case 4: // quarterly
                return ceil(daysBetweenSettlementAndMaturity / 90);
            case 6: // bimonthly
                return ceil(daysBetweenSettlementAndMaturity / 60);
            case 12: // monthly
                return ceil(daysBetweenSettlementAndMaturity / 30);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * COUPPCD
     *
     * Returns the previous coupon date before the settlement date.
     *
     * Excel Function:
     *        COUPPCD(settlement,maturity,frequency[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    frequency    the number of coupon payments per year.
     *                                    Valid frequency values are:
     *                                        1    Annual
     *                                        2    Semi-Annual
     *                                        4    Quarterly
     *                                    If working in Gnumeric Mode, the following frequency options are
     *                                    also available
     *                                        6    Bimonthly
     *                                        12    Monthly
     * @param    integer        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    mixed    Excel date/time serial value, PHP date/time serial value or PHP date/time object,
     *                        depending on the value of the ReturnDateType flag
     */
    public static function couppcd(var settlement, var maturity, var frequency, var basis = 0)
    {
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((settlement > maturity) || (!self::isValidFrequency(frequency)) || ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return self::couponFirstPeriodDate(settlement, maturity, frequency, false);
    }


    /**
     * CUMIPMT
     *
     * Returns the cumulative interest paid on a loan between the start and end periods.
     *
     * Excel Function:
     *        CUMIPMT(rate,nper,pv,start,end[,type])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    rate    The Interest rate
     * @param    integer    nper    The total number of payment periods
     * @param    float    pv        Present Value
     * @param    integer    start    The first period in the calculation.
     *                            Payment periods are numbered beginning with 1.
     * @param    integer    end    The last period in the calculation.
     * @param    integer    type    A number 0 or 1 and indicates when payments are due:
     *                                0 or omitted    At the end of the period.
     *                                1                At the beginning of the period.
     * @return    float
     */
    public static function cumipmt(double rate, int nper, double pv, int start, int end, int type = 0)
    {
        double interest;
        int per;
        
        let rate  = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let nper  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pv    = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let start = (int) \ZExcel\Calculation\Functions::flattenSingleValue(start);
        let end   = (int) \ZExcel\Calculation\Functions::flattenSingleValue(end);
        let type  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (start < 1 || start > end) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Calculate
        let interest = 0;
        for per in range(start, end) {
            let interest = interest + self::iPMT(rate, per, nper, pv, 0, type);
        }

        return interest;
    }


    /**
     * CUMPRINC
     *
     * Returns the cumulative principal paid on a loan between the start and end periods.
     *
     * Excel Function:
     *        CUMPRINC(rate,nper,pv,start,end[,type])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    rate    The Interest rate
     * @param    integer    nper    The total number of payment periods
     * @param    float    pv        Present Value
     * @param    integer    start    The first period in the calculation.
     *                            Payment periods are numbered beginning with 1.
     * @param    integer    end    The last period in the calculation.
     * @param    integer    type    A number 0 or 1 and indicates when payments are due:
     *                                0 or omitted    At the end of the period.
     *                                1                At the beginning of the period.
     * @return    float
     */
    public static function cumprinc(rate, nper, pv, start, end, type = 0)
    {
        double principal;
        int per;
        
        let rate  = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let nper  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pv    = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let start = (int) \ZExcel\Calculation\Functions::flattenSingleValue(start);
        let end   = (int) \ZExcel\Calculation\Functions::flattenSingleValue(end);
        let type  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (start < 1 || start > end) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        // Calculate
        let principal = 0;
        for per in range(start, end) {
            let principal = principal + self::ppmt(rate, per, nper, pv, 0, type);
        }

        return principal;
    }


    /**
     * DB
     *
     * Returns the depreciation of an asset for a specified period using the
     * fixed-declining balance method.
     * This form of depreciation is used if you want to get a higher depreciation value
     * at the beginning of the depreciation (as opposed to linear depreciation). The
     * depreciation value is reduced with every depreciation period by the depreciation
     * already deducted from the initial cost.
     *
     * Excel Function:
     *        DB(cost,salvage,life,period[,month])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    cost        Initial cost of the asset.
     * @param    float    salvage        Value at the end of the depreciation.
     *                                (Sometimes called the salvage value of the asset)
     * @param    integer    life        Number of periods over which the asset is depreciated.
     *                                (Sometimes called the useful life of the asset)
     * @param    integer    period        The period for which you want to calculate the
     *                                depreciation. Period must use the same units as life.
     * @param    integer    month        Number of months in the first year. If month is omitted,
     *                                it defaults to 12.
     * @return    float
     */
    public static function db(var cost, var salvage, var life, var period, var month = 12)
    {
        var fixedDepreciationRate, previousDepreciation, depreciation;
        int per;
        
        let cost    = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let salvage = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let life    = \ZExcel\Calculation\Functions::flattenSingleValue(life);
        let period  = \ZExcel\Calculation\Functions::flattenSingleValue(period);
        let month   = \ZExcel\Calculation\Functions::flattenSingleValue(month);
        
        //    Validate
        if ((is_numeric(cost)) && (is_numeric(salvage)) && (is_numeric(life)) && (is_numeric(period)) && (is_numeric(month))) {
            let life    = (int) life;
            let period  = (int) period;
            let cost    = (float) cost;
            let month   = (float) month;
            let salvage = (float) salvage;
            
            if (cost == 0.0) {
                return cost;
            }
            
            if ((cost < 0) || ((salvage / (double) cost) < 0) || (life <= 0) || (period < 1) || (month < 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            //    Set Fixed Depreciation Rate
            let fixedDepreciationRate = 1 - pow((salvage / (double) cost), (1 / life));
            let fixedDepreciationRate = round(fixedDepreciationRate, 3);

            //    Loop through each period calculating the depreciation
            let previousDepreciation = 0;
            
            for per in range(1, period) {
                if (per == 1) {
                    let depreciation = cost * (double)fixedDepreciationRate * month / 12;
                } else {
                    if (per == (life + 1)) {
                        let depreciation = (cost - previousDepreciation) * fixedDepreciationRate * (12 - month) / 12;
                    } else {
                        let depreciation = (cost - previousDepreciation) * fixedDepreciationRate;
                    }
                }
                
                let previousDepreciation = previousDepreciation + depreciation;
            }
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                let depreciation = round(depreciation, 2);
            }
            
            return depreciation;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * DDB
     *
     * Returns the depreciation of an asset for a specified period using the
     * double-declining balance method or some other method you specify.
     *
     * Excel Function:
     *        DDB(cost,salvage,life,period[,factor])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    cost        Initial cost of the asset.
     * @param    float    salvage        Value at the end of the depreciation.
     *                                (Sometimes called the salvage value of the asset)
     * @param    integer    life        Number of periods over which the asset is depreciated.
     *                                (Sometimes called the useful life of the asset)
     * @param    integer    period        The period for which you want to calculate the
     *                                depreciation. Period must use the same units as life.
     * @param    float    factor        The rate at which the balance declines.
     *                                If factor is omitted, it is assumed to be 2 (the
     *                                double-declining balance method).
     * @return    float
     */
    public static function ddb(var cost, double salvage, int life, int period, double factor = 2.0)
    {
        double fixedDepreciationRate, previousDepreciation, depreciation;
        int per;
        
        let cost    = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let salvage = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let life    = \ZExcel\Calculation\Functions::flattenSingleValue(life);
        let period  = \ZExcel\Calculation\Functions::flattenSingleValue(period);
        let factor  = \ZExcel\Calculation\Functions::flattenSingleValue(factor);
        
        if ((is_numeric(cost)) && (is_numeric(salvage)) && (is_numeric(life)) && (is_numeric(period)) && (is_numeric(factor))) {
            let cost    = (float) cost;
            let salvage = (float) salvage;
            let life    = (int) life;
            let period  = (int) period;
            let factor  = (float) factor;
            
            if ((cost <= 0) || ((salvage / cost) < 0) || (life <= 0) || (period < 1) || (factor <= 0.0) || (period > life)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            //    Set Fixed Depreciation Rate
            let fixedDepreciationRate = 1 - pow((salvage / cost), (1 / life));
            let fixedDepreciationRate = round(fixedDepreciationRate, 3);

            //    Loop through each period calculating the depreciation
            let previousDepreciation = 0;
            for per in range(1, period) {
                let depreciation =  0.0 + min(((double) cost - previousDepreciation) * (factor / life), ((double) cost - salvage - previousDepreciation));
                let previousDepreciation = previousDepreciation + depreciation;
            }
            
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                let depreciation = round(depreciation, 2);
            }
            
            return depreciation;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * DISC
     *
     * Returns the discount rate for a security.
     *
     * Excel Function:
     *        DISC(settlement,maturity,price,redemption[,basis])
     *
     * @access    public
     * @category Financial Functions
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue
     *                                date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    integer    price        The security"s price per 100 face value.
     * @param    integer    redemption    The security"s redemption value per 100 face value.
     * @param    integer    basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function disc(var settlement, var maturity, var price, var redemption, var basis = 0)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let price      = \ZExcel\Calculation\Functions::flattenSingleValue(price);
        let redemption = \ZExcel\Calculation\Functions::flattenSingleValue(redemption);
        let basis      = \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if ((is_numeric(price)) && (is_numeric(redemption)) && (is_numeric(basis))) {
            let price      = (float) price;
            let redemption = (float) redemption;
            let basis      = (int) basis;
            
            if ((price <= 0) || (redemption <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }

            return ((1 - (double) price / (double) redemption) / daysBetweenSettlementAndMaturity);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * DOLLARDE
     *
     * Converts a dollar price expressed as an integer part and a fraction
     *        part into a dollar price expressed as a decimal number.
     * Fractional dollar numbers are sometimes used for security prices.
     *
     * Excel Function:
     *        DOLLARDE(fractional_dollar,fraction)
     *
     * @access    public
     * @category Financial Functions
     * @param    float    fractional_dollar    Fractional Dollar
     * @param    integer    fraction            Fraction
     * @return    float
     */
    public static function dollarDE(double fractional_dollar = null, int fraction = 0) -> double
    {
        double dollars, cents;
        
        let fractional_dollar = \ZExcel\Calculation\Functions::flattenSingleValue(fractional_dollar);
        let fraction          = (int)\ZExcel\Calculation\Functions::flattenSingleValue(fraction);

        // Validate parameters
        if (is_null(fractional_dollar) || fraction < 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (fraction == 0) {
            return \ZExcel\Calculation\Functions::DiV0();
        }

        let dollars = 0.0 + floor(fractional_dollar);
        let cents = 0.0 + fmod(fractional_dollar, 1);
        let cents = cents / fraction;
        let cents = cents * pow(10, ceil(log10(fraction)));
        
        return dollars + cents;
    }


    /**
     * DOLLARFR
     *
     * Converts a dollar price expressed as a decimal number into a dollar price
     *        expressed as a fraction.
     * Fractional dollar numbers are sometimes used for security prices.
     *
     * Excel Function:
     *        DOLLARFR(decimal_dollar,fraction)
     *
     * @access    public
     * @category Financial Functions
     * @param    float    decimal_dollar        Decimal Dollar
     * @param    integer    fraction            Fraction
     * @return    float
     */
    public static function dollarFR(double decimal_dollar = null, int fraction = 0) -> double
    {
        double dollars, cents;
        
        let decimal_dollar = \ZExcel\Calculation\Functions::flattenSingleValue(decimal_dollar);
        let fraction          = (int)\ZExcel\Calculation\Functions::flattenSingleValue(fraction);

        // Validate parameters
        if (is_null(decimal_dollar) || fraction < 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (fraction == 0) {
            return \ZExcel\Calculation\Functions::DiV0();
        }
        
        let dollars = floor(decimal_dollar);
        let cents = 0.0 + fmod(decimal_dollar, 1);
        let cents = cents * fraction;
        let cents = cents * pow(10, -ceil(log10(fraction)));
        
        return dollars + cents;
    }


    /**
     * EFFECT
     *
     * Returns the effective interest rate given the nominal rate and the number of
     *        compounding payments per year.
     *
     * Excel Function:
     *        EFFECT(nominal_rate,npery)
     *
     * @access    public
     * @category Financial Functions
     * @param    float    nominal_rate        Nominal interest rate
     * @param    integer    npery                Number of compounding payments per year
     * @return    float
     */
    public static function effect(double nominal_rate = 0, int npery = 0) -> double
    {
        let nominal_rate = \ZExcel\Calculation\Functions::flattenSingleValue(nominal_rate);
        let npery        = (int) \ZExcel\Calculation\Functions::flattenSingleValue(npery);

        // Validate parameters
        if (nominal_rate <= 0 || npery < 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return (double) pow((1 + nominal_rate / npery), npery) - 1;
    }


    /**
     * FV
     *
     * Returns the Future Value of a cash flow with constant payments and interest rate (annuities).
     *
     * Excel Function:
     *        FV(rate,nper,pmt[,pv[,type]])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    rate    The interest rate per period
     * @param    int        nper    Total number of payment periods in an annuity
     * @param    float    pmt    The payment made each period: it cannot change over the
     *                            life of the annuity. Typically, pmt contains principal
     *                            and interest but no other fees or taxes.
     * @param    float    pv        Present Value, or the lump-sum amount that a series of
     *                            future payments is worth right now.
     * @param    integer    type    A number 0 or 1 and indicates when payments are due:
     *                                0 or omitted    At the end of the period.
     *                                1                At the beginning of the period.
     * @return    float
     */
    public static function fv(double rate = 0, int nper = 0, double pmt = 0, double pv = 0, int type = 0)
    {
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let nper = \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pmt  = \ZExcel\Calculation\Functions::flattenSingleValue(pmt);
        let pv   = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let type = \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Calculate
        if (!is_null(rate) && rate != 0) {
            return -pv * pow(1 + rate, nper) - pmt * (1 + rate * type) * (pow(1 + rate, nper) - 1) / rate;
        }
        
        return -pv - pmt * nper;
    }


    /**
     * FVSCHEDULE
     *
     * Returns the future value of an initial principal after applying a series of compound interest rates.
     * Use FVSCHEDULE to calculate the future value of an investment with a variable or adjustable rate.
     *
     * Excel Function:
     *        FVSCHEDULE(principal,schedule)
     *
     * @param    float    principal    The present value.
     * @param    float[]    schedule    An array of interest rates to apply.
     * @return    float
     */
    public static function fvSchedule(float principal, array schedule) -> double
    {
        var rate;
        
        let principal = \ZExcel\Calculation\Functions::flattenSingleValue(principal);
        let schedule  = \ZExcel\Calculation\Functions::flattenArray(schedule);

        for rate in schedule {
            let principal = principal * (1 + rate);
        }

        return principal;
    }


    /**
     * INTRATE
     *
     * Returns the interest rate for a fully invested security.
     *
     * Excel Function:
     *        INTRATE(settlement,maturity,investment,redemption[,basis])
     *
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    integer    investment    The amount invested in the security.
     * @param    integer    redemption    The amount to be received at maturity.
     * @param    integer    basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function intRate(var settlement, var maturity, var investment, var redemption, var basis = 0)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let investment = \ZExcel\Calculation\Functions::flattenSingleValue(investment);
        let redemption = \ZExcel\Calculation\Functions::flattenSingleValue(redemption);
        let basis      = \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if ((is_numeric(investment)) && (is_numeric(redemption)) && (is_numeric(basis))) {
            let investment    = (float) investment;
            let redemption    = (float) redemption;
            let basis        = (int) basis;
            if ((investment <= 0) || (redemption <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }

            return (((double) redemption / (double) investment) - 1) / (daysBetweenSettlementAndMaturity);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * IPMT
     *
     * Returns the interest payment for a given period for an investment based on periodic, constant payments and a constant interest rate.
     *
     * Excel Function:
     *        IPMT(rate,per,nper,pv[,fv][,type])
     *
     * @param    float    rate    Interest rate per period
     * @param    int        per    Period for which we want to find the interest
     * @param    int        nper    Number of periods
     * @param    float    pv        Present Value
     * @param    float    fv        Future Value
     * @param    int        type    Payment type: 0 = at the end of each period, 1 = at the beginning of each period
     * @return    float
     */
    public static function ipmt(double rate, int per, int nper, double pv, double fv = 0, int type = 0)
    {
        var interestAndPrincipal;
        
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let per  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(per);
        let nper = (int) \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pv   = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let fv   = \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type = (int) \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (per <= 0 || per > nper) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Calculate
        let interestAndPrincipal = self::interestAndPrincipal(rate, per, nper, pv, fv, type);
        return interestAndPrincipal[0];
    }

    /**
     * IRR
     *
     * Returns the internal rate of return for a series of cash flows represented by the numbers in values.
     * These cash flows do not have to be even, as they would be for an annuity. However, the cash flows must occur
     * at regular intervals, such as monthly or annually. The internal rate of return is the interest rate received
     * for an investment consisting of payments (negative values) and income (positive values) that occur at regular
     * periods.
     *
     * Excel Function:
     *        IRR(values[,guess])
     *
     * @param    float[]    values        An array or a reference to cells that contain numbers for which you want
     *                                    to calculate the internal rate of return.
     *                                Values must contain at least one positive value and one negative value to
     *                                    calculate the internal rate of return.
     * @param    float    guess        A number that you guess is close to the result of IRR
     * @return    float
     */
    public static function irr(array values, double guess = 0.1)
    {
        double x1, x2, f, f1, f2, rtb, dx, x_mid, f_mid;
        int i;
        
        if (!is_array(values)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let values = \ZExcel\Calculation\Functions::flattenArray(values);
        let guess  = \ZExcel\Calculation\Functions::flattenSingleValue(guess);

        // create an initial range, with a root somewhere between 0 and guess
        let x1 = 0.0;
        let x2 = guess;
        let f1 = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x1, values);
        let f2 = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x2, values);
        
        for i in range(0, self::FINANCIAL_MAX_ITERATIONS - 1) {
            if ((f1 * f2) < 0.0) {
                break;
            }
            
            if (abs(f1) < abs(f2)) {
                let x1 = x1 + (1.6 * (x1 - x2));
                let f1 = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x1, values);
            } else {
                let x2 = x2 + (1.6 * (x2 - x1));
                let f2 = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x2, values);
            }
        }
        
        if ((f1 * f2) > 0.0) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        let f = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x1, values);
        
        if (f < 0.0) {
            let rtb = x1;
            let dx = x2 - x1;
        } else {
            let rtb = x2;
            let dx = x1 - x2;
        }

        for i in range(0, self::FINANCIAL_MAX_ITERATIONS - 1) {
            let dx *= 0.5;
            let x_mid = rtb + dx;
            let f_mid = 0.0 + call_user_func(["\\ZExcel\\Calculation\\Financial", "npv"], x_mid, values);
            
            if (f_mid <= 0.0) {
                let rtb = x_mid;
            }
            
            if ((abs(f_mid) < self::FINANCIAL_PRECISION) || (abs(dx) < self::FINANCIAL_PRECISION)) {
                return x_mid;
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * ISPMT
     *
     * Returns the interest payment for an investment based on an interest rate and a constant payment schedule.
     *
     * Excel Function:
     *     =ISPMT(interest_rate, period, number_payments, PV)
     *
     * interest_rate is the interest rate for the investment
     *
     * period is the period to calculate the interest rate.  It must be betweeen 1 and number_payments.
     *
     * number_payments is the number of payments for the annuity
     *
     * PV is the loan amount or present value of the payments
     */
    public static function ispmt()
    {
        var aArgs, interestRate, period, numberPeriods, principleRemaining;
        double returnValue = 0, principlePayment;
        int i;

        // Get the parameters
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        let interestRate = array_shift(aArgs);
        let period = array_shift(aArgs);
        let numberPeriods = array_shift(aArgs);
        let principleRemaining = array_shift(aArgs);

        // Calculate
        let principlePayment = (principleRemaining * 1.0) / (numberPeriods * 1.0);
        
        for i in range(0, period) {
            let returnValue = interestRate * principleRemaining * -1;
            let principleRemaining = principleRemaining - (double) principlePayment;
            // principle needs to be 0 after the last payment, don"t let floating point screw it up
            if (i == numberPeriods) {
                let returnValue = 0;
            }
        }
        
        return(returnValue);
    }


    /**
     * MIRR
     *
     * Returns the modified internal rate of return for a series of periodic cash flows. MIRR considers both
     *        the cost of the investment and the interest received on reinvestment of cash.
     *
     * Excel Function:
     *        MIRR(values,finance_rate, reinvestment_rate)
     *
     * @param    float[]    values                An array or a reference to cells that contain a series of payments and
     *                                            income occurring at regular intervals.
     *                                        Payments are negative value, income is positive values.
     * @param    float    finance_rate        The interest rate you pay on the money used in the cash flows
     * @param    float    reinvestment_rate    The interest rate you receive on the cash flows as you reinvest them
     * @return    float
     */
    public static function mirr(array values, double finance_rate, double reinvestment_rate)
    {
        var i, v;
        double rr, fr, npv_pos, npv_neg, mirr;
        int n;
        
        if (!is_array(values)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let values            = \ZExcel\Calculation\Functions::flattenArray(values);
        let finance_rate      = \ZExcel\Calculation\Functions::flattenSingleValue(finance_rate);
        let reinvestment_rate = \ZExcel\Calculation\Functions::flattenSingleValue(reinvestment_rate);
        let n = 0 + count(values);

        let rr = 1.0 + reinvestment_rate;
        let fr = 1.0 + finance_rate;

        let npv_pos = 0.0;
        let npv_neg = 0.0;
        
        for i, v in values {
            if (v >= 0) {
                let npv_pos = npv_pos + (v / pow(rr, i));
            } else {
                let npv_neg = npv_neg + (v / pow(fr, i));
            }
        }

        if ((npv_neg == 0) || (npv_pos == 0) || (reinvestment_rate <= -1)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        let mirr = pow((-npv_pos * pow(rr, n)) / (npv_neg * (rr)), (1.0 / (n - 1))) - 1.0;

        return (is_finite(mirr) ? mirr : \ZExcel\Calculation\Functions::VaLUE());
    }


    /**
     * NOMINAL
     *
     * Returns the nominal interest rate given the effective rate and the number of compounding payments per year.
     *
     * @param    float    effect_rate    Effective interest rate
     * @param    int        npery            Number of compounding payments per year
     * @return    float
     */
    public static function nominal(double effect_rate = 0, int npery = 0)
    {
        let effect_rate = \ZExcel\Calculation\Functions::flattenSingleValue(effect_rate);
        let npery       = (int)\ZExcel\Calculation\Functions::flattenSingleValue(npery);

        // Validate parameters
        if (effect_rate <= 0 || npery < 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Calculate
        return (double) npery * (pow(effect_rate + 1, 1 / (double) npery) - 1);
    }


    /**
     * NPER
     *
     * Returns the number of periods for a cash flow with constant periodic payments (annuities), and interest rate.
     *
     * @param    float    rate    Interest rate per period
     * @param    int        pmt    Periodic payment (annuity)
     * @param    float    pv        Present Value
     * @param    float    fv        Future Value
     * @param    int        type    Payment type: 0 = at the end of each period, 1 = at the beginning of each period
     * @return    float
     */
    public static function nper(double rate = 0, double pmt = 0, double pv = 0, double fv = 0, double type = 0) -> double
    {
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let pmt  = \ZExcel\Calculation\Functions::flattenSingleValue(pmt);
        let pv   = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let fv   = \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type = \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Calculate
        if (!is_null(rate) && rate != 0) {
            if (pmt == 0 && pv == 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return log(((pmt * ((1 + (rate * type)) / rate)) - fv) / ((pv + (pmt * (1 + rate * type)) / rate))) / log(1 + rate);
        }
        
        if (pmt == 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        return (-pv -fv) / pmt;
    }

    /**
     * NPV
     *
     * Returns the Net Present Value of a cash flow series given a discount rate.
     *
     * @return    float
     */
    public static function npv() -> double
    {
        var aArgs, rate;
        double returnValue = 0;
        int i;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let rate = array_shift(aArgs);
        for i in range(1, count(aArgs)) {
            // Is it a numeric value?
            if (is_numeric(aArgs[i - 1])) {
                let returnValue = returnValue + aArgs[i - 1] / pow(1 + (double) rate, i);
            }
        }

        // Return
        return returnValue;
    }

    /**
     * PMT
     *
     * Returns the constant payment (annuity) for a cash flow with a constant interest rate.
     *
     * @param    float    rate    Interest rate per period
     * @param    int        nper    Number of periods
     * @param    float    pv        Present Value
     * @param    float    fv        Future Value
     * @param    int        type    Payment type: 0 = at the end of each period, 1 = at the beginning of each period
     * @return    float
     */
    public static function pmt(double rate = 0, int nper = 0, double pv = 0, double fv = 0, int type = 0) -> double
    {
        double tmp1, tmp2;
        
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let nper = \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pv   = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let fv   = \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type = \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Calculate
        if (!is_null(rate) && rate != 0) {
            let tmp1 = (double) pow(1 + rate, nper) - 1;
            let tmp1 = tmp1 * (1 / rate);
            let tmp2 = pv * pow(1 + rate, nper);
            
            return ((1 - fv - tmp2) / (1 + rate * type)) / tmp1;
        }
        
        return (-pv - fv) / nper;
    }


    /**
     * PPMT
     *
     * Returns the interest payment for a given period for an investment based on periodic, constant payments and a constant interest rate.
     *
     * @param    float    rate    Interest rate per period
     * @param    int        per    Period for which we want to find the interest
     * @param    int        nper    Number of periods
     * @param    float    pv        Present Value
     * @param    float    fv        Future Value
     * @param    int        type    Payment type: 0 = at the end of each period, 1 = at the beginning of each period
     * @return    float
     */
    public static function ppmt(double rate, int per, int nper, float pv, float fv = 0, int type = 0)
    {
        var interestAndPrincipal;
        
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let per  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(per);
        let nper = (int) \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pv   = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let fv   = \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type = (int) \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        if (per <= 0 || per > nper) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // Calculate
        let interestAndPrincipal = self::interestAndPrincipal(rate, per, nper, pv, fv, type);
        return interestAndPrincipal[1];
    }


    public static function price(var settlement, var maturity, double rate, double yield, double redemption, int frequency, var basis = 0)
    {
        var dsc, e, n, a, baseYF, rfp, de, result;
        int k;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let rate       = (float) \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let yield      = (float) \ZExcel\Calculation\Functions::flattenSingleValue(yield);
        let redemption = (float) \ZExcel\Calculation\Functions::flattenSingleValue(redemption);
        let frequency  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(frequency);
        let basis      = (is_null(basis))    ? 0 :    (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        let settlement = \ZExcel\Calculation\DateTime::getDateValue(settlement);

        if (is_string(settlement)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);
        
        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if ((settlement > maturity) ||
            (!self::isValidFrequency(frequency)) ||
            ((basis < 0) || (basis > 4))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let dsc = self::CoUPDAYSNC(settlement, maturity, frequency, basis);
        let e = self::CoUPDAYS(settlement, maturity, frequency, basis);
        let n = self::CoUPNUM(settlement, maturity, frequency, basis);
        let a = self::CoUPDAYBS(settlement, maturity, frequency, basis);

        let baseYF = 1.0 + (yield / frequency);
        let rfp    = 100 * (rate / frequency);
        let de     = dsc / e;

        let n = n - 1;
        let result = redemption / pow(baseYF, ((double) n + de));
        for k in range(0, n) {
            let result = result + (rfp / (pow(baseYF, (k + de))));
        }
        let result = result - (rfp * (a / e));

        return result;
    }


    /**
     * PRICEDISC
     *
     * Returns the price per 100 face value of a discounted security.
     *
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    int        discount    The security"s discount rate.
     * @param    int        redemption    The security"s redemption value per 100 face value.
     * @param    int        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function priceDisc(var settlement, var maturity, int discount, int redemption, int basis = 0)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let discount   = (float) \ZExcel\Calculation\Functions::flattenSingleValue(discount);
        let redemption = (float) \ZExcel\Calculation\Functions::flattenSingleValue(redemption);
        let basis      = (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if ((is_numeric(discount)) && (is_numeric(redemption)) && (is_numeric(basis))) {
            if ((discount <= 0) || (redemption <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }

            return redemption * (1 - discount * daysBetweenSettlementAndMaturity);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * PRICEMAT
     *
     * Returns the price per 100 face value of a security that pays interest at maturity.
     *
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security"s settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    mixed    issue        The security"s issue date.
     * @param    int        rate        The security"s interest rate at date of issue.
     * @param    int        yield        The security"s annual yield.
     * @param    int        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function priceMat(var settlement, var maturity, var issue, var rate, var yield, int basis = 0)
    {
        var daysPerYear, daysBetweenIssueAndSettlement, daysBetweenIssueAndMaturity, daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let issue      = \ZExcel\Calculation\Functions::flattenSingleValue(issue);
        let rate       = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let yield      = \ZExcel\Calculation\Functions::flattenSingleValue(yield);
        let basis      = (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if (is_numeric(rate) && is_numeric(yield)) {
            if ((rate <= 0) || (yield <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(settlement), basis);
            
            if (!is_numeric(daysPerYear)) {
                return daysPerYear;
            }
            
            let daysBetweenIssueAndSettlement = \ZExcel\Calculation\DateTime::YeARFRAC(issue, settlement, basis);
            
            if (!is_numeric(daysBetweenIssueAndSettlement)) {
                //    return date error
                return daysBetweenIssueAndSettlement;
            }
            
            let daysBetweenIssueAndSettlement = daysBetweenIssueAndSettlement * daysPerYear;
            let daysBetweenIssueAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(issue, maturity, basis);
            
            if (!is_numeric(daysBetweenIssueAndMaturity)) {
                //    return date error
                return daysBetweenIssueAndMaturity;
            }
            
            let daysBetweenIssueAndMaturity = daysBetweenIssueAndMaturity * daysPerYear;
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }
            
            let daysBetweenSettlementAndMaturity = daysBetweenSettlementAndMaturity *daysPerYear;

            return ((100 + ((daysBetweenIssueAndMaturity / daysPerYear) * rate * 100)) /
                   (1 + ((daysBetweenSettlementAndMaturity / daysPerYear) * yield)) -
                   ((daysBetweenIssueAndSettlement / daysPerYear) * rate * 100));
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * PV
     *
     * Returns the Present Value of a cash flow with constant payments and interest rate (annuities).
     *
     * @param    float    rate    Interest rate per period
     * @param    int        nper    Number of periods
     * @param    float    pmt    Periodic payment (annuity)
     * @param    float    fv        Future Value
     * @param    int        type    Payment type: 0 = at the end of each period, 1 = at the beginning of each period
     * @return    float
     */
    public static function pv(double rate = 0, int nper = 0, double pmt = 0, double fv = 0, int type = 0)
    {
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let nper = \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pmt  = \ZExcel\Calculation\Functions::flattenSingleValue(pmt);
        let fv   = \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type = \ZExcel\Calculation\Functions::flattenSingleValue(type);

        // Validate parameters
        if (type != 0 && type != 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // Calculate
        if (!is_null(rate) && rate != 0) {
            return (-pmt * (1 + rate * type) * ((pow(1 + rate, nper) - 1) / rate) - fv) / pow(1 + rate, nper);
        }
        return -fv - pmt * nper;
    }


    /**
     * RATE
     *
     * Returns the interest rate per period of an annuity.
     * RATE is calculated by iteration and can have zero or more solutions.
     * If the successive results of RATE do not converge to within 0.0000001 after 20 iterations,
     * RATE returns the #NUM! error value.
     *
     * Excel Function:
     *        RATE(nper,pmt,pv[,fv[,type[,guess]]])
     *
     * @access    public
     * @category Financial Functions
     * @param    float    nper        The total number of payment periods in an annuity.
     * @param    float    pmt            The payment made each period and cannot change over the life
     *                                    of the annuity.
     *                                Typically, pmt includes principal and interest but no other
     *                                    fees or taxes.
     * @param    float    pv            The present value - the total amount that a series of future
     *                                    payments is worth now.
     * @param    float    fv            The future value, or a cash balance you want to attain after
     *                                    the last payment is made. If fv is omitted, it is assumed
     *                                    to be 0 (the future value of a loan, for example, is 0).
     * @param    integer    type        A number 0 or 1 and indicates when payments are due:
     *                                        0 or omitted    At the end of the period.
     *                                        1                At the beginning of the period.
     * @param    float    guess        Your guess for what the rate will be.
     *                                    If you omit guess, it is assumed to be 10 percent.
     * @return    float
     **/
    public static function rate(double nper, double pmt, double pv, var fv = 0.0, var type = 0, var guess = 0.1) -> double
    {
        double rate, y, y0, y1, f, i, x0, x1;
        
        let nper  = (int) \ZExcel\Calculation\Functions::flattenSingleValue(nper);
        let pmt   = \ZExcel\Calculation\Functions::flattenSingleValue(pmt);
        let pv    = \ZExcel\Calculation\Functions::flattenSingleValue(pv);
        let fv    = (is_null(fv))    ? 0.0 : \ZExcel\Calculation\Functions::flattenSingleValue(fv);
        let type  = (is_null(type))  ? 0   : (int) \ZExcel\Calculation\Functions::flattenSingleValue(type);
        let guess = (is_null(guess)) ? 0.1 : \ZExcel\Calculation\Functions::flattenSingleValue(guess);

        let rate = (double) guess;
        
        if (abs(rate) < self::FINANCIAL_PRECISION) {
            let y = pv * (1 + nper * rate) + pmt * (1 + rate * (double) type) * nper + (double) fv;
        } else {
            let f = 0.0 + exp(nper * log(1 + rate));
            let y = pv * f + pmt * (1 / rate + (double) type) * (f - 1) + (double) fv;
        }
        
        let y0 = pv + pmt * nper + fv;
        let y1 = pv * f + pmt * (1 / rate + (double) type) * (f - 1) + (double) fv;

        // find root by secant method
        let i = 0.0;
        let x0 = 0.0;
        let x1 = rate;
        
        while ((abs(y0 - y1) > self::FINANCIAL_PRECISION) && (i < self::FINANCIAL_MAX_ITERATIONS)) {
            let rate = (y1 * x0 - y0 * x1) / (y1 - y0);
            let x0 = x1;
            let x1 = rate;
            
            if ((nper * abs(pmt)) > (pv - fv)) {
                let x1 = abs(x1);
            }
            if (abs(rate) < self::FINANCIAL_PRECISION) {
                let y = pv * (1 + nper * rate) + pmt * (1 + rate * (double) type) * nper + (double) fv;
            } else {
                let f = 0.0 + exp(nper * log(1 + rate));
                let y = pv * f + pmt * (1 / rate + (double) type) * (f - 1) + (double) fv;
            }

            let y0 = y1;
            let y1 = y;
            let i = i + 1;
        }
        
        return rate;
    }


    /**
     * RECEIVED
     *
     * Returns the price per 100 face value of a discounted security.
     *
     * @param    mixed    settlement    The security"s settlement date.
     *                                The security settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity    The security"s maturity date.
     *                                The maturity date is the date when the security expires.
     * @param    int        investment    The amount invested in the security.
     * @param    int        discount    The security"s discount rate.
     * @param    int        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function received(var settlement, var maturity, var investment, var discount, var basis = 0)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let investment = (float) \ZExcel\Calculation\Functions::flattenSingleValue(investment);
        let discount   = (float) \ZExcel\Calculation\Functions::flattenSingleValue(discount);
        let basis      = (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if ((is_numeric(investment)) && (is_numeric(discount)) && (is_numeric(basis))) {
            if ((investment <= 0) || (discount <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }

            return investment / ( 1 - (discount * daysBetweenSettlementAndMaturity));
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SLN
     *
     * Returns the straight-line depreciation of an asset for one period
     *
     * @param    cost        Initial cost of the asset
     * @param    salvage        Value at the end of the depreciation
     * @param    life        Number of periods over which the asset is depreciated
     * @return    float
     */
    public static function sln(var cost, var salvage, var life)
    {
        let cost    = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let salvage = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let life    = \ZExcel\Calculation\Functions::flattenSingleValue(life);

        // Calculate
        if ((is_numeric(cost)) && (is_numeric(salvage)) && (is_numeric(life))) {
            if (life < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return (cost - salvage) / life;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SYD
     *
     * Returns the sum-of-years" digits depreciation of an asset for a specified period.
     *
     * @param    cost        Initial cost of the asset
     * @param    salvage        Value at the end of the depreciation
     * @param    life        Number of periods over which the asset is depreciated
     * @param    period        Period
     * @return    float
     */
    public static function syd(var cost, var salvage, var life, var period)
    {
        let cost    = \ZExcel\Calculation\Functions::flattenSingleValue(cost);
        let salvage = \ZExcel\Calculation\Functions::flattenSingleValue(salvage);
        let life    = \ZExcel\Calculation\Functions::flattenSingleValue(life);
        let period  = \ZExcel\Calculation\Functions::flattenSingleValue(period);

        // Calculate
        if ((is_numeric(cost)) && (is_numeric(salvage)) && (is_numeric(life)) && (is_numeric(period))) {
            if ((life < 1) || (period > life)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return ((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1));
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * TBILLEQ
     *
     * Returns the bond-equivalent yield for a Treasury bill.
     *
     * @param    mixed    settlement    The Treasury bill"s settlement date.
     *                                The Treasury bill"s settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
     * @param    mixed    maturity    The Treasury bill"s maturity date.
     *                                The maturity date is the date when the Treasury bill expires.
     * @param    int        discount    The Treasury bill"s discount rate.
     * @return    float
     */
    public static function tBillEQ(var settlement, var maturity, int discount)
    {
        var testValue, daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let discount   = \ZExcel\Calculation\Functions::flattenSingleValue(discount);

        //    Use TBILLPRICE for validation
        let testValue = self::TBiLLPRICE(settlement, maturity, discount);
        
        if (is_string(testValue)) {
            return testValue;
        }

        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);

        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
            let maturity = maturity + 1;
            let daysBetweenSettlementAndMaturity = 360 * \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity);
        } else {
            let daysBetweenSettlementAndMaturity = (\ZExcel\Calculation\DateTime::getDateValue(maturity) - \ZExcel\Calculation\DateTime::getDateValue(settlement));
        }

        return (365 * discount) / (360 - discount * daysBetweenSettlementAndMaturity);
    }


    /**
     * TBILLPRICE
     *
     * Returns the yield for a Treasury bill.
     *
     * @param    mixed    settlement    The Treasury bill"s settlement date.
     *                                The Treasury bill"s settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
     * @param    mixed    maturity    The Treasury bill"s maturity date.
     *                                The maturity date is the date when the Treasury bill expires.
     * @param    int        discount    The Treasury bill"s discount rate.
     * @return    float
     */
    public static function tBillPrice(var settlement, var maturity, int discount)
    {
        var daysBetweenSettlementAndMaturity, price;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let discount   = \ZExcel\Calculation\Functions::flattenSingleValue(discount);

        let maturity = \ZExcel\Calculation\DateTime::getDateValue(maturity);

        if (is_string(maturity)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        //    Validate
        if (is_numeric(discount)) {
            if (discount <= 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let maturity = maturity + 1;
                let daysBetweenSettlementAndMaturity = 360 * \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity);
                
                if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                    //    return date error
                    return daysBetweenSettlementAndMaturity;
                }
            } else {
                let daysBetweenSettlementAndMaturity = (\ZExcel\Calculation\DateTime::getDateValue(maturity) - \ZExcel\Calculation\DateTime::getDateValue(settlement));
            }

            if (daysBetweenSettlementAndMaturity > 360) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            let price = 100 * (1 - ((discount * daysBetweenSettlementAndMaturity) / 360));
            
            if (price <= 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return price;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * TBILLYIELD
     *
     * Returns the yield for a Treasury bill.
     *
     * @param    mixed    settlement    The Treasury bill"s settlement date.
     *                                The Treasury bill"s settlement date is the date after the issue date when the Treasury bill is traded to the buyer.
     * @param    mixed    maturity    The Treasury bill"s maturity date.
     *                                The maturity date is the date when the Treasury bill expires.
     * @param    int        price        The Treasury bill"s price per 100 face value.
     * @return    float
     */
    public static function tBillYield(var settlement, var maturity, var price)
    {
        var daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let price      = \ZExcel\Calculation\Functions::flattenSingleValue(price);

        //    Validate
        if (is_numeric(price)) {
            if (price <= 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE) {
                let maturity = maturity + 1;
                let daysBetweenSettlementAndMaturity = 360 * \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity);
                
                if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                    //    return date error
                    return daysBetweenSettlementAndMaturity;
                }
            } else {
                let daysBetweenSettlementAndMaturity = (\ZExcel\Calculation\DateTime::getDateValue(maturity) - \ZExcel\Calculation\DateTime::getDateValue(settlement));
            }

            if (daysBetweenSettlementAndMaturity > 360) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            return ((100 - price) / price) * (360 / daysBetweenSettlementAndMaturity);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    public static function xirr(var values, var dates, var guess = 0.1)
    {
        var x1, x2, f1, f2, f, dx, rtb, f_mid, x_mid;
        int i;
        
        if ((!is_array(values)) && (!is_array(dates))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let values = \ZExcel\Calculation\Functions::flattenArray(values);
        let dates  = \ZExcel\Calculation\Functions::flattenArray(dates);
        let guess  = \ZExcel\Calculation\Functions::flattenSingleValue(guess);
        
        if (count(values) != count(dates)) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        // create an initial range, with a root somewhere between 0 and guess
        let x1 = 0.0;
        let x2 = guess;
        let f1 = self::xnpv(x1, values, dates);
        let f2 = self::xnpv(x2, values, dates);
        
        for i in range(0, self::FINANCIAL_MAX_ITERATIONS - 1) {
            if ((f1 * f2) < 0.0) {
                break;
            } else {
                if (abs(f1) < abs(f2)) {
                    let x1 = x1 + (1.6 * (x1 - (double) x2));
                    let f1 = self::xnpv(x1, values, dates);
                } else {
                    let x2 = x2 + (1.6 * (x2 - (double) x1));
                    let f2 = self::xnpv(x2, values, dates);
                }
            }
        }
        
        if ((f1 * f2) > 0.0) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        let f = self::xnpv(x1, values, dates);
        
        if (f < 0.0) {
            let rtb = x1;
            let dx = (double) x2 - x1;
        } else {
            let rtb = x2;
            let dx = (double) x1 - x2;
        }

        for i in range(0, self::FINANCIAL_MAX_ITERATIONS - 1) {
            let dx = dx * 0.5;
            let x_mid = rtb + dx;
            let f_mid = self::xnpv(x_mid, values, dates);
            
            if (f_mid <= 0.0) {
                let rtb = x_mid;
            }
            
            if ((abs(f_mid) < self::FINANCIAL_PRECISION) || (abs(dx) < self::FINANCIAL_PRECISION)) {
                return x_mid;
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * XNPV
     *
     * Returns the net present value for a schedule of cash flows that is not necessarily periodic.
     * To calculate the net present value for a series of cash flows that is periodic, use the NPV function.
     *
     * Excel Function:
     *        =XNPV(rate,values,dates)
     *
     * @param    float            rate        The discount rate to apply to the cash flows.
     * @param    array of float    values     A series of cash flows that corresponds to a schedule of payments in dates.
     *                                         The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment.
     *                                         If the first value is a cost or payment, it must be a negative value. All succeeding payments are discounted based on a 365-day year.
     *                                         The series of values must contain at least one positive value and one negative value.
     * @param    array of mixed    dates      A schedule of payment dates that corresponds to the cash flow payments.
     *                                         The first payment date indicates the beginning of the schedule of payments.
     *                                         All other dates must be later than this date, but they may occur in any order.
     * @return    float
     */
    public static function xnpv(double rate, var values, var dates)
    {
        double xnpv;
        int i, valCount;
        
        let rate = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        
        if (!is_numeric(rate)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        if ((!is_array(values)) || (!is_array(dates))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let values   = \ZExcel\Calculation\Functions::flattenArray(values);
        let dates    = \ZExcel\Calculation\Functions::flattenArray(dates);
        let valCount = 0 + count(values);
        
        if (valCount != count(dates)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        if ((min(values) > 0) || (max(values) < 0)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        let xnpv = 0.0;
        
        for i in range(0, valCount - 1) {
            if (!is_numeric(values[i])) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            let xnpv = xnpv + (values[i] / pow(1 + rate, \ZExcel\Calculation\DateTime::DaTEDIF(dates[0], dates[i], "d") / 365));
        }
        
        return (is_finite(xnpv)) ? xnpv : \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * YIELDDISC
     *
     * Returns the annual yield of a security that pays interest at maturity.
     *
     * @param    mixed    settlement      The security"s settlement date.
     *                                    The security"s settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity        The security"s maturity date.
     *                                    The maturity date is the date when the security expires.
     * @param    int        price         The security"s price per 100 face value.
     * @param    int        redemption    The security"s redemption value per 100 face value.
     * @param    int        basis         The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function yieldDisc(var settlement, var maturity, var price, var redemption, int basis = 0)
    {
        var daysPerYear, daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let price      = \ZExcel\Calculation\Functions::flattenSingleValue(price);
        let redemption = \ZExcel\Calculation\Functions::flattenSingleValue(redemption);
        let basis      = (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if (is_numeric(price) && is_numeric(redemption)) {
            if ((price <= 0) || (redemption <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(settlement), basis);
            
            if (!is_numeric(daysPerYear)) {
                return daysPerYear;
            }
            
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }
            
            let daysBetweenSettlementAndMaturity = daysBetweenSettlementAndMaturity * daysPerYear;

            return ((redemption - price) / price) * (daysPerYear / daysBetweenSettlementAndMaturity);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * YIELDMAT
     *
     * Returns the annual yield of a security that pays interest at maturity.
     *
     * @param    mixed    settlement     The security"s settlement date.
     *                                   The security"s settlement date is the date after the issue date when the security is traded to the buyer.
     * @param    mixed    maturity       The security"s maturity date.
     *                                   The maturity date is the date when the security expires.
     * @param    mixed    issue          The security"s issue date.
     * @param    int        rate         The security"s interest rate at date of issue.
     * @param    int        price        The security"s price per 100 face value.
     * @param    int        basis        The type of day count to use.
     *                                        0 or omitted    US (NASD) 30/360
     *                                        1                Actual/actual
     *                                        2                Actual/360
     *                                        3                Actual/365
     *                                        4                European 30/360
     * @return    float
     */
    public static function yieldMat(var settlement, var maturity, var issue, var rate, var price, int basis = 0)
    {
        var daysPerYear, daysBetweenIssueAndSettlement, daysBetweenIssueAndMaturity, daysBetweenSettlementAndMaturity;
        
        let settlement = \ZExcel\Calculation\Functions::flattenSingleValue(settlement);
        let maturity   = \ZExcel\Calculation\Functions::flattenSingleValue(maturity);
        let issue      = \ZExcel\Calculation\Functions::flattenSingleValue(issue);
        let rate       = \ZExcel\Calculation\Functions::flattenSingleValue(rate);
        let price      = \ZExcel\Calculation\Functions::flattenSingleValue(price);
        let basis      = (int) \ZExcel\Calculation\Functions::flattenSingleValue(basis);

        //    Validate
        if (is_numeric(rate) && is_numeric(price)) {
            if ((rate <= 0) || (price <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let daysPerYear = self::daysPerYear(\ZExcel\Calculation\DateTime::YeAR(settlement), basis);
            
            if (!is_numeric(daysPerYear)) {
                return daysPerYear;
            }
            
            let daysBetweenIssueAndSettlement = \ZExcel\Calculation\DateTime::YeARFRAC(issue, settlement, basis);
            
            if (!is_numeric(daysBetweenIssueAndSettlement)) {
                //    return date error
                return daysBetweenIssueAndSettlement;
            }
            
            let daysBetweenIssueAndSettlement = daysBetweenIssueAndSettlement * daysPerYear;
            let daysBetweenIssueAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(issue, maturity, basis);
            
            if (!is_numeric(daysBetweenIssueAndMaturity)) {
                //    return date error
                return daysBetweenIssueAndMaturity;
            }
            
            let daysBetweenIssueAndMaturity = daysBetweenIssueAndMaturity * daysPerYear;
            let daysBetweenSettlementAndMaturity = \ZExcel\Calculation\DateTime::YeARFRAC(settlement, maturity, basis);
            
            if (!is_numeric(daysBetweenSettlementAndMaturity)) {
                //    return date error
                return daysBetweenSettlementAndMaturity;
            }
            
            let daysBetweenSettlementAndMaturity = daysBetweenSettlementAndMaturity * daysPerYear;

            return ((1 + ((daysBetweenIssueAndMaturity / daysPerYear) * rate) - ((price / 100) + ((daysBetweenIssueAndSettlement / daysPerYear) * rate))) /
                   ((price / 100) + ((daysBetweenIssueAndSettlement / daysPerYear) * rate))) *
                   (daysPerYear / daysBetweenSettlementAndMaturity);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }
}
