namespace ZExcel\Worksheet\AutoFilter\Column;

class Rule
{
    const AUTOFILTER_RULETYPE_FILTER        = "filter";
    const AUTOFILTER_RULETYPE_DATEGROUP        = "dateGroupItem";
    const AUTOFILTER_RULETYPE_CUSTOMFILTER    = "customFilter";
    const AUTOFILTER_RULETYPE_DYNAMICFILTER    = "dynamicFilter";
    const AUTOFILTER_RULETYPE_TOPTENFILTER    = "top10Filter";

    const AUTOFILTER_RULETYPE_DATEGROUP_YEAR    = "year";
    const AUTOFILTER_RULETYPE_DATEGROUP_MONTH    = "month";
    const AUTOFILTER_RULETYPE_DATEGROUP_DAY        = "day";
    const AUTOFILTER_RULETYPE_DATEGROUP_HOUR    = "hour";
    const AUTOFILTER_RULETYPE_DATEGROUP_MINUTE    = "minute";
    const AUTOFILTER_RULETYPE_DATEGROUP_SECOND    = "second";

    const AUTOFILTER_RULETYPE_DYNAMIC_YESTERDAY        = "yesterday";
    const AUTOFILTER_RULETYPE_DYNAMIC_TODAY            = "today";
    const AUTOFILTER_RULETYPE_DYNAMIC_TOMORROW        = "tomorrow";
    const AUTOFILTER_RULETYPE_DYNAMIC_YEARTODATE    = "yearToDate";
    const AUTOFILTER_RULETYPE_DYNAMIC_THISYEAR        = "thisYear";
    const AUTOFILTER_RULETYPE_DYNAMIC_THISQUARTER    = "thisQuarter";
    const AUTOFILTER_RULETYPE_DYNAMIC_THISMONTH        = "thisMonth";
    const AUTOFILTER_RULETYPE_DYNAMIC_THISWEEK        = "thisWeek";
    const AUTOFILTER_RULETYPE_DYNAMIC_LASTYEAR        = "lastYear";
    const AUTOFILTER_RULETYPE_DYNAMIC_LASTQUARTER    = "lastQuarter";
    const AUTOFILTER_RULETYPE_DYNAMIC_LASTMONTH        = "lastMonth";
    const AUTOFILTER_RULETYPE_DYNAMIC_LASTWEEK        = "lastWeek";
    const AUTOFILTER_RULETYPE_DYNAMIC_NEXTYEAR        = "nextYear";
    const AUTOFILTER_RULETYPE_DYNAMIC_NEXTQUARTER    = "nextQuarter";
    const AUTOFILTER_RULETYPE_DYNAMIC_NEXTMONTH        = "nextMonth";
    const AUTOFILTER_RULETYPE_DYNAMIC_NEXTWEEK        = "nextWeek";
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_1        = "M1";
    const AUTOFILTER_RULETYPE_DYNAMIC_JANUARY        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_1;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_2        = "M2";
    const AUTOFILTER_RULETYPE_DYNAMIC_FEBRUARY        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_2;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_3        = "M3";
    const AUTOFILTER_RULETYPE_DYNAMIC_MARCH            = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_3;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_4        = "M4";
    const AUTOFILTER_RULETYPE_DYNAMIC_APRIL            = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_4;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_5        = "M5";
    const AUTOFILTER_RULETYPE_DYNAMIC_MAY            = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_5;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_6        = "M6";
    const AUTOFILTER_RULETYPE_DYNAMIC_JUNE            = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_6;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_7        = "M7";
    const AUTOFILTER_RULETYPE_DYNAMIC_JULY            = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_7;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_8        = "M8";
    const AUTOFILTER_RULETYPE_DYNAMIC_AUGUST        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_8;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_9        = "M9";
    const AUTOFILTER_RULETYPE_DYNAMIC_SEPTEMBER        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_9;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_10        = "M10";
    const AUTOFILTER_RULETYPE_DYNAMIC_OCTOBER        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_10;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_11        = "M11";
    const AUTOFILTER_RULETYPE_DYNAMIC_NOVEMBER        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_11;
    const AUTOFILTER_RULETYPE_DYNAMIC_MONTH_12        = "M12";
    const AUTOFILTER_RULETYPE_DYNAMIC_DECEMBER        = self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_12;
    const AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_1        = "Q1";
    const AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_2        = "Q2";
    const AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_3        = "Q3";
    const AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_4        = "Q4";
    const AUTOFILTER_RULETYPE_DYNAMIC_ABOVEAVERAGE    = "aboveAverage";
    const AUTOFILTER_RULETYPE_DYNAMIC_BELOWAVERAGE    = "belowAverage";

    /*
     *    The only valid filter rule operators for filter and customFilter types are:
     *        <xsd:enumeration value="equal"/>
     *        <xsd:enumeration value="lessThan"/>
     *        <xsd:enumeration value="lessThanOrEqual"/>
     *        <xsd:enumeration value="notEqual"/>
     *        <xsd:enumeration value="greaterThanOrEqual"/>
     *        <xsd:enumeration value="greaterThan"/>
     */
    const AUTOFILTER_COLUMN_RULE_EQUAL                = "equal";
    const AUTOFILTER_COLUMN_RULE_NOTEQUAL            = "notEqual";
    const AUTOFILTER_COLUMN_RULE_GREATERTHAN        = "greaterThan";
    const AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL    = "greaterThanOrEqual";
    const AUTOFILTER_COLUMN_RULE_LESSTHAN            = "lessThan";
    const AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL    = "lessThanOrEqual";

    const AUTOFILTER_COLUMN_RULE_TOPTEN_BY_VALUE    = "byValue";
    const AUTOFILTER_COLUMN_RULE_TOPTEN_PERCENT        = "byPercent";

    const AUTOFILTER_COLUMN_RULE_TOPTEN_TOP            = "top";
    const AUTOFILTER_COLUMN_RULE_TOPTEN_BOTTOM        = "bottom";


    /* Rule Operators (Numeric, Boolean etc) */
//    const AUTOFILTER_COLUMN_RULE_BETWEEN            = "between";        //    greaterThanOrEqual 1 && lessThanOrEqual 2
    /* Rule Operators (Numeric Special) which are translated to standard numeric operators with calculated values */
//    const AUTOFILTER_COLUMN_RULE_TOPTEN                = "topTen";            //    greaterThan calculated value
//    const AUTOFILTER_COLUMN_RULE_TOPTENPERCENT        = "topTenPercent";    //    greaterThan calculated value
//    const AUTOFILTER_COLUMN_RULE_ABOVEAVERAGE        = "aboveAverage";    //    Value is calculated as the average
//    const AUTOFILTER_COLUMN_RULE_BELOWAVERAGE        = "belowAverage";    //    Value is calculated as the average
    /* Rule Operators (String) which are set as wild-carded values */
//    const AUTOFILTER_COLUMN_RULE_BEGINSWITH            = "beginsWith";            // A*
//    const AUTOFILTER_COLUMN_RULE_ENDSWITH            = "endsWith";            // *Z
//    const AUTOFILTER_COLUMN_RULE_CONTAINS            = "contains";            // *B*
//    const AUTOFILTER_COLUMN_RULE_DOESNTCONTAIN        = "notEqual";            //    notEqual *B*
    /* Rule Operators (Date Special) which are translated to standard numeric operators with calculated values */
//    const AUTOFILTER_COLUMN_RULE_BEFORE                = "lessThan";
//    const AUTOFILTER_COLUMN_RULE_AFTER                = "greaterThan";
//    const AUTOFILTER_COLUMN_RULE_YESTERDAY            = "yesterday";
//    const AUTOFILTER_COLUMN_RULE_TODAY                = "today";
//    const AUTOFILTER_COLUMN_RULE_TOMORROW            = "tomorrow";
//    const AUTOFILTER_COLUMN_RULE_LASTWEEK            = "lastWeek";
//    const AUTOFILTER_COLUMN_RULE_THISWEEK            = "thisWeek";
//    const AUTOFILTER_COLUMN_RULE_NEXTWEEK            = "nextWeek";
//    const AUTOFILTER_COLUMN_RULE_LASTMONTH            = "lastMonth";
//    const AUTOFILTER_COLUMN_RULE_THISMONTH            = "thisMonth";
//    const AUTOFILTER_COLUMN_RULE_NEXTMONTH            = "nextMonth";
//    const AUTOFILTER_COLUMN_RULE_LASTQUARTER        = "lastQuarter";
//    const AUTOFILTER_COLUMN_RULE_THISQUARTER        = "thisQuarter";
//    const AUTOFILTER_COLUMN_RULE_NEXTQUARTER        = "nextQuarter";
//    const AUTOFILTER_COLUMN_RULE_LASTYEAR            = "lastYear";
//    const AUTOFILTER_COLUMN_RULE_THISYEAR            = "thisYear";
//    const AUTOFILTER_COLUMN_RULE_NEXTYEAR            = "nextYear";
//    const AUTOFILTER_COLUMN_RULE_YEARTODATE            = "yearToDate";            //    <dynamicFilter val="40909" type="yearToDate" maxVal="41113"/>
//    const AUTOFILTER_COLUMN_RULE_ALLDATESINMONTH    = "allDatesInMonth";    //    <dynamicFilter type="M2"/> for Month/February
//    const AUTOFILTER_COLUMN_RULE_ALLDATESINQUARTER    = "allDatesInQuarter";    //    <dynamicFilter type="Q2"/> for Quarter 2

    private static _ruleTypes = [
        //    Currently we're not handling
        //        colorFilter
        //        extLst
        //        iconFilter
        self::AUTOFILTER_RULETYPE_FILTER,
        self::AUTOFILTER_RULETYPE_DATEGROUP,
        self::AUTOFILTER_RULETYPE_CUSTOMFILTER,
        self::AUTOFILTER_RULETYPE_DYNAMICFILTER
    ];

    private static _dateTimeGroups = [
        self::AUTOFILTER_RULETYPE_DATEGROUP_YEAR,
        self::AUTOFILTER_RULETYPE_DATEGROUP_MONTH,
        self::AUTOFILTER_RULETYPE_DATEGROUP_DAY,
        self::AUTOFILTER_RULETYPE_DATEGROUP_HOUR,
        self::AUTOFILTER_RULETYPE_DATEGROUP_MINUTE,
        self::AUTOFILTER_RULETYPE_DATEGROUP_SECOND
    ];

    private static _dynamicTypes = [
        self::AUTOFILTER_RULETYPE_DYNAMIC_YESTERDAY,
        self::AUTOFILTER_RULETYPE_DYNAMIC_TODAY,
        self::AUTOFILTER_RULETYPE_DYNAMIC_TOMORROW,
        self::AUTOFILTER_RULETYPE_DYNAMIC_YEARTODATE,
        self::AUTOFILTER_RULETYPE_DYNAMIC_THISYEAR,
        self::AUTOFILTER_RULETYPE_DYNAMIC_THISQUARTER,
        self::AUTOFILTER_RULETYPE_DYNAMIC_THISMONTH,
        self::AUTOFILTER_RULETYPE_DYNAMIC_THISWEEK,
        self::AUTOFILTER_RULETYPE_DYNAMIC_LASTYEAR,
        self::AUTOFILTER_RULETYPE_DYNAMIC_LASTQUARTER,
        self::AUTOFILTER_RULETYPE_DYNAMIC_LASTMONTH,
        self::AUTOFILTER_RULETYPE_DYNAMIC_LASTWEEK,
        self::AUTOFILTER_RULETYPE_DYNAMIC_NEXTYEAR,
        self::AUTOFILTER_RULETYPE_DYNAMIC_NEXTQUARTER,
        self::AUTOFILTER_RULETYPE_DYNAMIC_NEXTMONTH,
        self::AUTOFILTER_RULETYPE_DYNAMIC_NEXTWEEK,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_1,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_2,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_3,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_4,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_5,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_6,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_7,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_8,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_9,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_10,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_11,
        self::AUTOFILTER_RULETYPE_DYNAMIC_MONTH_12,
        self::AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_1,
        self::AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_2,
        self::AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_3,
        self::AUTOFILTER_RULETYPE_DYNAMIC_QUARTER_4,
        self::AUTOFILTER_RULETYPE_DYNAMIC_ABOVEAVERAGE,
        self::AUTOFILTER_RULETYPE_DYNAMIC_BELOWAVERAGE
    ];

    private static _operators = [
        self::AUTOFILTER_COLUMN_RULE_EQUAL,
        self::AUTOFILTER_COLUMN_RULE_NOTEQUAL,
        self::AUTOFILTER_COLUMN_RULE_GREATERTHAN,
        self::AUTOFILTER_COLUMN_RULE_GREATERTHANOREQUAL,
        self::AUTOFILTER_COLUMN_RULE_LESSTHAN,
        self::AUTOFILTER_COLUMN_RULE_LESSTHANOREQUAL
    ];

    private static _topTenValue = [
        self::AUTOFILTER_COLUMN_RULE_TOPTEN_BY_VALUE,
        self::AUTOFILTER_COLUMN_RULE_TOPTEN_PERCENT
    ];

    private static _topTenType = [
        self::AUTOFILTER_COLUMN_RULE_TOPTEN_TOP,
        self::AUTOFILTER_COLUMN_RULE_TOPTEN_BOTTOM
    ];
    
    /**
     * Autofilter Column
     *
     * @var PHPExcel_Worksheet_AutoFilter_Column
     */
    private _parent = NULL;


    /**
     * Autofilter Rule Type
     *
     * @var string
     */
    private _ruleType = self::AUTOFILTER_RULETYPE_FILTER;


    /**
     * Autofilter Rule Value
     *
     * @var string
     */
    private _value = "";

    /**
     * Autofilter Rule Operator
     *
     * @var string
     */
    private _operator = self::AUTOFILTER_COLUMN_RULE_EQUAL;

    /**
     * DateTimeGrouping Group Value
     *
     * @var string
     */
    private _grouping = "";


    /**
     * Create a new PHPExcel_Worksheet_AutoFilter_Column_Rule
     *
     * @param PHPExcel_Worksheet_AutoFilter_Column $pParent
     */
    public function __construct(<\Zxcel\Worksheet\AutoFilter\Column> pParent = null)
    {
        let this->_parent = pParent;
    }

    /**
     * Get AutoFilter Rule Type
     *
     * @return string
     */
    public function getRuleType() -> string
    {
        return this->_ruleType;
    }

    /**
     *    Set AutoFilter Rule Type
     *
     *    @param    string        $pRuleType
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet_AutoFilter_Column
     */
    public function setRuleType(string pRuleType = self::AUTOFILTER_RULETYPE_FILTER) -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        if (!in_array(pRuleType, self::_ruleTypes)) {
            throw new \ZExcel\Exception("Invalid rule type for column AutoFilter Rule.");
        }

        let this->_ruleType = pRuleType;

        return this;
    }

    /**
     * Get AutoFilter Rule Value
     *
     * @return string
     */
    public function getValue() -> string
    {
        return this->_value;
    }

    /**
     *    Set AutoFilter Rule Value
     *
     *    @param    string|string[]        $pValue
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet_AutoFilter_Column_Rule
     */
    public function setValue(pValue = "") -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        var key, grouping;
    
        if (is_array(pValue)) {
            let grouping = -1;
            for key, _ in pValue {
                //    Validate array entries
                if (!in_array(key, self::_dateTimeGroups)) {
                    //    Remove any invalid entries from the value array
                    unset(pValue[key]);
                } else {
                    //    Work out what the dateTime grouping will be
                    let grouping = max(grouping, array_search(key, self::_dateTimeGroups));
                }
            }
            if (count(pValue) == 0) {
                throw new \ZExcel\Exception("Invalid rule value for column AutoFilter Rule.");
            }
            //    Set the dateTime grouping that we've anticipated
            this->setGrouping(self::_dateTimeGroups[grouping]);
        }
        
        let this->_value = pValue;

        return this;
    }

    /**
     * Get AutoFilter Rule Operator
     *
     * @return string
     */
    public function getOperator() -> string
    {
        return this->_operator;
    }

    /**
     *    Set AutoFilter Rule Operator
     *
     *    @param    string        $pOperator
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet_AutoFilter_Column_Rule
     */
    public function setOperator(string pOperator = self::AUTOFILTER_COLUMN_RULE_EQUAL) -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        if (empty(pOperator)) {
            let pOperator = self::AUTOFILTER_COLUMN_RULE_EQUAL;
        }
        
        if ((!in_array(pOperator,self::_operators)) &&
            (!in_array(pOperator,self::_topTenValue))) {
            throw new \ZExcel\Exception("Invalid operator for column AutoFilter Rule.");
        }
        
        let this->_operator = pOperator;

        return this;
    }

    /**
     * Get AutoFilter Rule Grouping
     *
     * @return string
     */
    public function getGrouping() -> string
    {
        return this->_grouping;
    }

    /**
     *    Set AutoFilter Rule Grouping
     *
     *    @param    string        $pGrouping
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet_AutoFilter_Column_Rule
     */
    public function setGrouping(string pGrouping = null) -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        if ((pGrouping !== null) &&
            (!in_array(pGrouping,self::_dateTimeGroups)) &&
            (!in_array(pGrouping,self::_dynamicTypes)) &&
            (!in_array(pGrouping,self::_topTenType))) {
            throw new \ZExcel\Exception("Invalid rule type for column AutoFilter Rule.");
        }

        let this->_grouping = pGrouping;

        return this;
    }

    /**
     *    Set AutoFilter Rule
     *
     *    @param    string                $pOperator
     *    @param    string|string[]        $pValue
     *    @param    string                $pGrouping
     *    @throws    PHPExcel_Exception
     *    @return PHPExcel_Worksheet_AutoFilter_Column_Rule
     */
    public function setRule(string pOperator = self::AUTOFILTER_COLUMN_RULE_EQUAL, pValue = "", string! pGrouping = null) -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        this->setOperator(pOperator);
        this->setValue(pValue);
        //    Only set grouping if it's been passed in as a user-supplied argument,
        //        otherwise we're calculating it when we setValue() and don't want to overwrite that
        //        If the user supplies an argumnet for grouping, then on their own head be it
        if (pGrouping !== null) {
            this->setGrouping(pGrouping);
        }
        
        return this;
    }

    /**
     * Get this Rule's AutoFilter Column Parent
     *
     * @return PHPExcel_Worksheet_AutoFilter_Column
     */
    public function getParent() -> <ZExcel\Worksheet\AutoFilter\Column>
    {
        return this->_parent;
    }

    /**
     * Set this Rule's AutoFilter Column Parent
     *
     * @param PHPExcel_Worksheet_AutoFilter_Column
     * @return PHPExcel_Worksheet_AutoFilter_Column_Rule
     */
    public function setParent(<\ZExcel\Worksheet\AutoFilter\Column> pParent = null) -> <ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        let this->_parent = pParent;

        return this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        var key, value, vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                if (key == "_parent") {
                    //    Detach from autofilter column parent
                    let this->{key} = null;
                } else {
                    let this->{key} = clone value;
                }
            } else {
                let this->{key} = value;
            }
        }
    }
}
