namespace ZExcel\Worksheet\AutoFilter;

class Column
{
    const AUTOFILTER_FILTERTYPE_FILTER          = "filters";
    const AUTOFILTER_FILTERTYPE_CUSTOMFILTER    = "customFilters";
    //    Supports no more than 2 rules, with an And/Or join criteria
    //        if more than 1 rule is defined
    const AUTOFILTER_FILTERTYPE_DYNAMICFILTER   = "dynamicFilter";
    //    Even though the filter rule is constant, the filtered data can vary
    //        e.g. filtered by date = TODAY
    const AUTOFILTER_FILTERTYPE_TOPTENFILTER    = "top10";

    /* Multiple Rule Connections */
    const AUTOFILTER_COLUMN_JOIN_AND    = "and";
    const AUTOFILTER_COLUMN_JOIN_OR     = "or";

    /**
     * Types of autofilter rules
     *
     * @var string[]
     */
    private static _filterTypes = [
        //    Currently we're not handling
        //        colorFilter
        //        extLst
        //        iconFilter
        self::AUTOFILTER_FILTERTYPE_FILTER,
        self::AUTOFILTER_FILTERTYPE_CUSTOMFILTER,
        self::AUTOFILTER_FILTERTYPE_DYNAMICFILTER,
        self::AUTOFILTER_FILTERTYPE_TOPTENFILTER
    ];

    /**
     * Join options for autofilter rules
     *
     * @var string[]
     */
    private static _ruleJoins = [
        self::AUTOFILTER_COLUMN_JOIN_AND,
        self::AUTOFILTER_COLUMN_JOIN_OR
    ];

    /**
     * Autofilter
     *
     * @var \ZExcel\Worksheet\AutoFilter
     */
    private _parent = NULL;


    /**
     * Autofilter Column Index
     *
     * @var string
     */
    private _columnIndex = "";


    /**
     * Autofilter Column Filter Type
     *
     * @var string
     */
    private _filterType = self::AUTOFILTER_FILTERTYPE_FILTER;


    /**
     * Autofilter Multiple Rules And/Or
     *
     * @var string
     */
    private _join = self::AUTOFILTER_COLUMN_JOIN_OR;


    /**
     * Autofilter Column Rules
     *
     * @var array of \ZExcel\Worksheet\AutoFilter\Column\Rule
     */
    private _ruleset = [];


    /**
     * Autofilter Column Dynamic Attributes
     *
     * @var array of mixed
     */
    private _attributes = [];


    /**
     * Create a new \ZExcel\Worksheet\AutoFilter\Column
     *
     *    @param    string                           $pColumn        Column (e.g. A)
     *    @param    \ZExcel\Worksheet\AutoFilter  $pParent        Autofilter for this column
     */
    public function __construct(string pColumn, <\ZExcel\Worksheet\AutoFilter> pParent = null)
    {
        let this->_columnIndex = pColumn;
        let this->_parent = pParent;
    }

    /**
     * Get AutoFilter Column Index
     *
     * @return string
     */
    public function getColumnIndex() -> string
    {
        return this->_columnIndex;
    }

    /**
     *    Set AutoFilter Column Index
     *
     *    @param    string        $pColumn        Column (e.g. A)
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setColumnIndex(string pColumn) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        // Uppercase coordinate
        let pColumn = strtoupper(pColumn);
        
        if (this->_parent !== null) {
            this->_parent->testColumnInRange(pColumn);
        }

        let this->_columnIndex = pColumn;

        return this;
    }

    /**
     * Get this Column's AutoFilter Parent
     *
     * @return \ZExcel\Worksheet\AutoFilter
     */
    public function getParent() -> <\ZExcel\Worksheet\AutoFilter>
    {
        return this->_parent;
    }

    /**
     * Set this Column's AutoFilter Parent
     *
     * @param \ZExcel\Worksheet\AutoFilter
     * @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setParent(<\ZExcel\Worksheet\AutoFilter> pParent = null) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        let this->_parent = pParent;

        return this;
    }

    /**
     * Get AutoFilter Type
     *
     * @return string
     */
    public function getFilterType() -> string
    {
        return this->_filterType;
    }

    /**
     *    Set AutoFilter Type
     *
     *    @param    string        $pFilterType
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setFilterType(string pFilterType = self::AUTOFILTER_FILTERTYPE_FILTER) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        if (!in_array(pFilterType, self::_filterTypes)) {
            throw new \ZExcel\Exception("Invalid filter type for column AutoFilter.");
        }

        let this->_filterType = pFilterType;

        return this;
    }

    /**
     * Get AutoFilter Multiple Rules And/Or Join
     *
     * @return string
     */
    public function getJoin() -> string
    {
        return this->_join;
    }

    /**
     *    Set AutoFilter Multiple Rules And/Or
     *
     *    @param    string        $pJoin        And/Or
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setJoin(string pJoin = self::AUTOFILTER_COLUMN_JOIN_OR) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        // Lowercase And/Or
        let pJoin = strtolower(pJoin);
        if (!in_array(pJoin, self::_ruleJoins)) {
            throw new \ZExcel\Exception("Invalid rule connection for column AutoFilter.");
        }

        let this->_join = $pJoin;

        return this;
    }

    /**
     *    Set AutoFilter Attributes
     *
     *    @param    string[]        $pAttributes
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setAttributes(array pAttributes = []) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        let this->_attributes = pAttributes;

        return this;
    }

    /**
     *    Set An AutoFilter Attribute
     *
     *    @param    string        $pName        Attribute Name
     *    @param    string        $pValue        Attribute Value
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function setAttribute(string pName, string pValue) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        let this->_attributes[pName] = pValue;

        return this;
    }

    /**
     * Get AutoFilter Column Attributes
     *
     * @return string
     */
    public function getAttributes() {
        return this->_attributes;
    }

    /**
     * Get specific AutoFilter Column Attribute
     *
     *    @param    string        $pName        Attribute Name
     * @return string
     */
    public function getAttribute(string pName)
    {
        if (isset(this->_attributes[pName])) {
            return this->_attributes[pName];
        }
        
        return null;
    }

    /**
     * Get all AutoFilter Column Rules
     *
     * @throws    \ZExcel\Exception
     * @return array of \ZExcel\Worksheet\AutoFilter\Column\Rule
     */
    public function getRules() -> array
    {
        return this->_ruleset;
    }

    /**
     * Get a specified AutoFilter Column Rule
     *
     * @param    integer    $pIndex        Rule index in the ruleset array
     * @return    \ZExcel\Worksheet\AutoFilter\Column\Rule
     */
    public function getRule(int pIndex) -> <\ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        if (!isset(this->_ruleset[pIndex])) {
            let this->_ruleset[pIndex] = new \ZExcel\Worksheet\AutoFilter\Column\Rule(this);
        }
        return this->_ruleset[pIndex];
    }

    /**
     * Create a new AutoFilter Column Rule in the ruleset
     *
     * @return    \ZExcel\Worksheet\AutoFilter\Column\Rule
     */
    public function createRule() -> <\ZExcel\Worksheet\AutoFilter\Column\Rule>
    {
        let this->_ruleset[] = new \ZExcel\Worksheet\AutoFilter\Column\Rule(this);

        return end(this->_ruleset);
    }

    /**
     * Add a new AutoFilter Column Rule to the ruleset
     *
     * @param    \ZExcel\Worksheet\AutoFilter\Column\Rule    $pRule
     * @param    boolean    $returnRule     Flag indicating whether the rule object or the column object should be returned
     * @return    \ZExcel\Worksheet\AutoFilter\Column|\ZExcel\Worksheet\AutoFilter\Column\Rule
     */
    public function addRule(<\ZExcel\Worksheet\AutoFilter\Column\Rule> pRule, boolean returnRule = true)
    {
        pRule->setParent(this);
        let this->_ruleset[] = pRule;

        return (returnRule) ? pRule : this;
    }

    /**
     * Delete a specified AutoFilter Column Rule
     *    If the number of rules is reduced to 1, then we reset And/Or logic to Or
     *
     * @param    integer    $pIndex        Rule index in the ruleset array
     * @return    \ZExcel\Worksheet\AutoFilter\Column
     */
    public function deleteRule(int pIndex) -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        if (isset(this->_ruleset[pIndex])) {
            unset(this->_ruleset[pIndex]);
            //    If we've just deleted down to a single rule, then reset And/Or joining to Or
            if (count(this->_ruleset) <= 1) {
                this->setJoin(self::AUTOFILTER_COLUMN_JOIN_OR);
            }
        }

        return this;
    }

    /**
     * Delete all AutoFilter Column Rules
     *
     * @return    \ZExcel\Worksheet\AutoFilter\Column
     */
    public function clearRules() -> <\ZExcel\Worksheet\AutoFilter\Column>
    {
        let this->_ruleset = [];
        
        this->setJoin(self::AUTOFILTER_COLUMN_JOIN_OR);

        return this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        var key, value, k, v, aTmp, vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                if (key == "_parent") {
                    //    Detach from autofilter parent
                    let this->{key} = null;
                } else {
                    let this->{key} = clone value;
                }
            } elseif ((is_array(value)) && (key == "_ruleset")) {
                //    The columns array of \ZExcel\Worksheet\AutoFilter objects
                let aTmp = [];
                
                for k, v in value {
                    let aTmp[k] = clone v;
                    // attach the new cloned Rule to this new cloned Autofilter Cloned object
                    aTmp[k]->setParent(this);
                }
                
                let this->{key} = aTmp;
            } else {
                let this->{key} = value;
            }
        }
    }
}
