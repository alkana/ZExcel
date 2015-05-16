namespace ZExcel\Style;

use ZExcel\IComparable as ZIComparable;

class Conditional implements ZIComparable
{
	/* Condition types */
    const CONDITION_NONE         = "none";
    const CONDITION_CELLIS       = "cellIs";
    const CONDITION_CONTAINSTEXT = "containsText";
    const CONDITION_EXPRESSION   = "expression";

    /* Operator types */
    const OPERATOR_NONE               = "";
    const OPERATOR_BEGINSWITH         = "beginsWith";
    const OPERATOR_ENDSWITH           = "endsWith";
    const OPERATOR_EQUAL              = "equal";
    const OPERATOR_GREATERTHAN        = "greaterThan";
    const OPERATOR_GREATERTHANOREQUAL = "greaterThanOrEqual";
    const OPERATOR_LESSTHAN           = "lessThan";
    const OPERATOR_LESSTHANOREQUAL    = "lessThanOrEqual";
    const OPERATOR_NOTEQUAL           = "notEqual";
    const OPERATOR_CONTAINSTEXT       = "containsText";
    const OPERATOR_NOTCONTAINS        = "notContains";
    const OPERATOR_BETWEEN            = "between";

    /**
     * Condition type
     *
     * @var int
     */
    private conditionType;

    /**
     * Operator type
     *
     * @var int
     */
    private operatorType;

    /**
     * Text
     *
     * @var string
     */
    private text;

    /**
     * Condition
     *
     * @var string[]
     */
    private condition = [];

    /**
     * Style
     *
     * @var \ZExcel\Style
     */
    private style;

    public function __construct()
    {
        // Initialise values
        let this->conditionType = \ZExcel\Style\Conditional::CONDITION_NONE;
        let this->operatorType  = \ZExcel\Style\Conditional::OPERATOR_NONE;
        let this->text          = null;
        let this->condition     = [];
        let this->style         = new \ZExcel\Style(false, true);
    }

    public function getConditionType()
    {
        return this->conditionType;
    }

    public function setConditionType(string pValue = \ZExcel\Style\Conditional::CONDITION_NONE)
    {
        let this->conditionType = pValue;
        
        return this;
    }

    public function getOperatorType()
    {
        return this->operatorType;
    }

    public function setOperatorType(string pValue = \ZExcel\Style\Conditional::OPERATOR_NONE)
    {
        let this->operatorType = pValue;
        
        return this;
    }

    public function getText()
    {
        return this->text;
    }

    public function setText(string value = null)
    {
        let this->text = value;
        
        return this;
    }

    public function getCondition()
    {
        if (isset(this->condition[0])) {
            return this->condition[0];
        }

        return "";
    }

    public function setCondition(var pValue = "")
    {
        if (!is_array(pValue)) {
            let pValue = [pValue];
        }

        return this->setConditions(pValue);
    }

    public function getConditions()
    {
        return this->condition;
    }

    public function setConditions(pValue)
    {
        if (!is_array(pValue)) {
            let pValue = [pValue];
        }
        
        let this->condition = pValue;
        
        return this;
    }

    public function addCondition(pValue = "")
    {
        let this->condition[] = pValue;
        return this;
    }

    public function getStyle()
    {
        return this->style;
    }

    public function setStyle(<\ZExcel\Style> pValue = null)
    {
       let this->style = pValue;
       return this;
    }

    public function getHashCode()
    {
        return md5(
            this->conditionType .
            this->operatorType .
            implode(";", this->condition) .
            this->style->getHashCode() .
            __CLASS__
        );
    }

    public function __clone()
    {
    	var vars, key, value;
    	
        let vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                let this->{key} = clone value;
            } else {
                let this->{key} = value;
            }
        }
    }
}
