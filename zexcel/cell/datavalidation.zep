namespace ZExcel\Cell;

class DataValidation
{
    /* Data validation types */
    const TYPE_NONE        = "none";
    const TYPE_CUSTOM      = "custom";
    const TYPE_DATE        = "date";
    const TYPE_DECIMAL     = "decimal";
    const TYPE_LIST        = "list";
    const TYPE_TEXTLENGTH  = "textLength";
    const TYPE_TIME        = "time";
    const TYPE_WHOLE       = "whole";

    /* Data validation error styles */
    const STYLE_STOP         = "stop";
    const STYLE_WARNING      = "warning";
    const STYLE_INFORMATION  = "information";

    /* Data validation operators */
    const OPERATOR_BETWEEN             = "between";
    const OPERATOR_EQUAL               = "equal";
    const OPERATOR_GREATERTHAN         = "greaterThan";
    const OPERATOR_GREATERTHANOREQUAL  = "greaterThanOrEqual";
    const OPERATOR_LESSTHAN            = "lessThan";
    const OPERATOR_LESSTHANOREQUAL     = "lessThanOrEqual";
    const OPERATOR_NOTBETWEEN          = "notBetween";
    const OPERATOR_NOTEQUAL            = "notEqual";

    /**
     * Formula 1
     *
     * @var string
     */
    private formula1;

    /**
     * Formula 2
     *
     * @var string
     */
    private formula2;

    /**
     * Type
     *
     * @var string
     */
    private type = \ZExcel\Cell\DataValidation::TYPE_NONE;

    /**
     * Error style
     *
     * @var string
     */
    private errorStyle = \ZExcel\Cell\DataValidation::STYLE_STOP;

    /**
     * Operator
     *
     * @var string
     */
    private operator;

    /**
     * Allow Blank
     *
     * @var boolean
     */
    private allowBlank;

    /**
     * Show DropDown
     *
     * @var boolean
     */
    private showDropDown;

    /**
     * Show InputMessage
     *
     * @var boolean
     */
    private showInputMessage;

    /**
     * Show ErrorMessage
     *
     * @var boolean
     */
    private showErrorMessage;

    /**
     * Error title
     *
     * @var string
     */
    private errorTitle;

    /**
     * Error
     *
     * @var string
     */
    private error;

    /**
     * Prompt title
     *
     * @var string
     */
    private promptTitle;

    /**
     * Prompt
     *
     * @var string
     */
    private prompt;

    /**
     * Create a new \ZExcel\Cell\DataValidation
     */
    public function __construct()
    {
        // Initialise member variables
        let this->formula1          = "";
        let this->formula2          = "";
        let this->type              = \ZExcel\Cell\DataValidation::TYPE_NONE;
        let this->errorStyle        = \ZExcel\Cell\DataValidation::STYLE_STOP;
        let this->operator          = "";
        let this->allowBlank        = false;
        let this->showDropDown      = false;
        let this->showInputMessage  = false;
        let this->showErrorMessage  = false;
        let this->errorTitle        = "";
        let this->error             = "";
        let this->promptTitle       = "";
        let this->prompt            = "";
    }

    /**
     * Get Formula 1
     *
     * @return string
     */
    public function getFormula1()
    {
        return this->formula1;
    }

    /**
     * Set Formula 1
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setFormula1(var value = "")
    {
        let this->formula1 = value;
        return this;
    }

    /**
     * Get Formula 2
     *
     * @return string
     */
    public function getFormula2()
    {
        return this->formula2;
    }

    /**
     * Set Formula 2
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setFormula2(var value = "")
    {
        let this->formula2 = value;
        return this;
    }

    /**
     * Get Type
     *
     * @return string
     */
    public function getType()
    {
        return this->type;
    }

    /**
     * Set Type
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setType(var value = \ZExcel\Cell\DataValidation::TYPE_NONE)
    {
        let this->type = value;
        return this;
    }

    /**
     * Get Error style
     *
     * @return string
     */
    public function getErrorStyle()
    {
        return this->errorStyle;
    }

    /**
     * Set Error style
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setErrorStyle(var value = \ZExcel\Cell\DataValidation::STYLE_STOP)
    {
        let this->errorStyle = value;
        return this;
    }

    /**
     * Get Operator
     *
     * @return string
     */
    public function getOperator()
    {
        return this->operator;
    }

    /**
     * Set Operator
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setOperator(var value = "")
    {
        let this->operator = value;
        return this;
    }

    /**
     * Get Allow Blank
     *
     * @return boolean
     */
    public function getAllowBlank()
    {
        return this->allowBlank;
    }

    /**
     * Set Allow Blank
     *
     * @param  boolean    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setAllowBlank(var value = false)
    {
        let this->allowBlank = value;
        return this;
    }

    /**
     * Get Show DropDown
     *
     * @return boolean
     */
    public function getShowDropDown()
    {
        return this->showDropDown;
    }

    /**
     * Set Show DropDown
     *
     * @param  boolean    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setShowDropDown(var value = false)
    {
        let this->showDropDown = value;
        return this;
    }

    /**
     * Get Show InputMessage
     *
     * @return boolean
     */
    public function getShowInputMessage()
    {
        return this->showInputMessage;
    }

    /**
     * Set Show InputMessage
     *
     * @param  boolean    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setShowInputMessage(var value = false)
    {
        let this->showInputMessage = value;
        return this;
    }

    /**
     * Get Show ErrorMessage
     *
     * @return boolean
     */
    public function getShowErrorMessage()
    {
        return this->showErrorMessage;
    }

    /**
     * Set Show ErrorMessage
     *
     * @param  boolean    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setShowErrorMessage(var value = false)
    {
        let this->showErrorMessage = value;
        return this;
    }

    /**
     * Get Error title
     *
     * @return string
     */
    public function getErrorTitle()
    {
        return this->errorTitle;
    }

    /**
     * Set Error title
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setErrorTitle(var value = "")
    {
        let this->errorTitle = value;
        return this;
    }

    /**
     * Get Error
     *
     * @return string
     */
    public function getError()
    {
        return this->error;
    }

    /**
     * Set Error
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setError(var value = "")
    {
        let this->error = value;
        return this;
    }

    /**
     * Get Prompt title
     *
     * @return string
     */
    public function getPromptTitle()
    {
        return this->promptTitle;
    }

    /**
     * Set Prompt title
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setPromptTitle(var value = "")
    {
        let this->promptTitle = value;
        return this;
    }

    /**
     * Get Prompt
     *
     * @return string
     */
    public function getPrompt()
    {
        return this->prompt;
    }

    /**
     * Set Prompt
     *
     * @param  string    value
     * @return \ZExcel\Cell\DataValidation
     */
    public function setPrompt(var value = "")
    {
        let this->prompt = value;
        return this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        let this->type = \ZExcel\Cell\DataValidation::TYPE_NONE;
        let this->errorStyle = \ZExcel\Cell\DataValidation::STYLE_STOP;
        
        return md5(
            this->formula1 .
            this->formula2 .
            this->type .
            this->errorStyle .
            this->operator .
            (this->allowBlank ? "t" : "f") .
            (this->showDropDown ? "t" : "f") .
            (this->showInputMessage ? "t" : "f") .
            (this->showErrorMessage ? "t" : "f") .
            this->errorTitle .
            this->error .
            this->promptTitle .
            this->prompt .
            get_class(this)
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
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
