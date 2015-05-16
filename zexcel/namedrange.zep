namespace ZExcel;

class NamedRange
{
	/**
     * Range name
     *
     * @var string
     */
    private _name;

    /**
     * Worksheet on which the named range can be resolved
     *
     * @var PHPExcel_Worksheet
     */
    private _worksheet;

    /**
     * Range of the referenced cells
     *
     * @var string
     */
    private _range;

    /**
     * Is the named range local? (i.e. can only be used on this->_worksheet)
     *
     * @var bool
     */
    private _localOnly;

    /**
     * Scope
     *
     * @var PHPExcel_Worksheet
     */
    private _scope;

    /**
     * Create a new NamedRange
     *
     * @param string pName
     * @param PHPExcel_Worksheet pWorksheet
     * @param string pRange
     * @param bool pLocalOnly
     * @param PHPExcel_Worksheet|null pScope    Scope. Only applies when pLocalOnly = true. Null for global scope.
     * @throws PHPExcel_Exception
     */
    public function __construct(pName = null, <\ZExcel\Worksheet> pWorksheet, pRange = "A1", boolean pLocalOnly = false, pScope = null)
    {
        // Validate data
        if ((pName === null) || (pWorksheet === null) || (pRange === null)) {
            throw new \ZExcel\Exception("Parameters can not be null.");
        }

        // Set local members
        let this->_name = pName;
        let this->_worksheet = pWorksheet;
        let this->_range = pRange;
        let this->_localOnly = pLocalOnly;
        let this->_scope = (pLocalOnly == true) ? ((pScope == null) ? pWorksheet : pScope) : null;
    }

    /**
     * Get name
     *
     * @return string
     */
    public function getName() {
        return this->_name;
    }

    /**
     * Set name
     *
     * @param string value
     * @return PHPExcel_NamedRange
     */
    public function setName(value = null) {
        var oldTitle, newTitle;
    
        if (value !== null) {
            // Old title
            let oldTitle = this->_name;

            // Re-attach
            if (this->_worksheet !== null) {
                this->_worksheet->getParent()->removeNamedRange(this->_name,this->_worksheet);
            }
            let this->_name = value;

            if (this->_worksheet !== null) {
                this->_worksheet->getParent()->addNamedRange(this);
            }

            // New title
            let newTitle = this->_name;
            \ZExcel\ReferenceHelper::getInstance()
            	->updateNamedFormulas(this->_worksheet->getParent(), oldTitle, newTitle);
        }
        
        return this;
    }

    /**
     * Get worksheet
     *
     * @return PHPExcel_Worksheet
     */
    public function getWorksheet() {
        return this->_worksheet;
    }

    /**
     * Set worksheet
     *
     * @param PHPExcel_Worksheet value
     * @return PHPExcel_NamedRange
     */
    public function setWorksheet(<\ZExcel\Worksheet> value = null) {
        if (value !== null) {
            let this->_worksheet = value;
        }
        
        return this;
    }

    /**
     * Get range
     *
     * @return string
     */
    public function getRange() {
        return this->_range;
    }

    /**
     * Set range
     *
     * @param string value
     * @return PHPExcel_NamedRange
     */
    public function setRange(value = null) {
        if (value !== NULL) {
            let this->_range = value;
        }
        
        return this;
    }

    /**
     * Get localOnly
     *
     * @return bool
     */
    public function getLocalOnly() {
        return this->_localOnly;
    }

    /**
     * Set localOnly
     *
     * @param bool value
     * @return PHPExcel_NamedRange
     */
    public function setLocalOnly(boolean value = false) {
        let this->_localOnly = value;
        let this->_scope = value ? this->_worksheet : null;
        
        return this;
    }

    /**
     * Get scope
     *
     * @return PHPExcel_Worksheet|null
     */
    public function getScope() {
        return this->_scope;
    }

    /**
     * Set scope
     *
     * @param PHPExcel_Worksheet|null value
     * @return PHPExcel_NamedRange
     */
    public function setScope(<\ZExcel\Worksheet> value = null) {
        let this->_scope = value;
        let this->_localOnly = (value == null) ? false : true;
        
        return this;
    }

    /**
     * Resolve a named range to a regular cell range
     *
     * @param string pNamedRange Named range
     * @param PHPExcel_Worksheet|null pSheet Scope. Use null for global scope
     * @return PHPExcel_NamedRange
     */
    public static function resolveRange(pNamedRange = "", <\ZExcel\Worksheet> pSheet) {
        return pSheet->getParent()->getNamedRange(pNamedRange, pSheet);
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
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
