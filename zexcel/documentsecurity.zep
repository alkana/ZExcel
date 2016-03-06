namespace ZExcel;

class DocumentSecurity
{
    /**
     * LockRevision
     *
     * @var boolean
     */
    private _lockRevision;

    /**
     * LockStructure
     *
     * @var boolean
     */
    private _lockStructure;

    /**
     * LockWindows
     *
     * @var boolean
     */
    private _lockWindows;

    /**
     * RevisionsPassword
     *
     * @var string
     */
    private _revisionsPassword;

    /**
     * WorkbookPassword
     *
     * @var string
     */
    private _workbookPassword;

    /**
     * Create a new \ZExcel\DocumentSecurity
     */
    public function __construct()
    {
        // Initialise values
        let this->_lockRevision      = false;
        let this->_lockStructure     = false;
        let this->_lockWindows       = false;
        let this->_revisionsPassword = "";
        let this->_workbookPassword  = "";
    }

    /**
     * Is some sort of document security enabled?
     *
     * @return boolean
     */
    public function isSecurityEnabled() -> boolean
    {
        return  this->_lockRevision || this->_lockStructure || this->_lockWindows;
    }

    /**
     * Get LockRevision
     *
     * @return boolean
     */
    public function getLockRevision()
    {
        return this->_lockRevision;
    }

    /**
     * Set LockRevision
     *
     * @param boolean pValue
     * @return \ZExcel\DocumentSecurity
     */
    public function setLockRevision(var pValue = false)
    {
        let this->_lockRevision = pValue;
        
        return this;
    }

    /**
     * Get LockStructure
     *
     * @return boolean
     */
    public function getLockStructure()
    {
        return this->_lockStructure;
    }

    /**
     * Set LockStructure
     *
     * @param boolean pValue
     * @return \ZExcel\DocumentSecurity
     */
    public function setLockStructure(var pValue = false) -> <\ZExcel\DocumentSecurity>
    {
        let this->_lockStructure = pValue;
        
        return this;
    }

    /**
     * Get LockWindows
     *
     * @return boolean
     */
    public function getLockWindows()
    {
        return this->_lockWindows;
    }

    /**
     * Set LockWindows
     *
     * @param boolean pValue
     * @return \ZExcel\DocumentSecurity
     */
    public function setLockWindows(var pValue = false) -> <\ZExcel\DocumentSecurity>
    {
        let this->_lockWindows = pValue;
        
        return this;
    }

    /**
     * Get RevisionsPassword (hashed)
     *
     * @return string
     */
    public function getRevisionsPassword()
    {
        return this->_revisionsPassword;
    }

    /**
     * Set RevisionsPassword
     *
     * @param string     pValue
     * @param boolean     pAlreadyHashed If the password has already been hashed, set this to true
     * @return \ZExcel\DocumentSecurity
     */
    public function setRevisionsPassword(pValue = "", pAlreadyHashed = false) -> <\ZExcel\DocumentSecurity>
    {
        if (!pAlreadyHashed) {
            let pValue = \ZExcel\Shared\PasswordHasher::hashPassword(pValue);
        }
        
        let this->_revisionsPassword = pValue;
        
        return this;
    }

    /**
     * Get WorkbookPassword (hashed)
     *
     * @return string
     */
    public function getWorkbookPassword()
    {
        return this->_workbookPassword;
    }

    /**
     * Set WorkbookPassword
     *
     * @param string     pValue
     * @param boolean     pAlreadyHashed If the password has already been hashed, set this to true
     * @return \ZExcel\DocumentSecurity
     */
    public function setWorkbookPassword(var pValue = "", var pAlreadyHashed = false) -> <\ZExcel\DocumentSecurity>
    {
        if (!pAlreadyHashed) {
            let pValue = \ZExcel\Shared\PasswordHasher::hashPassword(pValue);
        }
        
        let this->_workbookPassword = pValue;
        
        return this;
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
