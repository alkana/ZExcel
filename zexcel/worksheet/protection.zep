namespace ZExcel\Worksheet;

class Protection
{
    /**
     * Sheet
     *
     * @var boolean
     */
    private sheet = false;

    /**
     * Objects
     *
     * @var boolean
     */
    private objects = false;

    /**
     * Scenarios
     *
     * @var boolean
     */
    private scenarios = false;

    /**
     * Format cells
     *
     * @var boolean
     */
    private formatCells = false;

    /**
     * Format columns
     *
     * @var boolean
     */
    private formatColumns = false;

    /**
     * Format rows
     *
     * @var boolean
     */
    private formatRows = false;

    /**
     * Insert columns
     *
     * @var boolean
     */
    private insertColumns = false;

    /**
     * Insert rows
     *
     * @var boolean
     */
    private insertRows = false;

    /**
     * Insert hyperlinks
     *
     * @var boolean
     */
    private insertHyperlinks = false;

    /**
     * Delete columns
     *
     * @var boolean
     */
    private deleteColumns = false;

    /**
     * Delete rows
     *
     * @var boolean
     */
    private deleteRows = false;

    /**
     * Select locked cells
     *
     * @var boolean
     */
    private selectLockedCells = false;

    /**
     * Sort
     *
     * @var boolean
     */
    private sort = false;

    /**
     * AutoFilter
     *
     * @var boolean
     */
    private autoFilter = false;

    /**
     * Pivot tables
     *
     * @var boolean
     */
    private pivotTables = false;

    /**
     * Select unlocked cells
     *
     * @var boolean
     */
    private selectUnlockedCells = false;

    /**
     * Password
     *
     * @var string
     */
    private password = "";

    /**
     * Create a new \ZExcel\Worksheet\Protection
     */
    public function __construct()
    {
    }

    /**
     * Is some sort of protection enabled?
     *
     * @return boolean
     */
    public function isProtectionEnabled()
    {
        return this->sheet
            || this->objects
            || this->scenarios
            || this->formatCells
            || this->formatColumns
            || this->formatRows
            || this->insertColumns
            || this->insertRows
            || this->insertHyperlinks
            || this->deleteColumns
            || this->deleteRows
            || this->selectLockedCells
            || this->sort
            || this->autoFilter
            || this->pivotTables
            || this->selectUnlockedCells;
    }

    /**
     * Get Sheet
     *
     * @return boolean
     */
    public function getSheet()
    {
        return this->sheet;
    }

    /**
     * Set Sheet
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setSheet(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->sheet = pValue;
        
        return this;
    }

    /**
     * Get Objects
     *
     * @return boolean
     */
    public function getObjects()
    {
        return this->objects;
    }

    /**
     * Set Objects
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setObjects(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->objects = pValue;
        
        return this;
    }

    /**
     * Get Scenarios
     *
     * @return boolean
     */
    public function getScenarios()
    {
        return this->scenarios;
    }

    /**
     * Set Scenarios
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setScenarios(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->scenarios = pValue;
        
        return this;
    }

    /**
     * Get FormatCells
     *
     * @return boolean
     */
    public function getFormatCells()
    {
        return this->formatCells;
    }

    /**
     * Set FormatCells
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setFormatCells(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->formatCells = pValue;
        
        return this;
    }

    /**
     * Get FormatColumns
     *
     * @return boolean
     */
    public function getFormatColumns()
    {
        return this->formatColumns;
    }

    /**
     * Set FormatColumns
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setFormatColumns(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->formatColumns = pValue;
        
        return this;
    }

    /**
     * Get FormatRows
     *
     * @return boolean
     */
    public function getFormatRows()
    {
        return this->formatRows;
    }

    /**
     * Set FormatRows
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setFormatRows(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->formatRows = pValue;
        
        return this;
    }

    /**
     * Get InsertColumns
     *
     * @return boolean
     */
    public function getInsertColumns()
    {
        return this->insertColumns;
    }

    /**
     * Set InsertColumns
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setInsertColumns(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->insertColumns = pValue;
        
        return this;
    }

    /**
     * Get InsertRows
     *
     * @return boolean
     */
    public function getInsertRows()
    {
        return this->insertRows;
    }

    /**
     * Set InsertRows
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setInsertRows(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->insertRows = pValue;
        
        return this;
    }

    /**
     * Get InsertHyperlinks
     *
     * @return boolean
     */
    public function getInsertHyperlinks()
    {
        return this->insertHyperlinks;
    }

    /**
     * Set InsertHyperlinks
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setInsertHyperlinks(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->insertHyperlinks = pValue;
        
        return this;
    }

    /**
     * Get DeleteColumns
     *
     * @return boolean
     */
    public function getDeleteColumns()
    {
        return this->deleteColumns;
    }

    /**
     * Set DeleteColumns
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setDeleteColumns(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->deleteColumns = pValue;
        
        return this;
    }

    /**
     * Get DeleteRows
     *
     * @return boolean
     */
    public function getDeleteRows()
    {
        return this->deleteRows;
    }

    /**
     * Set DeleteRows
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setDeleteRows(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->deleteRows = pValue;
        
        return this;
    }

    /**
     * Get SelectLockedCells
     *
     * @return boolean
     */
    public function getSelectLockedCells()
    {
        return this->selectLockedCells;
    }

    /**
     * Set SelectLockedCells
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setSelectLockedCells(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->selectLockedCells = pValue;
        
        return this;
    }

    /**
     * Get Sort
     *
     * @return boolean
     */
    public function getSort()
    {
        return this->sort;
    }

    /**
     * Set Sort
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setSort(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->sort = pValue;
        
        return this;
    }

    /**
     * Get AutoFilter
     *
     * @return boolean
     */
    public function getAutoFilter()
    {
        return this->autoFilter;
    }

    /**
     * Set AutoFilter
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setAutoFilter(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->autoFilter = pValue;
        
        return this;
    }

    /**
     * Get PivotTables
     *
     * @return boolean
     */
    public function getPivotTables()
    {
        return this->pivotTables;
    }

    /**
     * Set PivotTables
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setPivotTables(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->pivotTables = pValue;
        
        return this;
    }

    /**
     * Get SelectUnlockedCells
     *
     * @return boolean
     */
    public function getSelectUnlockedCells()
    {
        return this->selectUnlockedCells;
    }

    /**
     * Set SelectUnlockedCells
     *
     * @param boolean pValue
     * @return \ZExcel\Worksheet\Protection
     */
    public function setSelectUnlockedCells(boolean pValue = false) -> <\ZExcel\Worksheet\Protection>
    {
        let this->selectUnlockedCells = pValue;
        
        return this;
    }

    /**
     * Get Password (hashed)
     *
     * @return string
     */
    public function getPassword()
    {
        return this->password;
    }

    /**
     * Set Password
     *
     * @param string     pValue
     * @param boolean     pAlreadyHashed If the password has already been hashed, set this to true
     * @return \ZExcel\Worksheet\Protection
     */
    public function setPassword(string pValue = "", boolean pAlreadyHashed = false) -> <\ZExcel\Worksheet\Protection>
    {
        if (!pAlreadyHashed) {
            let pValue = \ZExcel\Shared\PasswordHasher::hashPassword(pValue);
        }
        
        let this->password = pValue;
        
        return this;
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
