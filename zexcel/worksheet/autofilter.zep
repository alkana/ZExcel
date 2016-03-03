namespace ZExcel\Worksheet;

class AutoFilter
{
    /**
     *    Search/Replace arrays to convert Excel wildcard syntax to a regexp syntax for preg_matching
     *
     *    @var    array
     */
    private static _toReplace   = [".*", ".",  "~",  "\*",  "\?"];
    private static _fromReplace = ["\*", "\?", "~~", "~.*", "~.?"];
    
    /**
     * Autofilter Worksheet
     *
     * @var \ZExcel\Worksheet
     */
    private _workSheet = null;


    /**
     * Autofilter Range
     *
     * @var string
     */
    private _range = "";


    /**
     * Autofilter Column Ruleset
     *
     * @var array of \ZExcel\Worksheet\AutoFilter\Column
     */
    private _columns = [];


    /**
     * Create a new \ZExcel\Worksheet\AutoFilter
     *
     *    @param    string        pRange        Cell range (i.e. A1:E10)
     * @param \ZExcel\Worksheet pSheet
     */
    public function __construct(string pRange = "", <\ZExcel\Worksheet> pSheet = null)
    {
        let this->_range = pRange;
        let this->_workSheet = pSheet;
    }

    /**
     * Get AutoFilter Parent Worksheet
     *
     * @return \ZExcel\Worksheet
     */
    public function getParent()
    {
        return this->_workSheet;
    }

    /**
     * Set AutoFilter Parent Worksheet
     *
     * @param \ZExcel\Worksheet pSheet
     * @return \ZExcel\Worksheet\AutoFilter
     */
    public function setParent(<\ZExcel\Worksheet> pSheet = null)
    {
        let this->_workSheet = pSheet;

        return this;
    }

    /**
     * Get AutoFilter Range
     *
     * @return string
     */
    public function getRange()
    {
        return this->_range;
    }

    /**
     *    Set AutoFilter Range
     *
     *    @param    string        pRange        Cell range (i.e. A1:E10)
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter
     */
    public function setRange(var pRange = "")
    {
        var cellAddress, worksheet, key, rangeStart, rangeEnd, colIndex;
        
        let cellAddress = explode("!",strtoupper(pRange));
        
        if (count(cellAddress) > 1) {
            let worksheet = cellAddress[0];
            let pRange = cellAddress[1];
        }

        if (strpos(pRange,":") !== false) {
            let this->_range = pRange;
        } elseif(empty(pRange)) {
            let this->_range = "";
        } else {
            throw new \ZExcel\Exception("Autofilter must be set on a range of cells.");
        }

        if (empty(pRange)) {
            //    Discard all column rules
            let this->_columns = [];
        } else {
            //    Discard any column rules that are no longer valid within this range
            let key = \ZExcel\Cell::rangeBoundaries(this->_range);
            let rangeStart = key[0];
            let rangeEnd = key[1];

            for key, _ in this->_columns {
                let colIndex = \ZExcel\Cell::columnIndexFromString(key);
                
                if ((rangeStart[0] > colIndex) || (rangeEnd[0] < colIndex)) {
                    unset(this->_columns[key]);
                }
            }
        }

        return this;
    }

    /**
     * Get all AutoFilter Columns
     *
     * @throws    \ZExcel\Exception
     * @return array of \ZExcel\Worksheet\AutoFilter\Column
     */
    public function getColumns()
    {
        return this->_columns;
    }

    /**
     * Validate that the specified column is in the AutoFilter range
     *
     * @param    string    column            Column name (e.g. A)
     * @throws    \ZExcel\Exception
     * @return    integer    The column offset within the autofilter range
     */
    public function testColumnInRange(var column)
    {
        var columnIndex, rangeStart, rangeEnd;
        
        if (empty(this->_range)) {
            throw new \ZExcel\Exception("No autofilter range is defined.");
        }

        let columnIndex = \ZExcel\Cell::columnIndexFromString(column);
        let column = \ZExcel\Cell::rangeBoundaries(this->_range);
        let rangeStart = column[0];
        let rangeEnd = column[1];
        
        if ((rangeStart[0] > columnIndex) || (rangeEnd[0] < columnIndex)) {
            throw new \ZExcel\Exception("Column is outside of current autofilter range.");
        }

        return columnIndex - rangeStart[0];
    }

    /**
     * Get a specified AutoFilter Column Offset within the defined AutoFilter range
     *
     * @param    string    pColumn        Column name (e.g. A)
     * @throws    \ZExcel\Exception
     * @return integer    The offset of the specified column within the autofilter range
     */
    public function getColumnOffset(string pColumn)
    {
        return this->testColumnInRange(pColumn);
    }

    /**
     * Get a specified AutoFilter Column
     *
     * @param    string    pColumn        Column name (e.g. A)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function getColumn(string pColumn)
    {
        this->testColumnInRange(pColumn);

        if (!isset(this->_columns[pColumn])) {
            let this->_columns[pColumn] = new \ZExcel\Worksheet\AutoFilter\Column(pColumn, this);
        }

        return this->_columns[pColumn];
    }

    /**
     * Get a specified AutoFilter Column by it"s offset
     *
     * @param    integer    pColumnOffset        Column offset within range (starting from 0)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet\AutoFilter\Column
     */
    public function getColumnByOffset(int pColumnOffset = 0)
    {
        var tmp, rangeStart, rangeEnd, pColumn;
        
        let tmp = \ZExcel\Cell::rangeBoundaries(this->_range);
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];
        
        let pColumn = \ZExcel\Cell::stringFromColumnIndex(rangeStart[0] + pColumnOffset - 1);

        return this->getColumn(pColumn);
    }

    /**
     *    Set AutoFilter
     *
     *    @param    \ZExcel\Worksheet\AutoFilter\Column|string        pColumn
     *            A simple string containing a Column ID like "A" is permitted
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter
     */
    public function setColumn(pColumn)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Clear a specified AutoFilter Column
     *
     * @param    string  pColumn    Column name (e.g. A)
     * @throws    \ZExcel\Exception
     * @return \ZExcel\Worksheet\AutoFilter
     */
    public function clearColumn(string pColumn) {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Shift an AutoFilter Column Rule to a different column
     *
     *    Note: This method bypasses validation of the destination column to ensure it is within this AutoFilter range.
     *        Nor does it verify whether any column rule already exists at toColumn, but will simply overrideany existing value.
     *        Use with caution.
     *
     *    @param    string    fromColumn        Column name (e.g. A)
     *    @param    string    toColumn        Column name (e.g. B)
     *    @return \ZExcel\Worksheet\AutoFilter
     */
    public function shiftColumn(string fromColumn = null, string toColumn = null)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    Test if cell value is in the defined set of values
     *
     *    @param    mixed        cellValue
     *    @param    mixed[]        dataSet
     *    @return boolean
     */
    private static function _filterTestInSimpleDataSet(cellValue, array dataSet)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Test if cell value is in the defined set of Excel date values
     *
     *    @param    mixed        cellValue
     *    @param    mixed[]        dataSet
     *    @return boolean
     */
    private static function _filterTestInDateGroupSet(cellValue,array dataSet)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Test if cell value is within a set of values defined by a ruleset
     *
     *    @param    mixed        cellValue
     *    @param    mixed[]        ruleSet
     *    @return boolean
     */
    private static function _filterTestInCustomDataSet(cellValue, array ruleSet)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Test if cell date value is matches a set of values defined by a set of months
     *
     *    @param    mixed        cellValue
     *    @param    mixed[]        monthSet
     *    @return boolean
     */
    private static function _filterTestInPeriodDateSet(cellValue, array monthSet)
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    Convert a dynamic rule daterange to a custom filter range expression for ease of calculation
     *
     *    @FIXME filterColumn is a reference (not manage by Zephir)
     *
     *    @param    string                                        dynamicRuleType
     *    @param    \ZExcel\Worksheet\AutoFilter\Column        &filterColumn
     *    @return mixed[]
     */
    private function _dynamicFilterDateRange(string dynamicRuleType, <\ZExcel\Worksheet\AutoFilter\Column> filterColumn)
    {
        throw new \Exception("Not implemented yet!");
    }

    private function _calculateTopTenValue(columnID, startRow, endRow, ruleType, ruleValue)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Apply the AutoFilter rules to the AutoFilter Range
     *
     *    @throws    \ZExcel\Exception
     *    @return \ZExcel\Worksheet\AutoFilter
     */
    public function showHideRows()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        var key, value, k, v, aTmp, vars = get_object_vars(this);
        
        for key, value in vars {
            if (is_object(value)) {
                if (key == "_workSheet") {
                    //    Detach from autofilter parent
                    let this->{key} = null;
                } else {
                    let this->{key} = clone value;
                }
            } elseif ((is_array(value)) && (key == "_olumns")) {
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

    /**
     * toString method replicates previous behavior by returning the range if object is
     *    referenced as a property of its parent.
     */
    public function __toString() -> string
    {
        return (string) this->_range;
    }
}
