namespace ZExcel;

class Cell
{

    /**
     *  Default range variable constant
     *
     *  @var  string
     */
    const DEFAULT_RANGE = "A1:A1";
    
    protected static _columnLookup = [
        "A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8, "I": 9, "J": 10, "K": 11, "L": 12, "M": 13,
        "N": 14, "O": 15, "P": 16, "Q": 17, "R": 18, "S": 19, "T": 20, "U": 21, "V": 22, "W": 23, "X": 24, "Y": 25, "Z": 26,
        "a": 1, "b": 2, "c": 3, "d": 4, "e": 5, "f": 6, "g": 7, "h": 8, "i": 9, "j": 10, "k": 11, "l": 12, "m": 13,
        "n": 14, "o": 15, "p": 16, "q": 17, "r": 18, "s": 19, "t": 20, "u": 21, "v": 22, "w": 23, "x": 24, "y": 25, "z": 26
    ];
    
    protected static _indexCache = [];

    /**
     *    Value binder to use
     *
     *    @var    \ZExcel\Cell\IValueBinder
     */
    private static _valueBinder;
    
    /**
     *    Value of the cell
     *
     *    @var    mixed
     */
    private _value;

    /**
     *    Calculated value of the cell (used for caching)
     *    This returns the value last calculated by MS Excel or whichever spreadsheet program was used to
     *        create the original spreadsheet file.
     *    Note that this value is not guaranteed to reflect the actual calculated value because it is
     *        possible that auto-calculation was disabled in the original spreadsheet, and underlying data
     *        values used by the formula have changed since it was last calculated.
     *
     *    @var mixed
     */
    private _calculatedValue;

    /**
     *    Type of the cell data
     *
     *    @var    string
     */
    private _dataType;

    /**
     *    Parent worksheet
     *
     *    @var    \ZExcel\CachedObjectStorage_CacheBase
     */
    private _parent;

    /**
     *    Index to cellXf
     *
     *    @var    int
     */
    private _xfIndex = 0;

    /**
     *    Attributes of the formula
     *
     */
    private _formulaAttributes;


    /**
     *    Send notification to the cache controller
     *
     *    @return void
     **/
    public function notifyCacheController()
    {
        this->_parent->updateCacheData(this);

        return this;
    }

    public function detach()
    {
        let this->_parent = null;
    }

    public function attach(<\ZExcel\CachedObjectStorage\CacheBase> parent)
    {
        let this->_parent = parent;
    }


    /**
     *    Create a new Cell
     *
     *    @param    mixed                pValue
     *    @param    string                pDataType
     *    @param    \ZExcel\Worksheet    pSheet
     *    @throws    \ZExcel\Exception
     */
    public function __construct(pValue = null, pDataType = null, <\ZExcel\Worksheet> pSheet = null)
    {
        // Initialise cell value
        let this->_value = pValue;

        // Set worksheet cache
        let this->_parent = pSheet->getCellCacheController();

        // Set datatype?
        if (pDataType !== null) {
            if (pDataType == \ZExcel\Cell\DataType::TYPE_STRING2) {
                let pDataType = \ZExcel\Cell\DataType::TYPE_STRING;
            }
            let this->_dataType = pDataType;
        } elseif (!self::getValueBinder()->bindValue(this, pValue)) {
            throw new \ZExcel\Exception("Value could not be bound to cell.");
        }
    }

    /**
     *    Get cell coordinate column
     *
     *    @return    string
     */
    public function getColumn()
    {
        return this->_parent->getCurrentColumn();
    }

    /**
     *    Get cell coordinate row
     *
     *    @return    int
     */
    public function getRow()
    {
        return this->_parent->getCurrentRow();
    }

    /**
     *    Get cell coordinate
     *
     *    @return    string
     */
    public function getCoordinate()
    {
        return this->_parent->getCurrentAddress();
    }

    /**
     *    Get cell value
     *
     *    @return    mixed
     */
    public function getValue()
    {
        return this->_value;
    }

    /**
     *    Get cell value with formatting
     *
     *    @return    string
     */
    public function getFormattedValue()
    {
        return (string) \ZExcel\Style\NumberFormat::toFormattedString(
            this->getCalculatedValue(),
            this->getStyle()
                ->getNumberFormat()->getFormatCode()
        );
    }

    /**
     *    Set cell value
     *
     *    Sets the value for a cell, automatically determining the datatype using the value binder
     *
     *    @param    mixed    pValue                    Value
     *    @return    \ZExcel\Cell
     *    @throws    \ZExcel\Exception
     */
    public function setValue(pValue = null)
    {
        if (!self::getValueBinder()->bindValue(this, pValue)) {
            throw new \ZExcel\Exception("Value could not be bound to cell.");
        }
        return this;
    }

    /**
     *    Set the value for a cell, with the explicit data type passed to the method (bypassing any use of the value binder)
     *
     *    @param    mixed    pValue            Value
     *    @param    string    pDataType        Explicit data type
     *    @return    \ZExcel\Cell
     *    @throws    \ZExcel\Exception
     */
    public function setValueExplicit(var pValue = null, var pDataType = \ZExcel\Cell\DataType::TYPE_STRING)
    {
        // set the value according to data type
        switch (pDataType) {
            case \ZExcel\Cell\DataType::TYPE_NULL:
                let this->_value = pValue;
                break;
            case \ZExcel\Cell\DataType::TYPE_STRING2:
                let pDataType = \ZExcel\Cell\DataType::TYPE_STRING;
            case \ZExcel\Cell\DataType::TYPE_STRING:
                // Synonym for string
            case \ZExcel\Cell\DataType::TYPE_INLINE:
                // Rich text
                let this->_value = \ZExcel\Cell\DataType::checkString(pValue);
                break;
            case \ZExcel\Cell\DataType::TYPE_NUMERIC:
                let this->_value = (float) pValue;
                break;
            case \ZExcel\Cell\DataType::TYPE_FORMULA:
                let this->_value = (string) pValue;
                break;
            case \ZExcel\Cell\DataType::TYPE_BOOL:
                let this->_value = (boolean) pValue;
                break;
            case \ZExcel\Cell\DataType::TYPE_ERROR:
                let this->_value = \ZExcel\Cell\DataType::checkErrorCode(pValue);
                break;
            default:
                throw new \ZExcel\Exception("Invalid datatype: " . pDataType);
        }

        // set the datatype
        let this->_dataType = pDataType;

        return this->notifyCacheController();
    }

    /**
     *    Get calculated cell value
     *
     *    @deprecated        Since version 1.7.8 for planned changes to cell for array formula handling
     *
     *    @param    boolean resetLog  Whether the calculation engine logger should be reset or not
     *    @return    mixed
     *    @throws    \ZExcel\Exception
     */
    public function getCalculatedValue(resetLog = true)
    {
        var result, ex;
        
        if (this->_dataType == \ZExcel\Cell\DataType::TYPE_FORMULA) {
            try {
                let result = \ZExcel\Calculation::getInstance(
                    this->getWorksheet()->getParent()
                )->calculateCellValue(this, resetLog);
                //    We don't yet handle array returns
                if (is_array(result)) {
                    while (is_array(result)) {
                        let result = array_pop(result);
                    }
                }
            } catch ZExcel\Exception, ex {
                if ((ex->getMessage() === "Unable to access External Workbook") && (this->_calculatedValue !== null)) {
                    return this->_calculatedValue; // Fallback for calculations referencing external files.
                }
                let result = "#N/A";
                throw new \ZExcel\Calculation\Exception(
                    this->getWorksheet()->getTitle()."!".this->getCoordinate()." -> ".ex->getMessage()
                );
            }

            if (result === "#Not Yet Implemented") {
                return this->_calculatedValue; // Fallback if calculation engine does not support the formula.
            }

            return result;
        } elseif (this->_value instanceof \ZExcel\RichText) {
            return this->_value->getPlainText();
        }

        return this->_value;
    }

    /**
     *    Set old calculated value (cached)
     *
     *    @param    mixed pValue    Value
     *    @return    \ZExcel\Cell
     */
    public function setCalculatedValue(pValue = null)
    {
        if (pValue !== null) {
            let this->_calculatedValue = (is_numeric(pValue)) ? (float) pValue : pValue;
        }

        return this->notifyCacheController();
    }

    /**
     *    Get old calculated value (cached)
     *    This returns the value last calculated by MS Excel or whichever spreadsheet program was used to
     *        create the original spreadsheet file.
     *    Note that this value is not guaranteed to refelect the actual calculated value because it is
     *        possible that auto-calculation was disabled in the original spreadsheet, and underlying data
     *        values used by the formula have changed since it was last calculated.
     *
     *    @return    mixed
     */
    public function getOldCalculatedValue()
    {
        return this->_calculatedValue;
    }

    /**
     *    Get cell data type
     *
     *    @return string
     */
    public function getDataType()
    {
        return this->_dataType;
    }

    /**
     *    Set cell data type
     *
     *    @param    string pDataType
     *    @return    \ZExcel\Cell
     */
    public function setDataType(string pDataType = \ZExcel\Cell\DataType::TYPE_STRING)
    {
        if (pDataType == \ZExcel\Cell\DataType::TYPE_STRING2) {
            let pDataType = \ZExcel\Cell\DataType::TYPE_STRING;
        }
        let this->_dataType = pDataType;

        return this->notifyCacheController();
    }

    /**
     *  Identify if the cell contains a formula
     *
     *  @return boolean
     */
    public function isFormula()
    {
        return this->_dataType == \ZExcel\Cell\DataType::TYPE_FORMULA;
    }

    /**
     *    Does this cell contain Data validation rules?
     *
     *    @return    boolean
     *    @throws    \ZExcel\Exception
     */
    public function hasDataValidation()
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot check for data validation when cell is not bound to a worksheet");
        }

        return this->getWorksheet()->dataValidationExists(this->getCoordinate());
    }

    /**
     *    Get Data validation rules
     *
     *    @return    \ZExcel\Cell\DataValidation
     *    @throws    \ZExcel\Exception
     */
    public function getDataValidation()
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot get data validation for cell that is not bound to a worksheet");
        }

        return this->getWorksheet()->getDataValidation(this->getCoordinate());
    }

    /**
     *    Set Data validation rules
     *
     *    @param    \ZExcel\Cell\DataValidation    pDataValidation
     *    @return    \ZExcel\Cell
     *    @throws    \ZExcel\Exception
     */
    public function setDataValidation(<\ZExcel\Cell\DataValidation> pDataValidation = null)
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot set data validation for cell that is not bound to a worksheet");
        }

        this->getWorksheet()->setDataValidation(this->getCoordinate(), pDataValidation);

        return this->notifyCacheController();
    }

    /**
     *    Does this cell contain a Hyperlink?
     *
     *    @return boolean
     *    @throws    \ZExcel\Exception
     */
    public function hasHyperlink()
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot check for hyperlink when cell is not bound to a worksheet");
        }

        return this->getWorksheet()->hyperlinkExists(this->getCoordinate());
    }

    /**
     *    Get Hyperlink
     *
     *    @return    \ZExcel\Cell\Hyperlink
     *    @throws    \ZExcel\Exception
     */
    public function getHyperlink()
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot get hyperlink for cell that is not bound to a worksheet");
        }

        return this->getWorksheet()->getHyperlink(this->getCoordinate());
    }

    /**
     *    Set Hyperlink
     *
     *    @param    \ZExcel\Cell\Hyperlink    pHyperlink
     *    @return    \ZExcel\Cell
     *    @throws    \ZExcel\Exception
     */
    public function setHyperlink(<\ZExcel\Cell\Hyperlink> pHyperlink = null)
    {
        if (!isset(this->_parent)) {
            throw new \ZExcel\Exception("Cannot set hyperlink for cell that is not bound to a worksheet");
        }

        this->getWorksheet()->setHyperlink(this->getCoordinate(), pHyperlink);

        return this->notifyCacheController();
    }

    /**
     *    Get parent worksheet
     *
     *    @return \ZExcel\CachedObjectStorage_CacheBase
     */
    public function getParent()
    {
        return this->_parent;
    }

    /**
     *    Get parent worksheet
     *
     *    @return \ZExcel\Worksheet
     */
    public function getWorksheet()
    {
        return this->_parent->getParent();
    }

    /**
     *    Is this cell in a merge range
     *
     *    @return boolean
     */
    public function isInMergeRange()
    {
        return (this->getMergeRange() == true) ? true : false;
    }

    /**
     *    Is this cell the master (top left cell) in a merge range (that holds the actual data value)
     *
     *    @return boolean
     */
    public function isMergeRangeValueCell()
    {
        var mergeRange, startCell;
        
        let mergeRange = this->getMergeRange();
        
        if (mergeRange) {
            let mergeRange = \ZExcel\Cell::splitRange(mergeRange);
            let startCell = mergeRange[0][0];
            if (this->getCoordinate() === startCell) {
                return true;
            }
        }
        
        return false;
    }

    /**
     *    If this cell is in a merge range, then return the range
     *
     *    @return string
     */
    public function getMergeRange()
    {
        var mergeRange;
        
        for mergeRange in this->getWorksheet()->getMergeCells() {
            if (this->isInRange(mergeRange)) {
                return mergeRange;
            }
        }
        
        return false;
    }

    /**
     *    Get cell style
     *
     *    @return    \ZExcel\Style
     */
    public function getStyle()
    {
        return this->getWorksheet()->getStyle(this->getCoordinate());
    }

    /**
     *    Re-bind parent
     *
     *    @param    \ZExcel\Worksheet parent
     *    @return    \ZExcel\Cell
     */
    public function rebindParent(<\ZExcel\Worksheet> parent)
    {
        let this->_parent = parent->getCellCacheController();

        return this->notifyCacheController();
    }

    /**
     *    Is cell in a specific range?
     *
     *    @param    string    pRange        Cell range (e.g. A1:A1)
     *    @return    boolean
     */
    public function isInRange(string pRange = "A1:A1") -> boolean
    {
        var rangeStart, rangeEnd, myColumn, myRow, tmp;
        
        let tmp = self::rangeBoundaries(pRange);
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];

        // Translate properties
        let myColumn = self::columnIndexFromString(this->getColumn());
        let myRow    = this->getRow();

        // Verify if cell is in range
        return ((rangeStart[0] <= myColumn) && (rangeEnd[0] >= myColumn) && (rangeStart[1] <= myRow) && (rangeEnd[1] >= myRow));
    }

    /**
     *    Coordinate from string
     *
     *    @param    string    pCoordinateString
     *    @return    array    Array containing column and row (indexes 0 and 1)
     *    @throws    \ZExcel\Exception
     */
    public static function coordinateFromString(string pCoordinateString = "A1")
    {
        var matches = [];
        
        if (preg_match("/^([$]?[A-Z]{1,3})([$]?\d{1,7})$/", pCoordinateString, matches)) {
            return [matches[1],matches[2]];
        } elseif ((strpos(pCoordinateString,":") !== false) || (strpos(pCoordinateString,",") !== false)) {
            throw new \ZExcel\Exception("Cell coordinate string can not be a range of cells");
        } elseif (pCoordinateString == "") {
            throw new \ZExcel\Exception("Cell coordinate can not be zero-length string");
        }

        throw new \ZExcel\Exception("Invalid cell coordinate ".pCoordinateString);
    }

    /**
     *    Make string row, column or cell coordinate absolute
     *
     *    @param    string    pCoordinateString        e.g. 'A' or '1' or 'A1'
     *                    Note that this value can be a row or column reference as well as a cell reference
     *    @return    string    Absolute coordinate        e.g. 'A' or '1' or 'A1'
     *    @throws    \ZExcel\Exception
     */
    public static function absoluteReference(var pCoordinateString = "A1")
    {
        var worksheet, cellAddress;
        
        let pCoordinateString = strval(pCoordinateString);
        
        if (strpos(pCoordinateString, ":") === false && strpos(pCoordinateString, ",") === false) {
            // Split out any worksheet name from the reference
            let worksheet = "";
            let cellAddress = explode("!", pCoordinateString);
            
            if (count(cellAddress) > 1) {
                let worksheet = cellAddress[0];
                let pCoordinateString = cellAddress[1];
            }
            
            if (strlen(worksheet) > 0) {
                let worksheet = worksheet . "!";
            }

            // Create absolute coordinate
            if (ctype_digit(pCoordinateString)) {
                return worksheet . "$" . pCoordinateString;
            } elseif (ctype_alpha(pCoordinateString)) {
                return worksheet . "$" . strtoupper(pCoordinateString);
            }
            
            return worksheet . self::absoluteCoordinate(pCoordinateString);
        }

        throw new \ZExcel\Exception("Cell coordinate string can not be a range of cells");
    }

    /**
     *    Make string coordinate absolute
     *
     *    @param    string    pCoordinateString        e.g. 'A1'
     *    @return    string    Absolute coordinate        e.g. 'A1'
     *    @throws    \ZExcel\Exception
     */
    public static function absoluteCoordinate(var pCoordinateString = "A1") -> string
    {
        var worksheet, cellAddress, column, row, tmp;
        
        if (strpos(pCoordinateString, ":") === false && strpos(pCoordinateString, ",") === false) {
            // Split out any worksheet name from the coordinate
            let worksheet = "";
            let cellAddress = explode("!", pCoordinateString);
            
            if (count(cellAddress) > 1) {
                let worksheet = cellAddress[0];
                let pCoordinateString = cellAddress[1];
            }
            
            if (strlen(worksheet) > 0) {
                let worksheet .= "!";
            }

            // Create absolute coordinate
            let tmp = self::coordinateFromString(pCoordinateString);
            let column = ltrim(tmp[0], "$");
            let row = ltrim(tmp[1], "$");
            
            return worksheet . "$" . column . "$" . row;
        }

        throw new \ZExcel\Exception("Cell coordinate string can not be a range of cells");
    }

    /**
     *    Split range into coordinate strings
     *
     *    @param    string    pRange        e.g. 'B4:D9' or 'B4:D9,H2:O11' or 'B4'
     *    @return    array    Array containg one or more arrays containing one or two coordinate strings
     *                                e.g. array('B4','D9') or array(array('B4','D9'),array('H2','O11'))
     *                                        or array('B4')
     */
    public static function splitRange(string pRange = "A1:A1") -> array
    {
        var exploded, counter, i;
        
        // Ensure pRange is a valid range
        if (empty(pRange)) {
            let pRange = self::DEFAULT_RANGE;
        }

        let exploded = explode(",", pRange);
        let counter = count(exploded) - 1;
        
        for i in range(0, counter) {
            let exploded[i] = explode(":", exploded[i]);
        }
        
        return exploded;
    }

    /**
     *    Build range from coordinate strings
     *
     *    @param    array    pRange    Array containg one or more arrays containing one or two coordinate strings
     *    @return    string    String representation of pRange
     *    @throws    \ZExcel\Exception
     */
    public static function buildRange(array pRange) -> array
    {
        var imploded, counter, i;
        
        // Verify range
        if (!is_array(pRange) || empty(pRange) || !is_array(pRange[0])) {
            throw new \ZExcel\Exception("Range does not contain any information");
        }

        // Build range
        let imploded = [];
        let counter = count(pRange);
        
        for i in range(0, counter - 1) {
            let pRange[i] = implode(":", pRange[i]);
        }
        
        let imploded = implode(",", pRange);

        return imploded;
    }

    /**
     *    Calculate range boundaries
     *
     *    @param    string    pRange        Cell range (e.g. A1:A1)
     *    @return    array    Range coordinates array(Start Cell, End Cell)
     *                    where Start Cell and End Cell are arrays (Column Number, Row Number)
     */
    public static function rangeBoundaries(var pRange = "A1:A1")
    {
        var rangeA, rangeB, rangeStart, rangeEnd, tmp;
        
        // Ensure pRange is a valid range
        if (empty(pRange)) {
            let pRange = self::DEFAULT_RANGE;
        }

        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        // Extract range
        if (strpos(pRange, ":") === false) {
            let rangeA = pRange;
            let rangeB = pRange;
        } else {
            let tmp = explode(":", pRange);
            
            let rangeA = tmp[0];
            let rangeB = tmp[1];
        }

        // Calculate range outer borders
        let rangeStart = self::coordinateFromString(rangeA);
        let rangeEnd    = self::coordinateFromString(rangeB);

        // Translate column into index
        let rangeStart[0]    = self::columnIndexFromString(rangeStart[0]);
        let rangeEnd[0]    = self::columnIndexFromString(rangeEnd[0]);

        return [rangeStart, rangeEnd];
    }

    /**
     *    Calculate range dimension
     *
     *    @param    string    pRange        Cell range (e.g. A1:A1)
     *    @return    array    Range dimension (width, height)
     */
    public static function rangeDimension(var pRange = "A1:A1")
    {
        var rangeStart, rangeEnd, tmp;
        
        // Calculate range outer borders
        let tmp = self::rangeBoundaries(pRange);
        
        let rangeStart = tmp[0];
        let rangeEnd = tmp[1];

        return [(rangeEnd[0] - rangeStart[0] + 1), (rangeEnd[1] - rangeStart[1] + 1)];
    }

    /**
     *    Calculate range boundaries
     *
     *    @param    string    pRange        Cell range (e.g. A1:A1)
     *    @return    array    Range coordinates array(Start Cell, End Cell)
     *                    where Start Cell and End Cell are arrays (Column ID, Row Number)
     */
    public static function getRangeBoundaries(var pRange = "A1:A1")
    {
        var rangeA, rangeB, tmp;
        
        // Ensure pRange is a valid range
        if (empty(pRange)) {
            let pRange = self::DEFAULT_RANGE;
        }

        // Uppercase coordinate
        let pRange = strtoupper(pRange);

        // Extract range
        if (strpos(pRange, ":") === false) {
            let rangeA = pRange;
            let rangeB = pRange;
        } else {
            let tmp = explode(":", pRange);
            
            let rangeA = tmp[0];
            let rangeB = tmp[1];
        }

        return [self::coordinateFromString(rangeA), self::coordinateFromString(rangeB)];
    }
    
    /**
     *    Column index from string
     *
     *    @param    string pString
     *    @return    int Column index (base 1 !!!)
     */
    public static function columnIndexFromString(var pString = "A")
    {
        var len;
        
        if (isset(self::_indexCache[pString])) {
            return self::_indexCache[pString];
        }

        //    We also use the language construct isset() rather than the more costly strlen() function to match the length of pString
        //        for improved performance
        let len = strlen(pString);
        if (len > 0) {
            if (len < 2) {
                let self::_indexCache[pString] = self::_columnLookup[pString];
                return self::_indexCache[pString];
            } elseif (len < 3) {
                let self::_indexCache[pString] = self::_columnLookup[substr(pString, 0, 1)] * 26 + self::_columnLookup[substr(pString, 1, 1)];
                return self::_indexCache[pString];
            } elseif (len < 4) {
                let self::_indexCache[pString] = self::_columnLookup[substr(pString, 0, 1)] * 676 + self::_columnLookup[substr(pString, 1, 1)] * 26 + self::_columnLookup[substr(pString, 2, 1)];
                return self::_indexCache[pString];
            }
        }
        
        throw new \ZExcel\Exception("Column string index can not be " . ((strlen(pString) > 0) ? "longer than 3 characters" : "empty"));
    }

    /**
     *    String from columnindex
     *
     *    @param    int pColumnIndex Column index (base 0 !!!)
     *    @return    string
     */
    public static function stringFromColumnIndex(var pColumnIndex = 0)
    {
        if (!isset(self::_indexCache[pColumnIndex])) {
            // Determine column string
            if (pColumnIndex < 26) {
                let self::_indexCache[pColumnIndex] = chr(65 + pColumnIndex);
            } elseif (pColumnIndex < 702) {
                let self::_indexCache[pColumnIndex] = chr(64 + (pColumnIndex / 26)) . chr(65 + pColumnIndex % 26);
            } else {
                let self::_indexCache[pColumnIndex] = chr(64 + ((pColumnIndex - 26) / 676)) . chr(65 + (((pColumnIndex - 26) % 676) / 26)) . chr(65 + pColumnIndex % 26);
            }
        }
        
        return self::_indexCache[pColumnIndex];
    }

    /**
     *    Extract all cell references in range
     *
     *    @param    string    pRange        Range (e.g. A1 or A1:C10 or A1:E10 A20:E25)
     *    @return    array    Array containing single cell references
     */
    public static function extractAllCellReferencesInRange(var pRange = "A1")
    {
        var cellBlocks, cellBlock, ranges, range, rangeStart, rangeEnd,
            startCol = 0, startRow = 0, endCol = 0, endRow = 0,
            currentCol, currentRow, sortKeys, coord, tmp;
        
        // Returnvalue
        array returnValue = [];

        // Explode spaces
        let cellBlocks = explode(" ", str_replace("$", "", strtoupper(pRange)));
        
        for cellBlock in cellBlocks {
            // Single cell?
            if (strpos(cellBlock,":") === false && strpos(cellBlock,",") === false) {
                let returnValue[] = cellBlock;
                continue;
            }

            // Range...
            let ranges = self::splitRange(cellBlock);
            
            for range in ranges {
                // Single cell?
                if (count(range) === 1) {
                    let returnValue[] = range[0];
                    continue;
                }

                // Range...
                let rangeStart = range[0];
                let rangeEnd = range[1];
                
                let tmp = sscanf(rangeStart, "%[A-Z]%d");
                
                let startCol = tmp[0];
                let startRow = tmp[1];
                
                let tmp = sscanf(rangeEnd, "%[A-Z]%d");
                
                let endCol = tmp[0];
                let endRow = tmp[1];
                
                let endCol = chr(ord(endCol) + 1);
                
                // Current data
                let currentCol = startCol;
                let currentRow = startRow;

                // Loop cells
                while (currentCol != endCol) {
                    while (currentRow <= endRow) {
                        let returnValue[] = currentCol . currentRow;
                        let currentRow = currentRow + 1;
                    }
                    let currentCol = chr(ord(currentCol) + 1);
                    let currentRow = startRow;
                }
            }
        }

        //    Sort the result by column and row
        let sortKeys = [];
        
        for coord in array_unique(returnValue) {
            let tmp = sscanf(coord, "%[A-Z]%d");
            let sortKeys[sprintf("%3s%09d", tmp[0], tmp[1])] = coord;
        }
        
        ksort(sortKeys);

        // Return value
        return array_values(sortKeys);
    }

    /**
     * Compare 2 cells
     *
     * @param    \ZExcel\Cell    a    Cell a
     * @param    \ZExcel\Cell    b    Cell b
     * @return    int        Result of comparison (always -1 or 1, never zero!)
     */
    public static function compareCells(<\ZExcel\Cell> a, <\ZExcel\Cell> b)
    {
        if (a->getRow() < b->getRow()) {
            return -1;
        } elseif (a->getRow() > b->getRow()) {
            return 1;
        } elseif (self::columnIndexFromString(a->getColumn()) < self::columnIndexFromString(b->getColumn())) {
            return -1;
        } else {
            return 1;
        }
    }

    /**
     * Get value binder to use
     *
     * @return \ZExcel\Cell\IValueBinder
     */
    public static function getValueBinder()
    {
        if (self::_valueBinder === null) {
            let self::_valueBinder = new \ZExcel\Cell\DefaultValueBinder();
        }

        return self::_valueBinder;
    }

    /**
     * Set value binder to use
     *
     * @param \ZExcel\Cell\IValueBinder binder
     * @throws \ZExcel\Exception
     */
    public static function setValueBinder(<\ZExcel\Cell\IValueBinder> binder = null)
    {
        if (binder === null) {
            throw new \ZExcel\Exception("A \ZExcel\Cell\IValueBinder is required for PHPExcel to function correctly.");
        }

        let self::_valueBinder = binder;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        var key, value, vars;
        
        let vars = get_object_vars(this);
        for key, value in vars {
            if ((is_object(value)) && (key != "_parent")) {
                let this->{key} = clone value;
            } else {
                let this->{key} = value;
            }
        }
    }

    /**
     * Get index to cellXf
     *
     * @return int
     */
    public function getXfIndex()
    {
        return this->_xfIndex;
    }

    /**
     * Set index to cellXf
     *
     * @param int pValue
     * @return \ZExcel\Cell
     */
    public function setXfIndex(pValue = 0)
    {
        let this->_xfIndex = pValue;

        return this->notifyCacheController();
    }

    /**
     *    @deprecated        Since version 1.7.8 for planned changes to cell for array formula handling
     */
    public function setFormulaAttributes(pAttributes)
    {
        let this->_formulaAttributes = pAttributes;
        
        return this;
    }

    /**
     *    @deprecated        Since version 1.7.8 for planned changes to cell for array formula handling
     */
    public function getFormulaAttributes()
    {
        return this->_formulaAttributes;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return (string) this->getValue();
    }
}
