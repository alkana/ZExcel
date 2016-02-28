namespace ZExcel\Calculation;

class LookupRef
{
    /**
     * CELL_ADDRESS
     *
     * Creates a cell address as text, given specified row and column numbers.
     *
     * Excel Function:
     *        =ADDRESS(row, column, [relativity], [referenceStyle], [sheetText])
     *
     * @param    row                Row number to use in the cell reference
     * @param    column            Column number to use in the cell reference
     * @param    relativity        Flag indicating the type of reference to return
     *                                1 or omitted    Absolute
     *                                2                Absolute row; relative column
     *                                3                Relative row; absolute column
     *                                4                Relative
     * @param    referenceStyle    A logical value that specifies the A1 or R1C1 reference style.
     *                                TRUE or omitted        CELL_ADDRESS returns an A1-style reference
     *                                FALSE                CELL_ADDRESS returns an R1C1-style reference
     * @param    sheetText        Optional Name of worksheet to use
     * @return    string
     */
    public static function cell_address(var row, var column, var relativity = 1, var referenceStyle = true, var sheetText = "")
    {
        var rowRelative, columnRelative;
        
        let row        = \ZExcel\Calculation\Functions::flattenSingleValue(row);
        let column     = \ZExcel\Calculation\Functions::flattenSingleValue(column);
        let relativity = \ZExcel\Calculation\Functions::flattenSingleValue(relativity);
        let sheetText  = \ZExcel\Calculation\Functions::flattenSingleValue(sheetText);

        if ((row < 1) || (column < 1)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (strlen(sheetText) > 0) {
            if (strpos(sheetText, " ") !== false) {
                let sheetText = "'" . sheetText . "'";
            }
            let sheetText = sheetText . "!";
        }
        if ((!is_bool(referenceStyle)) || referenceStyle) {
            let rowRelative = "";
            let columnRelative = "";
            let column = \ZExcel\Cell::stringFromColumnIndex(column - 1);
            
            if ((relativity == 2) || (relativity == 4)) {
                let columnRelative = "";
            }
            
            if ((relativity == 3) || (relativity == 4)) {
                let rowRelative = "";
            }
            return sheetText.columnRelative.column.rowRelative.row;
        } else {
            if ((relativity == 2) || (relativity == 4)) {
                let column = "[".column."]";
            }
            if ((relativity == 3) || (relativity == 4)) {
                let row = "[".row."]";
            }
            return sheetText . "R" . row . "C" . column;
        }
    }


    /**
     * COLUMN
     *
     * Returns the column number of the given cell reference
     * If the cell reference is a range of cells, COLUMN returns the column numbers of each column in the reference as a horizontal array.
     * If cell reference is omitted, and the function is being called through the calculation engine, then it is assumed to be the
     *        reference of the cell in which the COLUMN function appears; otherwise this function returns 0.
     *
     * Excel Function:
     *        =COLUMN([cellAddress])
     *
     * @param    cellAddress        A reference to a range of cells for which you want the column numbers
     * @return    integer or array of integer
     */
    public static function column(var cellAddress = null)
    {
        var columnKey, returnValue, sheet, startAddress, endAddress;
        
        if (is_null(cellAddress) || trim(cellAddress) === "") {
            return 0;
        }

        if (is_array(cellAddress)) {
            for columnKey, _ in cellAddress {
                let columnKey = preg_replace("/[^a-z]/i", "", columnKey);
                return (int) \ZExcel\Cell::columnIndexFromString(columnKey);
            }
        } else {
            if (strpos(cellAddress, "!") !== false) {
                let returnValue = explode("!", cellAddress);
                let sheet = returnValue[0];
                let cellAddress = returnValue[1];
            }
            if (strpos(cellAddress, ":") !== false) {
                let returnValue = explode(":", cellAddress);
                let startAddress = returnValue[0];
                let endAddress = returnValue[1];
                let startAddress = preg_replace("/[^a-z]/i", "", startAddress);
                let endAddress = preg_replace("/[^a-z]/i", "", endAddress);
                let returnValue = [];
                
                do {
                    let returnValue[] = (int) \ZExcel\Cell::columnIndexFromString(startAddress);
                    let startAddress = startAddress + 1;
                } while (startAddress != endAddress);
                
                return returnValue;
            } else {
                let cellAddress = preg_replace("/[^a-z]/i", "", cellAddress);
                return (int) \ZExcel\Cell::columnIndexFromString(cellAddress);
            }
        }
    }


    /**
     * COLUMNS
     *
     * Returns the number of columns in an array or reference.
     *
     * Excel Function:
     *        =COLUMNS(cellAddress)
     *
     * @param    cellAddress        An array or array formula, or a reference to a range of cells for which you want the number of columns
     * @return    integer            The number of columns in cellAddress
     */
    public static function columns(var cellAddress = null)
    {
        var isMatrix, tmp, columns, rows;
        
        if (is_null(cellAddress) || cellAddress === "") {
            return 1;
        } elseif (!is_array(cellAddress)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        reset(cellAddress);
        
        let isMatrix = (is_numeric(key(cellAddress)));
        let tmp = \ZExcel\Calculation::_getMatrixDimensions(cellAddress);
        let columns = tmp[0];
        let rows = tmp[1];

        if (isMatrix) {
            return rows;
        } else {
            return columns;
        }
    }


    /**
     * ROW
     *
     * Returns the row number of the given cell reference
     * If the cell reference is a range of cells, ROW returns the row numbers of each row in the reference as a vertical array.
     * If cell reference is omitted, and the function is being called through the calculation engine, then it is assumed to be the
     *        reference of the cell in which the ROW function appears; otherwise this function returns 0.
     *
     * Excel Function:
     *        =ROW([cellAddress])
     *
     * @param    cellAddress        A reference to a range of cells for which you want the row numbers
     * @return    integer or array of integer
     */
    public static function row(var cellAddress = null)
    {
        var rowValue, rowKey, cellValue, sheet, tmp, startAddress, endAddress, returnValue;
        int i;
        
        if (is_null(cellAddress) || trim(cellAddress) === "") {
            return 0;
        }

        if (is_array(cellAddress)) {
            for rowValue in cellAddress {
                for rowKey, cellValue in rowValue {
                    return (int) preg_replace("/[^0-9]/i", "", rowKey);
                }
            }
        } else {
            if (strpos(cellAddress, "!") !== false) {
                let tmp = explode("!", cellAddress);
                let sheet = tmp[0];
                let cellAddress = tmp[1]; 
            }
            
            if (strpos(cellAddress, ":") !== false) {
                let tmp = explode(":", cellAddress);
                let startAddress = tmp[0];
                let endAddress = tmp[1];
                let startAddress = preg_replace("/[^0-9]/", "", startAddress);
                let endAddress = preg_replace("/[^0-9]/", "", endAddress);
                let returnValue = [];
                let i = 0;
                
                do {
                    let returnValue[i] = [];
                    let returnValue[i][] = (int) startAddress;
                    
                    let i = i + 1;
                    let startAddress = startAddress + 1;
                } while (startAddress != endAddress);
                
                return returnValue;
            } else {
                let tmp = explode(":", cellAddress);
                let cellAddress = tmp[0];
                
                return (int) preg_replace("/[^0-9]/", "", cellAddress);
            }
        }
    }


    /**
     * ROWS
     *
     * Returns the number of rows in an array or reference.
     *
     * Excel Function:
     *        =ROWS(cellAddress)
     *
     * @param    cellAddress        An array or array formula, or a reference to a range of cells for which you want the number of rows
     * @return    integer            The number of rows in cellAddress
     */
    public static function rows(var cellAddress = null)
    {
        var isMatrix, columns, rows;
        
        if (is_null(cellAddress) || cellAddress === "") {
            return 1;
        } elseif (!is_array(cellAddress)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        reset(cellAddress);
        
        let isMatrix = (is_numeric(key(cellAddress)));
        let cellAddress = \ZExcel\Calculation::_getMatrixDimensions(cellAddress);
        let columns = cellAddress[0];
        let rows = cellAddress[1];

        if (isMatrix) {
            return columns;
        } else {
            return rows;
        }
    }


    /**
     * HYPERLINK
     *
     * Excel Function:
     *        =HYPERLINK(linkURL,displayName)
     *
     * @access    public
     * @category Logical Functions
     * @param    string            linkURL        Value to check, is also the value returned when no error
     * @param    string            displayName    Value to return when testValue is an error condition
     * @param    \ZExcel\Cell    pCell            The cell to set the hyperlink in
     * @return    mixed    The value of displayName (or linkURL if displayName was blank)
     */
    public static function hyperlink(var linkURL = "", var displayName = null, <\ZExcel\Cell> pCell = null)
    {
        var args;
        
        let args = func_get_args();
        let pCell = array_pop(args);

        let linkURL     = (is_null(linkURL))     ? "" : \ZExcel\Calculation\Functions::flattenSingleValue(linkURL);
        let displayName = (is_null(displayName)) ? "" : \ZExcel\Calculation\Functions::flattenSingleValue(displayName);

        if ((!is_object(pCell)) || (trim(linkURL) == "")) {
            return \ZExcel\Calculation\Functions::ReF();
        }

        if ((is_object(displayName)) || trim(displayName) == "") {
            let displayName = linkURL;
        }

        pCell->getHyperlink()->setUrl(linkURL);

        return displayName;
    }


    /**
     * INDIRECT
     *
     * Returns the reference specified by a text string.
     * References are immediately evaluated to display their contents.
     *
     * Excel Function:
     *        =INDIRECT(cellAddress)
     *
     * NOTE - INDIRECT() does not yet support the optional a1 parameter introduced in Excel 2010
     *
     * @param    cellAddress        cellAddress    The cell address of the current cell (containing this formula)
     * @param    \ZExcel\Cell    pCell            The current cell (containing this formula)
     * @return    mixed            The cells referenced by cellAddress
     *
     * @todo    Support for the optional a1 parameter introduced in Excel 2010
     *
     */
    public static function indirect(var cellAddress = null, <\ZExcel\Cell> pCell = null)
    {
        var cellAddress1, cellAddress2, tmp, sheetName, pSheet;
        array matches = [];
        
        let cellAddress = \ZExcel\Calculation\Functions::flattenSingleValue(cellAddress);
        
        if (is_null(cellAddress) || cellAddress === "") {
            return \ZExcel\Calculation\Functions::ReF();
        }

        let cellAddress1 = cellAddress;
        let cellAddress2 = null;
        
        if (strpos(cellAddress, ":") !== false) {
            let tmp = explode(":", cellAddress);
            let cellAddress1 = tmp[0];
            let cellAddress2 = tmp[1];
        }

        if ((!preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_CELLREF . "/i", cellAddress1, matches)) || ((!is_null(cellAddress2)) && (!preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_CELLREF . "/i", cellAddress2, matches)))) {
            if (!preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_NAMEDRANGE ."/i", cellAddress1, matches)) {
                return \ZExcel\Calculation\Functions::ReF();
            }

            if (strpos(cellAddress, "!") !== false) {
                let tmp = explode("!", cellAddress);
                let sheetName = tmp[0];
                let cellAddress = tmp[1];
                let sheetName = trim(sheetName, "'");
                let pSheet = pCell->getWorksheet()->getParent()->getSheetByName(sheetName);
            } else {
                let pSheet = pCell->getWorksheet();
            }

            return \ZExcel\Calculation::getInstance()->extractNamedRange(cellAddress, pSheet, false);
        }

        if (strpos(cellAddress, "!") !== false) {
            let tmp = explode("!", cellAddress);
            let sheetName = tmp[0];
            let cellAddress = tmp[1];
            let sheetName = trim(sheetName, "'");
            let pSheet = pCell->getWorksheet()->getParent()->getSheetByName(sheetName);
        } else {
            let pSheet = pCell->getWorksheet();
        }

        return \ZExcel\Calculation::getInstance()->extractCellRange(cellAddress, pSheet, false);
    }


    /**
     * OFFSET
     *
     * Returns a reference to a range that is a specified number of rows and columns from a cell or range of cells.
     * The reference that is returned can be a single cell or a range of cells. You can specify the number of rows and
     * the number of columns to be returned.
     *
     * Excel Function:
     *        =OFFSET(cellAddress, rows, cols, [height], [width])
     *
     * @param    cellAddress        The reference from which you want to base the offset. Reference must refer to a cell or
     *                                range of adjacent cells; otherwise, OFFSET returns the #VALUE! error value.
     * @param    rows            The number of rows, up or down, that you want the upper-left cell to refer to.
     *                                Using 5 as the rows argument specifies that the upper-left cell in the reference is
     *                                five rows below reference. Rows can be positive (which means below the starting reference)
     *                                or negative (which means above the starting reference).
     * @param    cols            The number of columns, to the left or right, that you want the upper-left cell of the result
     *                                to refer to. Using 5 as the cols argument specifies that the upper-left cell in the
     *                                reference is five columns to the right of reference. Cols can be positive (which means
     *                                to the right of the starting reference) or negative (which means to the left of the
     *                                starting reference).
     * @param    height            The height, in number of rows, that you want the returned reference to be. Height must be a positive number.
     * @param    width            The width, in number of columns, that you want the returned reference to be. Width must be a positive number.
     * @return    string            A reference to a cell or range of cells
     */
    public static function offset(var cellAddress = null, var rows = 0, var columns = 0, var height = null, var width = null)
    {
        var args, pCell, tmp, sheetName, startCell, endCell, startCellColumn, startCellRow, endCellColumn, endCellRow, pSheet;
        
        let rows    = \ZExcel\Calculation\Functions::flattenSingleValue(rows);
        let columns = \ZExcel\Calculation\Functions::flattenSingleValue(columns);
        let height  = \ZExcel\Calculation\Functions::flattenSingleValue(height);
        let width   = \ZExcel\Calculation\Functions::flattenSingleValue(width);
        
        if (cellAddress == null) {
            return 0;
        }

        let args = func_get_args();
        let pCell = array_pop(args);
        
        if (!is_object(pCell)) {
            return \ZExcel\Calculation\Functions::ReF();
        }

        let sheetName = null;
        
        if (strpos(cellAddress, "!")) {
            let tmp = explode("!", cellAddress); 
            let sheetName = tmp[0];
            let cellAddress = tmp[1];
            let sheetName = trim(sheetName, "'");
        }
        
        if (strpos(cellAddress, ":")) {
            let tmp = explode(":", cellAddress); 
            let startCell = tmp[0];
            let endCell = tmp[1];
        } else {
            let startCell = cellAddress;
            let endCell = cellAddress;
        }
        
        let tmp  = \ZExcel\Cell::coordinateFromString(startCell);
        let startCellColumn = tmp[0];
        let startCellRow = tmp[1];
        let tmp = \ZExcel\Cell::coordinateFromString(endCell);
        let endCellColumn = tmp[0];
        let endCellRow = tmp[1];

        let startCellRow = startCellRow + rows;
        let startCellColumn = \ZExcel\Cell::columnIndexFromString(startCellColumn) - 1;
        let startCellColumn = startCellColumn + columns;

        if ((startCellRow <= 0) || (startCellColumn < 0)) {
            return \ZExcel\Calculation\Functions::ReF();
        }
        
        let endCellColumn = \ZExcel\Cell::columnIndexFromString(endCellColumn) - 1;
        
        if ((width != null) && (!is_object(width))) {
            let endCellColumn = startCellColumn + width - 1;
        } else {
            let endCellColumn = endCellColumn + columns;
        }
        
        let startCellColumn = \ZExcel\Cell::stringFromColumnIndex(startCellColumn);

        if ((height != null) && (!is_object(height))) {
            let endCellRow = startCellRow + height - 1;
        } else {
            let endCellRow = endCellRow + rows;
        }

        if ((endCellRow <= 0) || (endCellColumn < 0)) {
            return \ZExcel\Calculation\Functions::ReF();
        }
        
        let endCellColumn = \ZExcel\Cell::stringFromColumnIndex(endCellColumn);

        let cellAddress = startCellColumn.startCellRow;
        
        if ((startCellColumn != endCellColumn) || (startCellRow != endCellRow)) {
            let cellAddress = cellAddress . ":" . endCellColumn . endCellRow;
        }

        if (sheetName !== null) {
            let pSheet = pCell->getWorksheet()->getParent()->getSheetByName(sheetName);
        } else {
            let pSheet = pCell->getWorksheet();
        }

        return \ZExcel\Calculation::getInstance()->extractCellRange(cellAddress, pSheet, false);
    }


    /**
     * CHOOSE
     *
     * Uses lookup_value to return a value from the list of value arguments.
     * Use CHOOSE to select one of up to 254 values based on the lookup_value.
     *
     * Excel Function:
     *        =CHOOSE(index_num, value1, [value2], ...)
     *
     * @param    index_num        Specifies which value argument is selected.
     *                            Index_num must be a number between 1 and 254, or a formula or reference to a cell containing a number
     *                                between 1 and 254.
     * @param    value1...        Value1 is required, subsequent values are optional.
     *                            Between 1 to 254 value arguments from which CHOOSE selects a value or an action to perform based on
     *                                index_num. The arguments can be numbers, cell references, defined names, formulas, functions, or
     *                                text.
     * @return    mixed            The selected value
     */
    public static function choose()
    {
        var chooseArgs, chosenEntry, entryCount;
        
        let chooseArgs = func_get_args();
        let chosenEntry = \ZExcel\Calculation\Functions::flattenArray(array_shift(chooseArgs));
        let entryCount = count(chooseArgs) - 1;

        if (is_array(chosenEntry)) {
            let chosenEntry = array_shift(chosenEntry);
        }
        
        if ((is_numeric(chosenEntry)) && (!is_bool(chosenEntry))) {
            let chosenEntry = chosenEntry - 1;
        } else {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let chosenEntry = floor(chosenEntry);
        
        if ((chosenEntry < 0) || (chosenEntry > entryCount)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (is_array(chooseArgs[chosenEntry])) {
            return \ZExcel\Calculation\Functions::flattenArray(chooseArgs[chosenEntry]);
        } else {
            return chooseArgs[chosenEntry];
        }
    }


    /**
     * MATCH
     *
     * The MATCH function searches for a specified item in a range of cells
     *
     * Excel Function:
     *        =MATCH(lookup_value, lookup_array, [match_type])
     *
     * @param    lookup_value    The value that you want to match in lookup_array
     * @param    lookup_array    The range of cells being searched
     * @param    match_type        The number -1, 0, or 1. -1 means above, 0 means exact match, 1 means below. If match_type is 1 or -1, the list has to be ordered.
     * @return    integer            The relative position of the found item
     */
    public static function match(var lookup_value, var lookup_array, var match_type = 1)
    {
        var lookupArraySize, lookupArrayValue, i, keySet;
        
        let lookup_array = \ZExcel\Calculation\Functions::flattenArray(lookup_array);
        let lookup_value = \ZExcel\Calculation\Functions::flattenSingleValue(lookup_value);
        let match_type    = (is_null(match_type)) ? 1 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(match_type);
        //    MATCH is not case sensitive
        let lookup_value = strtolower(lookup_value);

        //    lookup_value type has to be number, text, or logical values
        if ((!is_numeric(lookup_value)) && (!is_string(lookup_value)) && (!is_bool(lookup_value))) {
            return \ZExcel\Calculation\Functions::Na();
        }

        //    match_type is 0, 1 or -1
        if ((match_type !== 0) && (match_type !== -1) && (match_type !== 1)) {
            return \ZExcel\Calculation\Functions::Na();
        }

        //    lookup_array should not be empty
        let lookupArraySize = count(lookup_array);
        
        if (lookupArraySize <= 0) {
            return \ZExcel\Calculation\Functions::Na();
        }

        //    lookup_array should contain only number, text, or logical values, or empty (null) cells
        for i, lookupArrayValue in lookup_array {
            //    check the type of the value
            if ((!is_numeric(lookupArrayValue)) && (!is_string(lookupArrayValue)) && (!is_bool(lookupArrayValue)) && (!is_null(lookupArrayValue))) {
                return \ZExcel\Calculation\Functions::Na();
            }
            
            //    convert strings to lowercase for case-insensitive testing
            if (is_string(lookupArrayValue)) {
                let lookup_array[i] = strtolower(lookupArrayValue);
            }
            
            if ((is_null(lookupArrayValue)) && ((match_type == 1) || (match_type == -1))) {
                let lookup_array = array_slice(lookup_array, 0, i - 1);
            }
        }

        // if match_type is 1 or -1, the list has to be ordered
        if (match_type == 1) {
            asort(lookup_array);
            let keySet = array_keys(lookup_array);
        } elseif (match_type == -1) {
            arsort(lookup_array);
            let keySet = array_keys(lookup_array);
        }

        // **
        // find the match
        // **
        for i, lookupArrayValue in lookup_array {
            if ((match_type == 0) && (lookupArrayValue == lookup_value)) {
                //    exact match
                return (i + 1);
            } elseif ((match_type == -1) && (lookupArrayValue <= lookup_value)) {
                let i = array_search(i, keySet);
                // if match_type is -1 <=> find the smallest value that is greater than or equal to lookup_value
                if (i < 1) {
                    // 1st cell was already smaller than the lookup_value
                    break;
                } else {
                    // the previous cell was the match
                    return keySet[i - 1] + 1;
                }
            } elseif ((match_type == 1) && (lookupArrayValue >= lookup_value)) {
                let i = array_search(i, keySet);
                // if match_type is 1 <=> find the largest value that is less than or equal to lookup_value
                if (i < 1) {
                    // 1st cell was already bigger than the lookup_value
                    break;
                } else {
                    // the previous cell was the match
                    return keySet[i - 1] + 1;
                }
            }
        }

        //    unsuccessful in finding a match, return #N/A error value
        return \ZExcel\Calculation\Functions::Na();
    }


    /**
     * INDEX
     *
     * Uses an index to choose a value from a reference or array
     *
     * Excel Function:
     *        =INDEX(range_array, row_num, [column_num])
     *
     * @param    range_array        A range of cells or an array constant
     * @param    row_num            The row in array from which to return a value. If row_num is omitted, column_num is required.
     * @param    column_num        The column in array from which to return a value. If column_num is omitted, row_num is required.
     * @return    mixed            the value of a specified cell or array of cells
     */
    public static function index(var arrayValues, var rowNum = 0, var columnNum = 0)
    {
        /*
        if ((rowNum < 0) || (columnNum < 0)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        if (!is_array(arrayValues)) {
            return \ZExcel\Calculation\Functions::ReF();
        }

        let rowKeys = array_keys(arrayValues);
        let columnKeys = @array_keys(arrayValues[rowKeys[0]]);

        if (columnNum > count(columnKeys)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        } elseif (columnNum == 0) {
            
            if (rowNum == 0) {
                return arrayValues;
            }
            
            let rowNum = rowNum - 1;
            let rowNum = rowKeys[rowNum];
            let returnArray = [];
            
            for arrayColumn in arrayValues {
                if (is_array(arrayColumn)) {
                    if (isset(arrayColumn[rowNum])) {
                        returnArray[] = arrayColumn[rowNum];
                    } else {
                        return arrayValues[rowNum];
                    }
                } else {
                    return arrayValues[rowNum];
                }
            }
            
            return returnArray;
        }
        
        let columnNum = columnNum - 1;
        let columnNum = columnKeys[columnNum];
        
        if (rowNum > count(rowKeys)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        } elseif (rowNum == 0) {
            return arrayValues[columnNum];
        }
        
        let rowNum = rowNum - 1;
        let rowNum = rowKeys[rowNum];

        return arrayValues[rowNum][columnNum];
        */
    }


    /**
     * TRANSPOSE
     *
     * @param    array    matrixData    A matrix of values
     * @return    array
     *
     * Unlike the Excel TRANSPOSE function, which will only work on a single row or column, this function will transpose a full matrix.
     */
    public static function transpose(array matrixData)
    {
        var matrixCell, matrixRow;
        array returnMatrix = [];
        int column, row;
        
        if (!is_array(matrixData)) {
            let matrixData = [[matrixData]];
        }

        let column = 0;
        
        for matrixRow in matrixData {
            let row = 0;
            
            for matrixCell in matrixRow {
                let returnMatrix[row][column] = matrixCell;
                let row = row + 1;
            }
            
            let column = column + 1;
        }
        
        return returnMatrix;
    }


    public static function vlookupSort(a, b)
    {
        var firstColumn;
        
        reset(a);
        
        let firstColumn = key(a);
        
        if (strtolower(a[firstColumn]) == strtolower(b[firstColumn])) {
            return 0;
        }
        
        return (strtolower(a[firstColumn]) < strtolower(b[firstColumn])) ? -1 : 1;
    }


    /**
     * VLOOKUP
     * The VLOOKUP function searches for value in the left-most column of lookup_array and returns the value in the same row based on the index_number.
     * @param    lookup_value    The value that you want to match in lookup_array
     * @param    lookup_array    The range of cells being searched
     * @param    index_number    The column number in table_array from which the matching value must be returned. The first column is 1.
     * @param    not_exact_match    Determines if you are looking for an exact match based on lookup_value.
     * @return    mixed            The value of the found cell
     */
    public static function vlookup(var lookup_value, var lookup_array, var index_number, var not_exact_match = true)
    {
        var f, firstRow, columnKeys, returnColumn, firstColumn, rowNumber, rowValue, rowKey, rowData;
        
        let lookup_value    = \ZExcel\Calculation\Functions::flattenSingleValue(lookup_value);
        let index_number    = \ZExcel\Calculation\Functions::flattenSingleValue(index_number);
        let not_exact_match    = \ZExcel\Calculation\Functions::flattenSingleValue(not_exact_match);

        // index_number must be greater than or equal to 1
        if (index_number < 1) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // index_number must be less than or equal to the number of columns in lookup_array
        if ((!is_array(lookup_array)) || (empty(lookup_array))) {
            return \ZExcel\Calculation\Functions::ReF();
        } else {
            let f = array_keys(lookup_array);
            let firstRow = array_pop(f);
            
            if ((!is_array(lookup_array[firstRow])) || (index_number > count(lookup_array[firstRow]))) {
                return \ZExcel\Calculation\Functions::ReF();
            } else {
                let columnKeys = array_keys(lookup_array[firstRow]);
                let index_number = index_number - 1;
                let returnColumn = columnKeys[index_number];
                let firstColumn = array_shift(columnKeys);
            }
        }

        if (!not_exact_match) {
            uasort(lookup_array, ["\\ZExcel\\Calculation\\LookupRef", "vlookupSort"]);
        }

        let rowNumber = false;
        let rowValue = false;
        
        for rowKey, rowData in lookup_array {
            
            if ((is_numeric(lookup_value) && is_numeric(rowData[firstColumn]) && (rowData[firstColumn] > lookup_value)) || (!is_numeric(lookup_value) && !is_numeric(rowData[firstColumn]) && (strtolower(rowData[firstColumn]) > strtolower(lookup_value)))) {
                break;
            }
            
            let rowNumber = rowKey;
            let rowValue = rowData[firstColumn];
        }

        if (rowNumber !== false) {
            if ((!not_exact_match) && (rowValue != lookup_value)) {
                //    if an exact match is required, we have what we need to return an appropriate response
                return \ZExcel\Calculation\Functions::Na();
            } else {
                //    otherwise return the appropriate value
                return lookup_array[rowNumber][returnColumn];
            }
        }

        return \ZExcel\Calculation\Functions::Na();
    }


    /**
     * HLOOKUP
     * The HLOOKUP function searches for value in the top-most row of lookup_array and returns the value in the same column based on the index_number.
     * @param    lookup_value    The value that you want to match in lookup_array
     * @param    lookup_array    The range of cells being searched
     * @param    index_number    The row number in table_array from which the matching value must be returned. The first row is 1.
     * @param    not_exact_match Determines if you are looking for an exact match based on lookup_value.
     * @return   mixed           The value of the found cell
     */
    public static function hlookup(lookup_value, lookup_array, index_number, not_exact_match = true)
    {
        var f, firstRow, firstKey, columnKeys, returnColumn, firstColumn, firstRowH, rowNumber, rowValue, rowKey, rowData;
        
        let lookup_value    = \ZExcel\Calculation\Functions::flattenSingleValue(lookup_value);
        let index_number    = \ZExcel\Calculation\Functions::flattenSingleValue(index_number);
        let not_exact_match = \ZExcel\Calculation\Functions::flattenSingleValue(not_exact_match);

        // index_number must be greater than or equal to 1
        if (index_number < 1) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        // index_number must be less than or equal to the number of columns in lookup_array
        if ((!is_array(lookup_array)) || (empty(lookup_array))) {
            return \ZExcel\Calculation\Functions::ReF();
        } else {
            let f = array_keys(lookup_array);
            let firstRow = array_pop(f);
            
            if ((!is_array(lookup_array[firstRow])) || (index_number > count(lookup_array[firstRow]))) {
                return \ZExcel\Calculation\Functions::ReF();
            } else {
                let columnKeys = array_keys(lookup_array[firstRow]);
                let firstKey = f[0] - 1;
                let returnColumn = firstKey + index_number;
                let firstColumn = array_shift(f);
            }
        }

        if (!not_exact_match) {
            let firstRowH = asort(lookup_array[firstColumn]);
        }

        let rowNumber = false;
        let rowValue = false;
        
        for rowKey, rowData in lookup_array[firstColumn] {
            if ((is_numeric(lookup_value) && is_numeric(rowData) && (rowData > lookup_value)) || (!is_numeric(lookup_value) && !is_numeric(rowData) && (strtolower(rowData) > strtolower(lookup_value)))) {
                break;
            }
            
            let rowNumber = rowKey;
            let rowValue = rowData;
        }

        if (rowNumber !== false) {
            if ((!not_exact_match) && (rowValue != lookup_value)) {
                //  if an exact match is required, we have what we need to return an appropriate response
                return \ZExcel\Calculation\Functions::Na();
            } else {
                //  otherwise return the appropriate value
                return lookup_array[returnColumn][rowNumber];
            }
        }

        return \ZExcel\Calculation\Functions::Na();
    }


    /**
     * LOOKUP
     * The LOOKUP function searches for value either from a one-row or one-column range or from an array.
     * @param    lookup_value    The value that you want to match in lookup_array
     * @param    lookup_vector    The range of cells being searched
     * @param    result_vector    The column from which the matching value must be returned
     * @return    mixed            The value of the found cell
     */
    public static function lookup(var lookup_value, var lookup_vector, var result_vector = null)
    {
        var lookupRows, l, lookupColumns, resultRows, resultColumns, r, k, value, key1, key2, dataValue1, dataValue2;
        
        let lookup_value = \ZExcel\Calculation\Functions::flattenSingleValue(lookup_value);

        if (!is_array(lookup_vector)) {
            return \ZExcel\Calculation\Functions::Na();
        }
        
        let lookupRows = count(lookup_vector);
        let l = array_keys(lookup_vector);
        let l = array_shift(l);
        let lookupColumns = count(lookup_vector[l]);
        
        if (((lookupRows == 1) && (lookupColumns > 1)) || ((lookupRows == 2) && (lookupColumns != 2))) {
            let lookup_vector = self::TRaNSPOSE(lookup_vector);
            let lookupRows = count(lookup_vector);
            let l = array_keys(lookup_vector);
            let lookupColumns = count(lookup_vector[array_shift(l)]);
        }

        if (is_null(result_vector)) {
            let result_vector = lookup_vector;
        }
        
        let resultRows = count(result_vector);
        let l = array_keys(result_vector);
        let l = array_shift(l);
        let resultColumns = count(result_vector[l]);
        
        if (((resultRows == 1) && (resultColumns > 1)) || ((resultRows == 2) && (resultColumns != 2))) {
            let result_vector = self::TRaNSPOSE(result_vector);
            let resultRows = count(result_vector);
            let r = array_keys(result_vector);
            let resultColumns = count(result_vector[array_shift(r)]);
        }

        if (lookupRows == 2) {
            let result_vector = array_pop(lookup_vector);
            let lookup_vector = array_shift(lookup_vector);
        }
        
        if (lookupColumns != 2) {
            for k, value in lookup_vector {
                if (is_array(value)) {
                    let k = array_keys(value);
                    let key1 = array_shift(k);
                    let key2 = key1 + 1;
                    let dataValue1 = value[key1];
                } else {
                    let key1 = 0;
                    let key2 = 1;
                    let dataValue1 = value;
                }
                
                let dataValue2 = array_shift(result_vector);
                
                if (is_array(dataValue2)) {
                    let dataValue2 = array_shift(dataValue2);
                }
                
                let lookup_vector[k] = [key1: dataValue1, key2: dataValue2];
            }
        }

        return self::VLoOKUP(lookup_value, lookup_vector, 2);
    }
}
