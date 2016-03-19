namespace ZExcel;

class ReferenceHelper
{
    private static _instance = null;

    public static function getInstance() -> <\ZExcel\ReferenceHelper>
    {
        if (self::_instance === null) {
            let self::_instance = new \ZExcel\ReferenceHelper();
        }
        return self::_instance;
    }

    protected function __construct() {}
    
    
    public static function columnSort(var a, var b)
    {
        return strcasecmp(strlen(a) . a, strlen(b) . b);
    }

    public static function columnReverseSort(var a, var b)
    {
        return 1 - strcasecmp(strlen(a) . a, strlen(b) . b);
    }

    public static function cellSort(var a, var b)
    {
        var ac, ar, bc, br;
        
        sscanf(a,"%[A-Z]%d", ac, ar);
        sscanf(b,"%[A-Z]%d", bc, br);

        if (ar == br) {
            return strcasecmp(strlen(ac) . ac, strlen(bc) . bc);
        }
        
        return (ar < br) ? -1 : 1;
    }

    public static function cellReverseSort(var a, var b)
    {
        var ac, ar, bc, br;
        
        sscanf(a,"%[A-Z]%d", ac, ar);
        sscanf(b,"%[A-Z]%d", bc, br);

        if (ar == br) {
            return 1 - strcasecmp(strlen(ac) . ac, strlen(bc) . bc);
        }
        
        return (ar < br) ? 1 : -1;
    }

    private static function cellAddressInDeleteRange(var cellAddress, var beforeRow, var pNumRows, var beforeColumnIndex, var pNumCols)
    {
        var cellColumn, cellRow, cellColumnIndex, tmp;
        
        let tmp = \ZExcel\Cell::coordinateFromString(cellAddress);
        let cellColumn = tmp[0];
        let cellRow = tmp[1];
        let cellColumnIndex = \ZExcel\Cell::columnIndexFromString(cellColumn);
        
        //    Is cell within the range of rows/columns if we're deleting
        if (pNumRows < 0 && (cellRow >= (beforeRow + pNumRows)) && (cellRow < beforeRow)) {
            return true;
        } else {
            if (pNumCols < 0 && (cellColumnIndex >= (beforeColumnIndex + pNumCols)) && (cellColumnIndex < beforeColumnIndex)) {
                return true;
            }
        }
        
        return false;
    }

    protected function _adjustPageBreaks(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustComments(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustHyperlinks(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustDataValidations(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustMergeCells(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustProtectedCells(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustColumnDimensions(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustRowDimensions(<\ZExcel\Worksheet> pSheet, var pBefore, var beforeColumnIndex, var pNumCols, var beforeRow, var pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function insertNewBefore(var pBefore = "A1", var pNumCols = 0, var pNumRows = 0, <ZExcel\Worksheet> pSheet = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function updateFormulaReferences(var pFormula = "", var pBefore = "A1", var pNumCols = 0, var pNumRows = 0, var sheetName = "")
    {
        throw new \Exception("Not implemented yet!");
    }

    public function updateCellReference(var pCellRange = "A1", var pBefore = "A1", var pNumCols = 0, var pNumRows = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function updateNamedFormulas(<\ZExcel> pPhpExcel, var oldName = "", var newName = "")
    {
        throw new \Exception("Not implemented yet!");
    }

    private function _updateCellRange(var pCellRange = "A1:A1", var pBefore = "A1", var pNumCols = 0, var pNumRows = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    private function _updateSingleCellReference(var pCellReference = "A1", var pBefore = "A1", var pNumCols = 0, var pNumRows = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    public final function __clone()
    {
        throw new \ZExcel\Exception("Cloning a Singleton is not allowed!");
    }
}
