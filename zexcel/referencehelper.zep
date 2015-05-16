namespace ZExcel;

class ReferenceHelper
{
    private static $_instance = null;

    public static function getInstance() -> <\ZExcel\ReferenceHelper>
    {
        if (self::$_instance === null) {
            let self::$_instance = new \ZExcel\ReferenceHelper();
        }
        return self::$_instance;
    }

    protected function __construct() {}
    
    
    public static function columnSort(a, b)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function columnReverseSort(a, b)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function cellSort(a, b)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function cellReverseSort(a, b)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function cellAddressInDeleteRange(cellAddress, beforeRow, pNumRows, beforeColumnIndex, pNumCols)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustPageBreaks(<\ZExcel\Worksheet> pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustComments(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustHyperlinks(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustDataValidations(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustMergeCells(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustProtectedCells(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustColumnDimensions(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    protected function _adjustRowDimensions(pSheet, pBefore, beforeColumnIndex, pNumCols, beforeRow, pNumRows)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function insertNewBefore(pBefore = "A1", pNumCols = 0, pNumRows = 0, <ZExcel\Worksheet> pSheet = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function updateFormulaReferences(pFormula = "", pBefore = "A1", pNumCols = 0, pNumRows = 0, sheetName = "") {
        throw new \Exception("Not implemented yet!");
    }

    public function updateCellReference(pCellRange = "A1", pBefore = "A1", pNumCols = 0, pNumRows = 0) {
        throw new \Exception("Not implemented yet!");
    }

    public function updateNamedFormulas(<\ZExcel> pPhpExcel, oldName = "", newName = "") {
        throw new \Exception("Not implemented yet!");
    }

    private function _updateCellRange(pCellRange = "A1:A1", pBefore = "A1", pNumCols = 0, pNumRows = 0) {
        throw new \Exception("Not implemented yet!");
    }

    private function _updateSingleCellReference(pCellReference = "A1", pBefore = "A1", pNumCols = 0, pNumRows = 0) {
        throw new \Exception("Not implemented yet!");
    }

    public final function __clone() {
        throw new \ZExcel\Exception("Cloning a Singleton is not allowed!");
    }
}
