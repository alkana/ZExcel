namespace ZExcel\Cell;

interface IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  PHPExcel_Cell $cell    Cell to bind value to
     * @param  mixed $value           Value to bind in cell
     * @return boolean
     */
    public function bindValue(<\ZExcel\Cell> cell, var value = null);
}
