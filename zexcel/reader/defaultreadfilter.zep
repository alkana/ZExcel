namespace ZExcel\Reader;

class DefaultReadFilter implements IReader
{
    /**
     * Should this cell be read?
     *
     * @param    $column           Column address (as a string value like "A", or "IV")
     * @param    $row              Row number
     * @param    $worksheetName    Optional worksheet name
     * @return   boolean
     */
    public function readCell(string column, int row, var worksheetName = "")
    {
        return true;
    }
}
