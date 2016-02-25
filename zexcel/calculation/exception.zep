namespace ZExcel\Calculation;

use ZExcel\Exception as ZExcelException;

class Exception extends ZExcelException
{
    /**
     * Error handler callback
     *
     * @param mixed code
     * @param mixed string
     * @param mixed file
     * @param mixed line
     * @param mixed context
     */
    public static function errorHandlerCallback(var code,var stringg, var file, var line, var context)
    {
        var e;
        
        let e = new self(stringg, code);
        
        let e->line = line;
        let e->file = file;
        
        throw e;
    }
}
