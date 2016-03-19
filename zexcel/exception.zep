namespace ZExcel;

class Exception extends \Exception
{
    /**
     * Error handler callback
     *
     * @param mixed $code
     * @param mixed $string
     * @param mixed $file
     * @param mixed $line
     * @param mixed $context
     */
    public static function errorHandlerCallback(var code, var stringg, var file, var line, varcontext)
    {
        var e;
        
        let e = new \ZExcel\Exception(stringg, code);
        let e->line = line;
        let e->file = file;
        
        throw e;
    }
}
