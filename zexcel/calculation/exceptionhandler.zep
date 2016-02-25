namespace ZExcel\Calculation;

class ExceptionHandler
{
    /**
     * Register errorhandler
     */
    public function __construct()
    {
        set_error_handler(["\\PHPExcel\\Calculation\\Exception", "errorHandlerCallback"], E_ALL);
    }

    /**
     * Unregister errorhandler
     */
    public function __destruct()
    {
        restore_error_handler();
    }
}
