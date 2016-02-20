namespace ZExcel;

class Autoloader
{
    /**
     * @Depreciated Only require declaration for phpunit test
     */
    public static function register()
    {
    }

    /**
     * @Depreciated Only require declaration for phpunit test
     *
     * Simulate autoload for compatibility
     */
    public static function load(string pClassName) -> boolean
    {
        if ((class_exists(pClassName, false)) || (strpos(pClassName, "ZExcel") !== 0)) {
            // Either already loaded, or not a PHPExcel class request
            return false;
        }
        
        return true;
    }
}
