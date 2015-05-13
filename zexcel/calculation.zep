namespace ZExcel;

class Calculation
{
	private function __construct(<ZExcel\ZExcel> workbook = null)
	{
		int setPrecision = 16;
		
		if (PHP_INT_SIZE == 4) {
			let setPrecision = 14;
		}
		
        let this->_savedPrecision = ini_get("precision");
        
        if (this->_savedPrecision < setPrecision) {
            ini_set("precision", setPrecision);
        }
        
        let this->delta = 1 * pow(10, -setPrecision);
        
        if (workbook !== null) {
            let self::_workbookSets[workbook->getID()] = this;
        }
        let this->_workbook = workbook;
        let this->_cyclicReferenceStack = new \ZExcel\CalcEngine\CyclicReferenceStack();
        let this->_debugLog = new \ZExcel\CalcEngine\Logger(this->_cyclicReferenceStack);
    }
    
	
    public function __destruct()
	{
        if (this->_savedPrecision != ini_get("precision")) {
            ini_set("precision", this->_savedPrecision);
        }
    }
    
    /**
     * Get an instance of this class
     *
     * @access  public
     * @param   PHPExcel $workbook  Injected workbook for working with a PHPExcel object,
     *                                    or NULL to create a standalone claculation engine
     * @return PHPExcel_Calculation
     */
    public static function getInstance(<\ZExcel\ZExcel> workbook = NULL) -> <\ZExcel\Calculation>
    {
        if (workbook !== null) {
            if (isset(self::_workbookSets[workbook->getID()])) {
                return self::_workbookSets[workbook->getID()];
            }
            return new \ZExcel\Calculation(workbook);
        }
        
        if (!isset(self::_instance) || (self::_instance === null)) {
            let self::_instance = new \ZExcel\Calculation();
        }
        
        return self::_instance;
    }
    
    /**
     * Unset an instance of this class
     *
     * @access    public
     * @param   PHPExcel $workbook  Injected workbook identifying the instance to unset
     */
    public static function unsetInstance(<ZExcel\ZExcel> workbook = null) {
        if (workbook !== null) {
            if (isset(self::_workbookSets[workbook->getID()])) {
                unset(self::_workbookSets[workbook->getID()]);
            }
        }
    }
}
