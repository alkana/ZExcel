namespace ZExcel\CalcEngine;

class Logger
{
    /**
     * Flag to determine whether a debug log should be generated by the calculation engine
     *        If true, then a debug log will be generated
     *        If false, then a debug log will not be generated
     *
     * @var boolean
     */
    private writeDebugLog = false;

    /**
     * Flag to determine whether a debug log should be echoed by the calculation engine
     *        If true, then a debug log will be echoed
     *        If false, then a debug log will not be echoed
     * A debug log can only be echoed if it is generated
     *
     * @var boolean
     */
    private echoDebugLog = false;

    /**
     * The debug log generated by the calculation engine
     *
     * @var string[]
     */
    private debugLog = [];

    /**
     * The calculation engine cell reference stack
     *
     * @var \ZExcel\CalcEngine\CyclicReferenceStack
     */
    private cellStack;

    /**
     * Instantiate a Calculation engine logger
     *
     * @param  \ZExcel\CalcEngine\CyclicReferenceStack stack
     */
    public function __construct(<\ZExcel\CalcEngine\CyclicReferenceStack> stack)
    {
        let this->cellStack = stack;
    }

    /**
     * Enable/Disable Calculation engine logging
     *
     * @param  boolean pValue
     */
    public function setWriteDebugLog(pValue = false)
    {
        let this->writeDebugLog = pValue;
    }

    /**
     * Return whether calculation engine logging is enabled or disabled
     *
     * @return  boolean
     */
    public function getWriteDebugLog()
    {
        return this->writeDebugLog;
    }

    /**
     * Enable/Disable echoing of debug log information
     *
     * @param  boolean pValue
     */
    public function setEchoDebugLog(pValue = false)
    {
        let this->echoDebugLog = pValue;
    }

    /**
     * Return whether echoing of debug log information is enabled or disabled
     *
     * @return  boolean
     */
    public function getEchoDebugLog()
    {
        return this->echoDebugLog;
    }

    /**
     * Write an entry to the calculation engine debug log
     */
    public function writeDebugLog()
    {
        var message, cellReference;
        
        //    Only write the debug log if logging is enabled
        if (this->writeDebugLog) {
            let message = implode("", func_get_args());
            let cellReference = implode(" -> ", this->cellStack->showStack());
            if (this->echoDebugLog) {
                echo cellReference, (this->cellStack->count() > 0 ? " => " : ""), message, PHP_EOL;
            }
            let this->debugLog[] = cellReference . (this->cellStack->count() > 0 ? " => " : "") . message;
        }
    }

    /**
     * Clear the calculation engine debug log
     */
    public function clearLog()
    {
        let this->debugLog = [];
    }

    /**
     * Return the calculation engine debug log
     *
     * @return  string[]
     */
    public function getLog()
    {
        return this->debugLog;
    }
}
