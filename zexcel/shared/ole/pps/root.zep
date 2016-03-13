namespace ZExcel\Shared\Ole\Pps;

use ZExcel\Shared\OLE\Pps as OlePps;

/**
 * @FIXME WARNING, LOT OF REFERENCES ARGUMENTS
 */
class Root extends OlePps
{

    /**
     * Directory for temporary files
     * @var string
     */
    protected tempDirectory = null;

    /**
     * @param integer time_1st A timestamp
     * @param integer time_2nd A timestamp
     */
    public function __construct(time_1st, time_2nd, raChild)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Method for saving the whole OLE container (including files).
    * In fact, if called with an empty argument (or '-'), it saves to a
    * temporary file and then outputs it's contents to stdout.
    * If a resource pointer to a stream created by fopen() is passed
    * it will be used, but you have to close such stream by yourself.
    *
    * @param string|resource filename The name of the file or stream where to save the OLE container.
    * @access public
    * @return mixed true on success
    */
    public function save(filename)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Calculate some numbers
    *
    * @access public
    * @param array raList Reference to an array of PPS's
    * @return array The array of numbers
    */
    public function _calcSize(raList)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Helper function for caculating a magic value for block sizes
    *
    * @access public
    * @param integer i2 The argument
    * @see save()
    * @return integer
    */
    private static function adjust2(i2)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Save OLE header
    *
    * @access public
    * @param integer iSBDcnt
    * @param integer iBBcnt
    * @param integer iPPScnt
    */
    public function _saveHeader(iSBDcnt, iBBcnt, iPPScnt)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Saving big data (PPS's with data bigger than PHPExcel_Shared_OLE::OLE_DATA_SIZE_SMALL)
    *
    * @access public
    * @param integer iStBlk
    * @param array &raList Reference to array of PPS's
    */
    public function _saveBigData(iStBlk, raList)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * get small data (PPS's with data smaller than PHPExcel_Shared_OLE::OLE_DATA_SIZE_SMALL)
    *
    * @access public
    * @param array &raList Reference to array of PPS's
    */
    public function _makeSmallData(raList)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Saves all the PPS's WKs
    *
    * @access public
    * @param array raList Reference to an array with all PPS's
    */
    public function _savePps(raList)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Saving Big Block Depot
    *
    * @access public
    * @param integer iSbdSize
    * @param integer iBsize
    * @param integer iPpsCnt
    */
    public function _saveBbd(iSbdSize, iBsize, iPpsCnt)
    {
        throw new \Exception("Not implemented yet!");
    }
}
