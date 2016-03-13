namespace ZExcel\Shared\Ole\Pps;

use ZExcel\Shared\OLE\Pps as OlePps;

class File extends OlePps
{
    /**
    * The constructor
    *
    * @access public
    * @param string name The name of the file (in Unicode)
    * @see OLE::Asc2Ucs()
    */
    public function __construct(name)
    {
        parent::__construct(null, name, \ZExcel\Shared\OLE::OLE_PPS_TYPE_FILE, null, null, null, null, null,"", []);
    }

    /**
    * Initialization method. Has to be called right after OLE_PPS_File().
    *
    * @access public
    * @return mixed true on success
    */
    public function init()
    {
        return true;
    }

    /**
    * Append data to PPS
    *
    * @access public
    * @param string data The data to append
    */
    public function append(data)
    {
        let this->_data = this->_data . data;
    }

    /**
     * Returns a stream for reading this file using fread() etc.
     * @return  resource  a read-only stream
     */
    public function getStream()
    {
        this->ole->getStream(this);
    }
}
