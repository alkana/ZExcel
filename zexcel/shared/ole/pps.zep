namespace ZExcel\Shared\OLE;

class Pps
{
    /**
    * The PPS index
    * @var integer
    */
    public No;

    /**
    * The PPS name (in Unicode)
    * @var string
    */
    public Name;

    /**
    * The PPS type. Dir, Root or File
    * @var integer
    */
    public Type;

    /**
    * The index of the previous PPS
    * @var integer
    */
    public PrevPps;

    /**
    * The index of the next PPS
    * @var integer
    */
    public NextPps;

    /**
    * The index of it's first child if this is a Dir or Root PPS
    * @var integer
    */
    public DirPps;

    /**
    * A timestamp
    * @var integer
    */
    public Time1st;

    /**
    * A timestamp
    * @var integer
    */
    public Time2nd;

    /**
    * Starting block (small or big) for this PPS's data  inside the container
    * @var integer
    */
    public _StartBlock;

    /**
    * The size of the PPS's data (in bytes)
    * @var integer
    */
    public Size;

    /**
    * The PPS's data (only used if it's not using a temporary file)
    * @var string
    */
    public _data;

    /**
    * Array of child PPS's (only used by Root and Dir PPS's)
    * @var array
    */
    public children = [];

    /**
    * Pointer to OLE container
    * @var OLE
    */
    public ole;

    /**
    * The constructor
    *
    * @access public
    * @param integer No   The PPS index
    * @param string  name The PPS name
    * @param integer type The PPS type. Dir, Root or File
    * @param integer prev The index of the previous PPS
    * @param integer next The index of the next PPS
    * @param integer dir  The index of it's first child if this is a Dir or Root PPS
    * @param integer time_1st A timestamp
    * @param integer time_2nd A timestamp
    * @param string  data  The (usually binary) source data of the PPS
    * @param array   children Array containing children PPS for this PPS
    */
    public function __construct(No, name, type, prev, next, dir, time_1st, time_2nd, data, children)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Returns the amount of data saved for this PPS
    *
    * @access public
    * @return integer The amount of data (in bytes)
    */
    public function _DataLen()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Returns a string with the PPS's WK (What is a WK?)
    *
    * @access public
    * @return string The binary string
    */
    public function _getPpsWk()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * Updates index and pointers to previous, next and children PPS's for this
    * PPS. I don't think it'll work with Dir PPS's.
    *
    * @FIXME REFERENCES ARGUMENTS (raList)
    *
    * @access public
    * @param array &raList Reference to the array of PPS's for the whole OLE
    *                          container
    * @return integer          The index for this PPS
    */
    public static function _savePpsSetPnt(raList, to_save, depth = 0)
    {
        throw new \Exception("Not implemented yet!");
    }
}
