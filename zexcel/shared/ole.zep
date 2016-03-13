namespace ZExcel\Shared;

class Ole
{
    const OLE_PPS_TYPE_ROOT   =      5;
    const OLE_PPS_TYPE_DIR    =      1;
    const OLE_PPS_TYPE_FILE   =      2;
    const OLE_DATA_SIZE_SMALL = 0x1000;
    const OLE_LONG_INT_SIZE   =      4;
    const OLE_PPS_SIZE        =   0x80;

    /**
     * The file handle for reading an OLE container
     * @var resource
    */
    public _file_handle;

    /**
    * Array of PPS"s found on the OLE container
    * @var array
    */
    public _list = [];

    /**
     * Root directory of OLE container
     * @var OLE\PPS\Root
    */
    public root;

    /**
     * Big Block Allocation Table
     * @var array  (blockId => nextBlockId)
    */
    public bbat;

    /**
     * Short Block Allocation Table
     * @var array  (blockId => nextBlockId)
    */
    public sbat;

    /**
     * Size of big blocks. This is usually 512.
     * @var  int  number of octets per block.
    */
    public bigBlockSize;

    /**
     * Size of small blocks. This is usually 64.
     * @var  int  number of octets per block
    */
    public smallBlockSize;

    public static instances = [];
    
    public static isRegistered = false;
    /**
     * Reads an OLE container from the contents of the file given.
     *
     * @acces public
     * @param string file
     * @return mixed true on success, \Pear\Error on failure
    */
    public function read(string file)
    {
        var fh, signature, bbatBlockCount, directoryFirstBlockId, sbatFirstBlockId, sbbatBlockCount,
            mbatBlocks, mbatFirstBlockId, mbbatBlockCount, pos, i, j,
            shortBlockCount, sbatFh, blockId;
        
        let fh = fopen(file, "r");
        
        if (!fh) {
            throw new \ZExcel\Reader\Exception("Can't open file file");
        }
        
        let this->_file_handle = fh;

        let signature = fread(fh, 8);
        
        if ("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" != signature) {
            throw new \ZExcel\Reader\Exception("File doesn't seem to be an OLE container.");
        }
        
        fseek(fh, 28);
        
        if (fread(fh, 2) != "\xFE\xFF") {
            // This shouldn"t be a problem in practice
            throw new \ZExcel\Reader\Exception("Only Little-Endian encoding is supported.");
        }
        
        // Size of blocks and short blocks in bytes
        let this->bigBlockSize = pow(2, self::_readInt2(fh));
        let this->smallBlockSize  = pow(2, self::_readInt2(fh));

        // Skip UID, revision number and version number
        fseek(fh, 44);
        // Number of blocks in Big Block Allocation Table
        let bbatBlockCount = self::_readInt4(fh);

        // Root chain 1st block
        let directoryFirstBlockId = self::_readInt4(fh);

        // Skip unused bytes
        fseek(fh, 56);
        // Streams shorter than this are stored using small blocks
        let this->bigBlockThreshold = self::_readInt4(fh);
        // Block id of first sector in Short Block Allocation Table
        let sbatFirstBlockId = self::_readInt4(fh);
        // Number of blocks in Short Block Allocation Table
        let sbbatBlockCount = self::_readInt4(fh);
        // Block id of first sector in Master Block Allocation Table
        let mbatFirstBlockId = self::_readInt4(fh);
        // Number of blocks in Master Block Allocation Table
        let mbbatBlockCount = self::_readInt4(fh);
        
        let this->bbat = [];

        // Remaining 4 * 109 bytes of current block is beginning of Master
        // Block Allocation Table
        let mbatBlocks = [];
        
        for i in range(0, 108) {
            let mbatBlocks[] = self::_readInt4(fh);
        }

        // Read rest of Master Block Allocation Table (if any is left)
        let pos = this->_getBlockOffset(mbatFirstBlockId);
        
        for i in range(0, mbbatBlockCount - 1) {
            fseek(fh, pos);
            
            for j in range(0, (this->bigBlockSize / 4) - 2) {
                let mbatBlocks[] = self::_readInt4(fh);
            }
            
            // Last block id in each block points to next block
            let pos = this->_getBlockOffset(self::_readInt4(fh));
        }

        // Read Big Block Allocation Table according to chain specified by
        // mbatBlocks
        for i in range(0, bbatBlockCount - 1) {
            let pos = this->_getBlockOffset(mbatBlocks[i]);
            
            fseek(fh, pos);
            
            for j in range(0, (this->bigBlockSize / 4) - 1) {
                let this->bbat[] = self::_readInt4(fh);
            }
        }

        // Read short block allocation table (SBAT)
        let this->sbat = [];
        
        let shortBlockCount = sbbatBlockCount * this->bigBlockSize / 4;
        let sbatFh = this->getStream(sbatFirstBlockId);
        
        for blockId in range(0, shortBlockCount - 1) {
            let this->sbat[blockId] = self::_readInt4(sbatFh);
        }
        
        fclose(sbatFh);

        this->_readPpsWks(directoryFirstBlockId);

        return true;
    }

    /**
     * @param  int  block id
     * @param  int  byte offset from beginning of file
     * @access public
     */
    public function _getBlockOffset(int blockId)
    {
        return 512 + blockId * this->bigBlockSize;
    }

    /**
    * Returns a stream for use with fread() etc. External callers should
    * use \ZExcel\Shared\OLE\PPS\File::getStream().
    * @param   int|PPS   block id or PPS
    * @return  resource  read-only stream
    */
    public function getStream(var blockIdOrPps)
    {
        var instanceId, path;
        
        if (!self::isRegistered) {
            stream_wrapper_register("ole-chainedblockstream", "\ZExcel\Shared\OLE\ChainedBlockStream");
            let self::isRegistered = true;
        }

        // Store current instance in global array, so that it can be accessed
        // in OLE\ChainedBlockStream::stream_open().
        // Object is removed from self::instances in OLE\Stream::close().
        array_push(self::instances, this);
        let instanceId = end(array_keys(self::instances));

        let path = "ole-chainedblockstream://oleInstanceId=" . instanceId;
        if (blockIdOrPps instanceof \ZExcel\Shared\OLE\Pps) {
            let path = path . "&blockId=" . blockIdOrPps->_StartBlock;
            let path = path . "&size=" . blockIdOrPps->Size;
        } else {
            let path = path . "&blockId=" . blockIdOrPps;
        }
        
        return fopen(path, "r");
    }

    /**
     * Reads a signed char.
     * @param   resource  file handle
     * @return  int
     * @access public
     */
    private static function _readInt1(var fh)
    {
        var tmp;
        
        let tmp = unpack("c", fread(fh, 1));
        
        return tmp[1];
    }

    /**
     * Reads an unsigned short (2 octets).
     * @param   resource  file handle
     * @return  int
     * @access public
     */
    private static function _readInt2(var fh)
    {
        var tmp;
        
        let tmp = unpack("v", fread(fh, 2));
        
        return tmp[1];
    }

    /**
     * Reads an unsigned long (4 octets).
     * @param   resource  file handle
     * @return  int
     * @access public
     */
    private static function _readInt4(var fh)
    {
        var tmp;
        
        let tmp = unpack("V", fread(fh, 4));
        
        return tmp[1];
    }

    /**
    * Gets information about all PPS"s on the OLE container from the PPS WK"s
    * creates an OLE\PPS object for each one.
    *
    * @access public
    * @param  integer  the block id of the first block
    * @return mixed true on success, \Pear\Error on failure
    */
    public function _readPpsWks(int blockId)
    {
        var fh, pos, nameLength, nameUtf16, name, type, pps, childPps, nos, no;
        
        let fh = this->getStream(blockId);
        let pos = 0;
        
        while (true) {
            fseek(fh, pos, SEEK_SET);
            
            let nameUtf16 = fread(fh, 64);
            let nameLength = self::_readInt2(fh);
            let nameUtf16 = substr(nameUtf16, 0, nameLength - 2);
            
            // Simple conversion from UTF-16LE to ISO-8859-1
            let name = str_replace(chr(0), "", nameUtf16);
            let type = self::_readInt1(fh);
            
            /* @FIXME compilation Error
            switch (type) {
                case self::OLE_PPS_TYPE_ROOT:
                    let pps = new \ZExcel\Shared\Ole\Pps\Root(null, null, []);
                    let this->root = pps;
                    break;
                case self::OLE_PPS_TYPE_DIR:
                    let pps = new \ZExcel\Shared\Ole\Pps(null, null, null, null, null, null, null, null, null, []);
                    break;
                case self::OLE_PPS_TYPE_FILE:
                    let pps = new \ZExcel\Shared\Ole\Pps\File(name);
                    break;
                default:
                    continue;
            }
            */
            
            fseek(fh, 1, SEEK_CUR);
            
            let pps->Type    = type;
            let pps->Name    = name;
            let pps->PrevPps = self::_readInt4(fh);
            let pps->NextPps = self::_readInt4(fh);
            let pps->DirPps  = self::_readInt4(fh);
            
            fseek(fh, 20, SEEK_CUR);
            
            let pps->Time1st = self::OLE2LocalDate(fread(fh, 8));
            let pps->Time2nd = self::OLE2LocalDate(fread(fh, 8));
            let pps->_StartBlock = self::_readInt4(fh);
            let pps->Size = self::_readInt4(fh);
            let pps->No = count(this->_list);
            let this->_list[] = pps;

            // check if the PPS tree (starting from root) is complete
            if (isset(this->root) && this->_ppsTreeComplete(this->root->No)) {
                break;
            }
        
            let pos = pos + 128;
        }
        
        fclose(fh);

        // Initialize pps->children on directories
        for pps in this->_list {
            if (pps->Type == self::OLE_PPS_TYPE_DIR || pps->Type == self::OLE_PPS_TYPE_ROOT) {
                let nos = [pps->DirPps];
                let pps->children = [];
                
                while (nos) {
                    let no = array_pop(nos);
                    
                    if (no != -1) {
                        let childPps = this->_list[no];
                        let nos[] = childPps->PrevPps;
                        let nos[] = childPps->NextPps;
                        let pps->children[] = childPps;
                    }
                }
            }
        }

        return true;
    }

    /**
    * It checks whether the PPS tree is complete (all PPS"s read)
    * starting with the given PPS (not necessarily root)
    *
    * @access public
    * @param integer index The index of the PPS from which we are checking
    * @return boolean Whether the PPS tree for the given PPS is complete
    */
    public function _ppsTreeComplete(int index)
    {
        var pps;
        
        if isset(this->_list[index]) {
            let pps = this->_list[index];
        }
        
        return !is_null(pps)
                    && (pps->PrevPps == -1 || this->_ppsTreeComplete(pps->PrevPps))
                    && (pps->NextPps == -1 || this->_ppsTreeComplete(pps->NextPps))
                    && (pps->DirPps == -1 || this->_ppsTreeComplete(pps->DirPps));
    }

    /**
    * Checks whether a PPS is a File PPS or not.
    * If there is no PPS for the index given, it will return false.
    *
    * @access public
    * @param integer index The index for the PPS
    * @return bool true if it"s a File PPS, false otherwise
    */
    public function isFile(int index)
    {
        if (isset(this->_list[index])) {
            return (this->_list[index]->Type == self::OLE_PPS_TYPE_FILE);
        }
        
        return false;
    }

    /**
    * Checks whether a PPS is a Root PPS or not.
    * If there is no PPS for the index given, it will return false.
    *
    * @access public
    * @param integer index The index for the PPS.
    * @return bool true if it"s a Root PPS, false otherwise
    */
    public function isRoot(int index)
    {
        if (isset(this->_list[index])) {
            return (this->_list[index]->Type == self::OLE_PPS_TYPE_ROOT);
        }
        
        return false;
    }

    /**
    * Gives the total number of PPS"s found in the OLE container.
    *
    * @access public
    * @return integer The total number of PPS"s found in the OLE container
    */
    public function ppsTotal()
    {
        return count(this->_list);
    }

    /**
    * Gets data from a PPS
    * If there is no PPS for the index given, it will return an empty string.
    *
    * @access public
    * @param integer index    The index for the PPS
    * @param integer position The position from which to start reading
    *                          (relative to the PPS)
    * @param integer length   The amount of bytes to read (at most)
    * @return string The binary string containing the data requested
    * @see OLE\PPS\File::getStream()
    */
    public function getData(int index, int position, int length)
    {
        var fh, data;
        
        // if position is not valid return empty string
        if (!isset(this->_list[index]) || (position >= this->_list[index]->Size) || (position < 0)) {
            return "";
        }
        
        let fh = this->getStream(this->_list[index]);
        let data = stream_get_contents(fh, length, position);
        
        fclose(fh);
        
        return data;
    }

    /**
    * Gets the data length from a PPS
    * If there is no PPS for the index given, it will return 0.
    *
    * @access public
    * @param integer index    The index for the PPS
    * @return integer The amount of bytes in data the PPS has
    */
    public function getDataLength(int index)
    {
        if (isset(this->_list[index])) {
            return this->_list[index]->Size;
        }
        
        return 0;
    }

    /**
    * Utility function to transform ASCII text to Unicode
    *
    * @access public
    * @static
    * @param string ascii The ASCII string to transform
    * @return string The string in Unicode
    */
    public static function Asc2Ucs(var ascii)
    {
        string rawname = "";
        int i;
        
        for i in range(0, strlen(ascii) - 1) {
            let rawname = rawname . substr(ascii, i, 1) . chr(0);
        }
        return rawname;
    }

    /**
    * Utility function
    * Returns a string for the OLE container with the date given
    *
    * @access public
    * @static
    * @param integer date A timestamp
    * @return string The string for the OLE container
    */
    public static function LocalDate2OLE(var date = null)
    {
        var factor, days, big_date, high_part, low_part, res, hex, i;
        
        if (empty(date)) {
            let i = chr(0);
            return i . i . i . i . i . i . i . i;
        }

        // factor used for separating numbers into 4 bytes parts
        let factor = pow(2, 32);

        // days from 1-1-1601 until the beggining of UNIX era
        let days = 134774;
        // calculate seconds
        let big_date = days*24*3600 + gmmktime(date("H", date), date("i", date), date("s", date), date("m", date), date("d", date), date("Y", date));
        // multiply just to make MS happy
        let big_date = big_date * 10000000;

        let high_part = floor(big_date / factor);
        // lower 4 bytes
        let low_part = floor(((big_date / factor) - high_part) * factor);

        // Make HEX string
        let res = "";

        for i in range(0, 3) {
            let hex = low_part % 0x100;
            let res = res . pack("c", hex);
            let low_part = low_part / 0x100;
        }
        
        for i in range(0, 3) {
            let hex = high_part % 0x100;
            let res = res . pack("c", hex);
            let high_part = high_part / 0x100;
        }
        
        return res;
    }

    /**
    * Returns a timestamp from an OLE container"s date
    *
    * @access public
    * @static
    * @param integer string A binary string with the encoded date
    * @return string The timestamp corresponding to the string
    */
    public static function OLE2LocalDate(var stringg)
    {
        var factor, high_part, low_part, big_date, days, tmp;
        
        if (strlen(stringg) != 8) {
            throw new \ZExcel\Exception("Expecting 8 byte string");
        }

        // factor used for separating numbers into 4 bytes parts
        let factor = pow(2, 32);
        
        let tmp = unpack("V", substr(stringg, 4, 4));
        let high_part = tmp[1];
        
        let tmp = unpack("V", substr(stringg, 0, 4));
        let low_part = tmp[1];

        let big_date = (high_part * factor) + low_part;
        // translate to seconds
        let big_date = big_date / 10000000;

        // days from 1-1-1601 until the beggining of UNIX era
        let days = 134774;

        // translate to seconds from beggining of UNIX era
        let big_date = big_date - (days * 24 * 3600);
        
        return floor(big_date);
    }
}
