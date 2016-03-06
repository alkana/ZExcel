namespace ZExcel;

class IOFactory
{
    /**
     * Search locations
     *
     * @var    array
     * @access    private
     * @static
     */
    private static _searchLocations = [
        [
            "type": "IWriter",
            "path": "ZExcel/Writer/{0}.php",
            "class": "\ZExcel\Writer_{0}"
        ],
        [
            "type": "IReader",
            "path": "ZExcel/Reader/{0}.php",
            "class": "\ZExcel\Reader\{0}"
        ]
    ];

    /**
     * Autoresolve classes
     *
     * @var    array
     * @access    private
     * @static
     */
    private static _autoResolveClasses = [
        "Excel2007",
        "Excel5",
        "Excel2003XML",
        "OOCalc",
        "SYLK",
        "Gnumeric",
        "HTML",
        "CSV"
    ];
    
    private function __construct() { }

    /**
     * Get search locations
     *
     * @static
     * @access    public
     * @return    array
     */
    public static function getSearchLocations()
    {
        return self::_searchLocations;
    }
    
    /**
     * Set search locations
     *
     * @static
     * @access    public
     * @param    array value
     * @throws    PHPExcel_Reader_Exception
     */
    public static function setSearchLocations(var value)
    {
        if (is_array(value)) {
            let self::_searchLocations = value;
        } else {
            throw new \ZExcel\Reader\Exception("Invalid parameter passed.");
        }
    }
    
    /**
     * Add search location
     *
     * @static
     * @access    public
     * @param    string type        Example: IWriter
     * @param    string location    Example: PHPExcel/Writer/{0}.php
     * @param    string classname     Example: PHPExcel_Writer_{0}
     */
    public static function addSearchLocation(var type = "", var location = "", var classname = "")
    {
        let self::_searchLocations[] = [
            "type": type,
            "path": location,
            "class": classname
        ];
    }
    
    /**
     * Create PHPExcel_Writer_IWriter
     *
     * @static
     * @access    public
     * @param    PHPExcel phpExcel
     * @param    string  writerType    Example: Excel2007
     * @return    PHPExcel_Writer_IWriter
     * @throws    PHPExcel_Reader_Exception
     */
    public static function createWriter(<ZExcel> zExcel, string writerType = "") -> <Writer\IWriter>
    {
        // Search type
        var instance;
        string className = "\\ZExcel\\Writer\\";
        
        let className = className . ucfirst(writerType);
        
        // Include class
        if class_exists(className) {
            let instance = <Writer\IWriter> new {className}(zExcel);
            
            if (instance !== NULL) {
                return instance;
            }
        }
        
        // Nothing found...
        throw new Reader\Exception("No IWriter found for type " . writerType);
    }

    public static function createReader(string readerType = "") -> <Reader\IReader>
    {
        // Search type
        var instance;
        string className = "\\ZExcel\\Reader\\";
        
        let className = className . ucfirst(readerType);
        
        // Include class
        if class_exists(className) {
            let instance = <Reader\IReader> new {className}();
            
            if (instance !== NULL) {
                return instance;
            }
        }
        
        // Nothing found...
        throw new Reader\Exception("No IReader found for type " . readerType);
    }

    /**
     * Loads PHPExcel from file using automatic PHPExcel_Reader_IReader resolution
     *
     * @static
     * @access public
     * @param     string         $pFilename        The name of the spreadsheet file
     * @return    PHPExcel
     * @throws    PHPExcel_Reader_Exception
     */
    public static function load(string pFilename) -> <Reader\IReader>
    {
        var reader;
        
        let reader = <Reader\IReader> self::createReaderForFile(pFilename);
        
        return reader->load(pFilename);
    }

    /**
     * Identify file type using automatic PHPExcel_Reader_IReader resolution
     *
     * @static
     * @access public
     * @param     string         $pFilename        The name of the spreadsheet file to identify
     * @return    string
     * @throws    PHPExcel_Reader_Exception
     */
    public static function identify(var pFilename)
    {
        var reader, className, classType;
        
        let reader = self::createReaderForFile(pFilename);
        let className = get_class(reader);
        let classType = explode("\\", className);
        
        return array_pop(classType);
    }

    /**
     * Create PHPExcel_Reader_IReader for file using automatic PHPExcel_Reader_IReader resolution
     *
     * @static
     * @access    public
     * @param     string         $pFilename        The name of the spreadsheet file
     * @return    PHPExcel_Reader_IReader
     * @throws    PHPExcel_Reader_Exception
     */
    public static function createReaderForFile(string pFilename) -> <Reader\IReader>
    {
        var pathinfo = null, reader = null;
        string extensionType = null;
        
        let pathinfo = pathinfo(pFilename);
        
        if (isset(pathinfo["extension"])) {
            switch (strtolower(pathinfo["extension"])) {
                case "xlsx":            //    Excel (OfficeOpenXML) Spreadsheet
                case "xlsm":            //    Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
                case "xltx":            //    Excel (OfficeOpenXML) Template
                case "xltm":            //    Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                    let extensionType = "Excel2007";
                    break;
                case "xls":                //    Excel (BIFF) Spreadsheet
                case "xlt":                //    Excel (BIFF) Template
                    let extensionType = "Excel5";
                    break;
                case "ods":                //    Open/Libre Offic Calc
                case "ots":                //    Open/Libre Offic Calc Template
                    let extensionType = "OOCalc";
                    break;
                case "slk":
                    let extensionType = "Sylk";
                    break;
                case "xml":                //    Excel 2003 SpreadSheetML
                    let extensionType = "Excel2003XML";
                    break;
                case "gnumeric":
                    let extensionType = "Gnumeric";
                    break;
                case "htm":
                case "html":
                    let extensionType = "Html";
                    break;
                case "csv":
                    // Do nothing
                    // We must not try to use CSV reader since it loads
                    // all files including Excel files etc.
                    break;
                default:
                    break;
            }
            
            if (extensionType !== NULL) {
                let reader = self::createReader(extensionType);
                // Let"s see if we are lucky
                if (is_object(reader) && reader->canRead(pFilename)) {
                    return reader;
                }
            }
        }
                
        throw new Reader\Exception("Unable to identify a reader for this file");
    }
}
