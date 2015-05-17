namespace ZExcel;

class ZExcel
{
    private uniqueID;

    /**
     * Document properties
     *
     * @var PHPExcel_DocumentProperties
     */
    private properties;

    /**
     * Document security
     *
     * @var PHPExcel_DocumentSecurity
     */
    private security;

    /**
     * Collection of Worksheet objects
     *
     * @var PHPExcel_Worksheet[]
     */
    private workSheetCollection = [];

    /**
     * Calculation Engine
     *
     * @var PHPExcel_Calculation
     */
    private calculationEngine;

    /**
     * Active sheet index
     *
     * @var integer
     */
    private activeSheetIndex = 0;

    /**
     * Named ranges
     *
     * @var PHPExcel_NamedRange[]
     */
    private namedRanges = [];

    /**
     * CellXf supervisor
     *
     * @var PHPExcel_Style
     */
    private cellXfSupervisor;

    /**
     * CellXf collection
     *
     * @var PHPExcel_Style[]
     */
    private cellXfCollection = [];

    /**
     * CellStyleXf collection
     *
     * @var PHPExcel_Style[]
     */
    private cellStyleXfCollection = [];

    /**
    * hasMacros : this workbook have macros ?
    *
    * @var bool
    */
    private hasMacros = false;

    /**
    * macrosCode : all macros code (the vbaProject.bin file, this include form, code,  etc.), null if no macro
    *
    * @var binary
    */
    private macrosCode;
    /**
    * macrosCertificate : if macros are signed, contains vbaProjectSignature.bin file, null if not signed
    *
    * @var binary
    */
    private macrosCertificate;

    /**
    * ribbonXMLData : null if workbook is'nt Excel 2007 or not contain a customized UI
    *
    * @var null|string
    */
    private ribbonXMLData;

    /**
    * ribbonBinObjects : null if workbook is'nt Excel 2007 or not contain embedded objects (picture(s)) for Ribbon Elements
    * ignored if ribbonXMLData is null
    *
    * @var null|array
    */
    private ribbonBinObjects;

    /**
    * The workbook has macros ?
    *
    * @return true if workbook has macros, false if not
    */
    public function hasMacros() -> boolean
    {
        return this->hasMacros;
    }

    /**
    * Define if a workbook has macros
    *
    * @param boolean hasMacros true|false
    */
    public function setHasMacros(boolean hasMacros = false)
    {
        let this->hasMacros = hasMacros;
    }

    /**
    * Set the macros code
    *
    * @param string MacrosCode string|null
    */
    public function setMacrosCode(string MacrosCode = null)
    {
        let this->macrosCode=MacrosCode;
        this->setHasMacros(!is_null(MacrosCode));
    }

    /**
    * Return the macros code
    *
    * @return string|null
    */
    public function getMacrosCode() -> string
    {
        return this->macrosCode;
    }

    /**
    * Set the macros certificate
    *
    * @param string|null Certificate
    */
    public function setMacrosCertificate(string Certificate = null)
    {
        let this->macrosCertificate = Certificate;
    }

    /**
    * Is the project signed ?
    *
    * @return boolean true|false
    */
    public function hasMacrosCertificate() -> boolean
    {
        return !is_null(this->macrosCertificate);
    }

    /**
    * Return the macros certificate
    *
    * @return string|null
    */
    public function getMacrosCertificate() -> string
    {
        return this->macrosCertificate;
    }

    /**
    * Remove all macros, certificate from spreadsheet
    *
    */
    public function discardMacros()
    {
        let this->hasMacros = false;
        let this->macrosCode = null;
        let this->macrosCertificate = null;
    }

    /**
    * set ribbon XML data
    *
    */
    public function setRibbonXMLData(var Target = null, var XMLData = null)
    {
        if (!is_null(Target) && !is_null(XMLData)) {
            let this->ribbonXMLData = ["target": Target, "data": XMLData];
        } else {
            let this->ribbonXMLData = null;
        }
    }

    /**
    * retrieve ribbon XML Data
    *
    * return string|null|array
    */
    public function getRibbonXMLData(string What = "all") //we need some constants here...
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * store binaries ribbon objects (pictures)
    *
    */
    public function setRibbonBinObjects(string BinObjectsNames = null, var BinObjectsData = null)
    {
        if (!is_null(BinObjectsNames) && !is_null(BinObjectsData)) {
            let this->ribbonBinObjects = ["names": BinObjectsNames, "data": BinObjectsData];
        } else {
            let this->ribbonBinObjects = null;
        }
    }
    /**
    * return the extension of a filename. Internal use for a array_map callback (php<5.3 don't like lambda function)
    *
    */
    private function getExtensionOnly(string thePath)
    {
        return pathinfo(thePath, PATHINFO_EXTENSION);
    }

    /**
    * retrieve Binaries Ribbon Objects
    *
    */
    public function getRibbonBinObjects(string What = "all")
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
    * This workbook have a custom UI ?
    *
    * @return true|false
    */
    public function hasRibbon() -> boolean
    {
        return !is_null(this->ribbonXMLData);
    }

    /**
    * This workbook have additionnal object for the ribbon ?
    *
    * @return true|false
    */
    public function hasRibbonBinObjects() -> boolean
    {
        return !is_null(this->ribbonBinObjects);
    }

    /**
     * Check if a sheet with a specified code name already exists
     *
     * @param string pSheetCodeName  Name of the worksheet to check
     * @return boolean
     */
    public function sheetCodeNameExists(string pSheetCodeName) -> boolean
    {
        return (this->getSheetByCodeName(pSheetCodeName) !== null);
    }

    /**
     * Get sheet by code name. Warning : sheet don't have always a code name !
     *
     * @param string pName Sheet name
     * @return PHPExcel_Worksheet
     */
    public function getSheetByCodeName(string pName = "")
    {
        throw new \Exception("Not implemented yet!");
    }

     /**
     * Create a new PHPExcel with one Worksheet
     */
    public function __construct()
    {
        let this->uniqueID = uniqid();
        let this->calculationEngine = \ZExcel\Calculation::getInstance(this);

        // Initialise worksheet collection and add one worksheet
        let this->workSheetCollection = [];
        let this->workSheetCollection[] = new \ZExcel\Worksheet(this);
        let this->activeSheetIndex = 0;

        // Create document properties
        let this->properties = new \ZExcel\DocumentProperties();

        // Create document security
        let this->security = new \ZExcel\DocumentSecurity();

        // Set named ranges
        let this->namedRanges = [];

        // Create the cellXf supervisor
        let this->cellXfSupervisor = new \ZExcel\Style(true);
        this->cellXfSupervisor->bindParent(this);

        // Create the default style
        this->addCellXf(new \ZExcel\Style());
        this->addCellStyleXf(new \ZExcel\Style());
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    public function __destruct()
    {
        \ZExcel\Calculation::unsetInstance(this);
        this->disconnectWorksheets();
    }

    /**
     * Disconnect all worksheets from this PHPExcel workbook object,
     *    typically so that the PHPExcel object can be unset
     *
     */
    public function disconnectWorksheets()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Return the calculation engine for this worksheet
     *
     * @return PHPExcel_Calculation
     */
    public function getCalculationEngine() -> <\ZExcel\Calculation>
    {
        return this->calculationEngine;
    }

    /**
     * Get properties
     *
     * @return PHPExcel_DocumentProperties
     */
    public function getProperties()
    {
        return this->properties;
    }

    /**
     * Set properties
     *
     * @param PHPExcel_DocumentProperties    pValue
     */
    public function setProperties(<ZExcel\DocumentProperties> pValue)
    {
        let this->properties = pValue;
    }

    /**
     * Get security
     *
     * @return ZExcel\DocumentSecurity
     */
    public function getSecurity() -> <\ZExcel\DocumentSecurity>
    {
        return this->security;
    }

    /**
     * Set security
     *
     * @param PHPExcel_DocumentSecurity    pValue
     */
    public function setSecurity(<ZExcel\DocumentSecurity> pValue)
    {
        let this->security = pValue;
    }

    /**
     * Get active sheet
     *
     * @return PHPExcel_Worksheet
     *
     * @throws PHPExcel_Exception
     */
    public function getActiveSheet() -> <\ZExcel\WorkSheet>
    {
        return this->getSheet(this->activeSheetIndex);
    }

    /**
     * Create sheet and add it to this workbook
     *
     * @param  int|null iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return PHPExcel_Worksheet
     * @throws PHPExcel_Exception
     */
    public function createSheet(int iSheetIndex = null) -> <\ZExcel\Worksheet>
    {
        var newSheet;
        
        let newSheet = new \ZExcel\Worksheet(this);
        this->addSheet(newSheet, iSheetIndex);
        
        return newSheet;
    }

    /**
     * Check if a sheet with a specified name already exists
     *
     * @param  string pSheetName  Name of the worksheet to check
     * @return boolean
     */
    public function sheetNameExists(string pSheetName) -> boolean
    {
        return (this->getSheetByName(pSheetName) !== null);
    }

    /**
     * Add sheet
     *
     * @param  PHPExcel_Worksheet pSheet
     * @param  int|null iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return ZExcel\Worksheet
     * @throws ZExcel\Exception
     */
    public function addSheet(<ZExcel\Worksheet> pSheet, int iSheetIndex = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove sheet by index
     *
     * @param  int pIndex Active sheet index
     * @throws PHPExcel_Exception
     */
    public function removeSheetByIndex(int pIndex = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get sheet by index
     *
     * @param  int pIndex Sheet index
     * @return PHPExcel_Worksheet
     * @throws PHPExcel_Exception
     */
    public function getSheet(int pIndex = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get all sheets
     *
     * @return PHPExcel_Worksheet[]
     */
    public function getAllSheets() -> array
    {
        return this->workSheetCollection;
    }

    /**
     * Get sheet by name
     *
     * @param  string pName Sheet name
     * @return PHPExcel_Worksheet
     */
    public function getSheetByName(string pName = "")
    {
        var worksheet, workSheetCount = 0, i = 0;
        
        let workSheetCount = count(this->workSheetCollection) - 1;
        
        for i in range(0, workSheetCount) {
            let worksheet = this->workSheetCollection[i]; 
            if (worksheet->getTitle() === pName) {
                return worksheet;
            }
        }
        
        return null;
    }

    /**
     * Get index for sheet
     *
     * @param  ZExcel\Worksheet pSheet
     * @return Sheet index
     * @throws PHPExcel_Exception
     */
    public function getIndex(<\ZExcel\Worksheet> pSheet) -> string
    {
        var key, value;
        
        for key, value in this->workSheetCollection {
            if (value->getHashCode() == pSheet->getHashCode()) {
                return key;
            }
        }

        throw new \ZExcel\Exception("Sheet does not exist.");
    }

    /**
     * Set index for sheet by sheet name.
     *
     * @param  string sheetName Sheet name to modify index for
     * @param  int newIndex New index for the sheet
     * @return New sheet index
     * @throws PHPExcel_Exception
     */
    public function setIndexByName(string sheetName, int newIndex)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get sheet count
     *
     * @return int
     */
    public function getSheetCount() -> int
    {
        return count(this->workSheetCollection);
    }

    /**
     * Get active sheet index
     *
     * @return int Active sheet index
     */
    public function getActiveSheetIndex() -> int
    {
        return this->activeSheetIndex;
    }

    /**
     * Set active sheet index
     *
     * @param  int pIndex Active sheet index
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function setActiveSheetIndex(int pIndex = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Set active sheet index by name
     *
     * @param  string pValue Sheet title
     * @return PHPExcel_Worksheet
     * @throws PHPExcel_Exception
     */
    public function setActiveSheetIndexByName(string pValue = "")
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get sheet names
     *
     * @return string[]
     */
    public function getSheetNames() -> array
    {
        var worksheetCount;
        array returnValue = [];
        int i = 0;
        
        let worksheetCount = this->getSheetCount() - 1;
        
        for i in range(1, worksheetCount) {
            let returnValue[] = this->getSheet(i)->getTitle();
        }

        return returnValue;
    }

    /**
     * Add external sheet
     *
     * @param  PHPExcel_Worksheet pSheet External sheet to add
     * @param  int|null iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @throws PHPExcel_Exception
     * @return PHPExcel_Worksheet
     */
    public function addExternalSheet(<\ZExcel\Worksheet> pSheet, int iSheetIndex = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get named ranges
     *
     * @return \ZExcel\NamedRange[]
     */
    public function getNamedRanges()
    {
        return this->namedRanges;
    }

    /**
     * Add named range
     *
     * @param  PHPExcel_NamedRange namedRange
     * @return PHPExcel
     */
    public function addNamedRange(<\ZExcel\NamedRange> namedRange)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get named range
     *
     * @param  string namedRange
     * @param  PHPExcel_Worksheet|null pSheet Scope. Use null for global scope
     * @return PHPExcel_NamedRange|null
     */
    public function getNamedRange(string namedRange, <ZExcel\Worksheet> pSheet = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove named range
     *
     * @param  string  namedRange
     * @param  PHPExcel_Worksheet|null  pSheet  Scope: use null for global scope.
     * @return PHPExcel
     */
    public function removeNamedRange(string namedRange, <\ZExcel\Worksheet> pSheet = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get worksheet iterator
     *
     * @return PHPExcel_WorksheetIterator
     */
    public function getWorksheetIterator()
    {
        return new \ZExcel\WorksheetIterator(this);
    }

    /**
     * Copy workbook (!= clone!)
     *
     * @return ZExcel\ZExcel
     */
    public function copy()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        var vars, key, val;
        
        let vars = get_object_vars(this);
        for key, val in vars {
            if (is_object(val) || (is_array(val))) {
                let this->{key} = unserialize(serialize(val));
            }
        }
    }

    /**
     * Get the workbook collection of cellXfs
     *
     * @return \ZExcel\Style[]
     */
    public function getCellXfCollection() -> array
    {
        return this->cellXfCollection;
    }

    /**
     * Get cellXf by index
     *
     * @param  int pIndex
     * @return ZExcel\Style
     */
    public function getCellXfByIndex(int pIndex = 0) -> <\ZExcel\Style>
    {
        return this->cellXfCollection[pIndex];
    }

    /**
     * Get cellXf by hash code
     *
     * @param  string pValue
     * @return ZExcel\Style|false
     */
    public function getCellXfByHashCode(string pValue = "")
    {
        var cellXf;
    
        for cellXf in this->cellXfCollection {
            if (cellXf->getHashCode() == pValue) {
                return cellXf;
            }
        }
        
        return false;
    }

    /**
     * Check if style exists in style collection
     *
     * @param  ZExcel\Style pCellStyle
     * @return boolean
     */
    public function cellXfExists(<\ZExcel\Style> pCellStyle = null) -> boolean
    {
        return in_array(pCellStyle, this->cellXfCollection, true);
    }

    /**
     * Get default style
     *
     * @return PHPExcel_Style
     * @throws PHPExcel_Exception
     */
    public function getDefaultStyle() -> <\ZExcel\Style>
    {
        if (isset(this->cellXfCollection[0])) {
            return this->cellXfCollection[0];
        }
        
        throw new \ZExcel\Exception("No default style found for this workbook");
    }

    /**
     * Add a cellXf to the workbook
     *
     * @param PHPExcel_Style style
     */
    public function addCellXf(<Style> style)
    {
        let this->cellXfCollection[] = style;
        style->setIndex(count(this->cellXfCollection) - 1);
    }

    /**
     * Remove cellXf by index. It is ensured that all cells get their xf index updated.
     *
     * @param integer pIndex Index to cellXf
     * @throws PHPExcel_Exception
     */
    public function removeCellXfByIndex(pIndex = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Get the cellXf supervisor
     *
     * @return PHPExcel_Style
     */
    public function getCellXfSupervisor()
    {
        return this->cellXfSupervisor;
    }

    /**
     * Get the workbook collection of cellStyleXfs
     *
     * @return PHPExcel_Style[]
     */
    public function getCellStyleXfCollection()
    {
        return this->cellStyleXfCollection;
    }

    /**
     * Get cellStyleXf by index
     *
     * @param integer pIndex Index to cellXf
     * @return PHPExcel_Style
     */
    public function getCellStyleXfByIndex(pIndex = 0)
    {
        return this->cellStyleXfCollection[pIndex];
    }

    /**
     * Get cellStyleXf by hash code
     *
     * @param  string pValue
     * @return PHPExcel_Style|false
     */
    public function getCellStyleXfByHashCode(pValue = "")
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Add a cellStyleXf to the workbook
     *
     * @param PHPExcel_Style pStyle
     */
    public function addCellStyleXf(<ZExcel\Style> pStyle)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Remove cellStyleXf by index
     *
     * @param integer pIndex Index to cellXf
     * @throws PHPExcel_Exception
     */
    public function removeCellStyleXfByIndex(pIndex = 0)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     * Eliminate all unneeded cellXf and afterwards update the xfIndex for all cells
     * and columns in the workbook
     */
    public function garbageCollect()
    {
        throw new \Exception("Not implemented yet!");
    }

    public function getID()
    {
        return this->uniqueID;
    }
}
