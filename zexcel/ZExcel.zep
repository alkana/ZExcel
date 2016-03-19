namespace ZExcel;

class ZExcel
{
    private static version = "0.1";
    
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
    public function getRibbonXMLData(string what = "all") //we need some constants here...
    {
        var returnData;
        
        let what = strtolower(what);
        
        switch (what){
            case "all":
                let returnData = this->ribbonXMLData;
                break;
            case "target":
            case "data":
                if (is_array(this->ribbonXMLData) && array_key_exists(what, this->ribbonXMLData)) {
                    let returnData = this->ribbonXMLData[what];
                }
                break;
        }

        return returnData;
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
    public function getRibbonBinObjects(string what = "all")
    {
        var returnData, tmpTypes;
        
        let what = strtolower(what);
        
        switch(what) {
            case "all":
                return this->ribbonBinObjects;
                break;
            case "names":
            case "data":
                if (is_array(this->ribbonBinObjects) && array_key_exists(what, this->ribbonBinObjects)) {
                    let returnData = this->ribbonBinObjects[what];
                }
                break;
            case "types":
                if (is_array(this->ribbonBinObjects) && array_key_exists("data", this->ribbonBinObjects) && is_array(this->ribbonBinObjects["data"])) {
                    let tmpTypes = array_keys(this->ribbonBinObjects["data"]);
                    let returnData = array_unique(array_map([this, "getExtensionOnly"], tmpTypes));
                } else {
                    let returnData = []; // the caller want an array... not null if empty
                }
                break;
        }
        
        return returnData;
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
     * Get sheet by code name. Warning : sheet don"t have always a code name !
     *
     * @param string pName Sheet name
     * @return PHPExcel_Worksheet
     */
    public function getSheetByCodeName(string pName = "")
    {
        var worksheetCount, i;
        
        let worksheetCount = count(this->workSheetCollection);
        for i in range(0, worksheetCount - 1) {
            if (this->workSheetCollection[i]->getCodeName() == pName) {
                return this->workSheetCollection[i];
            }
        }

        return null;
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
        let this->calculationEngine = null;
        this->disconnectWorksheets();
    }

    /**
     * Disconnect all worksheets from this PHPExcel workbook object,
     *    typically so that the PHPExcel object can be unset
     *
     */
    public function disconnectWorksheets()
    {
        var k;
        
        for k, _ in this->workSheetCollection {
            this->workSheetCollection[k]->disconnectCells();
        }
        
        let this->_workSheetCollection = [];
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
    public function setProperties(<\ZExcel\DocumentProperties> pValue)
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
    public function setSecurity(<\ZExcel\DocumentSecurity> pValue)
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
    public function addSheet(<\ZExcel\Worksheet> pSheet, int iSheetIndex = -1)
    {
        if (this->sheetNameExists(pSheet->getTitle())) {
            throw new \ZExcel\Exception("Workbook already contains a worksheet named '". pSheet->getTitle() . "'. Rename this worksheet first.");
        }

        if(iSheetIndex == -1) {
            if (this->activeSheetIndex < 0) {
                let this->activeSheetIndex = 0;
            }
            
            let this->workSheetCollection[] = pSheet;
        } else {
            // Insert the sheet at the requested index
            array_splice(
                this->workSheetCollection,
                iSheetIndex,
                0,
                [pSheet]
            );

            // Adjust active sheet index if necessary
            if (this->activeSheetIndex >= iSheetIndex) {
                let this->activeSheetIndex = this->activeSheetIndex + 1;
            }
        }

        if (pSheet->getParent() == null) {
            pSheet->rebindParent(this);
        }

        return pSheet;
    }

    /**
     * Remove sheet by index
     *
     * @param  int pIndex Active sheet index
     * @throws PHPExcel_Exception
     */
    public function removeSheetByIndex(int pIndex = 0)
    {
        var numSheets;
        
        let numSheets = count(this->workSheetCollection);

        if (pIndex > numSheets - 1) {
            throw new \ZExcel\Exception("You tried to remove a sheet by the out of bounds index: " . pIndex . ". The actual number of sheets is " . numSheets . ".");
        } else {
            array_splice(this->workSheetCollection, pIndex, 1);
        }
        // Adjust active sheet index if necessary
        if ((this->activeSheetIndex >= pIndex) && (pIndex > count(this->workSheetCollection) - 1)) {
            let this->activeSheetIndex = this->activeSheetIndex - 1;
        }
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
        var numSheets;
        
        if (!isset(this->workSheetCollection[pIndex])) {
            let numSheets = this->getSheetCount();
            throw new \ZExcel\Exception("Your requested sheet index: " . pIndex . " is out of bounds. The actual number of sheets is " . numSheets . ".");
        }

        return this->workSheetCollection[pIndex];
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
        var oldIndex, pSheet;
        
        let oldIndex = this->getIndex(this->getSheetByName(sheetName));
        
        let pSheet = array_splice(
            this->workSheetCollection,
            oldIndex,
            1
        );
        
        array_splice(
            this->workSheetCollection,
            newIndex,
            0,
            pSheet
        );
        
        return newIndex;
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
        int numSheets = 0;
    
        let numSheets = count(this->workSheetCollection);

        if (pIndex > numSheets - 1) {
            throw new \ZExcel\Exception("You tried to set a sheet active by the out of bounds index: " . pIndex . " The actual number of sheets is " . numSheets . ".");
        } else {
            let this->activeSheetIndex = pIndex;
        }
        
        return this->getActiveSheet();
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
        var worksheet;
        
        let worksheet = this->getSheetByName(pValue);
        
        if (is_object(worksheet) && worksheet instanceof \ZExcel\Worksheet) {
            this->setActiveSheetIndex(this->getIndex(worksheet));
            
            return worksheet;
        }

        throw new \ZExcel\Exception("Workbook does not contain sheet:" . pValue);
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
        
        for i in range(0, worksheetCount) {
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
        var countCellXfs, cellXf, cellID, cell;
        
        if (this->sheetNameExists(pSheet->getTitle())) {
            throw new \ZExcel\Exception("Workbook already contains a worksheet named '" . pSheet->getTitle() . "'. Rename the external sheet first.");
        }

        // count how many cellXfs there are in this workbook currently, we will need this below
        let countCellXfs = count(this->cellXfCollection);

        // copy all the shared cellXfs from the external workbook and append them to the current
        for cellXf in pSheet->getParent()->getCellXfCollection() {
            this->addCellXf(clone cellXf);
        }

        // move sheet to this workbook
        pSheet->rebindParent(this);

        // update the cellXfs
        for cellID in pSheet->getCellCollection(false) {
            let cell = pSheet->getCell(cellID);
            cell->setXfIndex(cell->getXfIndex() + countCellXfs);
        }

        return this->addSheet(pSheet, iSheetIndex);
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
        if (namedRange->getScope() == null) {
            // global scope
            let this->namedRanges[namedRange->getName()] = namedRange;
        } else {
            // local scope
            let this->namedRanges[namedRange->getScope()->getTitle() . "!" . namedRange->getName()] = namedRange;
        }
        
        return true;
    }

    /**
     * Get named range
     *
     * @param  string namedRange
     * @param  PHPExcel_Worksheet|null pSheet Scope. Use null for global scope
     * @return PHPExcel_NamedRange|null
     */
    public function getNamedRange(string namedRange, <\ZExcel\Worksheet> pSheet = null)
    {
        var returnValue;

        if (namedRange != "" && namedRange !== null) {
            // first look for global defined name
            if (isset(this->namedRanges[namedRange])) {
                let returnValue = this->namedRanges[namedRange];
            }

            // then look for local defined name (has priority over global defined name if both names exist)
            if ((pSheet !== null) && isset(this->namedRanges[pSheet->getTitle() . "!" . namedRange])) {
                let returnValue = this->namedRanges[pSheet->getTitle() . "!" . namedRange];
            }
        }

        return returnValue;
    }

    /**
     * Remove named range
     *
     * @param  string  namedRange
     * @param  PHPExcel_Worksheet|null  pSheet  Scope: use null for global scope.
     * @return PHPExcel
     */
    public function removeNamedRange(string namedRange, <\ZExcel\Worksheet> pSheet = null) -> <\ZExcel\ZExcel>
    {
        if (pSheet === null) {
            if (isset(this->namedRanges[namedRange])) {
                unset(this->namedRanges[namedRange]);
            }
        } else {
            if (isset(this->namedRanges[pSheet->getTitle() . "!" . namedRange])) {
                unset(this->namedRanges[pSheet->getTitle() . "!" . namedRange]);
            }
        }
        
        return this;
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
        var copied, worksheetCount, i;
        
        let copied = clone this;
        let worksheetCount = count(this->workSheetCollection);
        
        for i in range(0, worksheetCount - 1) {
            let this->workSheetCollection[i] = this->workSheetCollection[i]->copy();
            this->workSheetCollection[i]->rebindParent(this);
        }

        return copied;
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
        var worksheet, cellID, cell, xfIndex;
        
        if (pIndex > count(this->cellXfCollection) - 1) {
            throw new \ZExcel\Exception("CellXf index is out of bounds.");
        } else {
            // first remove the cellXf
            array_splice(this->cellXfCollection, pIndex, 1);

            // then update cellXf indexes for cells
            for worksheet in this->workSheetCollection {
                for cellID in worksheet->getCellCollection(false) {
                    let cell = worksheet->getCell(cellID);
                    let xfIndex = cell->getXfIndex();
                    
                    if (xfIndex > pIndex ) {
                        // decrease xf index by 1
                        cell->setXfIndex(xfIndex - 1);
                    } else {
                        if (xfIndex == pIndex) {
                            // set to default xf index 0
                            cell->setXfIndex(0);
                        }
                    }
                }
            }
        }
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
    public function getCellStyleXfByHashCode(var pValue = "")
    {
        var cellStyleXf;
        
        for cellStyleXf in this->cellStyleXfCollection {
            if (cellStyleXf->getHashCode() == pValue) {
                return cellStyleXf;
            }
        }
        
        return false;
    }

    /**
     * Add a cellStyleXf to the workbook
     *
     * @param PHPExcel_Style pStyle
     */
    public function addCellStyleXf(<Style> pStyle)
    {
        let this->cellStyleXfCollection[] = pStyle;
        pStyle->setIndex(count(this->cellStyleXfCollection) - 1);
    }

    /**
     * Remove cellStyleXf by index
     *
     * @param integer pIndex Index to cellXf
     * @throws PHPExcel_Exception
     */
    public function removeCellStyleXfByIndex(pIndex = 0)
    {
        if (pIndex > count(this->cellStyleXfCollection) - 1) {
            throw new \ZExcel\Exception("CellStyleXf index is out of bounds.");
        } else {
            array_splice(this->cellStyleXfCollection, pIndex, 1);
        }
    }

    /**
     * Eliminate all unneeded cellXf and afterwards update the xfIndex for all cells
     * and columns in the workbook
     */
    public function garbageCollect()
    {
        var i, map, index, cellXf, sheet, cell, cellID, rowDimension, columnDimension, countNeededCellXfs;
        
        // how many references are there to each cellXf ?
        array countReferencesCellXf = [];
        
        for index, cellXf in this->cellXfCollection {
            let countReferencesCellXf[index] = 0;
        }

        for sheet in this->getWorksheetIterator() {
            // from cells
            for cellID in sheet->getCellCollection(false) {
                let cell = sheet->getCell(cellID);
                let countReferencesCellXf[cell->getXfIndex()] = countReferencesCellXf[cell->getXfIndex()] + 1;
            }

            // from row dimensions
            for rowDimension in sheet->getRowDimensions() {
                if (rowDimension->getXfIndex() !== null) {
                    let countReferencesCellXf[rowDimension->getXfIndex()] = countReferencesCellXf[rowDimension->getXfIndex()] + 1;
                }
            }

            // from column dimensions
            for columnDimension in sheet->getColumnDimensions() {
                let countReferencesCellXf[columnDimension->getXfIndex()] = countReferencesCellXf[columnDimension->getXfIndex()] + 1;
            }
        }

        // remove cellXfs without references and create mapping so we can update xfIndex
        // for all cells and columns
        let countNeededCellXfs = 0;
        
        for index, cellXf in this->cellXfCollection {
            if (countReferencesCellXf[index] > 0 || index == 0) { // we must never remove the first cellXf
                let countNeededCellXfs = countNeededCellXfs + 1;
            } else {
                unset(this->cellXfCollection[index]);
            }
            
            let map[index] = countNeededCellXfs - 1;
        }
        
        let this->cellXfCollection = array_values(this->cellXfCollection);

        // update the index for all cellXfs
        for i, cellXf in this->cellXfCollection {
            cellXf->setIndex(i);
        }

        // make sure there is always at least one cellXf (there should be)
        if (empty(this->cellXfCollection)) {
            let this->cellXfCollection[] = new \ZExcel\Style();
        }

        // update the xfIndex for all cells, row dimensions, column dimensions
        for sheet in this->getWorksheetIterator() {
            // for all cells
            for cellID in sheet->getCellCollection(false) {
                let cell = sheet->getCell(cellID);
                cell->setXfIndex(map[cell->getXfIndex()]);
            }

            // for all row dimensions
            for rowDimension in sheet->getRowDimensions() {
                if (rowDimension->getXfIndex() !== null) {
                    rowDimension->setXfIndex(map[rowDimension->getXfIndex()]);
                }
            }

            // for all column dimensions
            for columnDimension in sheet->getColumnDimensions() {
                columnDimension->setXfIndex(map[columnDimension->getXfIndex()]);
            }

            // also do garbage collection for all the sheets
            sheet->garbageCollect();
        }
    }

    public function getID()
    {
        return this->uniqueID;
    }
    
    public static function getVersion()
    {
        return self::version;
    }
}
