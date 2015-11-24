namespace ZExcel\Reader;

class Excel2007 extends Abstrac implements IReader
{
    private _readFilter = null;
    
    private _referenceHelper = null;
    
    private static _theme = null;
    
    public function __construct() {
        let this->_readFilter = new DefaultReadFilter();
        let this->_referenceHelper = \ZExcel\ReferenceHelper::getInstance();
    }
    
    public function canRead(string pFilename) -> boolean
    {
        var zipClass, zip, rel, rels, relationships = [];
        boolean xl = false;
    
        // Check if file exists
        if (!file_exists(pFilename)) {
            throw new Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        // Load file
        let zip = new {zipClass}();
        if (zip->open(pFilename) === true) {
            // check if it is an OOXML archive
            let rels = simplexml_load_string(this->securityScan(this->_getFromZipArchive(zip, "_rels/.rels")), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());
            
            if (rels !== false) {
                let relationships = reset(rels);
                
                for rel in relationships {
                    let rel = reset(rel);
                    
                    switch (rel["Type"]) {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                            if (basename(rel["Target"]) == "workbook.xml") {
                                let xl = true;
                            }
                            break;
                    }
                }
            }
            zip->close();
        }

        return xl;
    }

    public function listWorksheetNames(string pFilename)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function listWorksheetInfo(pFilename)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _castToBool(c)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _castToError(c)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _castToString(c)
    {
        throw new \Exception("Not implemented yet!");
    }

    private function _castToFormula(c, r, cellDataType, value, calculatedValue, sharedFormulas, castBaseType)
    {
        throw new \Exception("Not implemented yet!");
    }

    public function _getFromZipArchive(<\ZipArchive> archive, fileName = "")
    {
        var contents;
        
        // Root-relative paths
        if (strpos(fileName, "//") !== false)
        {
            let fileName = substr(fileName, strpos(fileName, "//") + 1);
        }
        
        let fileName = \ZExcel\Shared\File::realpath(fileName);
        
        // Apache POI fixes
        let contents = archive->getFromName(fileName);
        
        if (contents === false)
        {
            let contents = archive->getFromName(substr(fileName, 1));
        }
        
        return contents;
    }

    public function load(string pFilename)
    {
        var excel, zipClass, zip, wbRels, rel, relationships,
            themeOrderArray, themeOrderAdditional,
            xmlTheme, xmlThemeName, themeName, themePos,
            colourScheme, colourSchemeName, themeColours,
            k, xmlColour, xmlColourData,
            dir, relsWorkbook, sharedStrings, xpath, xmlStrings, val,
            worksheets, macros, customUI, ele,
            sheetId, oldSheetId, countSkippedSheets, mapSheetId,
            charts, chartDetails,
            xmlCore, xmlWorkbook, arrWorkbook,
            sheets, eleSheet, docSheet, docProps,
            fileWorksheet, xmlSheet, sharedFormulas;
        
        if (!file_exists(pFilename)) {
            throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        // Initialisations
        let excel = new \ZExcel\ZExcel();
        // remove created sheet on new ZExcel Document
        excel->removeSheetByIndex(0);
        
        if (!this->_readDataOnly) {
            excel->removeCellStyleXfByIndex(0); // remove the default style
            excel->removeCellXfByIndex(0); // remove the default style
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        let zip = new {zipClass}();
        zip->open(pFilename);

        //    Read the theme first, because we need the colour scheme when reading the styles
        let wbRels = simplexml_load_string(this->securityScan(this->_getFromZipArchive(zip, "xl/_rels/workbook.xml.rels")), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());
        
        let relationships = wbRels->Relationship;
                
        for rel in relationships->xpath(".") {
            let rel = reset(rel);
            switch (rel["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                    let themeOrderArray = ["lt1","dk1","lt2","dk2"];
                    let themeOrderAdditional = count(themeOrderArray);

                    let xmlTheme = simplexml_load_string(this->securityScan(this->_getFromZipArchive(zip, "xl/" . rel["Target"])), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());
                    if (is_object(xmlTheme)) {
                        let xmlThemeName = xmlTheme->attributes();
                        let xmlTheme = xmlTheme->children("http://schemas.openxmlformats.org/drawingml/2006/main");
                        let themeName = (string) xmlThemeName["name"];

                        let colourScheme = xmlTheme->themeElements->clrScheme->attributes();
                        let colourSchemeName = (string) colourScheme["name"];
                        let colourScheme = xmlTheme->themeElements->clrScheme->children("http://schemas.openxmlformats.org/drawingml/2006/main");

                        let themeColours = [];
                        for k, xmlColour in colourScheme {
                            let themePos = array_search(k, themeOrderArray);
                            
                            if (themePos === false) {
                                let themeOrderAdditional = themeOrderAdditional + 1;
                                let themePos = themeOrderAdditional;
                            }
                            
                            if (isset(xmlColour->sysClr)) {
                                let xmlColourData = xmlColour->sysClr->attributes();
                                let themeColours[themePos] = xmlColourData["lastClr"];
                            } elseif (isset(xmlColour->srgbClr)) {
                                let xmlColourData = xmlColour->srgbClr->attributes();
                                let themeColours[themePos] = xmlColourData["val"];
                            }
                        }

                        let self::_theme = new \ZExcel\Reader\Excel2007\Theme(themeName,colourSchemeName,themeColours);
                    }
                    break;
            }
        }

        let relationships = reset(simplexml_load_string(
            this->securityScan(this->_getFromZipArchive(zip, "_rels/.rels")),
            "SimpleXMLElement",
            \ZExcel\Settings::getLibXmlLoaderOptions()
        ));
        
        for rel in relationships {
            let rel = reset(rel);
            
            switch (rel["Type"]) {
                case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties":
                    let xmlCore = simplexml_load_string(
                        this->securityScan(this->_getFromZipArchive(zip, rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if (is_object(xmlCore)) {
                        // @TODO require checking
                        
                        xmlCore->registerXPathNamespace("dc", "http://purl.org/dc/elements/1.1/");
                        xmlCore->registerXPathNamespace("dcterms", "http://purl.org/dc/terms/");
                        xmlCore->registerXPathNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                        let docProps = excel->getProperties();
                        
                        docProps->setCreator((string) self::array_item(xmlCore->xpath("dc:creator")));
                        docProps->setLastModifiedBy((string) self::array_item(xmlCore->xpath("cp:lastModifiedBy")));
                        docProps->setCreated(strtotime(self::array_item(xmlCore->xpath("dcterms:created")))); //! respect xsi:type
                        docProps->setModified(strtotime(self::array_item(xmlCore->xpath("dcterms:modified")))); //! respect xsi:type
                        docProps->setTitle((string) self::array_item(xmlCore->xpath("dc:title")));
                        docProps->setDescription((string) self::array_item(xmlCore->xpath("dc:description")));
                        docProps->setSubject((string) self::array_item(xmlCore->xpath("dc:subject")));
                        docProps->setKeywords((string) self::array_item(xmlCore->xpath("cp:keywords")));
                        docProps->setCategory((string) self::array_item(xmlCore->xpath("cp:category")));
                        
                        excel->setProperties(docProps);
                    }
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties":
                    let xmlCore = simplexml_load_string(
                        this->securityScan(this->_getFromZipArchive(zip, rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if (is_object(xmlCore)) {
                        let xmlCore = json_decode(json_encode(xmlCore), 1);
                        
                        let docProps = excel->getProperties();
                        
                        if isset(xmlCore["Company"]) {
                            docProps->setCompany(xmlCore["Company"]);
                        }
                        
                        if isset(xmlCore["Manager"]) {
                            docProps->setManager(xmlCore["Manager"]);
                        }
                    }
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties":
                    throw new \Exception("Not implemented yet! (" . rel["Type"] . ")");
                    break;
                //Ribbon
                case "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility":
                    throw new \Exception("Not implemented yet! (" . rel["Type"] . ")");
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                    let dir = dirname(rel["Target"]);
                    
                    let relsWorkbook = simplexml_load_string(this->securityScan(
                        this->_getFromZipArchive(zip, dir . "/_rels/" . basename(rel["Target"]) . ".rels")),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                    let sharedStrings = [];
                    let xpath = reset(self::array_item(relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']")));
                    
                    let xmlStrings = simplexml_load_string(
                        this->securityScan(this->_getFromZipArchive(zip, dir . "/" . xpath["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if (xmlStrings !== false) {
                        let xmlStrings = reset($xmlStrings);
                        if (isset(xmlStrings["si"])) {
                            for val in xmlStrings {
                                let val = reset(val);
                                if (isset(val["t"])) {
                                    let sharedStrings[] = \ZExcel\Shared\Stringg::ControlCharacterOOXML2PHP( (string) val["t"] );
                                } elseif (isset(val["r"])) {
                                    let sharedStrings[] = this->_parseRichText(val);
                                }
                            }
                        }
                    }
                    
                    let macros = null;
                    let customUI = null;
                    let worksheets = [];
                    let wbRels = reset(relsWorkbook);
                    
                    for ele in wbRels {
                        let ele = reset(ele);
                        switch(rel["Type"]) {
                            case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                                let worksheets[(string) ele["Id"]] = ele["Target"];
                                break;
                            // a vbaProject ? (: some macros)
                            case "http://schemas.microsoft.com/office/2006/relationships/vbaProject":
                                let macros = ele["Target"];
                                break;
                        }
                    }
                    
                    // @TODO code from lines 492 - 608
                    
                    let xmlWorkbook = simplexml_load_string(
                        this->securityScan(this->_getFromZipArchive(zip, rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    let arrWorkbook = json_decode(json_encode(xmlWorkbook), 1);
                    
                    if isset(arrWorkbook["workbookPr"]) {
                        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_WINDOWS_1900);
                        if (isset(arrWorkbook["workbookPr"]["@attributes"]["date1904"])) {
                            if (self::boolea((string) arrWorkbook["workbookPr"]["@attributes"]["date1904"])) {
                                \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_MAC_1904);
                            }
                        }
                    }
                    
                    let sheetId = 0; // keep track of new sheet id in final workbook
                    let oldSheetId = -1; // keep track of old sheet id in final workbook
                    let countSkippedSheets = 0; // keep track of number of skipped sheets
                    let mapSheetId = []; // mapping of sheet ids from old to new

                    let charts = [];
                    let chartDetails = [];
                    
                    if isset(arrWorkbook["sheets"]) {
                        let sheets = arrWorkbook["sheets"];
                        
                        // if we have more than one sheet, it's not the same mapping
                        if isset(sheets["sheet"]) {
                            let sheets = sheets["sheet"];
                        }
                    
                        for eleSheet in sheets {
                            let oldSheetId = oldSheetId + 1;
                            
                            if (this->_loadSheetsOnly == true && !in_array((string) eleSheet["@attributes"]["name"], this->_loadSheetsOnly)) {
                                let countSkippedSheets = countSkippedSheets + 1;
                                let mapSheetId[oldSheetId] = null;
                                continue;
                            }
                            
                            let mapSheetId[oldSheetId] = oldSheetId - countSkippedSheets;
                            
                            let docSheet = new \ZExcel\Worksheet(excel);
                            
                            docSheet->setTitle((string) eleSheet["@attributes"]["name"], false);
                            /*
                            let fileWorksheet = worksheets[(string) self::array_item(eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];
                            let xmlSheet = simplexml_load_string(
                                this->securityScan(this->_getFromZipArchive(zip, dir . "/" . fileWorksheet)),
                                "SimpleXMLElement",
                                \ZExcel\Settings::getLibXmlLoaderOptions()
                            );

                            let sharedFormulas = [];

                            if (isset(eleSheet["@attributes"]["state"]) && eleSheet["@attributes"]["state"] != "") {
                                docSheet->setSheetState((string) eleSheet["@attributes"]["state"] );
                            }
                            */
                            // @TODO add code from lines 662 - 1553

                            excel->addSheet(docSheet, eleSheet["@attributes"]["sheetId"]);
                        }
                    }
                    
                    break;
            }

        }

        // @TODO simplexml_load_string Relationship (line 402)
        
        zip->close();
        
        return excel;
    }

    private static function _readColor(color, background = false)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _readStyle(docStyle, style)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _readBorder(docBorder, eleBorder)
    {
        throw new \Exception("Not implemented yet!");
    }

    private function _parseRichText(is = null) {
        throw new \Exception("Not implemented yet!");
    }

    private function _readRibbon(excel, customUITarget, zip)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function array_item(array arry, int key = 0) -> string
    {
        var data = null;
        
        if isset arry[key] {
            let data = arry[key];
        }
        
        return data;
    }

    private static function dir_add(string base, string add) -> string
    {
        return preg_replace("~[^/]+/\.\./~", "", dirname(base) . "/" . add);
    }

    private static function toCSSArray(style)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function boolea(value = null)
    {
        if (is_object(value)) {
            let value = (string) value;
        }
        if (is_numeric(value)) {
            return (bool) value;
        }
        return (value === "true" || value === "TRUE");
    }
}
