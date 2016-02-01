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
            xmlCore, xmlWorkbook,
            eleSheet, docSheet, docProps,
            fileWorksheet, xmlSheet, sharedFormulas,
            xSplit, ySplit, sqref,
            sheetViewAttr, paneAttr, selectionAttr,
            activeTab;
        
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
        
        for rel in iterator(wbRels->Relationship) {
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

        let relationships = simplexml_load_string(
            this->securityScan(this->_getFromZipArchive(zip, "_rels/.rels")),
            "SimpleXMLElement",
            \ZExcel\Settings::getLibXmlLoaderOptions()
        );
        
        for rel in iterator(relationships) {
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
                        let xmlStrings = reset(xmlStrings);
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
                    
                    for ele in iterator(relsWorkbook->Relationship) {
                        let ele = reset(ele);
                        
                        switch(ele["Type"]) {
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
                    
                    if isset(xmlWorkbook->workbookPr) {
                        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_WINDOWS_1900);
                        if !empty(xmlWorkbook->workbookPr->attributes()->{"date1904"}) && self::boolea((string) xmlWorkbook->workbookPr->attributes()->{"date1904"}) {
                            \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_MAC_1904);
                        }
                    }
                    
                    let sheetId = 0; // keep track of new sheet id in final workbook
                    let oldSheetId = -1; // keep track of old sheet id in final workbook
                    let countSkippedSheets = 0; // keep track of number of skipped sheets
                    let mapSheetId = []; // mapping of sheet ids from old to new

                    let charts = [];
                    let chartDetails = [];
                    
                    if isset(xmlWorkbook->sheets) {
                        for eleSheet in iterator(xmlWorkbook->sheets->sheet) {
                            let oldSheetId = oldSheetId + 1;
                            
                            if (this->_loadSheetsOnly == true && !in_array((string) eleSheet->attributes()->name, this->_loadSheetsOnly)) {
                                let countSkippedSheets = countSkippedSheets + 1;
                                let mapSheetId[oldSheetId] = null;
                                continue;
                            }
                            
                            let mapSheetId[oldSheetId] = oldSheetId - countSkippedSheets;
                            
                            let docSheet = new \ZExcel\Worksheet(excel);
                            
                            docSheet->setTitle((string) eleSheet->attributes()->name, false);
                            
                            let fileWorksheet = worksheets[(string) eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships")->id];
                            let xmlSheet = simplexml_load_string(
                                this->securityScan(this->_getFromZipArchive(zip, dir . "/" . fileWorksheet)),
                                "SimpleXMLElement",
                                \ZExcel\Settings::getLibXmlLoaderOptions()
                            );
    
                            let sharedFormulas = [];
    
                            if (!empty(eleSheet->attributes()->state) && ((string) eleSheet->attributes()->state) != "") {
                                docSheet->setSheetState((string) eleSheet->attributes()->state);
                            }
                            
                            if (isset(xmlSheet->sheetViews) && isset(xmlSheet->sheetViews->sheetView)) {
                                let sheetViewAttr = json_decode(json_encode(xmlSheet->sheetViews->sheetView->attributes()), true);
                               
                                if (is_array(sheetViewAttr)) {
                                    let sheetViewAttr = sheetViewAttr["@attributes"];
                            
                                    if (isset(sheetViewAttr["zoomScale"]) && is_numeric(sheetViewAttr["zoomScale"])) {
                                        docSheet->getSheetView()->setZoomScale(intval(sheetViewAttr["zoomScale"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["zoomScaleNormal"]) && is_numeric(sheetViewAttr["zoomScaleNormal"])) {
                                        docSheet->getSheetView()->setZoomScaleNormal( intval(sheetViewAttr["zoomScaleNormal"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["view"]) && sheetViewAttr["view"] != "") {
                                        docSheet->getSheetView()->setView(sheetViewAttr["view"]);
                                    }
                                    
                                    if (isset(sheetViewAttr["showGridLines"])) {
                                        docSheet->setShowGridLines(self::boolea(sheetViewAttr["showGridLines"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["showRowColHeaders"])) {
                                        docSheet->setShowRowColHeaders(self::boolea(sheetViewAttr["showRowColHeaders"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["rightToLeft"])) {
                                        docSheet->setRightToLeft(self::boolea(sheetViewAttr["rightToLeft"]));
                                    }
                                }
                                
                                if (isset(xmlSheet->sheetViews->sheetView->pane)) {
                                    let paneAttr = json_decode(json_encode(xmlSheet->sheetViews->sheetView->pane), true);
                                    
                                    if (is_array(paneAttr)) {
                                        let paneAttr = paneAttr["@attributes"];
                                    
                                        if (isset(paneAttr["topLeftCell"])) {
                                            docSheet->freezePane(paneAttr["topLeftCell"]);
                                        } else {
                                            let xSplit = 0;
                                            let ySplit = 0;
                                    
                                            if (isset(paneAttr["xSplit"]) && is_numeric(paneAttr["xSplit"])) {
                                                let xSplit = 1 + intval(paneAttr["xSplit"]);
                                            }
                                    
                                            if (isset(paneAttr["ySplit"]) && is_numeric(paneAttr["ySplit"])) {
                                                let ySplit = 1 + intval(paneAttr["ySplit"]);
                                            }
                                    
                                            docSheet->freezePaneByColumnAndRow(xSplit, ySplit);
                                        }
                                    }
                                }
                                
                                if (isset(xmlSheet->sheetViews->sheetView->selection)) {
                                    let selectionAttr = json_decode(json_encode(xmlSheet->sheetViews->sheetView->selection->attributes()), true);
                                    
                                    if (is_array(selectionAttr)) {
                                        let selectionAttr = selectionAttr["@attributes"];
                                    }
                                
                                    if (isset(selectionAttr["sqref"])) {
                                        let sqref = explode(" ", selectionAttr["sqref"]);
                                        let sqref = sqref[0];
                                        
                                        docSheet->setSelectedCells(sqref);
                                    }
                                }
                            }
                            
                            if (count(xmlSheet->sheetPr) > 0) {
                                if (count(xmlSheet->sheetPr->tabColor) > 0 && xmlSheet->sheetPr->tabColor->getAttributes()->rgb !== null) {
                                    docSheet->getTabColor()->setARGB((string) xmlSheet->sheetPr->tabColor->getAttributes()->rgb);
                                }
                                
                                if (xmlSheet->sheetPr->attributes()->codeName !== null) {
                                    docSheet->setCodeName((string) xmlSheet->sheetPr->attributes()->codeName);
                                }
                                
                                if (count(xmlSheet->sheetPr->outlinePr) > 0) {
                                    if (xmlSheet->sheetPr->outlinePr->attributes()->summaryRight !== null
                                           && !self::boolea((string) xmlSheet->sheetPr->outlinePr->attributes()->summaryRight)) {
                                        docSheet->setShowSummaryRight(false);
                                    } else {
                                        docSheet->setShowSummaryRight(true);
                                    }
        
                                    if (xmlSheet->sheetPr->outlinePr->attributes()->summaryBelow !== null
                                           && !self::boolea((string) xmlSheet->sheetPr->outlinePr->attributes()->summaryBelow)) {
                                        docSheet->setShowSummaryBelow(false);
                                    } else {
                                        docSheet->setShowSummaryBelow(true);
                                    }
                                }

                                if (count(xmlSheet->sheetPr->pageSetUpPr) > 0) {
                                
                                    if (xmlSheet->sheetPr->pageSetUpPr->attributes()->fitToPage !== null
                                           && !self::boolea((string) xmlSheet->sheetPr->pageSetUpPr->attributes()->fitToPage)) {
                                        docSheet->getPageSetup()->setFitToPage(false);
                                    } else {
                                        docSheet->getPageSetup()->setFitToPage(true);
                                    }
                                }
                            }
    
                            if (count(xmlSheet->sheetFormatPr) > 0) {
                                if (xmlSheet->sheetFormatPr->attributes()->customHeight !== null
                                        && self::boolea((string) xmlSheet->attributes()->customHeight)
                                        && xmlSheet->sheetFormatPr->attributes()->defaultRowHeight !== null) {
                                    docSheet->getDefaultRowDimension()->setRowHeight((float) xmlSheet->sheetFormatPr->attributes()->defaultRowHeight);
                                }
                                
                                if (xmlSheet->sheetFormatPr->attributes()->defaultColWidth !== null) {
                                    docSheet->getDefaultColumnDimension()->setWidth((float) xmlSheet->sheetFormatPr->attributes()->defaultColWidth);
                                }
                                if (xmlSheet->sheetFormatPr->attributes()->zeroHeight !== null) {
                                    let k = (string) xmlSheet->sheetFormatPr->attributes()->zeroHeight;
                                    
                                    if (k === "1") {
                                        docSheet->getDefaultRowDimension()->setZeroHeight(true);
                                    }
                                }
                            }

                            // @TODO add code from lines 766 - 1553
                            
                            excel->addSheet(docSheet, eleSheet->attributes()->sheetId);
                        }
                    }
                    
                    // Select active sheet
                    if ((!this->_readDataOnly) || (!empty(this->_loadSheetsOnly))) {
                        // active sheet index
                        let activeTab = intval((string) xmlWorkbook->bookViews->workbookView->attributes()->activeTab); // refers to old sheet index

                        // keep active sheet index if sheet is still loaded, else first sheet is set as the active
                        if (isset(mapSheetId[activeTab]) && mapSheetId[activeTab] !== null) {
                            excel->setActiveSheetIndex(mapSheetId[activeTab]);
                        } else {
                            if (excel->getSheetCount() == 0) {
                                excel->createSheet();
                            }
                            
                            excel->setActiveSheetIndex(0);
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
