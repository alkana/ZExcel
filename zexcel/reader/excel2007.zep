namespace ZExcel\Reader;

class Excel2007 extends Abstrac implements IReader
{
    private readFilter = null;
    
    private referenceHelper = null;
    
    private static theme = null;
    
    public function __construct() {
        let this->readFilter = new \ZEXcel\Reader\DefaultReadFilter();
        let this->referenceHelper = \ZExcel\ReferenceHelper::getInstance();
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
            let rels = simplexml_load_string(this->securityScan(this->getFromZipArchive(zip, "_rels/.rels")), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());
            
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
        var zipClass, zip, rels, rel, xmlWorkbook, eleSheet;
        array worksheetNames = [];
        
        // Check if file exists
        if (!file_exists(pFilename)) {
            throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        let zip = new {zipClass}();
        zip->open(pFilename);

        //    The files we"re looking at here are small enough that simpleXML is more efficient than XMLReader
        let rels = simplexml_load_string(
            this->securityScan(this->getFromZipArchive(zip, "_rels/.rels")),
            "SimpleXMLElement",
            \ZExcel\Settings::getLibXmlLoaderOptions()
        );
        
        for rel in iterator(rels->Relationship) {
            let rel = reset(rel);
            
            switch (rel["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                    let xmlWorkbook = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, (string) rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );

                    if (xmlWorkbook->sheets) {
                        for eleSheet in iterator(xmlWorkbook->sheets->sheet) {
                            let eleSheet = reset(eleSheet);
                            
                            // Check if sheet should be skipped
                            let worksheetNames[] = (string) eleSheet["name"];
                        }
                    }
                    break;
            }
        }

        zip->close();

        return worksheetNames;
    }

    public function listWorksheetInfo(string pFilename)
    {
        var zipClass, zip, relsWorkbook, worksheets, xmlWorkbook, tmpInfo, fileWorksheet,
        rels, rel, dir, xml, res, ele, eleSheet, currCells, row;
        array worksheetInfo = [];
        
        // Check if file exists
        if (!file_exists(pFilename)) {
            throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        let zip = new {zipClass}();
        zip->open(pFilename);

        let rels = simplexml_load_string(this->securityScan(this->getFromZipArchive(zip, "_rels/.rels")), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions()); //~ http://schemas.openxmlformats.org/package/2006/relationships");
        
        for rel in iterator(rels->Relationship) {
            let rel = reset(rel);
            
            if (rel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") {
                let dir = dirname(rel["Target"]);
                
                let relsWorkbook = simplexml_load_string(
                    this->securityScan(this->getFromZipArchive(zip, dir . "/_rels/" . basename(rel["Target"]) . ".rels")),
                    "SimpleXMLElement",
                    \ZExcel\Settings::getLibXmlLoaderOptions()
                );
                
                relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                let worksheets = [];
                
                for ele in iterator(relsWorkbook->Relationship) {
                    let ele = reset(ele);
                    
                    if (ele["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
                        let worksheets[(string) ele["Id"]] = ele["Target"];
                    }
                }

                let xmlWorkbook = simplexml_load_string(this->securityScan(this->getFromZipArchive(zip, (string) rel["Target"])), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());  //~ http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                
                if (xmlWorkbook->sheets) {
                    let dir = dirname(rel["Target"]);
                    
                    for eleSheet in iterator(xmlWorkbook->sheets->sheet) {
                        let xml = reset(eleSheet);
                        
                        let tmpInfo = [
                            "worksheetName": (string) xml["name"],
                            "lastColumnLetter": "A",
                            "lastColumnIndex": 0,
                            "totalRows": 0,
                            "totalColumns": 0
                        ];
                        
                        let fileWorksheet = worksheets[(string) self::getArrayItem(eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "id")];

                        let xml = new \XMLReader();
                        let res = xml->xml(
                            this->securityScanFile("zip://" . \ZExcel\Shared\File::realpath(pFilename) . "#" . dir . "/" . fileWorksheet),
                            null,
                            \ZExcel\Settings::getLibXmlLoaderOptions()
                        );
                        
                        xml->setParserProperty(2, true);

                        let currCells = 0;
                        
                        while (xml->read()) {
                            if (xml->name == "row" && xml->nodeType == \XMLReader::ELEMENT) {
                                let row = xml->getAttribute("r");
                                let tmpInfo["totalRows"] = row;
                                let tmpInfo["totalColumns"] = max(tmpInfo["totalColumns"], currCells);
                                let currCells = 0;
                            } elseif (xml->name == "c" && xml->nodeType == \XMLReader::ELEMENT) {
                                let currCells = currCells + 1;
                            }
                        }
                        
                        let tmpInfo["totalColumns"] = max(tmpInfo["totalColumns"], currCells);
                        xml->close();

                        let tmpInfo["lastColumnIndex"] = tmpInfo["totalColumns"] - 1;
                        let tmpInfo["lastColumnLetter"] = \ZExcel\Cell::stringFromColumnIndex(tmpInfo["lastColumnIndex"]);

                        let worksheetInfo[] = tmpInfo;
                    }
                }
            }
        }

        zip->close();

        return worksheetInfo;
    }

    public static function castToBoolean(var c)
    {
        var value;
        
        let value = isset(c->v) ? (string) c->v : null;
        
        if (value == "0") {
            return false;
        } elseif (value == "1") {
            return true;
        } else {
            return (boolean) c->v;
        }
        
        return value;
    }

    public static function castToError(var c)
    {
        return isset(c->v) ? (string) c->v : null;
    }

    public static function castToString(var c)
    {
        return isset(c->v) ? (string) c->v : null;
    }

    public function castToFormula(var c, var r, var cellDataType, var value, var calculatedValue, var sharedFormulas, var castBaseType)
    {
        var instance, sharedFormulas, master, current, difference;
        
        let cellDataType    = "f";
        let value           = "=" . c->f;
        let calculatedValue = call_user_func(["\\ZExcel\\Reader\\Excel2007", castBaseType], c);

        // Shared formula?
        if (isset(c->f["t"]) && strtolower((string)c->f["t"]) == "shared") {
            let instance = (string) c->f["si"];

            if (!isset(sharedFormulas[(string)c->f["si"]])) {
                let sharedFormulas[instance] = [
                    "master": r,
                    "formula": value
                ];
            } else {
                let master = \ZExcel\Cell::coordinateFromString(sharedFormulas[instance]["master"]);
                let current = \ZExcel\Cell::coordinateFromString(r);

                let difference = [0, 0];
                let difference[0] = \ZExcel\Cell::columnIndexFromString(current[0]) - \ZExcel\Cell::columnIndexFromString(master[0]);
                let difference[1] = current[1] - master[1];

                let value = this->referenceHelper->updateFormulaReferences(sharedFormulas[instance]["formula"], "A1", difference[0], difference[1]);
            }
        }
        
        // @TODO return references
        return [
            cellDataType,
            value,
            calculatedValue
        ];
    }

    public function getFromZipArchive(<\ZipArchive> archive, fileName = "")
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
            k, xmlColour, xmlColourData, cell, styles, cellStyles,
            dir, relsWorkbook, sharedStrings, xpath, xmlStrings, val,
            worksheets, macros, customUI, ele,
            sheetId, oldSheetId, countSkippedSheets, mapSheetId,
            charts, chartDetails, roww, row, tmp,
            xmlCore, xmlWorkbook, xmlProperty, propertyName, attributeType, attributeValue,
            cellDataOfficeAttributes, cellDataOfficeChildren, cellDataType, value, calculatedValue,
            eleSheet, docSheet, docProps, coordinates,
            fileWorksheet, xmlSheet, sharedFormulas,
            xSplit, ySplit, sqref, col, i, c, r, att,
            sheetViewAttr, paneAttr, selectionAttr,
            activeTab;
        
        if (!file_exists(pFilename)) {
            throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        // Initialisations
        let excel = new \ZExcel\ZExcel();
        // remove created sheet on new ZExcel Document
        excel->removeSheetByIndex(0);
        
        if (!this->readDataOnly) {
            excel->removeCellStyleXfByIndex(0); // remove the default style
            excel->removeCellXfByIndex(0); // remove the default style
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        let zip = new {zipClass}();
        zip->open(pFilename);

        //    Read the theme first, because we need the colour scheme when reading the styles
        let wbRels = simplexml_load_string(
            this->securityScan(this->getFromZipArchive(zip, "xl/_rels/workbook.xml.rels")),
            "SimpleXMLElement",
            \ZExcel\Settings::getLibXmlLoaderOptions()
        );
        
        for rel in iterator(wbRels->Relationship) {
            let rel = reset(rel);
            
            switch (rel["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                    let themeOrderArray = ["lt1","dk1","lt2","dk2"];
                    let themeOrderAdditional = count(themeOrderArray);

                    let xmlTheme = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, "xl/" . rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
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

                        let self::theme = new \ZExcel\Reader\Excel2007\Theme(themeName,colourSchemeName,themeColours);
                    }
                    break;
            }
        }

        let relationships = simplexml_load_string(
            this->securityScan(this->getFromZipArchive(zip, "_rels/.rels")),
            "SimpleXMLElement",
            \ZExcel\Settings::getLibXmlLoaderOptions()
        );
        
        for rel in iterator(relationships) {
            let rel = reset(rel);
            
            switch (rel["Type"]) {
                case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties":
                    let xmlCore = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if (is_object(xmlCore)) {
                        // @TODO require checking
                        
                        xmlCore->registerXPathNamespace("dc", "http://purl.org/dc/elements/1.1/");
                        xmlCore->registerXPathNamespace("dcterms", "http://purl.org/dc/terms/");
                        xmlCore->registerXPathNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                        let docProps = excel->getProperties();
                        
                        docProps->setCreator((string) self::getArrayItem(xmlCore->xpath("dc:creator")));
                        docProps->setLastModifiedBy((string) self::getArrayItem(xmlCore->xpath("cp:lastModifiedBy")));
                        docProps->setCreated(strtotime(self::getArrayItem(xmlCore->xpath("dcterms:created")))); //! respect xsi:type
                        docProps->setModified(strtotime(self::getArrayItem(xmlCore->xpath("dcterms:modified")))); //! respect xsi:type
                        docProps->setTitle((string) self::getArrayItem(xmlCore->xpath("dc:title")));
                        docProps->setDescription((string) self::getArrayItem(xmlCore->xpath("dc:description")));
                        docProps->setSubject((string) self::getArrayItem(xmlCore->xpath("dc:subject")));
                        docProps->setKeywords((string) self::getArrayItem(xmlCore->xpath("cp:keywords")));
                        docProps->setCategory((string) self::getArrayItem(xmlCore->xpath("cp:category")));
                        
                        excel->setProperties(docProps);
                    }
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties":
                    let xmlCore = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, rel["Target"])),
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
                    let xmlCore = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, (string) rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if (is_object(xmlCore)) {
                        let docProps = excel->getProperties();
                        
                        for xmlProperty in iterator(xmlCore) {
                            let cellDataOfficeAttributes = xmlProperty->attributes();
                            
                            if (isset(cellDataOfficeAttributes["name"])) {
                                let propertyName = (string) cellDataOfficeAttributes["name"];
                                let cellDataOfficeChildren = xmlProperty->children("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                                let attributeType = cellDataOfficeChildren->getName();
                                let attributeValue = (string) cellDataOfficeChildren->{attributeType};
                                let attributeValue = \ZExcel\DocumentProperties::convertProperty(attributeValue, attributeType);
                                let attributeType = \ZExcel\DocumentProperties::convertPropertyType(attributeType);
                                
                                docProps->setCustomProperty(propertyName, attributeValue, attributeType);
                            }
                        }
                    }
                    break;
                //Ribbon
                case "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility":
                    let customUI = rel["Target"];
                    if (!is_null(customUI)) {
                        this->readRibbon(excel, customUI, zip);
                    }
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                    let dir = dirname(rel["Target"]);
                    
                    let relsWorkbook = simplexml_load_string(this->securityScan(
                        this->getFromZipArchive(zip, dir . "/_rels/" . basename(rel["Target"]) . ".rels")),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    relsWorkbook->registerXPathNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                    let xpath = reset(self::getArrayItem(relsWorkbook->xpath("rel:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings']")));
                    
                    let xmlStrings = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, dir . "/" . xpath["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    let sharedStrings = [];
                    
                    if (is_object(xmlStrings) && isset(xmlStrings->si)) {
                        for val in iterator(xmlStrings) {
                            if (isset(val->t)) {
                                let sharedStrings[] = \ZExcel\Shared\Stringg::ControlCharacterOOXML2PHP( (string) val->t );
                            } elseif (isset(val->r)) {
                                let sharedStrings[] = this->parseRichText(val);
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
                    
                    // @TODO code from lines 482 - 603
                    
                    let styles = [];
                    let cellStyles = [];
                    
                    let xmlWorkbook = simplexml_load_string(
                        this->securityScan(this->getFromZipArchive(zip, rel["Target"])),
                        "SimpleXMLElement",
                        \ZExcel\Settings::getLibXmlLoaderOptions()
                    );
                    
                    if isset(xmlWorkbook->workbookPr) {
                        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_WINDOWS_1900);
                        if !empty(xmlWorkbook->workbookPr->attributes()->{"date1904"}) && self::booleann((string) xmlWorkbook->workbookPr->attributes()->{"date1904"}) {
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
                            
                            if (this->loadSheetsOnly == true && !in_array((string) eleSheet->attributes()->name, this->loadSheetsOnly)) {
                                let countSkippedSheets = countSkippedSheets + 1;
                                let mapSheetId[oldSheetId] = null;
                                continue;
                            }
                            
                            let mapSheetId[oldSheetId] = oldSheetId - countSkippedSheets;
                            
                            let docSheet = new \ZExcel\Worksheet(excel);
                            
                            docSheet->setTitle((string) eleSheet->attributes()->name, false);
                            
                            let fileWorksheet = worksheets[(string) eleSheet->attributes("http://schemas.openxmlformats.org/officeDocument/2006/relationships")->id];
                            let xmlSheet = simplexml_load_string(
                                this->securityScan(this->getFromZipArchive(zip, dir . "/" . fileWorksheet)),
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
                                        docSheet->setShowGridLines(self::booleann(sheetViewAttr["showGridLines"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["showRowColHeaders"])) {
                                        docSheet->setShowRowColHeaders(self::booleann(sheetViewAttr["showRowColHeaders"]));
                                    }
                                    
                                    if (isset(sheetViewAttr["rightToLeft"])) {
                                        docSheet->setRightToLeft(self::booleann(sheetViewAttr["rightToLeft"]));
                                    }
                                }
                                
                                if (isset(xmlSheet->sheetViews->sheetView->pane)) {
                                    let paneAttr = reset(xmlSheet->sheetViews->sheetView->pane);
                                    
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
                                           && !self::booleann((string) xmlSheet->sheetPr->outlinePr->attributes()->summaryRight)) {
                                        docSheet->setShowSummaryRight(false);
                                    } else {
                                        docSheet->setShowSummaryRight(true);
                                    }
        
                                    if (xmlSheet->sheetPr->outlinePr->attributes()->summaryBelow !== null
                                           && !self::booleann((string) xmlSheet->sheetPr->outlinePr->attributes()->summaryBelow)) {
                                        docSheet->setShowSummaryBelow(false);
                                    } else {
                                        docSheet->setShowSummaryBelow(true);
                                    }
                                }

                                if (count(xmlSheet->sheetPr->pageSetUpPr) > 0) {
                                
                                    if (xmlSheet->sheetPr->pageSetUpPr->attributes()->fitToPage !== null
                                           && !self::booleann((string) xmlSheet->sheetPr->pageSetUpPr->attributes()->fitToPage)) {
                                        docSheet->getPageSetup()->setFitToPage(false);
                                    } else {
                                        docSheet->getPageSetup()->setFitToPage(true);
                                    }
                                }
                            }
    
                            if (count(xmlSheet->sheetFormatPr) > 0) {
                                if (xmlSheet->sheetFormatPr->attributes()->customHeight !== null
                                        && self::booleann((string) xmlSheet->attributes()->customHeight)
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

                            if (isset(xmlSheet->cols) && !this->readDataOnly) {
                                for col in iterator(xmlSheet->cols->col) {
                                    let col = reset(col);
                                    
                                    for i in range(intval(col["min"]) - 1, intval(col["max"]) - 1) {
                                        if (col["style"] && !this->readDataOnly) {
                                            docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setXfIndex(intval(col["style"]));
                                        }
                                        
                                        // if (self::booleann(col["bestFit"])) {
                                        //     docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setAutoSize(true);
                                        // }
                                        
                                        if (self::booleann(col["hidden"])) {
                                            docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setVisible(false);
                                        }
                                        
                                        if (self::booleann(col["collapsed"])) {
                                            docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setCollapsed(true);
                                        }
                                        
                                        if (isset(col["outlineLevel"]) && col["outlineLevel"] > 0) {
                                            docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setOutlineLevel(intval(col["outlineLevel"]));
                                        }
                                        
                                        docSheet->getColumnDimension(\ZExcel\Cell::stringFromColumnIndex(i))->setWidth(floatval(col["width"]));

                                        if (intval(col["max"]) == 16384) {
                                            break;
                                        }
                                    }
                                }
                            }

                            if (isset(xmlSheet->printOptions) && !this->readDataOnly) {
                                let r = reset(xmlSheet->printOptions);

                                if (self::booleann((string) r["gridLinesSet"])) {
                                    docSheet->setShowGridlines(true);
                                }
                                if (self::booleann((string) r["gridLines"])) {
                                    docSheet->setPrintGridlines(true);
                                }
                                if (self::booleann((string) r["horizontalCentered"])) {
                                    docSheet->getPageSetup()->setHorizontalCentered(true);
                                }
                                if (self::booleann((string) r["verticalCentered"])) {
                                    docSheet->getPageSetup()->setVerticalCentered(true);
                                }
                            }
                            
                            if (xmlSheet && xmlSheet->sheetData && xmlSheet->sheetData->row) {
                                for roww in iterator(xmlSheet->sheetData->row) {
                                    let row = reset(roww);
                                    
                                    if (isset(row["ht"]) && !this->readDataOnly) {
                                        docSheet->getRowDimension(intval(row["r"]))->setRowHeight(floatval(row["ht"]));
                                    }
                                    
                                    if (self::booleann(row["hidden"]) && !this->readDataOnly) {
                                        docSheet->getRowDimension(intval(row["r"]))->setVisible(false);
                                    }
                                    
                                    if (self::booleann(row["collapsed"])) {
                                        docSheet->getRowDimension(intval(row["r"]))->setCollapsed(true);
                                    }
                                    
                                    if (isset(row["outlineLevel"]) && row["outlineLevel"] > 0) {
                                        docSheet->getRowDimension(intval(row["r"]))->setOutlineLevel(intval(row["outlineLevel"]));
                                    }
                                    
                                    if (isset(row["s"]) && !this->readDataOnly) {
                                        docSheet->getRowDimension(intval(row["r"]))->setXfIndex(intval(row["s"]));
                                    }
                                    
                                    for c in iterator(roww->c) {
                                        let tmp = reset(c);
                                        
                                        let r               = (string) tmp["r"];
                                        let cellDataType    = (string) tmp["t"];
                                        let value           = null;
                                        let calculatedValue = null;

                                        // Read cell?
                                        if (this->getReadFilter() !== null) {
                                            let coordinates = \ZExcel\Cell::coordinateFromString(r);

                                            if (!this->getReadFilter()->readCell(coordinates[0], coordinates[1], docSheet->getTitle())) {
                                                continue;
                                            }
                                        }

                                        // Read cell!
                                        switch (cellDataType) {
                                            case "s":
                                                if ((string) c->v != "") {
                                                    let value = sharedStrings[(string) c->v];
                                                    
                                                    if (is_object(value) && value instanceof \ZExcel\RichText) {
                                                        let value = clone value;
                                                    }
                                                } else {
                                                    let value = "";
                                                }
                                                break;
                                            case "b":
                                                if (!isset(c->f)) {
                                                    let value = self::castToBoolean(c);
                                                } else {
                                                    // Formula
                                                    this->castToFormula(c, r, cellDataType, value, calculatedValue, sharedFormulas, "castToBoolean");
                                                    
                                                    if (isset(c->f["t"])) {
                                                        let att = c->f;
                                                        docSheet->getCell(r)->setFormulaAttributes(att);
                                                    }
                                                }
                                                break;
                                            case "inlineStr":
                                                if (isset(c->f)) {
                                                    this->castToFormula(c, r, cellDataType, value, calculatedValue, sharedFormulas, "castToError");
                                                } else {
                                                    let value = this->parseRichText(c->is);
                                                }
                                                break;
                                            case "e":
                                                if (!isset(c->f)) {
                                                    let value = self::castToError(c);
                                                } else {
                                                    // Formula
                                                    this->castToFormula(c, r, cellDataType, value, calculatedValue, sharedFormulas, "castToError");
                                                }
                                                break;
                                            default:
                                                if (!isset(c->f)) {
                                                    let value = self::castToString(c);
                                                } else {
                                                    // Formula
                                                    this->castToFormula(c, r, cellDataType, value, calculatedValue, sharedFormulas, "castToString");
                                                }
                                                break;
                                        }

                                        // Check for numeric values
                                        if (is_numeric(value) && cellDataType != "s") {
                                            if (value == (int) value) {
                                                let value = (int) value;
                                            } elseif (value == (float) value) {
                                                let value = (float) value;
                                            } elseif (value == (double)value) {
                                                let value = (double) value;
                                            }
                                        }
                                        
                                        // Rich text?
                                        if (is_object(value) && value instanceof \ZExcel\RichText && this->readDataOnly) {
                                            let value = value->getPlainText();
                                        }

                                        let cell = docSheet->getCell(r);
                                        
                                        // Assign value
                                        if (cellDataType != "") {
                                            cell->setValueExplicit(value, cellDataType);
                                        } else {
                                            cell->setValue(value);
                                        }
                                        
                                        if (calculatedValue !== null) {
                                            cell->setCalculatedValue(calculatedValue);
                                        }
                                        
                                        // Style information?
                                        if (tmp["s"] && !this->readDataOnly) {
                                            let tmp = intval(tmp["s"]);
                                            
                                            // no style index means 0, it seems
                                            if (isset(styles[tmp])) {
                                                cell->setXfIndex(tmp);
                                            } else {
                                                cell->setXfIndex(0);
                                            }
                                        }
                                    }
                                }
                            }
                            
                            // @TODO add code from lines 792 - 1553
                            
                            excel->addSheet(docSheet);
                        }
                    }
                    
                    // Select active sheet
                    if ((!this->readDataOnly) || (!empty(this->loadSheetsOnly))) {
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

    private static function readColor(var color, boolean background = false)
    {
        var returnColour, tintAdjust;
        
        if (isset(color["rgb"])) {
            return (string) color["rgb"];
        } elseif (isset(color["indexed"])) {
            return \ZExcel\Style\Color::indexedColor(color["indexed"] - 7, background)->getARGB();
        } elseif (isset(color["theme"])) {
            if (self::theme !== null) {
                let returnColour = self::theme->getColourByIndex((int) color["theme"]);
                
                if (isset(color["tint"])) {
                    let tintAdjust = (float) color["tint"];
                    let returnColour = \ZExcel\Style\Color::changeBrightness(returnColour, tintAdjust);
                }
                
                return "FF" . returnColour;
            }
        }

        if (background) {
            return "FFFFFFFF";
        }
        
        return "FF000000";
    }

    private static function readStyle(var docStyle, var style)
    {
        var vertAlign, gradientFill, patternType, diagonalUp, diagonalDown, textRotation;
        
        docStyle->getNumberFormat()->setFormatCode(style->numFmt);

        // font
        if (isset(style->font)) {
            docStyle->getFont()->setName((string) style->font->name["val"]);
            docStyle->getFont()->setSize((string) style->font->sz["val"]);
            
            if (isset(style->font->b)) {
                docStyle->getFont()->setBold(!isset(style->font->b["val"]) || self::booleann((string) style->font->b["val"]));
            }
            
            if (isset(style->font->i)) {
                docStyle->getFont()->setItalic(!isset(style->font->i["val"]) || self::booleann((string) style->font->i["val"]));
            }
            
            if (isset(style->font->strike)) {
                docStyle->getFont()->setStrikethrough(!isset(style->font->strike["val"]) || self::booleann((string) style->font->strike["val"]));
            }
            
            docStyle->getFont()->getColor()->setARGB(self::readColor(style->font->color));

            if (isset(style->font->u) && !isset(style->font->u["val"])) {
                docStyle->getFont()->setUnderline(\ZExcel\Style\Font::UNDERLINE_SINGLE);
            } elseif (isset(style->font->u) && isset(style->font->u["val"])) {
                docStyle->getFont()->setUnderline((string)style->font->u["val"]);
            }

            if (isset(style->font->vertAlign) && isset(style->font->vertAlign["val"])) {
                let vertAlign = strtolower((string)style->font->vertAlign["val"]);
                
                if (vertAlign == "superscript") {
                    docStyle->getFont()->setSuperScript(true);
                }
                
                if (vertAlign == "subscript") {
                    docStyle->getFont()->setSubScript(true);
                }
            }
        }

        // fill
        if (isset(style->fill)) {
            if (style->fill->gradientFill) {
                let gradientFill = style->fill->gradientFill[0];
                
                if (!empty(gradientFill["type"])) {
                    docStyle->getFill()->setFillType((string) gradientFill["type"]);
                }
                
                docStyle->getFill()->setRotation(floatval(gradientFill["degree"]));
                gradientFill->registerXPathNamespace("sml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                docStyle->getFill()->getStartColor()->setARGB(self::readColor(self::getArrayItem(gradientFill->xpath("sml:stop[@position=0]"))->color));
                docStyle->getFill()->getEndColor()->setARGB(self::readColor(self::getArrayItem(gradientFill->xpath("sml:stop[@position=1]"))->color));
            } elseif (style->fill->patternFill) {
                let patternType = (string) style->fill->patternFill["patternType"] != "" ? (string)style->fill->patternFill["patternType"] : "solid";
                
                docStyle->getFill()->setFillType(patternType);
                
                if (style->fill->patternFill->fgColor) {
                    docStyle->getFill()->getStartColor()->setARGB(self::readColor(style->fill->patternFill->fgColor, true));
                } else {
                    docStyle->getFill()->getStartColor()->setARGB("FF000000");
                }
                
                if (style->fill->patternFill->bgColor) {
                    docStyle->getFill()->getEndColor()->setARGB(self::readColor(style->fill->patternFill->bgColor, true));
                }
            }
        }

        // border
        if (isset(style->border)) {
            let diagonalUp = self::booleann((string) style->border["diagonalUp"]);
            let diagonalDown = self::booleann((string) style->border["diagonalDown"]);
            
            if (!diagonalUp && !diagonalDown) {
                docStyle->getBorders()->setDiagonalDirection(\ZExcel\Style\Borders::DIAGONAL_NONE);
            } elseif (diagonalUp && !diagonalDown) {
                docStyle->getBorders()->setDiagonalDirection(\ZExcel\Style\Borders::DIAGONAL_UP);
            } elseif (!diagonalUp && diagonalDown) {
                docStyle->getBorders()->setDiagonalDirection(\ZExcel\Style\Borders::DIAGONAL_DOWN);
            } else {
                docStyle->getBorders()->setDiagonalDirection(\ZExcel\Style\Borders::DIAGONAL_BOTH);
            }
            
            self::readBorder(docStyle->getBorders()->getLeft(), style->border->left);
            self::readBorder(docStyle->getBorders()->getRight(), style->border->right);
            self::readBorder(docStyle->getBorders()->getTop(), style->border->top);
            self::readBorder(docStyle->getBorders()->getBottom(), style->border->bottom);
            self::readBorder(docStyle->getBorders()->getDiagonal(), style->border->diagonal);
        }

        // alignment
        if (isset(style->alignment)) {
            docStyle->getAlignment()->setHorizontal((string) style->alignment["horizontal"]);
            docStyle->getAlignment()->setVertical((string) style->alignment["vertical"]);

            let textRotation = 0;
            
            if ((int)style->alignment["textRotation"] <= 90) {
                let textRotation = (int) style->alignment["textRotation"];
            } elseif ((int)style->alignment["textRotation"] > 90) {
                let textRotation = 90 - (int)style->alignment["textRotation"];
            }

            docStyle->getAlignment()->setTextRotation(intval(textRotation));
            docStyle->getAlignment()->setWrapText(self::booleann((string) style->alignment["wrapText"]));
            docStyle->getAlignment()->setShrinkToFit(self::booleann((string) style->alignment["shrinkToFit"]));
            docStyle->getAlignment()->setIndent(intval((string)style->alignment["indent"]) > 0 ? intval((string)style->alignment["indent"]) : 0);
            docStyle->getAlignment()->setReadorder(intval((string)style->alignment["readingOrder"]) > 0 ? intval((string)style->alignment["readingOrder"]) : 0);
        }

        // protection
        if (isset(style->protection)) {
            if (isset(style->protection["locked"])) {
                if (self::booleann((string) style->protection["locked"])) {
                    docStyle->getProtection()->setLocked(\ZExcel\Style\Protection::PROTECTION_PROTECTED);
                } else {
                    docStyle->getProtection()->setLocked(\ZExcel\Style\Protection::PROTECTION_UNPROTECTED);
                }
            }

            if (isset(style->protection["hidden"])) {
                if (self::booleann((string) style->protection["hidden"])) {
                    docStyle->getProtection()->setHidden(\ZExcel\Style\Protection::PROTECTION_PROTECTED);
                } else {
                    docStyle->getProtection()->setHidden(\ZExcel\Style\Protection::PROTECTION_UNPROTECTED);
                }
            }
        }

        // top-level style settings
        if (isset(style->quotePrefix)) {
            docStyle->setQuotePrefix(style->quotePrefix);
        }
    }

    private static function readBorder(var docBorder, var eleBorder)
    {
        if (isset(eleBorder["style"])) {
            docBorder->setBorderStyle((string) eleBorder["style"]);
        }
        if (isset(eleBorder->color)) {
            docBorder->getColor()->setARGB(self::readColor(eleBorder->color));
        }
    }

    private function parseRichText(var is = null)
    {
        var value, run, objText, vertAlign;
        
        let value = new \ZExcel\RichText();

        if (isset(is->t)) {
            value->createText(\ZExcel\Shared\Stringg::ControlCharacterOOXML2PHP((string) is->t));
        } else {
            if (is_object(is->r)) {
                for run in is->r {
                    if (!isset(run->rPr)) {
                        let objText = value->createText(\ZExcel\Shared\Stringg::ControlCharacterOOXML2PHP((string) run->t));

                    } else {
                        let objText = value->createTextRun(\ZExcel\Shared\Stringg::ControlCharacterOOXML2PHP((string) run->t));

                        if (isset(run->rPr->rFont["val"])) {
                            objText->getFont()->setName((string) run->rPr->rFont["val"]);
                        }
                        
                        if (isset(run->rPr->sz["val"])) {
                            objText->getFont()->setSize((string) run->rPr->sz["val"]);
                        }
                        
                        if (isset(run->rPr->color)) {
                            objText->getFont()->setColor(new \ZExcel\Style\Color(self::readColor(run->rPr->color)));
                        }
                        
                        if ((isset(run->rPr->b["val"]) && self::booleann((string) run->rPr->b["val"])) ||
                            (isset(run->rPr->b) && !isset(run->rPr->b["val"]))) {
                            objText->getFont()->setBold(true);
                        }
                        
                        if ((isset(run->rPr->i["val"]) && self::booleann((string) run->rPr->i["val"])) ||
                            (isset(run->rPr->i) && !isset(run->rPr->i["val"]))) {
                            objText->getFont()->setItalic(true);
                        }
                        
                        if (isset(run->rPr->vertAlign) && isset(run->rPr->vertAlign["val"])) {
                            let vertAlign = strtolower((string) run->rPr->vertAlign["val"]);
                            
                            if (vertAlign == "superscript") {
                                objText->getFont()->setSuperScript(true);
                            }
                            
                            if (vertAlign == "subscript") {
                                objText->getFont()->setSubScript(true);
                            }
                        }
                        
                        if (isset(run->rPr->u) && !isset(run->rPr->u["val"])) {
                            objText->getFont()->setUnderline(\ZExcel\Style\Font::UNDERLINE_SINGLE);
                        } elseif (isset(run->rPr->u) && isset(run->rPr->u["val"])) {
                            objText->getFont()->setUnderline((string)run->rPr->u["val"]);
                        }
                        
                        if ((isset(run->rPr->strike["val"]) && self::booleann((string) run->rPr->strike["val"]))
                                || (isset(run->rPr->strike) && !isset(run->rPr->strike["val"]))) {
                            objText->getFont()->setStrikethrough(true);
                        }
                    }
                }
            }
        }

        return value;
    }

    private function readRibbon(var excel, var customUITarget, var zip)
    {
        var baseDir, nameCustomUI, localRibbon, pathRels, dataRels, ele, UIRels;
        array customUIImagesNames = [], customUIImagesBinaries = [];
        
        let baseDir = dirname(customUITarget);
        let nameCustomUI = basename(customUITarget);
        
        // get the xml file (ribbon)
        let localRibbon = this->getFromZipArchive(zip, customUITarget);
        
        // something like customUI/_rels/customUI.xml.rels
        let pathRels = baseDir . "/_rels/" . nameCustomUI . ".rels";
        let dataRels = this->getFromZipArchive(zip, pathRels);
        
        if (dataRels) {
            // exists and not empty if the ribbon have some pictures (other than internal MSO)
            let UIRels = simplexml_load_string(this->securityScan(dataRels), "SimpleXMLElement", \ZExcel\Settings::getLibXmlLoaderOptions());
            
            if (UIRels) {
                // we need to save id and target to avoid parsing customUI.xml and "guess" if it"s a pseudo callback who load the image
                for ele in iterator(UIRels->Relationship) {
                    if (ele["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                        // an image ?
                        let customUIImagesNames[(string) ele["Id"]] = (string)ele["Target"];
                        let customUIImagesBinaries[(string)ele["Target"]] = this->getFromZipArchive(zip, baseDir . "/" . (string) ele["Target"]);
                    }
                }
            }
        }
        
        if (localRibbon) {
            excel->setRibbonXMLData(customUITarget, localRibbon);
            
            if (count(customUIImagesNames) > 0 && count(customUIImagesBinaries) > 0) {
                excel->setRibbonBinObjects(customUIImagesNames, customUIImagesBinaries);
            } else {
                excel->setRibbonBinObjects(null);
            }
        } else {
            excel->setRibbonXMLData(null);
            excel->setRibbonBinObjects(null);
        }
    }

    private static function getArrayItem(var arry, var key = 0) -> string
    {
        var data = null;
        
        // SimpleXMLElement is not an array
        if (is_object(arry) && arry instanceof \SimpleXMLElement) {
            let arry = reset(arry);
        }
        
        if isset arry[key] {
            let data = arry[key];
        }
        
        return data;
    }

    private static function dir_add(string base, string add) -> string
    {
        return preg_replace("~[^/]+/\.\./~", "", dirname(base) . "/" . add);
    }

    private static function toCSSArray(var style) -> array
    {
        var temp, item;
        
        let style = str_replace(["\r","\n"], "", style);

        let temp = explode(";", style);
        let style = [];
        for item in temp {
            let item = explode(":", item);

            if (strpos(item[1], "px") !== false) {
                let item[1] = str_replace("px", "", item[1]);
            }
            if (strpos(item[1], "pt") !== false) {
                let item[1] = str_replace("pt", "", item[1]);
                let item[1] = \ZExcel\Shared\Font::fontSizeToPixels(item[1]);
            }
            if (strpos(item[1], "in") !== false) {
                let item[1] = str_replace("in", "", item[1]);
                let item[1] = \ZExcel\Shared\Font::inchSizeToPixels(item[1]);
            }
            if (strpos(item[1], "cm") !== false) {
                let item[1] = str_replace("cm", "", item[1]);
                let item[1] = \ZExcel\Shared\Font::centimeterSizeToPixels(item[1]);
            }

            let style[item[0]] = item[1];
        }

        return style;
    }

    private static function booleann(value = null)
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
