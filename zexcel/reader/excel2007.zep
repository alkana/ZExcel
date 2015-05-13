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
				let relationships = json_decode(json_encode(rels), 1);
				
				for rel in relationships["Relationship"] {
					switch (rel["@attributes"]["Type"]) {
						case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
							if (basename(rel["@attributes"]["Target"]) == "workbook.xml") {
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
		var excel, zipClass, zip;
		
		if (!file_exists(pFilename)) {
            throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
        }

        // Initialisations
        let excel = new \ZExcel\ZExcel();
        excel->removeSheetByIndex(0);
        
        if (!this->_readDataOnly) {
            excel->removeCellStyleXfByIndex(0); // remove the default style
            excel->removeCellXfByIndex(0); // remove the default style
        }

        let zipClass = \ZExcel\Settings::getZipClass();

        let zip = new {zipClass}();
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
		throw new \Exception("Not implemented yet!");
	}
}
