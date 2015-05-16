namespace ZExcel\Reader;

abstract class Abstrac implements IReader
{
	protected _readDataOnly = false;
	
	protected _includeCharts = false;
	
	protected _loadSheetsOnly = null;
	
	protected _readFilter = null;
	
	protected _fileHandle = null;

	public function getReadDataOnly() -> boolean
	{
		return this->_readDataOnly;
	}

	public function setReadDataOnly(boolean pValue = false) -> <\ZExcel\Reader\Abstrac>
	{
		let this->_readDataOnly = pValue;
		
		return this;
	}

	public function getIncludeCharts() -> boolean
	{
		return this->_includeCharts;
	}

	public function setIncludeCharts(boolean pValue = false) -> <\ZExcel\Reader\Abstrac>
	{
		let this->_includeCharts = pValue;
		
		return this;
	}

	public function getLoadSheetsOnly()
	{
		return this->_loadSheetsOnly;
	}

	public function setLoadSheetsOnly(string value = null) -> <\ZExcel\Reader\Abstrac>
	{
        if (value === null) {
            return this->setLoadAllSheets();
        }
        
        if (is_array(value)) {
        	let this->_loadSheetsOnly = value;
        } else {
         	let this->_loadSheetsOnly = [value];
         }
         
		return this;
	}

	public function setLoadAllSheets()
	{
		let this->_loadSheetsOnly = null;
		
		return this;
	}

	public function getReadFilter() -> <\ZExcel\Reader\IReadFilter>
	{
		return this->_readFilter;
	}

	public function setReadFilter(<\ZExcel\Reader\IReadFilter> pValue) -> <\ZExcel\Reader\Abstrac>
	{
		let this->_readFilter = pValue;
		
		return this;
	}

	protected function _openFile(pFilename)
	{
		// Check if file exists
		if (!file_exists(pFilename) || !is_readable(pFilename)) {
			throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
		}
		
		// Open file
		let this->_fileHandle = fopen(pFilename, "r");
		
		if (this->_fileHandle === false) {
			throw new \ZExcel\Reader\Exception("Could not open file " . pFilename . " for reading.");
		}
	}

	public function canRead(pFilename) -> boolean
	{
		// Check if file exists
		try {
			this->_openFile(pFilename);
		} catch \Exception {
			return false;
		}
		
		fclose (this->_fileHandle);
		
		return false;
	}

	public function securityScan(xml)
	{
        string pattern;
        
        let pattern = "/\\0?" . implode("\\0?", str_split("<!DOCTYPE")) . "\\0?/";
        
        if (preg_match(pattern, xml)) { 
            throw new \ZExcel\Reader\Exception("Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks");
        }
        
        return xml;
    }

	public function securityScanFile(filestream)
	{
        return this->securityScan(file_get_contents(filestream));
    }
}
