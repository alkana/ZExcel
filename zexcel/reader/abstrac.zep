namespace ZExcel\Reader;

abstract class Abstrac implements IReader
{
    /**
     * Read data only?
     * Identifies whether the Reader should only read data values for cells, and ignore any formatting information;
     *        or whether it should read both data and formatting
     *
     * @var    boolean
     */
	protected readDataOnly = false;
	
    /**
     * Read empty cells?
     * Identifies whether the Reader should read data values for cells all cells, or should ignore cells containing
     *         null value or empty string
     *
     * @var    boolean
     */
    protected readEmptyCells = true;
    
    /**
     * Read charts that are defined in the workbook?
     * Identifies whether the Reader should read the definitions for any charts that exist in the workbook;
     *
     * @var    boolean
     */
	protected includeCharts = false;
	
    /**
     * Restrict which sheets should be loaded?
     * This property holds an array of worksheet names to be loaded. If null, then all worksheets will be loaded.
     *
     * @var array of string
     */
	protected loadSheetsOnly = null;
	
    /**
     * PHPExcel_Reader_IReadFilter instance
     *
     * @var PHPExcel_Reader_IReadFilter
     */
	protected readFilter = null;
	
	protected fileHandle = null;

    /**
     * Read data only?
     *        If this is true, then the Reader will only read data values for cells, it will not read any formatting information.
     *        If false (the default) it will read data and formatting.
     *
     * @return    boolean
     */
	public function getReadDataOnly() -> boolean
	{
		return this->readDataOnly;
	}

    /**
     * Set read data only
     *        Set to true, to advise the Reader only to read data values for cells, and to ignore any formatting information.
     *        Set to false (the default) to advise the Reader to read both data and formatting for cells.
     *
     * @param    boolean    pValue
     *
     * @return    \ZExcel\Reader\IReader
     */
	public function setReadDataOnly(boolean pValue = false) -> <\ZExcel\Reader\Abstrac>
	{
		let this->readDataOnly = pValue;
		
		return this;
	}

    /**
     * Read empty cells?
     *        If this is true (the default), then the Reader will read data values for all cells, irrespective of value.
     *        If false it will not read data for cells containing a null value or an empty string.
     *
     * @return    boolean
     */
    public function getReadEmptyCells()
    {
        return this->readEmptyCells;
    }

    /**
     * Set read empty cells
     *        Set to true (the default) to advise the Reader read data values for all cells, irrespective of value.
     *        Set to false to advise the Reader to ignore cells containing a null value or an empty string.
     *
     * @param    boolean    pValue
     *
     * @return    \ZExcel\Reader\IReader
     */
    public function setReadEmptyCells(boolean pValue = true) -> <\ZExcel\Reader\Abstrac>
    {
        let this->readEmptyCells = pValue;
        
        return this;
    }

    /**
     * Read charts in workbook?
     *        If this is true, then the Reader will include any charts that exist in the workbook.
     *      Note that a ReadDataOnly value of false overrides, and charts won't be read regardless of the IncludeCharts value.
     *        If false (the default) it will ignore any charts defined in the workbook file.
     *
     * @return    boolean
     */
	public function getIncludeCharts() -> boolean
	{
		return this->includeCharts;
	}

    /**
     * Set read charts in workbook
     *        Set to true, to advise the Reader to include any charts that exist in the workbook.
     *      Note that a ReadDataOnly value of false overrides, and charts won't be read regardless of the IncludeCharts value.
     *        Set to false (the default) to discard charts.
     *
     * @param    boolean    pValue
     *
     * @return    \ZExcel\Reader\IReader
     */
	public function setIncludeCharts(boolean pValue = false) -> <\ZExcel\Reader\Abstrac>
	{
		let this->includeCharts = pValue;
		
		return this;
	}

    /**
     * Get which sheets to load
     * Returns either an array of worksheet names (the list of worksheets that should be loaded), or a null
     *        indicating that all worksheets in the workbook should be loaded.
     *
     * @return mixed
     */
	public function getLoadSheetsOnly()
	{
		return this->loadSheetsOnly;
	}

    /**
     * Set which sheets to load
     *
     * @param mixed value
     *        This should be either an array of worksheet names to be loaded, or a string containing a single worksheet name.
     *        If NULL, then it tells the Reader to read all worksheets in the workbook
     *
     * @return \ZExcel\Reader\IReader
     */
	public function setLoadSheetsOnly(string value = null) -> <\ZExcel\Reader\Abstrac>
	{
        if (value === null) {
            return this->setLoadAllSheets();
        }
        
        if (is_array(value)) {
        	let this->loadSheetsOnly = value;
        } else {
         	let this->loadSheetsOnly = [value];
         }
         
		return this;
	}

    /**
     * Set all sheets to load
     *        Tells the Reader to load all worksheets from the workbook.
     *
     * @return \ZExcel\Reader\IReader
     */
	public function setLoadAllSheets()
	{
		let this->loadSheetsOnly = null;
		
		return this;
	}

    /**
     * Read filter
     *
     * @return \ZExcel\Reader\IReadFilter
     */
	public function getReadFilter() -> <\ZExcel\Reader\IReadFilter>
	{
		return this->readFilter;
	}

    /**
     * Set read filter
     *
     * @param \ZExcel\Reader\IReadFilter pValue
     * @return \ZExcel\Reader\IReader
     */
	public function setReadFilter(<\ZExcel\Reader\IReadFilter> pValue) -> <\ZExcel\Reader\Abstrac>
	{
		let this->readFilter = pValue;
		
		return this;
	}

    /**
     * Open file for reading
     *
     * @param string pFilename
     * @throws    \ZExcel\Reader\Exception
     * @return resource
     */
	protected function _openFile(string pFilename)
	{
		// Check if file exists
		if (!file_exists(pFilename) || !is_readable(pFilename)) {
			throw new \ZExcel\Reader\Exception("Could not open " . pFilename . " for reading! File does not exist.");
		}
		
		// Open file
		let this->fileHandle = fopen(pFilename, "r");
		
		if (this->fileHandle === false) {
			throw new \ZExcel\Reader\Exception("Could not open file " . pFilename . " for reading.");
		}
	}

    /**
     * Can the current \ZExcel\Reader\IReader read the file?
     *
     * @param     string         pFilename
     * @return boolean
     * @throws \ZExcel\Reader\Exception
     */
	public function canRead(string pFilename) -> boolean
	{
		// Check if file exists
		try {
			this->_openFile(pFilename);
		} catch \Exception {
			return false;
		}
		
		fclose (this->fileHandle);
		
		return false;
	}

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param     string         xml
     * @throws \ZExcel\Reader\Exception
     */
	public function securityScan(string xml)
	{
        string pattern;
        
        let pattern = "/\\0?" . implode("\\0?", str_split("<!DOCTYPE")) . "\\0?/";
        
        if (preg_match(pattern, xml)) { 
            throw new \ZExcel\Reader\Exception("Detected use of ENTITY in XML, spreadsheet file load() aborted to prevent XXE/XEE attacks");
        }
        
        return xml;
    }

    /**
     * Scan theXML for use of <!ENTITY to prevent XXE/XEE attacks
     *
     * @param  string filestream
     * @throws \ZExcel\Reader\Exception
     */
	public function securityScanFile(string filestream)
	{
        return this->securityScan(file_get_contents(filestream));
    }
}
