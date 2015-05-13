namespace ZExcel;

class IOFactory
{
	private function __construct() { }

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
				return $instance;
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
				return $instance;
			}
		}
		
		// Nothing found...
		throw new Reader\Exception("No IReader found for type " . readerType);
	}

	public static function load(string pFilename) -> <Reader\IReader>
	{
		var reader;
		
		let reader = <Reader\IReader> self::createReaderForFile(pFilename);
		
		return reader->load(pFilename);
	}

	public static function identify($pFilename)
	{
	}

	public static function createReaderForFile(string pFilename) -> <Reader\IReader>
	{
		var pathinfo = null, reader = null;
		string extensionType = null;
		
		let pathinfo = pathinfo($pFilename);
		
		if (isset($pathinfo["extension"])) {
			switch (strtolower(pathinfo["extension"])) {
				case "xlsx":			//	Excel (OfficeOpenXML) Spreadsheet
				case "xlsm":			//	Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
				case "xltx":			//	Excel (OfficeOpenXML) Template
				case "xltm":			//	Excel (OfficeOpenXML) Macro Template (macros will be discarded)
					let extensionType = "Excel2007";
					break;
				case "xls":				//	Excel (BIFF) Spreadsheet
				case "xlt":				//	Excel (BIFF) Template
					let extensionType = "Excel5";
					break;
				case "ods":				//	Open/Libre Offic Calc
				case "ots":				//	Open/Libre Offic Calc Template
					let extensionType = "OOCalc";
					break;
				case "slk":
					let extensionType = "Sylk";
					break;
				case "xml":				//	Excel 2003 SpreadSheetML
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
				// Let's see if we are lucky
				if (is_object(reader) && reader->canRead(pFilename)) {
					return reader;
				}
			}
		}
		
		/*
		// If we reach here then "lucky guess" didn't give any result
		// Try walking through all the options in self::$_autoResolveClasses
		foreach (self::$_autoResolveClasses as $autoResolveClass) {
			//	Ignore our original guess, we know that won't work
			if ($autoResolveClass !== $extensionType) {
				$reader = self::createReader($autoResolveClass);
				if ($reader->canRead($pFilename)) {
					return $reader;
				}
			}
		}
		*/
		
		throw new Reader\Exception("Unable to identify a reader for this file");
	}
}
