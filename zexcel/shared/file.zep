namespace ZExcel\Shared;

class File
{
	protected static _useUploadTempDirectory	= false;
	
	public static function setUseUploadTempDirectory(useUploadTempDir = false) {
		let self::_useUploadTempDirectory = (boolean) useUploadTempDir;
	}

	public static function getUseUploadTempDirectory() -> boolean
	{
		return self::_useUploadTempDirectory;
	}

	public static function file_exists(string pFilename) -> boolean
	{
		throw new \Exception("Not implemented yet!");
		/*
		var zip;
		boolean returnValue = false;
		string zipFile, archiveFile;
		
		// Sick construction, but it seems that
		// file_exists returns strange values when
		// doing the original file_exists on ZIP archives...
		if ( strtolower(substr(pFilename, 0, 3)) == "zip" ) {
			// Open ZIP file and verify if the file exists
			let zipFile 		= substr(pFilename, 6, strpos(pFilename, "#") - 6);
			let archiveFile 	= substr(pFilename, strpos(pFilename, "#") + 1);
			let zip = new ZipArchive();
			
			if (zip->open(zipFile) === true) {
				let returnValue = (zip->getFromName(archiveFile) !== false);
				zip->close();
				
				return returnValue;
			} else {
				return false;
			}
		} else {
			// Regular file_exists
			return file_exists(pFilename);
		}
		*/
	}

	public static function realpath(string pFilename) -> string
	{
		var returnValue = "", pathArray = null;
		int i = 0;
		
		if (file_exists(pFilename)) {
			let returnValue = realpath(pFilename);
		}

		if (returnValue == "" || (returnValue === NULL)) {
			let pathArray = explode("/" , pFilename);
			
			while(in_array("..", pathArray) && pathArray[0] != "..") {
				for i in range(0, count(pathArray)) {
					if (pathArray[i] == ".." && i > 0) {
						unset(pathArray[i]);
						unset(pathArray[i - 1]);
						break;
					}
				}
			}
			let returnValue = implode("/", pathArray);
		}

		return returnValue;
	}

	public static function sys_get_temp_dir() -> string
	{
		var temp;
	
		if (self::_useUploadTempDirectory) {
			//  use upload-directory when defined to allow running on environments having very restricted
			//      open_basedir configs
			if (ini_get("upload_tmp_dir") !== false) {
				let temp = ini_get("upload_tmp_dir");
				
				if (is_string(temp) && file_exists(temp)) {
					let temp = realpath(temp);
					
					return temp;
				}
			}
		}

		return realpath(sys_get_temp_dir());
	}
}
