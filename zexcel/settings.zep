namespace ZExcel;

class Settings
{
	const PCLZIP     = "\ZExcel\Shared\ZipArchive";
    const ZIPARCHIVE = "\ZipArchive";

    const CHART_RENDERER_JPGRAPH = "jpgraph";

    const PDF_RENDERER_TCPDF  = "tcPDF";
    const PDF_RENDERER_DOMPDF = "DomPDF";
    const PDF_RENDERER_MPDF   = "mPDF";
    
    private static zipClass = self::ZIPARCHIVE;

    private static chartRendererName;

    private static chartRendererPath;

    private static pdfRendererName;

    private static pdfRendererPath;

    private static libXmlLoaderOptions = null;

    public static function setZipClass(zipClass)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getZipClass() -> string
    {
        return self::zipClass;
    }

    public static function getCacheStorageMethod()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getCacheStorageClass()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setCacheStorageMethod(var method = null, arguments = [])
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setLocale(locale = "en_us")
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setChartRenderer(libraryName, libraryBaseDir)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setChartRendererName(libraryName)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setChartRendererPath(libraryBaseDir)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getChartRendererName()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getChartRendererPath()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setPdfRenderer(libraryName, libraryBaseDir)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setPdfRendererName(libraryName)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setPdfRendererPath(libraryBaseDir)
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getPdfRendererName()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function getPdfRendererPath()
    {
        throw new \Exception("Not implemented yet!");
    }

    public static function setLibXmlLoaderOptions(var options = null)
    {
    	var tmp = 0;
    	
        if (is_null(options) && defined(LIBXML_DTDLOAD)) {
        	let tmp = LIBXML_DTDLOAD | LIBXML_DTDATTR;
            let options = LIBXML_DTDLOAD | LIBXML_DTDATTR;
        }
        
        let tmp = (options == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        
		libxml_disable_entity_loader(tmp);
        
		let self::libXmlLoaderOptions = options;
    }

    public static function getLibXmlLoaderOptions()
    {
		var tmp;
		
		if (is_null(self::libXmlLoaderOptions) && defined(LIBXML_DTDLOAD)) {
            let tmp = LIBXML_DTDLOAD | LIBXML_DTDATTR;
            
            self::setLibXmlLoaderOptions(tmp);
        }
        
        // libxml_disable_entity_loader(self::libXmlLoaderOptions == (LIBXML_DTDLOAD | LIBXML_DTDATTR));
        
        return self::libXmlLoaderOptions;
    }
}
