namespace ZExcel;

class Settings
{
    const PCLZIP     = "\\ZExcel\\Shared\\ZipArchive";
    const ZIPARCHIVE = "\\ZipArchive";

    const CHART_RENDERER_JPGRAPH = "jpgraph";

    const PDF_RENDERER_TCPDF  = "tcPDF";
    const PDF_RENDERER_DOMPDF = "DomPDF";
    const PDF_RENDERER_MPDF   = "mPDF";
    
    private static chartRenderers = [
        self::CHART_RENDERER_JPGRAPH
    ];

    private static pdfRenderers = [
        self::PDF_RENDERER_TCPDF,
        self::PDF_RENDERER_DOMPDF,
        self::PDF_RENDERER_MPDF
    ];
    
    /**
     * Name of the class used for Zip file management
     *    e.g.
     *        ZipArchive
     *
     * @var string
     */
    private static zipClass = self::ZIPARCHIVE;

    /**
     * Name of the external Library used for rendering charts
     *    e.g.
     *        jpgraph
     *
     * @var string
     */
    private static chartRendererName;

    /**
     * Directory Path to the external Library used for rendering charts
     *
     * @var string
     */
    private static chartRendererPath;

    /**
     * Name of the external Library used for rendering PDF files
     *    e.g.
     *         mPDF
     *
     * @var string
     */
    private static pdfRendererName;

    /**
     * Directory Path to the external Library used for rendering PDF files
     *
     * @var string
     */
    private static pdfRendererPath;

    /**
     * Default options for libxml loader
     *
     * @var int
     */
    private static libXmlLoaderOptions = null;

    /**
     * Set the Zip handler Class that PHPExcel should use for Zip file management (PCLZip or ZipArchive)
     *
     * @param string zipClass    The Zip handler class that PHPExcel should use for Zip file management
     *      e.g. \ZExcel\Settings::PCLZip or \ZExcel\Settings::ZipArchive
     * @return    boolean    Success or failure
     */
    public static function setZipClass(var zipClass)
    {
        if ((zipClass === self::PCLZIP) || (zipClass === self::ZIPARCHIVE)) {
            let self::zipClass = zipClass;
            
            return true;
        }
        
        return false;
    }

    /**
     * Return the name of the Zip handler Class that PHPExcel is configured to use (PCLZip or ZipArchive)
     *    or Zip file management
     *
     * @return string Name of the Zip handler Class that PHPExcel is configured to use
     *    for Zip file management
     *    e.g. \ZExcel\Settings::PCLZip or \ZExcel\Settings::ZipArchive
     */
    public static function getZipClass() -> string
    {
        return self::zipClass;
    }

    /**
     * Return the name of the method that is currently configured for cell cacheing
     *
     * @return string Name of the cacheing method
     */
    public static function getCacheStorageMethod()
    {
        return \ZExcel\CachedObjectStorageFactory::getCacheStorageMethod();
    }

    /**
     * Return the name of the class that is currently being used for cell cacheing
     *
     * @return string Name of the class currently being used for cacheing
     */
    public static function getCacheStorageClass()
    {
        return \ZExcel\CachedObjectStorageFactory::getCacheStorageClass();
    }

    /**
     * Set the method that should be used for cell cacheing
     *
     * @param string method Name of the cacheing method
     * @param array arguments Optional configuration arguments for the cacheing method
     * @return boolean Success or failure
     */
    public static function setCacheStorageMethod(var method = null, var arguments = [])
    {
        return \ZExcel\CachedObjectStorageFactory::initialize(method, arguments);
    }

    /**
     * Set the locale code to use for formula translations and any special formatting
     *
     * @param string locale The locale code to use (e.g. "fr" or "pt_br" or "en_uk")
     * @return boolean Success or failure
     */
    public static function setLocale(var locale = "en_us")
    {
        return \ZExcel\Calculation::getInstance()->setLocale(locale);
    }

    /**
     * Set details of the external library that PHPExcel should use for rendering charts
     *
     * @param string libraryName    Internal reference name of the library
     *    e.g. \ZExcel\Settings::CHART_RENDERER_JPGRAPH
     * @param string libraryBaseDir Directory path to the library's base folder
     *
     * @return    boolean    Success or failure
     */
    public static function setChartRenderer(var libraryName, var libraryBaseDir)
    {
        if (!self::setChartRendererName(libraryName)) {
            return false;
        }
        
        return self::setChartRendererPath(libraryBaseDir);
    }

    /**
     * Identify to PHPExcel the external library to use for rendering charts
     *
     * @param string libraryName    Internal reference name of the library
     *    e.g. \ZExcel\Settings::CHART_RENDERER_JPGRAPH
     *
     * @return    boolean    Success or failure
     */
    public static function setChartRendererName(var libraryName)
    {
        if (!in_array(libraryName, self::chartRenderers)) {
            return false;
        }
        
        let self::chartRendererName = libraryName;

        return true;
    }

    /**
     * Tell PHPExcel where to find the external library to use for rendering charts
     *
     * @param string libraryBaseDir    Directory path to the library's base folder
     * @return    boolean    Success or failure
     */
    public static function setChartRendererPath(libraryBaseDir)
    {
        if ((file_exists(libraryBaseDir) === false) || (is_readable(libraryBaseDir) === false)) {
            return false;
        }
        
        let self::chartRendererPath = libraryBaseDir;

        return true;
    }

    /**
     * Return the Chart Rendering Library that PHPExcel is currently configured to use (e.g. jpgraph)
     *
     * @return string|NULL Internal reference name of the Chart Rendering Library that PHPExcel is
     *    currently configured to use
     *    e.g. \ZExcel\Settings::CHART_RENDERER_JPGRAPH
     */
    public static function getChartRendererName()
    {
        return self::chartRendererName;
    }

    /**
     * Return the directory path to the Chart Rendering Library that PHPExcel is currently configured to use
     *
     * @return string|NULL Directory Path to the Chart Rendering Library that PHPExcel is
     *     currently configured to use
     */
    public static function getChartRendererPath()
    {
        return self::chartRendererPath;
    }

    /**
     * Set details of the external library that PHPExcel should use for rendering PDF files
     *
     * @param string libraryName Internal reference name of the library
     *     e.g. \ZExcel\Settings::PDF_RENDERER_TCPDF,
     *     \ZExcel\Settings::PDF_RENDERER_DOMPDF
     *  or \ZExcel\Settings::PDF_RENDERER_MPDF
     * @param string libraryBaseDir Directory path to the library's base folder
     *
     * @return boolean Success or failure
     */
    public static function setPdfRenderer(var libraryName, var libraryBaseDir)
    {
        if (!self::setPdfRendererName(libraryName)) {
            return false;
        }
        
        return self::setPdfRendererPath(libraryBaseDir);
    }

    /**
     * Identify to PHPExcel the external library to use for rendering PDF files
     *
     * @param string libraryName Internal reference name of the library
     *     e.g. \ZExcel\Settings::PDF_RENDERER_TCPDF,
     *    \ZExcel\Settings::PDF_RENDERER_DOMPDF
     *     or \ZExcel\Settings::PDF_RENDERER_MPDF
     *
     * @return boolean Success or failure
     */
    public static function setPdfRendererName(var libraryName)
    {
        if (!in_array(libraryName, self::pdfRenderers)) {
            return false;
        }
        
        let self::pdfRendererName = libraryName;

        return true;
    }

    /**
     * Tell PHPExcel where to find the external library to use for rendering PDF files
     *
     * @param string libraryBaseDir Directory path to the library's base folder
     * @return boolean Success or failure
     */
    public static function setPdfRendererPath(var libraryBaseDir)
    {
        if ((file_exists(libraryBaseDir) === false) || (is_readable(libraryBaseDir) === false)) {
            return false;
        }
        
        let self::pdfRendererPath = libraryBaseDir;

        return true;
    }

    /**
     * Return the PDF Rendering Library that PHPExcel is currently configured to use (e.g. dompdf)
     *
     * @return string|NULL Internal reference name of the PDF Rendering Library that PHPExcel is
     *     currently configured to use
     *  e.g. \ZExcel\Settings::PDF_RENDERER_TCPDF,
     *  \ZExcel\Settings::PDF_RENDERER_DOMPDF
     *  or \ZExcel\Settings::PDF_RENDERER_MPDF
     */
    public static function getPdfRendererName()
    {
        return self::pdfRendererName;
    }

    /**
     * Return the directory path to the PDF Rendering Library that PHPExcel is currently configured to use
     *
     * @return string|NULL Directory Path to the PDF Rendering Library that PHPExcel is
     *        currently configured to use
     */
    public static function getPdfRendererPath()
    {
        return self::pdfRendererPath;
    }

    /**
     * Set default options for libxml loader
     *
     * @param int options Default options for libxml loader
     */
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

    /**
     * Get default options for libxml loader.
     * Defaults to LIBXML_DTDLOAD | LIBXML_DTDATTR when not set explicitly.
     *
     * @return int Default options for libxml loader
     */
    public static function getLibXmlLoaderOptions()
    {
        var tmp;
        
        if (is_null(self::libXmlLoaderOptions) && defined(LIBXML_DTDLOAD)) {
            let tmp = LIBXML_DTDLOAD | LIBXML_DTDATTR;
            
            self::setLibXmlLoaderOptions(tmp);
        }
        
        return self::libXmlLoaderOptions;
    }
}
