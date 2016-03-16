namespace ZExcel;

class CachedObjectStorageFactory
{
    const CACHE_IN_MEMORY            = "Memory";
    const CACHE_IN_MEMORY_GZIP       = "MemoryGZip";
    const CACHE_IN_MEMORY_SERIALIZED = "MemorySerialized";
    const CACHE_IGBINARY             = "Igbinary";
    const CACHE_TO_DISCISAM          = "DiscISAM";
    const CACHE_TO_APC               = "APC";
    const CACHE_TO_MEMCACHE          = "Memcache";
    const CACHE_TO_PHPTEMP           = "PHPTemp";
    const CACHE_TO_WINCACHE          = "Wincache";
    const CACHE_TO_SQLITE            = "SQLite";
    const CACHE_TO_SQLITE3           = "SQLite3";

    /**
     * Name of the method used for cell cacheing
     *
     * @var string
     */
    private static cacheStorageMethod = null;

    /**
     * Name of the class used for cell cacheing
     *
     * @var string
     */
    private static cacheStorageClass = null;

    /**
     * List of all possible cache storage methods
     *
     * @var string[]
     */
    private static storageMethods = [
        self::CACHE_IN_MEMORY,
        self::CACHE_IN_MEMORY_GZIP,
        self::CACHE_IN_MEMORY_SERIALIZED,
        self::CACHE_IGBINARY,
        self::CACHE_TO_PHPTEMP,
        self::CACHE_TO_DISCISAM,
        self::CACHE_TO_APC,
        self::CACHE_TO_MEMCACHE,
        self::CACHE_TO_WINCACHE,
        self::CACHE_TO_SQLITE,
        self::CACHE_TO_SQLITE3
    ];

    /**
     * Default arguments for each cache storage method
     *
     * @var array of mixed array
     */
    private static storageMethodDefaultParameters = [
        "Memory": [],
        "MemoryGZip": [],
        "MemorySerialized": [],
        "Igbinary": [],
        "PHPTemp": ["memoryCacheSize": "1MB"],
        "DiscISAM": ["dir": null],
        "APC": ["cacheTime": 600],
        "Memcache": [
            "memcacheServer": "localhost",
            "memcachePort": 11211,
            "cacheTime": 600
        ],
        "Wincache": ["cacheTime": 600],
        "SQLite": [],
        "SQLite3": []
    ];

    /**
     * Arguments for the active cache storage method
     *
     * @var array of mixed array
     */
    private static storageMethodParameters = [];
    
    /**
     * Return the current cache storage method
     *
     * @return string|null
     **/
    public static function getCacheStorageMethod()
    {
        return self::cacheStorageMethod;
    }

    /**
     * Return the current cache storage class
     *
     * @return PHPExcel_CachedObjectStorage_ICache|null
     **/
    public static function getCacheStorageClass()
    {
        return self::cacheStorageClass;
    }

    /**
     * Return the list of all possible cache storage methods
     *
     * @return string[]
     **/
    public static function getAllCacheStorageMethods()
    {
        return self::storageMethods;
    }

    /**
     * Return the list of all available cache storage methods
     *
     * @return string[]
     **/
    public static function getCacheStorageMethods()
    {
        var storageMethod, cacheStorageClass;
        array activeMethods = [];
        
        for storageMethod in self::storageMethods {
            let cacheStorageClass = "\\ZExcel\\CachedObjectStorage\\" . storageMethod;
            if (call_user_func([cacheStorageClass, "cacheMethodIsAvailable"])) {
                let activeMethods[] = storageMethod;
            }
        }
        
        return activeMethods;
    }

    /**
     * Identify the cache storage method to use
     *
     * @param    string            method        Name of the method to use for cell cacheing
     * @param    array of mixed    arguments    Additional arguments to pass to the cell caching class
     *                                        when instantiating
     * @return boolean
     **/
    public static function initialize(var method = self::CACHE_IN_MEMORY, arguments = []) -> boolean
    {
        var cacheStorageClass, k, v;
        
        if (!in_array(method, self::storageMethods)) {
            return false;
        }

        let cacheStorageClass = "\\ZExcel\\CachedObjectStorage\\" . method;
        if (!call_user_func([cacheStorageClass, "cacheMethodIsAvailable"])) {
            return false;
        }

        let self::storageMethodParameters[method] = self::storageMethodDefaultParameters[method];
        for k, v in arguments {
            if (array_key_exists(k, self::storageMethodParameters[method])) {
                let self::storageMethodParameters[method][k] = v;
            }
        }

        if (self::cacheStorageMethod === null) {
            let self::cacheStorageClass = "\\ZExcel\\CachedObjectStorage\\" . method;
            let self::cacheStorageMethod = method;
        }
        return true;
    }

    /**
     * Initialise the cache storage
     *
     * @param    PHPExcel_Worksheet     parent        Enable cell caching for this worksheet
     * @return    PHPExcel_CachedObjectStorage_ICache
     **/
    public static function getInstance(<\ZExcel\Worksheet> parent)
    {
        var instance, cacheMethodIsAvailable, functionn;
        
        let cacheMethodIsAvailable = true;
        if (self::cacheStorageMethod === null) {
            let cacheMethodIsAvailable = self::initialize();
        }

        if (cacheMethodIsAvailable) {
            let functionn = self::cacheStorageClass;
            
            // @FIXME Check the second parameter (cache not need, but trying to send it)
            let instance = new {functionn}(parent);
            
            if (instance !== null) {
                return instance;
            }
        }

        return false;
    }

    /**
     * Clear the cache storage
     *
     **/
    public static function finalize()
    {
        let self::cacheStorageMethod = null;
        let self::cacheStorageClass = null;
        let self::storageMethodParameters = [];
    }
}
