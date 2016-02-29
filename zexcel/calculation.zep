namespace ZExcel;

class Calculation
{
    //    @FIXME conditional constants (@see PHPExcel_Calculation class)
    //    Cell reference (cell or range of cells, with or without a sheet reference)
    const CALCULATION_REGEXP_CELLREF = "((([^\s,!&%^\/\*\+<>=-]*)|(\'[^\']*\')|(\"[^\"]*\"))!)?\$?([a-z]{1,3})\$?(\d{1,7})";
    //    Named Range of cells
    const CALCULATION_REGEXP_NAMEDRANGE = "((([^\s,!&%^\/\*\+<>=-]*)|(\'[^\']*\')|(\"[^\"]*\"))!)?([_A-Z][_A-Z0-9\.]*)";
            
            
    //    Numeric operand
    const CALCULATION_REGEXP_NUMBER        = "[-+]?\d*\.?\d+(e[-+]?\d+)?";
    //    String operand
    const CALCULATION_REGEXP_STRING        = "\"(?:[^\"]|\"\")*\"\"";
    //    Opening bracket
    const CALCULATION_REGEXP_OPENBRACE    = "\(";
    //    Function (allow for the old @ symbol that could be used to prefix a function, but we'll ignore it)
    const CALCULATION_REGEXP_FUNCTION    = "@?([A-Z][A-Z0-9\.]*)[\s]*\(";
    //    Error
    const CALCULATION_REGEXP_ERROR        = "\#[A-Z][A-Z0_\/]*[!\?]?";


    /** constants */
    const RETURN_ARRAY_AS_ERROR = "error";
    const RETURN_ARRAY_AS_VALUE = "value";
    const RETURN_ARRAY_AS_ARRAY = "array";

    private static returnArrayAsType = self::RETURN_ARRAY_AS_VALUE;


    /**
     * Instance of this class
     *
     * @access    private
     * @var \ZExcel\Calculation
     */
    private static instance;


    /**
     * Instance of the workbook this Calculation Engine is using
     *
     * @access    private
     * @var PHPExcel
     */
    private workbook;

    private delta = null;
    
    private debugLog = null;

    /**
     * List of instances of the calculation engine that we've instantiated for individual workbooks
     *
     * @access    private
     * @var \ZExcel\Calculation[]
     */
    private static workbookSets;

    /**
     * Calculation cache
     *
     * @access    private
     * @var array
     */
    private _calculationCache = [];


    /**
     * Calculation cache enabled
     *
     * @access    private
     * @var boolean
     */
    private calculationCacheEnabled = true;

    /**
     * Flag to determine how formula errors should be handled
     *        If true, then a user error will be triggered
     *        If false, then an exception will be thrown
     *
     * @access    public
     * @var boolean
     *
     */
    public suppressFormulaErrors = false;

    /**
     * Error message for any error that was raised/thrown by the calculation engine
     *
     * @access    public
     * @var string
     *
     */
    public formulaError = null;

    /**
     * An array of the nested cell references accessed by the calculation engine, used for the debug log
     *
     * @access    private
     * @var array of string
     *
     */
    private cyclicReferenceStack;

    private cellStack = [];

    /**
     * Number of iterations for cyclic formulae
     *
     * @var integer
     *
     */
    private cyclicFormulaCounter = 1;
    
    private cyclicFormulaCell = "";

    /**
     * Number of iterations for cyclic formulae
     *
     * @var integer
     *
     */
    public cyclicFormulaCount = 1;
    
    /**
     * Precision used for calculations
     *
     * @var integer
     *
     */
    private savedPrecision = 14;


    /**
     * The current locale setting
     *
     * @var string
     *
     */
    private static _localeLanguage = "en_us";                    //    US English    (default locale)
    
    /**
     * Locale-specific argument separator for function arguments
     *
     * @var string
     *
     */
    private static localeArgumentSeparator = ",";
    
    private static localeFunctions = [];
    
    private static binaryOperators = [
        "+": true, "-": true, "*": true, "/": true,
        "^": true, "&": true, ">": true, "<": true,
        "=": true, ">=": true, "<=": true, "<>": true,
        "|": true, ":": true
    ];
    
    public static localeBoolean = [
        "TRUE": "TRUE",
        "FALSE": "FALSE",
        "NULL": "NULL"
    ];
    
    private static PHPExcelFunctions = [
        "ABS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "abs",
            "argumentCount": "1"
        ],
        "ACCRINT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::ACCRINT",
            "argumentCount": "4-7"
        ],
        "ACCRINTM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::ACCRINTM",
            "argumentCount": "3-5"
        ],
        "ACOS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "acos",
            "argumentCount": "1"
        ],
        "ACOSH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "acosh",
            "argumentCount": "1"
        ],
        "ADDRESS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::CELL_ADDRESS",
            "argumentCount": "2-5"
        ],
        "AMORDEGRC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::AMORDEGRC",
            "argumentCount": "6,7"
        ],
        "AMORLINC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::AMORLINC",
            "argumentCount": "6,7"
        ],
        "AND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::LOGICAL_AND",
            "argumentCount": "1+"
        ],
        "AREAS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "ASC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "ASIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "asin",
            "argumentCount": "1"
        ],
        "ASINH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "asinh",
            "argumentCount": "1"
        ],
        "ATAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "atan",
            "argumentCount": "1"
        ],
        "ATAN2": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::ATAN2",
            "argumentCount": "2"
        ],
        "ATANH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "atanh",
            "argumentCount": "1"
        ],
        "AVEDEV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::AVEDEV",
            "argumentCount": "1+"
        ],
        "AVERAGE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::AVERAGE",
            "argumentCount": "1+"
        ],
        "AVERAGEA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::AVERAGEA",
            "argumentCount": "1+"
        ],
        "AVERAGEIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::AVERAGEIF",
            "argumentCount": "2,3"
        ],
        "AVERAGEIFS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "3+"
        ],
        "BAHTTEXT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "BESSELI": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BESSELI",
            "argumentCount": "2"
        ],
        "BESSELJ": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::desselj",
            "argumentCount": "2"
        ],
        "BESSELK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BESSELK",
            "argumentCount": "2"
        ],
        "BESSELY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BESSELY",
            "argumentCount": "2"
        ],
        "BETADIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::BETADIST",
            "argumentCount": "3-5"
        ],
        "BETAINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::BETAINV",
            "argumentCount": "3-5"
        ],
        "BIN2DEC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BINTODEC",
            "argumentCount": "1"
        ],
        "BIN2HEX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BINTOHEX",
            "argumentCount": "1,2"
        ],
        "BIN2OCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::BINTOOCT",
            "argumentCount": "1,2"
        ],
        "BINOMDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::BINOMDIST",
            "argumentCount": "4"
        ],
        "CEILING": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::CEILING",
            "argumentCount": "2"
        ],
        "CELL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1,2"
        ],
        "CHAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::CHARACTER",
            "argumentCount": "1"
        ],
        "CHIDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CHIDIST",
            "argumentCount": "2"
        ],
        "CHIINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CHIINV",
            "argumentCount": "2"
        ],
        "CHITEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "CHOOSE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::CHOOSE",
            "argumentCount": "2+"
        ],
        "CLEAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::TRIMNONPRINTABLE",
            "argumentCount": "1"
        ],
        "CODE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::ASCIICODE",
            "argumentCount": "1"
        ],
        "COLUMN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::COLUMN",
            "argumentCount": "-1",
            "passByReference": [
                TRUE
            ]
        ],
        "COLUMNS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::COLUMNS",
            "argumentCount": "1"
        ],
        "COMBIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::COMBIN",
            "argumentCount": "2"
        ],
        "COMPLEX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::COMPLEX",
            "argumentCount": "2,3"
        ],
        "CONCATENATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::CONCATENATE",
            "argumentCount": "1+"
        ],
        "CONFIDENCE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CONFIDENCE",
            "argumentCount": "3"
        ],
        "CONVERT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::CONVERTUOM",
            "argumentCount": "3"
        ],
        "CORREL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CORREL",
            "argumentCount": "2"
        ],
        "COS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "cos",
            "argumentCount": "1"
        ],
        "COSH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "cosh",
            "argumentCount": "1"
        ],
        "COUNT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::COUNT",
            "argumentCount": "1+"
        ],
        "COUNTA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::COUNTA",
            "argumentCount": "1+"
        ],
        "COUNTBLANK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::COUNTBLANK",
            "argumentCount": "1"
        ],
        "COUNTIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::COUNTIF",
            "argumentCount": "2"
        ],
        "COUNTIFS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "COUPDAYBS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPDAYBS",
            "argumentCount": "3,4"
        ],
        "COUPDAYS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPDAYS",
            "argumentCount": "3,4"
        ],
        "COUPDAYSNC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPDAYSNC",
            "argumentCount": "3,4"
        ],
        "COUPNCD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPNCD",
            "argumentCount": "3,4"
        ],
        "COUPNUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPNUM",
            "argumentCount": "3,4"
        ],
        "COUPPCD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::COUPPCD",
            "argumentCount": "3,4"
        ],
        "COVAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::COVAR",
            "argumentCount": "2"
        ],
        "CRITBINOM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CRITBINOM",
            "argumentCount": "3"
        ],
        "CUBEKPIMEMBER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBEMEMBER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBEMEMBERPROPERTY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBERANKEDMEMBER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBESET": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBESETCOUNT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUBEVALUE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_CUBE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "CUMIPMT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::CUMIPMT",
            "argumentCount": "6"
        ],
        "CUMPRINC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::CUMPRINC",
            "argumentCount": "6"
        ],
        "DATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DATE",
            "argumentCount": "3"
        ],
        "DATEDIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DATEDIF",
            "argumentCount": "2,3"
        ],
        "DATEVALUE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DATEVALUE",
            "argumentCount": "1"
        ],
        "DAVERAGE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DAVERAGE",
            "argumentCount": "3"
        ],
        "DAY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DAYOFMONTH",
            "argumentCount": "1"
        ],
        "DAYS360": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DAYS360",
            "argumentCount": "2,3"
        ],
        "DB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::DB",
            "argumentCount": "4,5"
        ],
        "DCOUNT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DCOUNT",
            "argumentCount": "3"
        ],
        "DCOUNTA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DCOUNTA",
            "argumentCount": "3"
        ],
        "DDB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::DDB",
            "argumentCount": "4,5"
        ],
        "DEC2BIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::DECTOBIN",
            "argumentCount": "1,2"
        ],
        "DEC2HEX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::DECTOHEX",
            "argumentCount": "1,2"
        ],
        "DEC2OCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::DECTOOCT",
            "argumentCount": "1,2"
        ],
        "DEGREES": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "rad2deg",
            "argumentCount": "1"
        ],
        "DELTA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::DELTA",
            "argumentCount": "1,2"
        ],
        "DEVSQ": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::DEVSQ",
            "argumentCount": "1+"
        ],
        "DGET": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DGET",
            "argumentCount": "3"
        ],
        "DISC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::DISC",
            "argumentCount": "4,5"
        ],
        "DMAX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DMAX",
            "argumentCount": "3"
        ],
        "DMIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DMIN",
            "argumentCount": "3"
        ],
        "DOLLAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::DOLLAR",
            "argumentCount": "1,2"
        ],
        "DOLLARDE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::DOLLARDE",
            "argumentCount": "2"
        ],
        "DOLLARFR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::DOLLARFR",
            "argumentCount": "2"
        ],
        "DPRODUCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DPRODUCT",
            "argumentCount": "3"
        ],
        "DSTDEV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DSTDEV",
            "argumentCount": "3"
        ],
        "DSTDEVP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DSTDEVP",
            "argumentCount": "3"
        ],
        "DSUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DSUM",
            "argumentCount": "3"
        ],
        "DURATION": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "5,6"
        ],
        "DVAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DVAR",
            "argumentCount": "3"
        ],
        "DVARP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATABASE,
            "functionCall": "\\ZExcel\\Calculation\\Database::DVARP",
            "argumentCount": "3"
        ],
        "EDATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::EDATE",
            "argumentCount": "2"
        ],
        "EFFECT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::EFFECT",
            "argumentCount": "2"
        ],
        "EOMONTH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::EOMONTH",
            "argumentCount": "2"
        ],
        "ERF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::ERF",
            "argumentCount": "1,2"
        ],
        "ERFC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::ERFC",
            "argumentCount": "1"
        ],
        "ERROR.TYPE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::ERROR_TYPE",
            "argumentCount": "1"
        ],
        "EVEN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::EVEN",
            "argumentCount": "1"
        ],
        "EXACT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "EXP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "exp",
            "argumentCount": "1"
        ],
        "EXPONDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::EXPONDIST",
            "argumentCount": "3"
        ],
        "FACT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::FACT",
            "argumentCount": "1"
        ],
        "FACTDOUBLE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::FACTDOUBLE",
            "argumentCount": "1"
        ],
        "FALSE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::FALSEE",
            "argumentCount": "0"
        ],
        "FDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "3"
        ],
        "FIND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::SEARCHSENSITIVE",
            "argumentCount": "2,3"
        ],
        "FINDB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::SEARCHSENSITIVE",
            "argumentCount": "2,3"
        ],
        "FINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "3"
        ],
        "FISHER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::FISHER",
            "argumentCount": "1"
        ],
        "FISHERINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::FISHERINV",
            "argumentCount": "1"
        ],
        "FIXED": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::FIXEDFORMAT",
            "argumentCount": "1-3"
        ],
        "FLOOR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::FLOOR",
            "argumentCount": "2"
        ],
        "FORECAST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::FORECAST",
            "argumentCount": "3"
        ],
        "FREQUENCY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "FTEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "FV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::FV",
            "argumentCount": "3-5"
        ],
        "FVSCHEDULE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::FVSCHEDULE",
            "argumentCount": "2"
        ],
        "GAMMADIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::GAMMADIST",
            "argumentCount": "4"
        ],
        "GAMMAINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::GAMMAINV",
            "argumentCount": "3"
        ],
        "GAMMALN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::GAMMALN",
            "argumentCount": "1"
        ],
        "GCD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::GCD",
            "argumentCount": "1+"
        ],
        "GEOMEAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::GEOMEAN",
            "argumentCount": "1+"
        ],
        "GESTEP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::GESTEP",
            "argumentCount": "1,2"
        ],
        "GETPIVOTDATA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2+"
        ],
        "GROWTH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::GROWTH",
            "argumentCount": "1-4"
        ],
        "HARMEAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::HARMEAN",
            "argumentCount": "1+"
        ],
        "HEX2BIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::HEXTOBIN",
            "argumentCount": "1,2"
        ],
        "HEX2DEC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::HEXTODEC",
            "argumentCount": "1"
        ],
        "HEX2OCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::HEXTOOCT",
            "argumentCount": "1,2"
        ],
        "HLOOKUP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::HLOOKUP",
            "argumentCount": "3,4"
        ],
        "HOUR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::HOUROFDAY",
            "argumentCount": "1"
        ],
        "HYPERLINK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::HYPERLINK",
            "argumentCount": "1,2",
            "passCellReference": TRUE
        ],
        "HYPGEOMDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::HYPGEOMDIST",
            "argumentCount": "4"
        ],
        "IF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::STATEMENT_IF",
            "argumentCount": "1-3"
        ],
        "IFERROR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::IFERROR",
            "argumentCount": "2"
        ],
        "IMABS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMABS",
            "argumentCount": "1"
        ],
        "IMAGINARY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMAGINARY",
            "argumentCount": "1"
        ],
        "IMARGUMENT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMARGUMENT",
            "argumentCount": "1"
        ],
        "IMCONJUGATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMCONJUGATE",
            "argumentCount": "1"
        ],
        "IMCOS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMCOS",
            "argumentCount": "1"
        ],
        "IMDIV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMDIV",
            "argumentCount": "2"
        ],
        "IMEXP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMEXP",
            "argumentCount": "1"
        ],
        "IMLN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMLN",
            "argumentCount": "1"
        ],
        "IMLOG10": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMLOG10",
            "argumentCount": "1"
        ],
        "IMLOG2": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMLOG2",
            "argumentCount": "1"
        ],
        "IMPOWER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMPOWER",
            "argumentCount": "2"
        ],
        "IMPRODUCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMPRODUCT",
            "argumentCount": "1+"
        ],
        "IMREAL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMREAL",
            "argumentCount": "1"
        ],
        "IMSIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMSIN",
            "argumentCount": "1"
        ],
        "IMSQRT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMSQRT",
            "argumentCount": "1"
        ],
        "IMSUB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMSUB",
            "argumentCount": "2"
        ],
        "IMSUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::IMSUM",
            "argumentCount": "1+"
        ],
        "INDEX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::INDEX",
            "argumentCount": "1-4"
        ],
        "INDIRECT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::INDIRECT",
            "argumentCount": "1,2",
            "passCellReference": TRUE
        ],
        "INFO": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "INT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::INTT",
            "argumentCount": "1"
        ],
        "INTERCEPT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::INTERCEPT",
            "argumentCount": "2"
        ],
        "INTRATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::INTRATE",
            "argumentCount": "4,5"
        ],
        "IPMT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::IPMT",
            "argumentCount": "4-6"
        ],
        "IRR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::IRR",
            "argumentCount": "1,2"
        ],
        "ISBLANK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_BLANK",
            "argumentCount": "1"
        ],
        "ISERR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_ERR",
            "argumentCount": "1"
        ],
        "ISERROR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_ERROR",
            "argumentCount": "1"
        ],
        "ISEVEN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_EVEN",
            "argumentCount": "1"
        ],
        "ISLOGICAL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_LOGICAL",
            "argumentCount": "1"
        ],
        "ISNA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_NA",
            "argumentCount": "1"
        ],
        "ISNONTEXT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_NONTEXT",
            "argumentCount": "1"
        ],
        "ISNUMBER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_NUMBER",
            "argumentCount": "1"
        ],
        "ISODD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_ODD",
            "argumentCount": "1"
        ],
        "ISPMT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::ISPMT",
            "argumentCount": "4"
        ],
        "ISREF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "ISTEXT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::IS_TEXT",
            "argumentCount": "1"
        ],
        "JIS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "KURT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::KURT",
            "argumentCount": "1+"
        ],
        "LARGE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::LARGE",
            "argumentCount": "2"
        ],
        "LCM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::LCM",
            "argumentCount": "1+"
        ],
        "LEFT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::LEFT",
            "argumentCount": "1,2"
        ],
        "LEFTB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::LEFT",
            "argumentCount": "1,2"
        ],
        "LEN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::STRINGLENGTH",
            "argumentCount": "1"
        ],
        "LENB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::STRINGLENGTH",
            "argumentCount": "1"
        ],
        "LINEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::LINEST",
            "argumentCount": "1-4"
        ],
        "LN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "log",
            "argumentCount": "1"
        ],
        "LOG": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::LOG_BASE",
            "argumentCount": "1,2"
        ],
        "LOG10": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "log10",
            "argumentCount": "1"
        ],
        "LOGEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::LOGEST",
            "argumentCount": "1-4"
        ],
        "LOGINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::LOGINV",
            "argumentCount": "3"
        ],
        "LOGNORMDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::LOGNORMDIST",
            "argumentCount": "3"
        ],
        "LOOKUP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::LOOKUP",
            "argumentCount": "2,3"
        ],
        "LOWER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::LOWERCASE",
            "argumentCount": "1"
        ],
        "MATCH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::MATCH",
            "argumentCount": "2,3"
        ],
        "MAX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MAX",
            "argumentCount": "1+"
        ],
        "MAXA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MAXA",
            "argumentCount": "1+"
        ],
        "MAXIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MAXIF",
            "argumentCount": "2+"
        ],
        "MDETERM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MDETERM",
            "argumentCount": "1"
        ],
        "MDURATION": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "5,6"
        ],
        "MEDIAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MEDIAN",
            "argumentCount": "1+"
        ],
        "MEDIANIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2+"
        ],
        "MID": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::MID",
            "argumentCount": "3"
        ],
        "MIDB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::MID",
            "argumentCount": "3"
        ],
        "MIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MIN",
            "argumentCount": "1+"
        ],
        "MINA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MINA",
            "argumentCount": "1+"
        ],
        "MINIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MINIF",
            "argumentCount": "2+"
        ],
        "MINUTE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::MINUTEOFHOUR",
            "argumentCount": "1"
        ],
        "MINVERSE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MINVERSE",
            "argumentCount": "1"
        ],
        "MIRR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::MIRR",
            "argumentCount": "3"
        ],
        "MMULT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MMULT",
            "argumentCount": "2"
        ],
        "MOD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MOD",
            "argumentCount": "2"
        ],
        "MODE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::MODE",
            "argumentCount": "1+"
        ],
        "MONTH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::MONTHOFYEAR",
            "argumentCount": "1"
        ],
        "MROUND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MROUND",
            "argumentCount": "2"
        ],
        "MULTINOMIAL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::MULTINOMIAL",
            "argumentCount": "1+"
        ],
        "N": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::N",
            "argumentCount": "1"
        ],
        "NA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::NA",
            "argumentCount": "0"
        ],
        "NEGBINOMDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::NEGBINOMDIST",
            "argumentCount": "3"
        ],
        "NETWORKDAYS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::NETWORKDAYS",
            "argumentCount": "2+"
        ],
        "NOMINAL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::NOMINAL",
            "argumentCount": "2"
        ],
        "NORMDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::NORMDIST",
            "argumentCount": "4"
        ],
        "NORMINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::NORMINV",
            "argumentCount": "3"
        ],
        "NORMSDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::NORMSDIST",
            "argumentCount": "1"
        ],
        "NORMSINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::NORMSINV",
            "argumentCount": "1"
        ],
        "NOT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::NOT",
            "argumentCount": "1"
        ],
        "NOW": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DATETIMENOW",
            "argumentCount": "0"
        ],
        "NPER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::NPER",
            "argumentCount": "3-5"
        ],
        "NPV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::NPV",
            "argumentCount": "2+"
        ],
        "OCT2BIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::OCTTOBIN",
            "argumentCount": "1,2"
        ],
        "OCT2DEC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::OCTTODEC",
            "argumentCount": "1"
        ],
        "OCT2HEX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_ENGINEERING,
            "functionCall": "\\ZExcel\\Calculation\\Engineering::OCTTOHEX",
            "argumentCount": "1,2"
        ],
        "ODD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::ODD",
            "argumentCount": "1"
        ],
        "ODDFPRICE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "8,9"
        ],
        "ODDFYIELD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "8,9"
        ],
        "ODDLPRICE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "7,8"
        ],
        "ODDLYIELD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "7,8"
        ],
        "OFFSET": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::OFFSET",
            "argumentCount": "3,5",
            "passCellReference": TRUE,
            "passByReference": [
                TRUE
            ]
        ],
        "OR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::LOGICAL_OR",
            "argumentCount": "1+"
        ],
        "PEARSON": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::CORREL",
            "argumentCount": "2"
        ],
        "PERCENTILE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::PERCENTILE",
            "argumentCount": "2"
        ],
        "PERCENTRANK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::PERCENTRANK",
            "argumentCount": "2,3"
        ],
        "PERMUT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::PERMUT",
            "argumentCount": "2"
        ],
        "PHONETIC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1"
        ],
        "PI": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "pi",
            "argumentCount": "0"
        ],
        "PMT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PMT",
            "argumentCount": "3-5"
        ],
        "POISSON": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::POISSON",
            "argumentCount": "3"
        ],
        "POWER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::POWER",
            "argumentCount": "2"
        ],
        "PPMT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PPMT",
            "argumentCount": "4-6"
        ],
        "PRICE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PRICE",
            "argumentCount": "6,7"
        ],
        "PRICEDISC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PRICEDISC",
            "argumentCount": "4,5"
        ],
        "PRICEMAT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PRICEMAT",
            "argumentCount": "5,6"
        ],
        "PROB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "3,4"
        ],
        "PRODUCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::PRODUCT",
            "argumentCount": "1+"
        ],
        "PROPER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::PROPERCASE",
            "argumentCount": "1"
        ],
        "PV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::PV",
            "argumentCount": "3-5"
        ],
        "QUARTILE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::QUARTILE",
            "argumentCount": "2"
        ],
        "QUOTIENT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::QUOTIENT",
            "argumentCount": "2"
        ],
        "RADIANS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "deg2rad",
            "argumentCount": "1"
        ],
        "RAND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::RAND",
            "argumentCount": "0"
        ],
        "RANDBETWEEN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::RAND",
            "argumentCount": "2"
        ],
        "RANK": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::RANK",
            "argumentCount": "2,3"
        ],
        "RATE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::RATE",
            "argumentCount": "3-6"
        ],
        "RECEIVED": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::RECEIVED",
            "argumentCount": "4-5"
        ],
        "REPLACE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::REPLACE",
            "argumentCount": "4"
        ],
        "REPLACEB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::REPLACE",
            "argumentCount": "4"
        ],
        "REPT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "str_repeat",
            "argumentCount": "2"
        ],
        "RIGHT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::RIGHT",
            "argumentCount": "1,2"
        ],
        "RIGHTB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::RIGHT",
            "argumentCount": "1,2"
        ],
        "ROMAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::ROMAN",
            "argumentCount": "1,2"
        ],
        "ROUND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "round",
            "argumentCount": "2"
        ],
        "ROUNDDOWN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::ROUNDDOWN",
            "argumentCount": "2"
        ],
        "ROUNDUP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::ROUNDUP",
            "argumentCount": "2"
        ],
        "ROW": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::ROW",
            "argumentCount": "-1",
            "passByReference": [
                TRUE
            ]
        ],
        "ROWS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::ROWS",
            "argumentCount": "1"
        ],
        "RSQ": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::RSQ",
            "argumentCount": "2"
        ],
        "RTD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "1+"
        ],
        "SEARCH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::SEARCHINSENSITIVE",
            "argumentCount": "2,3"
        ],
        "SEARCHB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::SEARCHINSENSITIVE",
            "argumentCount": "2,3"
        ],
        "SECOND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::SECONDOFMINUTE",
            "argumentCount": "1"
        ],
        "SERIESSUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SERIESSUM",
            "argumentCount": "4"
        ],
        "SIGN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SIGN",
            "argumentCount": "1"
        ],
        "SIN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "sin",
            "argumentCount": "1"
        ],
        "SINH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "sinh",
            "argumentCount": "1"
        ],
        "SKEW": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::SKEW",
            "argumentCount": "1+"
        ],
        "SLN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::SLN",
            "argumentCount": "3"
        ],
        "SLOPE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::SLOPE",
            "argumentCount": "2"
        ],
        "SMALL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::SMALL",
            "argumentCount": "2"
        ],
        "SQRT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "sqrt",
            "argumentCount": "1"
        ],
        "SQRTPI": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SQRTPI",
            "argumentCount": "1"
        ],
        "STANDARDIZE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STANDARDIZE",
            "argumentCount": "3"
        ],
        "STDEV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STDEV",
            "argumentCount": "1+"
        ],
        "STDEVA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STDEVA",
            "argumentCount": "1+"
        ],
        "STDEVP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STDEVP",
            "argumentCount": "1+"
        ],
        "STDEVPA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STDEVPA",
            "argumentCount": "1+"
        ],
        "STEYX": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::STEYX",
            "argumentCount": "2"
        ],
        "SUBSTITUTE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::SUBSTITUTE",
            "argumentCount": "3,4"
        ],
        "SUBTOTAL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUBTOTAL",
            "argumentCount": "2+"
        ],
        "SUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUM",
            "argumentCount": "1+"
        ],
        "SUMIF": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMIF",
            "argumentCount": "2,3"
        ],
        "SUMIFS": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "?"
        ],
        "SUMPRODUCT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMPRODUCT",
            "argumentCount": "1+"
        ],
        "SUMSQ": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMSQ",
            "argumentCount": "1+"
        ],
        "SUMX2MY2": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMX2MY2",
            "argumentCount": "2"
        ],
        "SUMX2PY2": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMX2PY2",
            "argumentCount": "2"
        ],
        "SUMXMY2": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::SUMXMY2",
            "argumentCount": "2"
        ],
        "SYD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::SYD",
            "argumentCount": "4"
        ],
        "T": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::RETURNSTRING",
            "argumentCount": "1"
        ],
        "TAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "tan",
            "argumentCount": "1"
        ],
        "TANH": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "tanh",
            "argumentCount": "1"
        ],
        "TBILLEQ": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::TBILLEQ",
            "argumentCount": "3"
        ],
        "TBILLPRICE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::TBILLPRICE",
            "argumentCount": "3"
        ],
        "TBILLYIELD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::TBILLYIELD",
            "argumentCount": "3"
        ],
        "TDIST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::TDIST",
            "argumentCount": "3"
        ],
        "TEXT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::TEXTFORMAT",
            "argumentCount": "2"
        ],
        "TIME": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::TIME",
            "argumentCount": "3"
        ],
        "TIMEVALUE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::TIMEVALUE",
            "argumentCount": "1"
        ],
        "TINV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::TINV",
            "argumentCount": "2"
        ],
        "TODAY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DATENOW",
            "argumentCount": "0"
        ],
        "TRANSPOSE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::TRANSPOSE",
            "argumentCount": "1"
        ],
        "TREND": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::TREND",
            "argumentCount": "1-4"
        ],
        "TRIM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::TRIMSPACES",
            "argumentCount": "1"
        ],
        "TRIMMEAN": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::TRIMMEAN",
            "argumentCount": "2"
        ],
        "TRUE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOGICAL,
            "functionCall": "\\ZExcel\\Calculation\\Logical::TRUEE",
            "argumentCount": "0"
        ],
        "TRUNC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_MATH_AND_TRIG,
            "functionCall": "\\ZExcel\\Calculation\\MathTrig::TRUNC",
            "argumentCount": "1,2"
        ],
        "TTEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "4"
        ],
        "TYPE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::TYPE",
            "argumentCount": "1"
        ],
        "UPPER": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::UPPERCASE",
            "argumentCount": "1"
        ],
        "USDOLLAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "2"
        ],
        "VALUE": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_TEXT_AND_DATA,
            "functionCall": "\\ZExcel\\Calculation\\TextData::VALUE",
            "argumentCount": "1"
        ],
        "VAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::VARFunc",
            "argumentCount": "1+"
        ],
        "VARA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::VARA",
            "argumentCount": "1+"
        ],
        "VARP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::VARP",
            "argumentCount": "1+"
        ],
        "VARPA": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::VARPA",
            "argumentCount": "1+"
        ],
        "VDB": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "5-7"
        ],
        "VERSION": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_INFORMATION,
            "functionCall": "\\ZExcel\\Calculation\\Functions::VERSION",
            "argumentCount": "0"
        ],
        "VLOOKUP": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_LOOKUP_AND_REFERENCE,
            "functionCall": "\\ZExcel\\Calculation\\LookupRef::VLOOKUP",
            "argumentCount": "3,4"
        ],
        "WEEKDAY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::DAYOFWEEK",
            "argumentCount": "1,2"
        ],
        "WEEKNUM": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::WEEKOFYEAR",
            "argumentCount": "1,2"
        ],
        "WEIBULL": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::WEIBULL",
            "argumentCount": "4"
        ],
        "WORKDAY": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::WORKDAY",
            "argumentCount": "2+"
        ],
        "XIRR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::XIRR",
            "argumentCount": "2,3"
        ],
        "XNPV": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::XNPV",
            "argumentCount": "3"
        ],
        "YEAR": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::YEAR",
            "argumentCount": "1"
        ],
        "YEARFRAC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_DATE_AND_TIME,
            "functionCall": "\\ZExcel\\Calculation\\DateTime::YEARFRAC",
            "argumentCount": "2,3"
        ],
        "YIELD": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Functions::DUMMY",
            "argumentCount": "6,7"
        ],
        "YIELDDISC": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::YIELDDISC",
            "argumentCount": "4,5"
        ],
        "YIELDMAT": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_FINANCIAL,
            "functionCall": "\\ZExcel\\Calculation\\Financial::YIELDMAT",
            "argumentCount": "5,6"
        ],
        "ZTEST": [
            "category": \ZExcel\Calculation\Functionn::CATEGORY_STATISTICAL,
            "functionCall": "\\ZExcel\\Calculation\\Statistical::ZTEST",
            "argumentCount": "2-3"
        ]
    ];
    
    private static _controlFunctions = [
        "MKMATRIX": [
            "argumentCount": "*",
            "functionCall": "self::_mkMatrix"
        ]
    ];
    
    private static excelConstants = [
        "TRUE": true,
        "FALSE": false,
        "NULL": null
    ];
    
    //    Binary Operators
    //    These operators always work on two values
    //    Array key is the operator, the value indicates whether this is a left or right associative operator
    private static operatorAssociativity= [
        "^": 0,                                                            //    Exponentiation
        "*": 0, "/": 0,                                                 //    Multiplication and Division
        "+": 0, "-": 0,                                                    //    Addition and Subtraction
        "&": 0,                                                            //    Concatenation
        "|": 0, ":": 0,                                                    //    Intersect and Range
        ">": 0, "<": 0, "=": 0, ">=": 0, "<=": 0, "<>": 0        //    Comparison
    ];

    //    Comparison (Boolean) Operators
    //    These operators work on two values, but always return a boolean result
    private static comparisonOperators = [">": true, "<": true, "=": true, ">=": true, "<=": true, "<>": true];

    //    Operator Precedence
    //    This list includes all valid operators, whether binary (including boolean) or unary (such as %)
    //    Array key is the operator, the value is its precedence
    private static operatorPrecedence = [
        ":": 8,                                                                //    Range
        "|": 7,                                                                //    Intersect
        "~": 6,                                                                //    Negation
        "%": 5,                                                                //    Percentage
        "^": 4,                                                                //    Exponentiation
        "*": 3, "/": 3,                                                     //    Multiplication and Division
        "+": 2, "-": 2,                                                        //    Addition and Subtraction
        "&": 1,                                                                //    Concatenation
        ">": 0, "<": 0, "=": 0, ">=": 0, "<=": 0, "<>": 0            //    Comparison
    ];

    private static matrixReplaceFrom = ["{", ";", "}"];
    
    private static matrixReplaceTo = ["MKMATRIX(MKMATRIX(", "),MKMATRIX(", "))"];
    
    private function __construct(<\ZExcel\ZExcel> workbook = null)
    {
        int setPrecision = 16;
        
        if (PHP_INT_SIZE == 4) {
            let setPrecision = 14;
        }
        
        let this->savedPrecision = ini_get("precision");
        
        if (this->savedPrecision < setPrecision) {
            ini_set("precision", setPrecision);
        }
        
        let this->delta = 1 * pow(10, -setPrecision);
        
        if (workbook !== null) {
            let self::workbookSets[workbook->getID()] = this;
        }
        
        let this->workbook = workbook;
        let this->cyclicReferenceStack = new \ZExcel\CalcEngine\CyclicReferenceStack();
        let this->debugLog = new \ZExcel\CalcEngine\Logger(this->cyclicReferenceStack);
    }
    
    
    public function __destruct()
    {
        if (this->savedPrecision != ini_get("precision")) {
            ini_set("precision", this->savedPrecision);
        }
    }
    
    private static function _loadLocales() {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Get an instance of this class
     *
     * @access  public
     * @param   PHPExcel workbook  Injected workbook for working with a PHPExcel object,
     *                                    or null to create a standalone claculation engine
     * @return \ZExcel\Calculation
     */
    public static function getInstance(<\ZExcel\ZExcel> workbook = null) -> <\ZExcel\Calculation>
    {
        if (workbook !== null) {
            if (isset(self::workbookSets[workbook->getID()])) {
                return self::workbookSets[workbook->getID()];
            }
            return new \ZExcel\Calculation(workbook);
        }
        
        if (!isset(self::instance) || (self::instance === null)) {
            let self::instance = new \ZExcel\Calculation();
        }
        
        return self::instance;
    }
    
    /**
     * Unset an instance of this class
     *
     * @access    public
     * @param   PHPExcel workbook  Injected workbook identifying the instance to unset
     */
    public static function unsetInstance(<\ZExcel\ZExcel> workbook = null)
    {
        if (workbook !== null) {
            if (isset(self::workbookSets[workbook->getID()])) {
                unset(self::workbookSets[workbook->getID()]);
            }
        }
    }

    /**
     * Flush the calculation cache for any existing instance of this class
     *        but only if a \ZExcel\Calculation instance exists
     *
     * @access    public
     * @return null
     */
    public function flushInstance()
    {
        this->clearCalculationCache();
    }

    /**
     * Get the debuglog for this claculation engine instance
     *
     * @access    public
     * @return \ZExcel\CalcEngine\Logger
     */
    public function getDebugLog() -> <\ZExcel\CalcEngine\Logger>
    {
        return this->debugLog;
    }

    /**
     * __clone implementation. Cloning should not be allowed in a Singleton!
     *
     * @access    public
     * @throws    \ZExcel\Calculation\Exception
     */
    final public function __clone() -> <\ZExcel\Calculation\Exception>
    {
        throw new \ZExcel\Calculation\Exception("Cloning the calculation engine is not allowed!");
    }


    /**
     * Return the locale-specific translation of TRUE
     *
     * @access    public
     * @return     string        locale-specific translation of TRUE
     */
    public static function getTRUE() -> string
    {
        return self::localeBoolean["TRUE"];
    }

    /**
     * Return the locale-specific translation of FALSE
     *
     * @access    public
     * @return     string        locale-specific translation of FALSE
     */
    public static function getFALSE() -> string
    {
        return self::localeBoolean["FALSE"];
    }

    /**
     * Set the Array Return Type (Array or Value of first element in the array)
     *
     * @access    public
     * @param     string    returnType            Array return type
     * @return     boolean                    Success or failure
     */
    public static function setArrayReturnType(string returnType) -> boolean
    {
        if ((returnType == self::RETURN_ARRAY_AS_VALUE)
                || (returnType == self::RETURN_ARRAY_AS_ERROR)
                || (returnType == self::RETURN_ARRAY_AS_ARRAY)) {
            let self::returnArrayAsType = returnType;
            
            return true;
        }
        return false;
    }


    /**
     * Return the Array Return Type (Array or Value of first element in the array)
     *
     * @access    public
     * @return     string        returnType            Array return type
     */
    public static function getArrayReturnType() -> string
    {
        return self::returnArrayAsType;
    }


    /**
     * Is calculation caching enabled?
     *
     * @access    public
     * @return boolean
     */
    public function getCalculationCacheEnabled() -> boolean
    {
        return this->calculationCacheEnabled;
    }

    /**
     * Enable/disable calculation cache
     *
     * @access    public
     * @param boolean pValue
     */
    public function setCalculationCacheEnabled(boolean pValue = true)
    {
        let this->calculationCacheEnabled = pValue;
        this->clearCalculationCache();
    }


    /**
     * Enable calculation cache
     */
    public function enableCalculationCache()
    {
        this->setCalculationCacheEnabled(true);
    }


    /**
     * Disable calculation cache
     */
    public function disableCalculationCache()
    {
        this->setCalculationCacheEnabled(false);
    }


    /**
     * Clear calculation cache
     */
    public function clearCalculationCache()
    {
        let this->_calculationCache = [];
    }

    /**
     * Clear calculation cache for a specified worksheet
     *
     * @param string $worksheetName
     */
    public function clearCalculationCacheForWorksheet(var worksheetName)
    {
        if (isset(this->_calculationCache[worksheetName])) {
            unset(this->_calculationCache[worksheetName]);
        }
    }
    
    public function renameCalculationCacheForWorksheet(fromWorksheetName, toWorksheetName)
    {
        if (isset(this->_calculationCache[fromWorksheetName])) {
            let this->_calculationCache[toWorksheetName] = this->_calculationCache[fromWorksheetName];
            unset(this->_calculationCache[fromWorksheetName]);
        }
    }

    /**
     * Get the currently defined locale code
     *
     * @return string
     */
    public function getLocale()
    {
        return self::_localeLanguage;
    }

    /**
     * Set the locale code
     *
     * @param string $locale  The locale to use for formula translation
     * @return boolean
     */
    public function setLocale($locale = "en_us")
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public static function _translateSeparator(fromSeparator, toSeparator, formula, inBraces)
    {
        throw new \Exception("Not implemented yet!");
    }

    private static function _translateFormula(from, to, formula, fromSeparator, toSeparator)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function _translateFormulaToLocale(formula)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function _translateFormulaToEnglish(formula)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public static function _localeFunc(functionn)
    {
        var functionName, brace;
    
        if (self::_localeLanguage !== "en_us") {
            let functionName = trim(functionn, "(");
            if (isset(self::localeFunctions[functionName])) {
                let brace = (functionName != functionn);
                let functionn = self::localeFunctions[functionName];
                
                if (brace) {
                    let functionn = functionn . "(";
                }
            }
        }
        return functionn;
    }
    /**
     * Wrap string values in quotes
     *
     * @param mixed $value
     * @return mixed
     */
    public static function _wrapResult(value)
    {
        if (is_string(value)) {
            //    Error values cannot be "wrapped"
            if (preg_match("/^" . self::CALCULATION_REGEXP_ERROR . "/i", value)) {
                //    Return Excel errors "as is"
                return value;
            }
            //    Return strings wrapped in quotes
            return "\"" . value . "\"";
        //    Convert numeric errors to NaN error
        } elseif ((is_float(value)) && ((is_nan(value)) || (is_infinite(value)))) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return value;
    }
    /**
     * Remove quotes used as a wrapper to identify string values
     *
     * @param mixed $value
     * @return mixed
     */
    public static function _unwrapResult(value)
    {
        if (is_string(value)) {
            if ((isset(value[0])) && (value[0] == '"') && (substr(value,-1) == '"')) {
                return substr(value,1,-1);
            }
        //    Convert numeric errors to NaN error
        } elseif((is_float(value)) && ((is_nan(value)) || (is_infinite(value)))) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        return value;
    }
    /**
     * Calculate cell value (using formula from a cell ID)
     * Retained for backward compatibility
     *
     * @access    public
     * @param    PHPExcel_Cell    $pCell    Cell to calculate
     * @return    mixed
     * @throws    PHPExcel_Calculation_Exception
     */
    public function calculate(<\ZExcel\Cell> pCell = null)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function calculateCellValue(<\ZExcel\Cell> pCell = null, resetLog = null)
    {
        var returnArrayAsType, result, cellAddress, e, testResult, r, c;
        
        if (pCell === null) {
            return null;
        }

        let returnArrayAsType = self::returnArrayAsType;
        
        if (resetLog) {
            //    Initialise the logging settings if requested
            let this->formulaError = null;
            this->debugLog->clearLog();
            this->cyclicReferenceStack->clear();
            let this->cyclicFormulaCount = 1;

            let self::returnArrayAsType = self::RETURN_ARRAY_AS_ARRAY;
        }

        //    Execute the calculation for the cell formula
        let this->cellStack[] = [
            "sheet": pCell->getWorksheet()->getTitle(),
            "cell": pCell->getCoordinate()
        ];
        try {
            let result = self::_unwrapResult(this->_calculateFormulaValue(pCell->getValue(), pCell->getCoordinate(), pCell));
            let cellAddress = array_pop(this->cellStack);
            this->workbook->getSheetByName(cellAddress["sheet"])->getCell(cellAddress["cell"]);
        } catch \ZExcel\Exception, e {
            let cellAddress = array_pop(this->cellStack);
            this->workbook->getSheetByName(cellAddress["sheet"])->getCell(cellAddress["cell"]);
            throw new \ZExcel\Calculation\Exception(e->getMessage());
        }

        if ((is_array(result)) && (self::returnArrayAsType != self::RETURN_ARRAY_AS_ARRAY)) {
            let self::returnArrayAsType = returnArrayAsType;
            let testResult = \ZExcel\Calculation\Functions::flattenArray(result);
            
            if (self::returnArrayAsType == self::RETURN_ARRAY_AS_ERROR) {
                return \ZExcel\Calculation\Functions::value();
            }
            
            //    If there's only a single cell in the array, then we allow it
            if (count(testResult) != 1) {
                //    If keys are numeric, then it's a matrix result rather than a cell range result, so we permit it
                let r = array_keys(result);
                let r = array_shift(r);
                if (!is_numeric(r)) {
                    return \ZExcel\Calculation\Functions::value();
                }
                
                if (is_array(result[r])) {
                    let c = array_keys(result[r]);
                    let c = array_shift(c);
                    if (!is_numeric(c)) {
                        return \ZExcel\Calculation\Functions::value();
                    }
                }
            }
            
            let result = array_shift(testResult);
        }
        let self::returnArrayAsType = returnArrayAsType;

        if (result === null) {
            return 0;
        } elseif((is_float(result)) && ((is_nan(result)) || (is_infinite(result)))) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        return result;
    }
    /**
     * Validate and parse a formula string
     *
     * @param    string        $formula        Formula to parse
     * @return    array
     * @throws    PHPExcel_Calculation_Exception
     */
    public function parseFormula(formula)
    {
        throw new \Exception("Not implemented yet!");
    }
    /**
     * Calculate the value of a formula
     *
     * @param    string            $formula    Formula to parse
     * @param    string            $cellID        Address of the cell to calculate
     * @param    PHPExcel_Cell    $pCell        Cell to calculate
     * @return    mixed
     * @throws    PHPExcel_Calculation_Exception
     */
    public function calculateFormula(formula, cellID = null, <\ZExcel\Cell> pCell = null)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function getValueFromCache(cellReference, cellValue)
    {
        var returnValue;
        
        // Is calculation cacheing enabled?
        // Is the value present in calculation cache?
        this->debugLog->writeDebugLog("Testing cache value for cell ", cellReference);
        if ((this->calculationCacheEnabled) && (isset(this->_calculationCache[cellReference]))) {
            this->debugLog->writeDebugLog("Retrieving value for cell ", cellReference, " from cache");
            // Return the cached result
            return (returnValue === false) ? true : this->_calculationCache[cellReference];
        }
        
        return (returnValue === false) ? false : null;
    }
    
    public function saveValueToCache(cellReference, cellValue)
    {
        if (this->calculationCacheEnabled) {
            let this->_calculationCache[cellReference] = cellValue;
        }
    }
    
    public function _calculateFormulaValue(string formula, var cellID = null, <\ZExcel\Cell> pCell = null)
    {
        var cellValue = null, pCellParent, wsTitle, wsCellReference;

        //    Basic validation that this is indeed a formula
        //    We simply return the cell value if not
        let formula = trim(formula);
        
        if (substr(formula, 0, 1) != "=") {
            return self::_wrapResult(formula);
        }
        
        let formula = ltrim(substr(formula, 1));
        
        if (strlen(formula) === 0) {
            return self::_wrapResult(formula);
        }

        let pCellParent = (pCell !== null) ? pCell->getWorksheet() : null;
        let wsTitle = (pCellParent !== null) ? pCellParent->getTitle() : "\x00Wrk";
        let wsCellReference = wsTitle . "!" . cellID;

        if ((cellID !== null) && (this->getValueFromCache(wsCellReference, cellValue))) {
            return cellValue;
        }

        if ((substr(wsTitle, 0, 1) !== "\x00") && (this->cyclicReferenceStack->onStack(wsCellReference))) {
            if (this->cyclicFormulaCount <= 0) {
                let this->cyclicFormulaCell = "";
                return this->_raiseFormulaError("Cyclic Reference in Formula");
            } elseif (this->cyclicFormulaCell === wsCellReference) {
                let this->cyclicFormulaCounter = this->cyclicFormulaCounter + 1;
                if (this->cyclicFormulaCounter >= this->cyclicFormulaCount) {
                    let this->cyclicFormulaCell = "";
                    return cellValue;
                }
            } elseif (this->cyclicFormulaCell == "") {
                if (this->cyclicFormulaCounter >= this->cyclicFormulaCount) {
                    return cellValue;
                }
                let this->cyclicFormulaCell = wsCellReference;
            }
        }

        //    Parse the formula onto the token stack and calculate the value
        this->cyclicReferenceStack->push(wsCellReference);
        let cellValue = this->processTokenStack(this->_parseFormula(formula, pCell), cellID, pCell);
        this->cyclicReferenceStack->pop();

        // Save to calculation cache
        if (cellID !== null) {
            this->saveValueToCache(wsCellReference, cellValue);
        }

        //    Return the calculated value
        return cellValue;
    }
    
    /**
     * Ensure that paired matrix operands are both matrices and of the same size
     *
     * @param    mixed        &$operand1    First matrix operand
     * @param    mixed        &$operand2    Second matrix operand
     * @param    integer        $resize        Flag indicating whether the matrices should be resized to match
     *                                        and (if so), whether the smaller dimension should grow or the
     *                                        larger should shrink.
     *                                            0 = no resize
     *                                            1 = shrink to fit
     *                                            2 = extend to fit
     */
    private static function _checkMatrixOperands(operand1, operand2, resize = 1)
    {
        throw new \Exception("Not implemented yet!");
    }
    /**
     * Read the dimensions of a matrix, and re-index it with straight numeric keys starting from row 0, column 0
     *
     * @param    mixed        &$matrix        matrix operand
     * @return    array        An array comprising the number of rows, and number of columns
     */
    public static function _getMatrixDimensions(matrix)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Ensure that paired matrix operands are both matrices of the same size
     *
     * @param    mixed        &$matrix1        First matrix operand
     * @param    mixed        &$matrix2        Second matrix operand
     * @param    integer        $matrix1Rows    Row size of first matrix operand
     * @param    integer        $matrix1Columns    Column size of first matrix operand
     * @param    integer        $matrix2Rows    Row size of second matrix operand
     * @param    integer        $matrix2Columns    Column size of second matrix operand
     */
    private static function _resizeMatricesShrink(matrix1, matrix2, matrix1Rows, matrix1Columns, matrix2Rows, matrix2Columns)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Ensure that paired matrix operands are both matrices of the same size
     *
     * @param    mixed        &$matrix1    First matrix operand
     * @param    mixed        &$matrix2    Second matrix operand
     * @param    integer        $matrix1Rows    Row size of first matrix operand
     * @param    integer        $matrix1Columns    Column size of first matrix operand
     * @param    integer        $matrix2Rows    Row size of second matrix operand
     * @param    integer        $matrix2Columns    Column size of second matrix operand
     */
    private static function _resizeMatricesExtend(matrix1, matrix2, matrix1Rows, matrix1Columns, matrix2Rows, matrix2Columns)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     * Format details of an operand for display in the log (based on operand type)
     *
     * @param    mixed        $value    First matrix operand
     * @return    mixed
     */
    private function _showValue(value)
    {
        var row, testArray, returnMatrix, pad, rpad;
        
        if (this->debugLog->getWriteDebugLog()) {
            let testArray = \ZExcel\Calculation\Functions::flattenArray(value);
            if (count(testArray) == 1) {
                let value = array_pop(testArray);
            }

            if (is_array(value)) {
                let returnMatrix = [];
                let pad = ", ";
                let rpad = ", ";
                for row in value {
                    if (is_array(row)) {
                        let returnMatrix[] = implode(pad,array_map([this, "showValue"],row));
                        let rpad = "; ";
                    } else {
                        let returnMatrix[] = this->_showValue(row);
                    }
                }
                return "{ " . implode(rpad,returnMatrix) . " }";
            } elseif(is_string(value) && (trim(value,"\"") == value)) {
                return "\"" . value . "\"";
            } elseif(is_bool(value)) {
                return (value) ? self::localeBoolean["TRUE"] : self::localeBoolean["FALSE"];
            }
        }
        
        return \ZExcel\Calculation\Functions::flattenSingleValue(value);
    }
    
    private function _showTypeDetails(var value) -> string
    {
        var testArray, typeString;
        string returnValue = null;
        
        if (this->debugLog->getWriteDebugLog()) {
            let testArray = \ZExcel\Calculation\Functions::flattenArray(value);
            if (count(testArray) == 1) {
                let value = array_pop(testArray);
            }

            if (value === null) {
                return "a NULL value";
            } elseif (is_float(value)) {
                let typeString = "a floating point number";
            } elseif(is_int(value)) {
                let typeString = "an integer number";
            } elseif(is_bool(value)) {
                let typeString = "a boolean";
            } elseif(is_array(value)) {
                let typeString = "a matrix";
            } else {
                if (value == "") {
                    return "an empty string";
                } elseif (value[0] == "#") {
                    return "a " . value . " error";
                } else {
                    let typeString = "a string";
                }
            }
            
            let returnValue = typeString . " with a value of " . this->_showValue(value);
        }
        
        return returnValue;
    }

    private function _convertMatrixReferences(var formula)
    {
        var temp, i, key, value, openCount, closeCount;
        
        //    Convert any Excel matrix references to the MKMATRIX() function
        if (strpos(formula, "{") !== false) {
            //    If there is the possibility of braces within a quoted string, then we don"t treat those as matrix indicators
            if (strpos(formula, "\"") !== false) {
                //    So instead we skip replacing in any quoted strings by only replacing in every other array element after we"ve exploded
                //        the formula
                let temp = explode("\"", formula);
                //    Open and Closed counts used for trapping mismatched braces in the formula
                let openCount = 0;
                let closeCount = 0;
                let i = false;
                for key, value in temp {
                    //    Only count/replace in alternating array entries
                    let i = !i;
                    if i {
                        let openCount = openCount + substr_count(value, "{");
                        let closeCount = closeCount + substr_count(value, "}");
                        let temp[key] = str_replace(self::matrixReplaceFrom, self::matrixReplaceTo, value);
                    }
                }

                //    Then rebuild the formula string
                let formula = implode("\"", temp);
            } else {
                //    If there"s no quoted strings, then we do a simple count/replace
                let openCount = substr_count(formula, "{");
                let closeCount = substr_count(formula, "}");
                let formula = str_replace(self::matrixReplaceFrom, self::matrixReplaceTo, formula);
            }
            //    Trap for mismatched braces and trigger an appropriate error
            if (openCount < closeCount) {
                if (openCount > 0) {
                    return this->_raiseFormulaError("Formula Error: Mismatched matrix braces '}'");
                } else {
                    return this->_raiseFormulaError("Formula Error: Unexpected '}' encountered");
                }
            } elseif (openCount > closeCount) {
                if (closeCount > 0) {
                    return this->_raiseFormulaError("Formula Error: Mismatched matrix braces '{'");
                } else {
                    return this->_raiseFormulaError("Formula Error: Unexpected '{' encountered");
                }
            }
        }

        return formula;
    }
    
    private static function _mkMatrix()
    {
        return func_get_args();
    }

    // Convert infix to postfix notation
    private function _parseFormula(var formula, <\ZExcel\Cell> pCell = null)
    {
        throw new \Exception("Not implemented yet!");
        /*
        let formula = this->convertMatrixReferences(trim(formula));
        
        if (formula === false) {
            return false;
        }

        //    If we're using cell caching, then pCell may well be flushed back to the cache (which detaches the parent worksheet),
        //        so we store the parent worksheet so that we can re-attach it when necessary
        let pCellParent = (pCell !== null) ? pCell->getWorksheet() : null;

        let regexpMatchString = "/^(" . self::CALCULATION_REGEXP_FUNCTION .
           "|" . self::CALCULATION_REGEXP_CELLREF .
           "|" . self::CALCULATION_REGEXP_NUMBER .
           "|" . self::CALCULATION_REGEXP_STRING .
           "|" . self::CALCULATION_REGEXP_OPENBRACE .
           "|" . self::CALCULATION_REGEXP_NAMEDRANGE .
           "|" . self::CALCULATION_REGEXP_ERROR . ")/si";

        //    Start with initialisation
        let index = 0;
        let stack = new \ZExcel\Calculation\Token\Stack;
        let output = [];
        let expectingOperator = false;                    //    We use this test in syntax-checking the expression to determine when a
                                                    //        - is a negation or + is a positive operator rather than an operation
        let expectingOperand = false;                    //    We use this test in syntax-checking the expression to determine whether an operand
                                                    //        should be null in a function call
        //    The guts of the lexical parser
        //    Loop through the formula extracting each operator and operand in turn
        while (true) {
            let opCharacter = substr(formula, index, 1);    //    Get the first character of the value at the current index position
            if ((isset(self::comparisonOperators[opCharacter])) && (strlen(formula) > index) && (isset(self::comparisonOperators[formula{index+1}]))) {
                let index = index + 1;
                let opCharacter = opCharacter . substr(formula, index, 1);
            }

            //    Find out if we"re currently at the beginning of a number, variable, cell reference, function, parenthesis or operand
            let isOperandOrFunction = preg_match(regexpMatchString, substr(formula, index), match);

            if (opCharacter == "-" && !expectingOperator) {                //    Is it a negation instead of a minus?
                stack->push("Unary Operator", "~");                            //    Put a negation on the stack
                let index = index + 1;                                                    //        and drop the negation symbol
            } elseif (opCharacter == "%" && expectingOperator) {
                stack->push("Unary Operator", "%");                            //    Put a percentage on the stack
                let index = index + 1;
            } elseif (opCharacter == "+" && !expectingOperator) {            //    Positive (unary plus rather than binary operator plus) can be discarded?
                let index = index + 1;                                                    //    Drop the redundant plus symbol
            } elseif (((opCharacter == "~") || (opCharacter == "|")) && (!isOperandOrFunction)) {    //    We have to explicitly deny a tilde or pipe, because they are legal
                return this->_raiseFormulaError("Formula Error: Illegal character "~"");                //        on the stack but not in the input expression

            } elseif ((isset(self::operators[opCharacter]) or isOperandOrFunction) && expectingOperator) {    //    Are we putting an operator on the stack?
                while (stack->count() > 0 &&
                    (o2 = stack->last()) &&
                    isset(self::operators[o2["value"]]) &&
                    @(self::operatorAssociativity[opCharacter] ? self::operatorPrecedence[opCharacter] < self::operatorPrecedence[o2["value"]] : self::operatorPrecedence[opCharacter] <= self::operatorPrecedence[o2["value"]])) {
                    let output[] = stack->pop();                                //    Swap operands and higher precedence operators from the stack to the output
                }
                stack->push("Binary Operator", opCharacter);    //    Finally put our current operator onto the stack
                let index = index + 1;
                let expectingOperator = false;

            } elseif (opCharacter == ")" && expectingOperator) {            //    Are we expecting to close a parenthesis?
                let expectingOperand = false;
                while ((o2 = stack->pop()) && o2["value"] != "(") {        //    Pop off the stack back to the last (
                    if (o2 === null) {
                        return this->_raiseFormulaError("Formula Error: Unexpected closing brace ")"");
                    } else {
                        let output[] = o2;
                    }
                }
                let d = stack->last(2);
                if (preg_match("/^".self::CALCULATION_REGEXP_FUNCTION."/i", d["value"], matches)) {    //    Did this parenthesis just close a function?
                    let functionName = matches[1]; // Get the function name

                    let d = stack->pop();
                    let argumentCount = d["value"];        //    See how many arguments there were (argument count is the next value stored on the stack)

                    let output[] = d;                        //    Dump the argument count on the output
                    let output[] = stack->pop();            //    Pop the function and push onto the output
                    if (isset(self::controlFunctions[functionName])) {
                        let expectedArgumentCount = self::controlFunctions[functionName]["argumentCount"];
                        let functionCall = self::controlFunctions[functionName]["functionCall"];
                    } elseif (isset(self::PHPExcelFunctions[functionName])) {
                        let expectedArgumentCount = self::PHPExcelFunctions[functionName]["argumentCount"];
                        let functionCall = self::PHPExcelFunctions[functionName]["functionCall"];
                    } else {    // did we somehow push a non-function on the stack? this should never happen
                        return this->_raiseFormulaError("Formula Error: Internal error, non-function on stack");
                    }
                    //    Check the argument count
                    let argumentCountError = false;
                    if (is_numeric(expectedArgumentCount)) {
                        if (expectedArgumentCount < 0) {
                            if (argumentCount > abs(expectedArgumentCount)) {
                                let argumentCountError = true;
                                let expectedArgumentCountString = "no more than ".abs(expectedArgumentCount);
                            }
                        } else {
                            if (argumentCount != expectedArgumentCount) {
                                let argumentCountError = true;
                                let expectedArgumentCountString = expectedArgumentCount;
                            }
                        }
                    } elseif (expectedArgumentCount != "*") {
                        let isOperandOrFunction = preg_match("/(\d*)([-+,])(\d*)/", expectedArgumentCount, argMatch);
                        switch (argMatch[2]) {
                            case "+":
                                if (argumentCount < argMatch[1]) {
                                    let argumentCountError = true;
                                    let expectedArgumentCountString = argMatch[1]." or more ";
                                }
                                break;
                            case "-":
                                if ((argumentCount < argMatch[1]) || (argumentCount > argMatch[3])) {
                                    let argumentCountError = true;
                                    let expectedArgumentCountString = "between ".argMatch[1]." and ".argMatch[3];
                                }
                                break;
                            case ",":
                                if ((argumentCount != argMatch[1]) && (argumentCount != argMatch[3])) {
                                    let argumentCountError = true;
                                    let expectedArgumentCountString = "either ".argMatch[1]." or ".argMatch[3];
                                }
                                break;
                        }
                    }
                    if (argumentCountError) {
                        return this->_raiseFormulaError("Formula Error: Wrong number of arguments for " . functionName . "() function: " . argumentCount . " given, ".expectedArgumentCountString." expected");
                    }
                }
                let index = index + 1;

            } elseif (opCharacter == ",") {            //    Is this the separator for function arguments?
                while ((o2 = stack->pop()) && o2["value"] != "(") {        //    Pop off the stack back to the last (
                    if (o2 === null) {
                        return this->_raiseFormulaError("Formula Error: Unexpected ,");
                    } else {
                        let output[] = o2;    // pop the argument expression stuff and push onto the output
                    }
                }
                //    If we"ve a comma when we"re expecting an operand, then what we actually have is a null operand;
                //        so push a null onto the stack
                if ((expectingOperand) || (!expectingOperator)) {
                    let output[] = ["type": "NULL Value", "value": self::excelConstants["NULL"], "reference": null];
                }
                // make sure there was a function
                let d = stack->last(2);
                if (!preg_match("/^".self::CALCULATION_REGEXP_FUNCTION."/i", d["value"], matches)) {
                    return this->_raiseFormulaError("Formula Error: Unexpected ,");
                }
                let d = stack->pop();
                stack->push(d["type"], ++d["value"], d["reference"]);    // increment the argument count
                stack->push("Brace", "(");    // put the ( back on, we"ll need to pop back to it again
                let expectingOperator = false;
                let expectingOperand = true;
                let index = index + 1;

            } elseif (opCharacter == "(" && !expectingOperator) {
                stack->push("Brace", "(");
                let index = index + 1;

            } elseif (isOperandOrFunction && !expectingOperator) {    // do we now have a function/variable/number?
                let expectingOperator = true;
                let expectingOperand = false;
                let val = match[1];
                let length = strlen(val);

                if (preg_match("/^".self::CALCULATION_REGEXP_FUNCTION."/i", val, matches)) {
                    let val = preg_replace("/\s/u", "", val);
                    if (isset(self::PHPExcelFunctions[strtoupper(matches[1])]) || isset(self::controlFunctions[strtoupper(matches[1])])) {    // it"s a function
                        stack->push("Function", strtoupper(val));
                        let ax = preg_match("/^\s*(\s*\))/ui", substr(formula, index+length), amatch);
                        if (ax) {
                            stack->push("Operand Count for Function ".strtoupper(val).")", 0);
                            let expectingOperator = true;
                        } else {
                            stack->push("Operand Count for Function ".strtoupper(val).")", 1);
                            let expectingOperator = false;
                        }
                        stack->push("Brace", "(");
                    } else {    // it"s a var w/ implicit multiplication
                        let output[] = ["type": "Value", "value": matches[1], "reference": null];
                    }
                } elseif (preg_match("/^".self::CALCULATION_REGEXP_CELLREF."/i", val, matches)) {
                    // Watch for this case-change when modifying to allow cell references in different worksheets...
                    // Should only be applied to the actual cell column, not the worksheet name

                    // If the last entry on the stack was a : operator, then we have a cell range reference
                    let testPrevOp = stack->last(1);
                    if (testPrevOp["value"] == ":") {
                        // If we have a worksheet reference, then we"re playing with a 3D reference
                        if (matches[2] == "") {
                            // Otherwise, we "inherit" the worksheet reference from the start cell reference
                            // The start of the cell range reference should be the last entry in output
                            let startCellRef = output[count(output)-1]["value"];
                            preg_match("/^".self::CALCULATION_REGEXP_CELLREF."/i", startCellRef, startMatches);
                            if (startMatches[2] > "") {
                                let val = startMatches[2] . "!" . val;
                            }
                        } else {
                            return this->_raiseFormulaError("3D Range references are not yet supported");
                        }
                    }

                    let output[] = ["type": "Cell Reference", "value": val, "reference": val];
                } else {    // it"s a variable, constant, string, number or boolean
                    //    If the last entry on the stack was a : operator, then we may have a row or column range reference
                    let testPrevOp = stack->last(1);
                    if (testPrevOp["value"] == ":") {
                        let startRowColRef = output[count(output)-1]["value"];
                        let rangeWS1 = "";
                        
                        if (strpos("!", startRowColRef) !== false) {
                            list(rangeWS1, startRowColRef) = explode("!", startRowColRef);
                        }
                        
                        if (rangeWS1 != "") {
                            let rangeWS1 = rangeWS1 . "!";
                        }
                        
                        let rangeWS2 = rangeWS1;
                        
                        if (strpos("!", val) !== false) {
                            list(rangeWS2, val) = explode("!", val);
                        }
                        
                        if (rangeWS2 != "") {
                            let rangeWS2 = rangeWS2 . "!";
                        }
                        
                        if ((is_integer(startRowColRef)) && (ctype_digit(val)) && (startRowColRef <= 1048576) && (val <= 1048576)) {
                            //    Row range
                            let endRowColRef = (pCellParent !== null) ? pCellParent->getHighestColumn() : "XFD";    //    Max 16,384 columns for Excel2007
                            let output[count(output)-1]["value"] = rangeWS1."A".startRowColRef;
                            let val = rangeWS2.endRowColRef.val;
                        } elseif ((ctype_alpha(startRowColRef)) && (ctype_alpha(val)) && (strlen(startRowColRef) <= 3) && (strlen(val) <= 3)) {
                            //    Column range
                            let endRowColRef = (pCellParent !== null) ? pCellParent->getHighestRow() : 1048576;        //    Max 1,048,576 rows for Excel2007
                            let output[count(output)-1]["value"] = rangeWS1.strtoupper(startRowColRef)."1";
                            let val = rangeWS2 . val . endRowColRef;
                        }
                    }

                    let localeConstant = false;
                    
                    if (opCharacter == "\"") {
                        //    UnEscape any quotes within the string
                        let val = self::_wrapResult(str_replace("\"\"", "\"", self::_unwrapResult(val)));
                    } elseif (is_numeric(val)) {
                        if ((strpos(val, ".") !== false) || (stripos(val, "e") !== false) || (val > PHP_INT_MAX) || (val < -PHP_INT_MAX)) {
                            let val = (float) val;
                        } else {
                            let val = (integer) val;
                        }
                    } elseif (isset(self::excelConstants[trim(strtoupper(val))])) {
                        let excelConstant = trim(strtoupper(val));
                        let val = self::excelConstants[excelConstant];
                    } elseif ((localeConstant = array_search(trim(strtoupper(val)), self::localeBoolean)) !== false) {
                        let val = self::excelConstants[localeConstant];
                    }
                    
                    let details = ["type": "Value", "value": val, "reference": null];
                    
                    if (localeConstant) {
                        let details["localeValue"] = localeConstant;
                    }
                    
                    let output[] = details;
                }
                let index = index + length;

            } elseif (opCharacter == "") {    // absolute row or column range
                let index = index + 1;
            } elseif (opCharacter == ")") {    // miscellaneous error checking
                if (expectingOperand) {
                    let output[] = ["type": "NULL Value", "value": self::excelConstants["NULL"], "reference": null];
                    let expectingOperand = false;
                    let expectingOperator = true;
                } else {
                    return this->_raiseFormulaError("Formula Error: Unexpected ")"");
                }
            } elseif (isset(self::operators[opCharacter]) && !expectingOperator) {
                return this->_raiseFormulaError("Formula Error: Unexpected operator "opCharacter"");
            } else {    // I don"t even want to know what you did to get here
                return this->_raiseFormulaError("Formula Error: An unexpected error occured");
            }
            
            //    Test for end of formula string
            if (index == strlen(formula)) {
                //    Did we end with an operator?.
                //    Only valid for the % unary operator
                if ((isset(self::operators[opCharacter])) && (opCharacter != "%")) {
                    return this->_raiseFormulaError("Formula Error: Operator "opCharacter" has no operands");
                } else {
                    break;
                }
            }
            
            //    Ignore white space
            while ((formula{index} == "\n") || (formula{index} == "\r")) {
                let index = index + 1;
            }
            
            if (formula{index} == " ") {
                while (formula{index} == " ") {
                    let index = index + 1;
                }
                // If we're expecting an operator, but only have a space between the previous and next operands (and both are Cell References) then we have an INTERSECTION operator
                if ((expectingOperator) && (preg_match("/^" . self::CALCULATION_REGEXP_CELLREF . "./Ui", substr(formula, index), match)) && (output[count(output)-1]["type"] == "Cell Reference")) {
                    while (stack->count() > 0 && (o2 = stack->last()) && isset(self::operators[o2["value"]]) &&
                            @(self::operatorAssociativity[opCharacter] ? self::operatorPrecedence[opCharacter] < self::operatorPrecedence[o2["value"]] : self::operatorPrecedence[opCharacter] <= self::operatorPrecedence[o2["value"]])) {
                        let output[] = stack->pop();                                //    Swap operands and higher precedence operators from the stack to the output
                    }
                    stack->push("Binary Operator", "|");    //    Put an Intersect Operator on the stack
                    let expectingOperator = false;
                }
            }
        }

        while ((op = stack->pop()) !== null) {    // pop everything off the stack and push onto output
            if ((is_array(op) && op["value"] == "(") || (op === "(")) {
                return this->_raiseFormulaError("Formula Error: Expecting ')'");    // if there are any opening braces on the stack, then braces were unbalanced
            }
            let output[] = op;
        }
        
        return output;
        */
    }
    
    private static function dataTestReference(operandData)
    {
        var operand, rKeys, rowKey, cKeys, colKey;
        
        let operand = operandData["value"];
        
        if ((operandData["reference"] === NULL) && (is_array(operand))) {
            let rKeys = array_keys(operand);
            let rowKey = array_shift(rKeys);
            let cKeys = array_keys(array_keys(operand[rowKey]));
            let colKey = array_shift(cKeys);
            
            if (ctype_upper(colKey)) {
                let operandData["reference"] = colKey . rowKey;
            }
        }
        
        return [operand, operandData];
    }
    
    private function processTokenStack(tokens, cellID = null, <\ZExcel\Cell> pCell = null)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    private function _validateBinaryOperand(cellID, operand, stack)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    private function _executeBinaryComparisonOperation(cellID, operand1, operand2, operation, stack, recursingArrays = false)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    private function strcmpLowercaseFirst(str1, str2)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    private function _executeNumericBinaryOperation(cellID, operand1, operand2, operation, matrixFunction, stack)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    protected function _raiseFormulaError(errorMessage)
    {
        let this->formulaError = errorMessage;
        
        this->cyclicReferenceStack->clear();
        
        if (!this->suppressFormulaErrors) {
            throw new \ZExcel\Calculation\Exception(errorMessage);
        }
        
        trigger_error(errorMessage, E_USER_ERROR);
    }
    
    public function extractNamedRange(pRange = "A1", <\ZExcel\Worksheet> pSheet = NULL, resetLog = true)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function extractCellRange(pRange = "A1", <\ZExcel\Worksheet> pSheet = null, resetLog = true)
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function isImplemented(pFunction = "")
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function listFunctions()
    {
        throw new \Exception("Not implemented yet!");
    }
    
    public function listFunctionNames()
    {
        throw new \Exception("Not implemented yet!");
    }
}
