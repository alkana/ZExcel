namespace ZExcel;

class DocumentProperties
{
    /** constants */
    const PROPERTY_TYPE_BOOLEAN    = "b";
    const PROPERTY_TYPE_INTEGER = "i";
    const PROPERTY_TYPE_FLOAT   = "f";
    const PROPERTY_TYPE_DATE    = "d";
    const PROPERTY_TYPE_STRING  = "s";
    const PROPERTY_TYPE_UNKNOWN = "u";

    /**
    * Creator
    *
    * @var string
    */
    private _creator = "Unknown Creator";

    /**
    * LastModifiedBy
    *
    * @var string
    */
    private _lastModifiedBy;

    /**
    * Created
    *
    * @var datetime
    */
    private _created;

    /**
    * Modified
    *
    * @var datetime
    */
    private _modified;

    /**
    * Title
    *
    * @var string
    */
    private _title = "Untitled Spreadsheet";

    /**
    * Description
    *
    * @var string
    */
    private _description = "";

    /**
    * Subject
    *
    * @var string
    */
    private _subject = "";

    /**
    * Keywords
    *
    * @var string
    */
    private _keywords = "";

    /**
    * Category
    *
    * @var string
    */
    private _category = "";

    /**
    * Manager
    *
    * @var string
    */
    private _manager = "";

    /**
    * Company
    *
    * @var string
    */
    private _company = "Microsoft Corporation";

    /**
    * Custom Properties
    *
    * @var string
    */
    private _customProperties = [];
    
    /**
     * Create a new PHPExcel_DocumentProperties
     */
    public function __construct()
    {
        // Initialise values
        let this->_lastModifiedBy = this->_creator;
        let this->_created = time();
        let this->_modified    = time();
    }

    /**
     * Get Creator
     *
     * @return string
     */
    public function getCreator() {
        return this->_creator;
    }

    /**
     * Set Creator
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setCreator(string pValue = "") {
        let this->_creator = pValue;
        return this;
    }

    /**
     * Get Last Modified By
     *
     * @return string
     */
    public function getLastModifiedBy() {
        return this->_lastModifiedBy;
    }

    /**
     * Set Last Modified By
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setLastModifiedBy(string pValue = "") {
        let this->_lastModifiedBy = pValue;
        return this;
    }

    /**
     * Get Created
     *
     * @return datetime
     */
    public function getCreated() {
        return this->_created;
    }

    /**
     * Set Created
     *
     * @param datetime pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setCreated(<\DateTime> pValue = null) {
        if (pValue == null) {
            let pValue = time();
        } else {
            if (is_string(pValue)) {
                if (is_numeric(pValue)) {
                    let pValue = intval(pValue);
                } else {
                    let pValue = strtotime(pValue);
                }
            }
       }

        let this->_created = pValue;
        return this;
    }

    /**
     * Get Modified
     *
     * @return datetime
     */
    public function getModified() {
        return this->_modified;
    }

    /**
     * Set Modified
     *
     * @param datetime pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setModified(pValue = null) {
        if (pValue == null) {
            let pValue = time();
        } else {
            if (is_string(pValue)) {
                if (is_numeric(pValue)) {
                    let pValue = intval(pValue);
                } else {
                    let pValue = strtotime(pValue);
                }
            }
        }
        
        let this->_modified = pValue;
        
        return this;
    }

    /**
     * Get Title
     *
     * @return string
     */
    public function getTitle() {
        return this->_title;
    }

    /**
     * Set Title
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setTitle(string pValue = "") {
        let this->_title = pValue;
        return this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription() {
        return this->_description;
    }

    /**
     * Set Description
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setDescription(string pValue = "") {
        let this->_description = pValue;
        return this;
    }

    /**
     * Get Subject
     *
     * @return string
     */
    public function getSubject() {
        return this->_subject;
    }

    /**
     * Set Subject
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setSubject(string pValue = "") {
        let this->_subject = pValue;
        return this;
    }

    /**
     * Get Keywords
     *
     * @return string
     */
    public function getKeywords() {
        return this->_keywords;
    }

    /**
     * Set Keywords
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setKeywords(string pValue = "") {
        let this->_keywords = pValue;
        return this;
    }

    /**
     * Get Category
     *
     * @return string
     */
    public function getCategory() {
        return this->_category;
    }

    /**
     * Set Category
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setCategory(string pValue = "") {
        let this->_category = pValue;
        return this;
    }

    /**
     * Get Company
     *
     * @return string
     */
    public function getCompany() {
        return this->_company;
    }

    /**
     * Set Company
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setCompany(string pValue = "") {
        let this->_company = pValue;
        return this;
    }

    /**
     * Get Manager
     *
     * @return string
     */
    public function getManager() {
        return this->_manager;
    }

    /**
     * Set Manager
     *
     * @param string pValue
     * @return PHPExcel_DocumentProperties
     */
    public function setManager(pValue = "") {
        let this->_manager = pValue;
        return this;
    }

    /**
     * Get a List of Custom Property Names
     *
     * @return array of string
     */
    public function getCustomProperties() {
        return array_keys(this->_customProperties);
    }

    /**
     * Check if a Custom Property is defined
     *
     * @param string propertyName
     * @return boolean
     */
    public function isCustomPropertySet(propertyName) {
        return isset(this->_customProperties[propertyName]);
    }

    /**
     * Get a Custom Property Value
     *
     * @param string propertyName
     * @return string
     */
    public function getCustomPropertyValue(propertyName) {
        if (isset(this->_customProperties[propertyName])) {
            return this->_customProperties[propertyName]["value"];
        }
    }

    /**
     * Get a Custom Property Type
     *
     * @param string propertyName
     * @return string
     */
    public function getCustomPropertyType(propertyName) {
        if (isset(this->_customProperties[propertyName])) {
            return this->_customProperties[propertyName]["type"];
        }
    }

    /**
     * Set a Custom Property
     *
     * @param string propertyName
     * @param mixed propertyValue
     * @param string propertyType
     *      "i"    : Integer
     *   "f" : Floating Point
     *   "s" : String
     *   "d" : Date/Time
     *   "b" : Boolean
     * @return PHPExcel_DocumentProperties
     */
    public function setCustomProperty(string propertyName,string propertyValue = "", var propertyType = null) {
        
        if ((propertyType === null) || (!in_array(propertyType, [self::PROPERTY_TYPE_INTEGER, self::PROPERTY_TYPE_FLOAT, self::PROPERTY_TYPE_STRING, self::PROPERTY_TYPE_DATE, self::PROPERTY_TYPE_BOOLEAN]))) {

            if (propertyValue === null) {
                let propertyType = self::PROPERTY_TYPE_STRING;
            } else {
                if (is_float(propertyValue)) {
                    let propertyType = self::PROPERTY_TYPE_FLOAT;
                } else {
                    if(is_int(propertyValue)) {
                        let propertyType = self::PROPERTY_TYPE_INTEGER;
                    } else {
                        if (is_bool(propertyValue)) {
                            let propertyType = self::PROPERTY_TYPE_BOOLEAN;
                        } else {
                            let propertyType = self::PROPERTY_TYPE_STRING;
                        }
                    }
                }
            }
        }

        let this->_customProperties[propertyName] = [
            "value": propertyValue,
            "type" : propertyType
        ];
        
        return this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        var vars, value;
        
        let vars = get_object_vars(this);
        
        for value in vars {
            if (is_object(value)) {
                let this->key = clone value;
            } else {
                let this->key = value;
            }
        }
    }

    public static function convertProperty(propertyValue,propertyType) {
        var value;
        
        switch (propertyType) {
            case "empty"    :    //    Empty
                let value = "";
                break;
            case "null"        :    //    Null
                let value = null;
                break;
            case "i1"        :    //    1-Byte Signed Integer
            case "i2"        :    //    2-Byte Signed Integer
            case "i4"        :    //    4-Byte Signed Integer
            case "i8"        :    //    8-Byte Signed Integer
            case "int"        :    //    Integer
                let value = (int) propertyValue;
                break;
            case "ui1"        :    //    1-Byte Unsigned Integer
            case "ui2"        :    //    2-Byte Unsigned Integer
            case "ui4"        :    //    4-Byte Unsigned Integer
            case "ui8"        :    //    8-Byte Unsigned Integer
            case "uint"        :    //    Unsigned Integer
                let value = abs((int) propertyValue);
                break;
            case "r4"        :    //    4-Byte Real Number
            case "r8"        :    //    8-Byte Real Number
            case "decimal"    :    //    Decimal
                let value = (float) propertyValue;
                break;
            case "lpstr"    :    //    LPSTR
            case "lpwstr"    :    //    LPWSTR
            case "bstr"        :    //    Basic String
                let value = propertyValue;
                break;
            case "date"        :    //    Date and Time
            case "filetime"    :    //    File Time
                let value = strtotime(propertyValue);
                break;
            case "bool"        :    //    Boolean
                let value = (propertyValue == "true") ? True : False;
                break;
            default:
                let value = propertyValue;
                break;
        }
        
        return value;
    }

    public static function convertPropertyType(propertyType) {
        var value;
        
        switch (propertyType) {
            case "i1"        :    //    1-Byte Signed Integer
            case "i2"        :    //    2-Byte Signed Integer
            case "i4"        :    //    4-Byte Signed Integer
            case "i8"        :    //    8-Byte Signed Integer
            case "int"        :    //    Integer
            case "ui1"        :    //    1-Byte Unsigned Integer
            case "ui2"        :    //    2-Byte Unsigned Integer
            case "ui4"        :    //    4-Byte Unsigned Integer
            case "ui8"        :    //    8-Byte Unsigned Integer
            case "uint"        :    //    Unsigned Integer
                let value = self::PROPERTY_TYPE_INTEGER;
                break;
            case "r4"        :    //    4-Byte Real Number
            case "r8"        :    //    8-Byte Real Number
            case "decimal"    :    //    Decimal
                let value = self::PROPERTY_TYPE_FLOAT;
                break;
            case "empty"    :    //    Empty
            case "null"        :    //    Null
            case "lpstr"    :    //    LPSTR
            case "lpwstr"    :    //    LPWSTR
            case "bstr"        :    //    Basic String
                let value = self::PROPERTY_TYPE_STRING;
                break;
            case "date"        :    //    Date and Time
            case "filetime"    :    //    File Time
                let value = self::PROPERTY_TYPE_DATE;
                break;
            case "bool"        :    //    Boolean
                let value = self::PROPERTY_TYPE_BOOLEAN;
                break;
            default:
                let value = self::PROPERTY_TYPE_UNKNOWN;
                break;
        }
        return value;
    }
    
}
