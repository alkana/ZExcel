namespace ZExcel\Cell;

class DefaultValueBinder implements IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  \ZExcel\Cell  $cell   Cell to bind value to
     * @param  mixed          $value  Value to bind in cell
     * @return boolean
     */
    public function bindValue(<\ZExcel\Cell> cell, value = null)
    {
        // sanitize UTF-8 strings
        if (is_string(value)) {
            let value = \ZExcel\Shared\Stringg::SanitizeUTF8(value);
        } elseif (is_object(value)) {
            // Handle any objects that might be injected
            if (value instanceof \DateTime) {
                let value = value->format("Y-m-d H:i:s");
            } elseif (!(value instanceof \ZExcel\RichText)) {
                let value = (string) value;
            }
        }

        // Set value explicit
        cell->setValueExplicit(value, self::dataTypeForValue(value));

        // Done!
        return true;
    }

    /**
     * DataType for value
     *
     * @param   mixed  $pValue
     * @return  string
     */
    public static function dataTypeForValue(pValue = null)
    {
        var tValue = [];
        
        // Match the value against a few data types
        if (pValue === null) {
            return \ZExcel\Cell\DataType::TYPE_NULL;
        } elseif (pValue === "") {
            return \ZExcel\Cell\DataType::TYPE_STRING;
        } elseif (pValue instanceof \ZExcel\RichText) {
            return \ZExcel\Cell\DataType::TYPE_INLINE;
        } elseif (pValue[0] === "=" && strlen(pValue) > 1) {
            return \ZExcel\Cell\DataType::TYPE_FORMULA;
        } elseif (is_bool(pValue)) {
            return \ZExcel\Cell\DataType::TYPE_BOOL;
        } elseif (is_float(pValue) || is_int(pValue)) {
            return \ZExcel\Cell\DataType::TYPE_NUMERIC;
        } elseif (preg_match("/^[\+\-]?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)([Ee][\-\+]?[0-2]?\d{1,3})?$/", pValue)) {
            let tValue = ltrim(pValue, "+-");
            if (is_string(pValue) && tValue[0] === "0" && strlen(tValue) > 1 && tValue[1] !== ".") {
                return \ZExcel\Cell\DataType::TYPE_STRING;
            } elseif ((strpos(pValue, ".") === false) && (pValue > PHP_INT_MAX)) {
                return \ZExcel\Cell\DataType::TYPE_STRING;
            }
            return \ZExcel\Cell\DataType::TYPE_NUMERIC;
        } elseif (is_string(pValue) && array_key_exists(pValue, \ZExcel\Cell\DataType::getErrorCodes())) {
            return \ZExcel\Cell\DataType::TYPE_ERROR;
        }

        return \ZExcel\Cell\DataType::TYPE_STRING;
    }
}
