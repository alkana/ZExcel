namespace ZExcel\Cell;

class AdvancedValueBinder extends DefaultValueBinder implements IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  PHPExcel_Cell  cell  Cell to bind value to
     * @param  mixed $value          Value to bind in cell
     * @return boolean
     */
    public function bindValue(<\ZExcel\Cell> cell, var value = null)
    {
        var dataType, tmp, d, h, m, s, days, formatCode,
            currencyCode, decimalSeparator, thousandsSeparator;
        array matches = [];
        
        // sanitize UTF-8 strings
        if (is_string(value)) {
            let value = \ZExcel\Shared\Stringg::SanitizeUTF8(value);
        }

        // Find out data type
        let dataType = parent::dataTypeForValue(value);

        // Style logic - strings
        if (dataType === \ZExcel\Cell\DataType::TYPE_STRING && (!is_object(value) || !(value instanceof \ZExcel\RichText))) {
            //    Test for booleans using locale-setting
            if (value == \ZExcel\Calculation::getTRUE()) {
                cell->setValueExplicit(true, \ZExcel\Cell\DataType::TYPE_BOOL);
                
                return true;
            } elseif (value == \ZExcel\Calculation::getFALSE()) {
                cell->setValueExplicit(false, \ZExcel\Cell\DataType::TYPE_BOOL);
                
                return true;
            }

            // Check for number in scientific format
            if (preg_match("/^" . \ZExcel\Calculation::CALCULATION_REGEXP_NUMBER . "$/", value)) {
                cell->setValueExplicit((float) value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                return true;
            }

            // Check for fraction
            if (preg_match("/^([+-]?)\s*([0-9]+)\s?\/\s*([0-9]+)$/", value, matches)) {
                // Convert value to number
                let value = matches[2] / matches[3];
                
                if (matches[1] == "-") {
                    let value = 0 - value;
                }
                
                cell->setValueExplicit((float) value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode("??/??");
                
                return true;
            } elseif (preg_match("/^([+-]?)([0-9]*) +([0-9]*)\s?\/\s*([0-9]*)$/", value, matches)) {
                // Convert value to number
                let value = matches[2] + (matches[3] / matches[4]);
                
                if (matches[1] == "-") {
                    let value = 0 - value;
                }
                
                cell->setValueExplicit((float) value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode("# ??/??");
                
                return true;
            }

            // Check for percentage
            if (preg_match("/^\-?[0-9]*\.?[0-9]*\s?\%$/", value)) {
                // Convert value to number
                let value = (float) str_replace("%", "", value) / 100;
                
                cell->setValueExplicit(value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode(\ZExcel\Style\NumberFormat::FORMAT_PERCENTAGE_00);
                    
                return true;
            }

            // Check for currency
            let currencyCode = \ZExcel\Shared\Stringg::getCurrencyCode();
            let decimalSeparator = \ZExcel\Shared\Stringg::getDecimalSeparator();
            let thousandsSeparator = \ZExcel\Shared\Stringg::getThousandsSeparator();
            
            if (preg_match("/^" . preg_quote(currencyCode) . " *(\d{1,3}(" . preg_quote(thousandsSeparator) . "\d{3})*|(\d+))(" . preg_quote(decimalSeparator) . "\d{2})?$/", value)) {
                // Convert value to number
                let value = (float) trim(str_replace(
                    [
                        currencyCode,
                        thousandsSeparator,
                        decimalSeparator
                    ],
                    [
                        "",
                        "",
                        "."
                    ],
                    value
                ));
                
                cell->setValueExplicit(value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode(str_replace("", currencyCode, \ZExcel\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE));
                
                return true;
            } elseif (preg_match("/^\$ *(\d{1,3}(\,\d{3})*|(\d+))(\.\d{2})?$/", value)) {
                // Convert value to number
                let value = (float) trim(str_replace(["",","], "", value));
                
                cell->setValueExplicit(value, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode(\ZExcel\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
                
                return true;
            }

            // Check for time without seconds e.g. "9:45", "09:45"
            if (preg_match("/^(\d|[0-1]\d|2[0-3]):[0-5]\d$/", value)) {
                // Convert value to number
                let tmp = explode(":", value);
                let h = tmp[0];
                let m = tmp[1];
                let days = h / 24 + m / 1440;
                
                cell->setValueExplicit(days, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()->getStyle(cell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\ZExcel\Style\NumberFormat::FORMAT_DATE_TIME3);
                
                return true;
            }

            // Check for time with seconds "9:45:59", "09:45:59"
            if (preg_match("/^(\d|[0-1]\d|2[0-3]):[0-5]\d:[0-5]\d$/", value)) {
                // Convert value to number
                let tmp = explode(":", value);
                let h = tmp[0];
                let m = tmp[1];
                let s = tmp[2];
                let days = h / 24 + m / 1440 + s / 86400;
                
                // Convert value to number
                cell->setValueExplicit(days, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Set style
                cell->getWorksheet()->getStyle(cell->getCoordinate())
                    ->getNumberFormat()->setFormatCode(\ZExcel\Style\NumberFormat::FORMAT_DATE_TIME4);
                
                return true;
            }

            // Check for datetime, e.g. "2008-12-31", "2008-12-31 15:59", "2008-12-31 15:59:10"
            let d = \ZExcel\Shared\Date::stringToExcel(value);
            if (d !== false) {
                // Convert value to number
                cell->setValueExplicit(d, \ZExcel\Cell\DataType::TYPE_NUMERIC);
                // Determine style. Either there is a time part or not. Look for ":"
                
                if (strpos(value, ":") !== false) {
                    let formatCode = "yyyy-mm-dd h:mm";
                } else {
                    let formatCode = "yyyy-mm-dd";
                }
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getNumberFormat()
                    ->setFormatCode(formatCode);
                
                return true;
            }

            // Check for newline character "\n"
            if (strpos(value, "\n") !== false) {
                let value = \ZExcel\Shared\Stringg::SanitizeUTF8(value);
                cell->setValueExplicit(value, \ZExcel\Cell\DataType::TYPE_STRING);
                // Set style
                cell->getWorksheet()
                    ->getStyle(cell->getCoordinate())
                    ->getAlignment()
                    ->setWrapText(true);
                
                return true;
            }
        }

        // Not bound yet? Use parent...
        return parent::bindValue(cell, value);
    }
}
