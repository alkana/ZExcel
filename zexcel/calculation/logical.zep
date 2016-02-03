namespace ZExcel\Calculation;

class Logical
{
    /**
     * TRUE
     *
     * Returns the boolean TRUE.
     *
     * Excel Function:
     *        =TRUE()
     *
     * @access    public
     * @category Logical Functions
     * @return    boolean        True
     */
    public static function truee() -> boolean
    {
        return true;
    }


    /**
     * FALSE
     *
     * Returns the boolean FALSE.
     *
     * Excel Function:
     *        =FALSE()
     *
     * @access    public
     * @category Logical Functions
     * @return    boolean        False
     */
    public static function falsee() -> boolean
    {
        return false;
    }


    /**
     * LOGICAL_AND
     *
     * Returns boolean TRUE if all its arguments are TRUE; returns FALSE if one or more argument is FALSE.
     *
     * Excel Function:
     *        =AND(logical1[,logical2[, ...]])
     *
     *        The arguments must evaluate to logical values such as TRUE or FALSE, or the arguments must be arrays
     *            or references that contain logical values.
     *
     *        Boolean arguments are treated as True or False as appropriate
     *        Integer or floating point arguments are treated as True, except for 0 or 0.0 which are False
     *        If any argument value is a string, or a Null, the function returns a #VALUE! error, unless the string holds
     *            the value TRUE or FALSE, in which case it is evaluated as the corresponding boolean value
     *
     * @access    public
     * @category Logical Functions
     * @param    mixed        $arg,...        Data values
     * @return    boolean        The logical AND of the arguments.
     */
    public static function logical_and() -> boolean
    {
        var returnValue, aArgs, arg, argCount;
        
        // Return value
        let returnValue = true;

        // Loop through the arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        let argCount = -1;
        for argCount, arg in aArgs {
            // Is it a boolean value?
            if (is_bool(arg)) {
                let returnValue = returnValue && arg;
            } elseif ((is_numeric(arg)) && (!is_string(arg))) {
                let returnValue = returnValue && (arg != 0);
            } elseif (is_string(arg)) {
                let arg = strtoupper(arg);
                if ((arg == "TRUE") || (arg == \ZExcel\Calculation::getTRUE())) {
                    let arg = true;
                } elseif ((arg == "FALSE") || (arg == \ZExcel\Calculation::getFALSE())) {
                    let arg = false;
                } else {
                    return \ZExcel\Calculation\Functions::value();
                }
                
                let returnValue = returnValue && (arg != 0);
            }
        }

        // Return
        if (argCount < 0) {
            return \ZExcel\Calculation\Functions::value();
        }
        
        return returnValue;
    }


    /**
     * LOGICAL_OR
     *
     * Returns boolean TRUE if any argument is TRUE; returns FALSE if all arguments are FALSE.
     *
     * Excel Function:
     *        =OR(logical1[,logical2[, ...]])
     *
     *        The arguments must evaluate to logical values such as TRUE or FALSE, or the arguments must be arrays
     *            or references that contain logical values.
     *
     *        Boolean arguments are treated as True or False as appropriate
     *        Integer or floating point arguments are treated as True, except for 0 or 0.0 which are False
     *        If any argument value is a string, or a Null, the function returns a #VALUE! error, unless the string holds
     *            the value TRUE or FALSE, in which case it is evaluated as the corresponding boolean value
     *
     * @access    public
     * @category Logical Functions
     * @param    mixed        $arg,...        Data values
     * @return    boolean        The logical OR of the arguments.
     */
    public static function logical_or() -> boolean
    {
        var returnValue, aArgs, arg, argCount;
        
        // Return value
        let returnValue = false;

        // Loop through the arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        let argCount = -1;
        for argCount, arg in aArgs {
            // Is it a boolean value?
            if (is_bool(arg)) {
                let returnValue = returnValue || arg;
            } elseif ((is_numeric(arg)) && (!is_string(arg))) {
                let returnValue = returnValue || (arg != 0);
            } elseif (is_string(arg)) {
                let arg = strtoupper(arg);
                if ((arg == "TRUE") || (arg == \ZExcel\Calculation::getTRUE())) {
                    let arg = true;
                } elseif ((arg == "FALSE") || (arg == \ZExcel\Calculation::getFALSE())) {
                    let arg = false;
                } else {
                    return \ZExcel\Calculation\Functions::value();
                }
                let returnValue = returnValue || (arg != 0);
            }
        }

        // Return
        if (argCount < 0) {
            return \ZExcel\Calculation\Functions::value();
        }
        
        return returnValue;
    }


    /**
     * NOT
     *
     * Returns the boolean inverse of the argument.
     *
     * Excel Function:
     *        =NOT(logical)
     *
     *        The argument must evaluate to a logical value such as TRUE or FALSE
     *
     *        Boolean arguments are treated as True or False as appropriate
     *        Integer or floating point arguments are treated as True, except for 0 or 0.0 which are False
     *        If any argument value is a string, or a Null, the function returns a #VALUE! error, unless the string holds
     *            the value TRUE or FALSE, in which case it is evaluated as the corresponding boolean value
     *
     * @access    public
     * @category Logical Functions
     * @param    mixed        $logical    A value or expression that can be evaluated to TRUE or FALSE
     * @return    boolean        The boolean inverse of the argument.
     */
    public static function not(logical = false)
    {
        let logical = \ZExcel\Calculation\Functions::flattenSingleValue(logical);
        if (is_string(logical)) {
            let logical = strtoupper(logical);
            if ((logical == "TRUE") || (logical == \ZExcel\Calculation::getTRUE())) {
                return false;
            } elseif ((logical == "FALSE") || (logical == \ZExcel\Calculation::getFALSE())) {
                return true;
            } else {
                return \ZExcel\Calculation\Functions::value();
            }
        }

        return !logical;
    }

    /**
     * STATEMENT_IF
     *
     * Returns one value if a condition you specify evaluates to TRUE and another value if it evaluates to FALSE.
     *
     * Excel Function:
     *        =IF(condition[,returnIfTrue[,returnIfFalse]])
     *
     *        Condition is any value or expression that can be evaluated to TRUE or FALSE.
     *            For example, A10=100 is a logical expression; if the value in cell A10 is equal to 100,
     *            the expression evaluates to TRUE. Otherwise, the expression evaluates to FALSE.
     *            This argument can use any comparison calculation operator.
     *        ReturnIfTrue is the value that is returned if condition evaluates to TRUE.
     *            For example, if this argument is the text string "Within budget" and the condition argument evaluates to TRUE,
     *            then the IF function returns the text "Within budget"
     *            If condition is TRUE and ReturnIfTrue is blank, this argument returns 0 (zero). To display the word TRUE, use
     *            the logical value TRUE for this argument.
     *            ReturnIfTrue can be another formula.
     *        ReturnIfFalse is the value that is returned if condition evaluates to FALSE.
     *            For example, if this argument is the text string "Over budget" and the condition argument evaluates to FALSE,
     *            then the IF function returns the text "Over budget".
     *            If condition is FALSE and ReturnIfFalse is omitted, then the logical value FALSE is returned.
     *            If condition is FALSE and ReturnIfFalse is blank, then the value 0 (zero) is returned.
     *            ReturnIfFalse can be another formula.
     *
     * @access    public
     * @category Logical Functions
     * @param    mixed    $condition        Condition to evaluate
     * @param    mixed    $returnIfTrue    Value to return when condition is true
     * @param    mixed    $returnIfFalse    Optional value to return when condition is false
     * @return    mixed    The value of returnIfTrue or returnIfFalse determined by condition
     */
    public static function statement_if(condition = true, returnIfTrue = 0, returnIfFalse = false)
    {
        let condition     = (is_null(condition))     ? true :  (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(condition);
        let returnIfTrue  = (is_null(returnIfTrue))  ? 0 :     \ZExcel\Calculation\Functions::flattenSingleValue(returnIfTrue);
        let returnIfFalse = (is_null(returnIfFalse)) ? false : \ZExcel\Calculation\Functions::flattenSingleValue(returnIfFalse);

        return (condition) ? returnIfTrue : returnIfFalse;
    }


    /**
     * IFERROR
     *
     * Excel Function:
     *        =IFERROR(testValue,errorpart)
     *
     * @access    public
     * @category Logical Functions
     * @param    mixed    $testValue    Value to check, is also the value returned when no error
     * @param    mixed    $errorpart    Value to return when testValue is an error condition
     * @return    mixed    The value of errorpart or testValue determined by error condition
     */
    public static function iferror(testValue = "", errorpart = "")
    {
        let testValue = (is_null(testValue)) ? "" : \ZExcel\Calculation\Functions::flattenSingleValue(testValue);
        let errorpart = (is_null(errorpart)) ? "" : \ZExcel\Calculation\Functions::flattenSingleValue(errorpart);

        return self::statement_if(\ZExcel\Calculation\Functions::is_error(testValue), errorpart, testValue);
    }
}
