namespace ZExcel\Calculation;

class Database
{
    /**
     * fieldExtract
     *
     * Extracts the column ID to use for the data field.
     *
     * @access    private
     * @param    mixed[]        database        The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    mixed        field            Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @return    string|NULL
     *
     */
    private static function fieldExtract(var database, var field)
    {
        var fieldNames, keys, key;
        
        let field = strtoupper(\ZExcel\Calculation\Functions::flattenSingleValue(field));
        let fieldNames = array_map("strtoupper", array_shift(database));

        if (is_numeric(field)) {
            let keys = array_keys(fieldNames);
            return keys[field - 1];
        }
        
        let key = array_search(field, fieldNames);
        
        return (key) ? key : null;
    }

    /**
     * filter
     *
     * Parses the selection criteria, extracts the database rows that match those criteria, and
     * returns that subset of rows.
     *
     * @access    private
     * @param    mixed[]        database        The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    mixed[]        criteria        The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    array of mixed
     *
     */
    private static function filter(var database, var criteria)
    {
        var fieldNames, criteriaNames, testConditions, testConditionsCount, testCondition,
            testConditionCount, testConditionSet, row, criterion, key, criteriaName, k, testConditionList,
            dataRow, dataValues, dataValue, result;
        
        let fieldNames = array_shift(database);
        let criteriaNames = array_shift(criteria);

        //    Convert the criteria into a set of AND/OR conditions with [:placeholders]
        let testConditions = [];
        let testConditionsCount = 0;
        
        for key, criteriaName in criteriaNames {
            let testCondition = [];
            let testConditionCount = 0;
            
            for row, criterion in criteria {
                if (strlen(criterion[key]) > 0) {
                    let testCondition[] = "[:" . criteriaName . "]" . \ZExcel\Calculation\Functions::ifCondition(criterion[key]);
                    let testConditionCount = testConditionCount + 1;
                }
            }
            
            if (testConditionCount > 1) {
                let testConditions[] = "OR(" . implode(",", testCondition) . ")";
                let testConditionCount = testConditionCount + 1;
            } elseif (testConditionCount == 1) {
                let testConditions[] = testCondition[0];
                let testConditionCount = testConditionCount + 1;
            }
        }

        if (testConditionsCount > 1) {
            let testConditionSet = "AND(" . implode(",", testConditions) . ")";
        } elseif (testConditionsCount == 1) {
            let testConditionSet = testConditions[0];
        }

        //    Loop through each row of the database
        for dataRow, dataValues in database {
            //    Substitute actual values from the database row for our [:placeholders]
            let testConditionList = testConditionSet;
            
            for key, criteriaName in criteriaNames {
                let k = array_search(criteriaName, fieldNames);
                
                if (isset(dataValues[k])) {
                    let dataValue = dataValues[k];
                    let dataValue = (is_string(dataValue)) ? \ZExcel\Calculation::wrapResult(strtoupper(dataValue)) : dataValue;
                    let testConditionList = str_replace("[:" . criteriaName . "]", dataValue, testConditionList);
                }
            }
            //    evaluate the criteria against the row data
            let result = \ZExcel\Calculation::getInstance()->_calculateFormulaValue("=" . testConditionList);
            //    If the row failed to meet the criteria, remove it from the database
            if (!result) {
                unset(database[dataRow]);
            }
        }

        return database;
    }


    private static function getFilteredColumn(var database, var field, var criteria)
    {
        var row;
        //    extract an array of values for the requested column
        array colData = [];
        //    reduce the database to a set of rows that match all the criteria
        let database = self::filter(database, criteria);
        
        for row in database {
            let colData[] = row[field];
        }
        
        return colData;
    }

    /**
     * DAVERAGE
     *
     * Averages the values in a column of a list or database that match conditions you specify.
     *
     * Excel Function:
     *        DAVERAGE(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dAverage(var database, var field, var criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }

        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::AVERAGE", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DCOUNT
     *
     * Counts the cells that contain numbers in a column of a list or database that match conditions
     * that you specify.
     *
     * Excel Function:
     *        DCOUNT(database,[field],criteria)
     *
     * Excel Function:
     *        DAVERAGE(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    integer
     *
     * @TODO    The field argument is optional. If field is omitted, DCOUNT counts all records in the
     *            database that match the criteria.
     *
     */
    public static function dCount(var database, var field, var criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }

        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::COUNT", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DCOUNTA
     *
     * Counts the nonblank cells in a column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DCOUNTA(database,[field],criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    integer
     *
     * @TODO    The field argument is optional. If field is omitted, DCOUNTA counts all records in the
     *            database that match the criteria.
     *
     */
    public static function dCountA(database, field, criteria)
    {
        var row;
        array colData = [];
        
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }

        //    reduce the database to a set of rows that match all the criteria
        let database = self::filter(database, criteria);
        
        for row in database {
            let colData[] = row[field];
        }

        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::COUNTA", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DGET
     *
     * Extracts a single value from a column of a list or database that matches conditions that you
     * specify.
     *
     * Excel Function:
     *        DGET(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    mixed
     *
     */
    public static function dGet(database, field, criteria)
    {
        var colData;
        
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        let colData = self::getFilteredColumn(database, field, criteria);
        
        if (count(colData) > 1) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        return colData[0];
    }


    /**
     * DMAX
     *
     * Returns the largest number in a column of a list or database that matches conditions you that
     * specify.
     *
     * Excel Function:
     *        DMAX(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dMax(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::MAX", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DMIN
     *
     * Returns the smallest number in a column of a list or database that matches conditions you that
     * specify.
     *
     * Excel Function:
     *        DMIN(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dMin(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::MIN", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DPRODUCT
     *
     * Multiplies the values in a column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DPRODUCT(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function pProduct(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\MathTrig::PRODUCT", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DSTDEV
     *
     * Estimates the standard deviation of a population based on a sample by using the numbers in a
     * column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DSTDEV(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dStdEv(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::STDEV", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DSTDEVP
     *
     * Calculates the standard deviation of a population based on the entire population by using the
     * numbers in a column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DSTDEVP(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dStdEvp(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::STDEVP", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DSUM
     *
     * Adds the numbers in a column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DSUM(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dSum(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\MathTrig::SUM", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DVAR
     *
     * Estimates the variance of a population based on a sample by using the numbers in a column
     * of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DVAR(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dVar(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::VARFunc", self::getFilteredColumn(database, field, criteria));
    }


    /**
     * DVARP
     *
     * Calculates the variance of a population based on the entire population by using the numbers
     * in a column of a list or database that match conditions that you specify.
     *
     * Excel Function:
     *        DVARP(database,field,criteria)
     *
     * @access    public
     * @category Database Functions
     * @param    mixed[]            database    The range of cells that makes up the list or database.
     *                                        A database is a list of related data in which rows of related
     *                                        information are records, and columns of data are fields. The
     *                                        first row of the list contains labels for each column.
     * @param    string|integer    field        Indicates which column is used in the function. Enter the
     *                                        column label enclosed between double quotation marks, such as
     *                                        "Age" or "Yield," or a number (without quotation marks) that
     *                                        represents the position of the column within the list: 1 for
     *                                        the first column, 2 for the second column, and so on.
     * @param    mixed[]            criteria    The range of cells that contains the conditions you specify.
     *                                        You can use any range for the criteria argument, as long as it
     *                                        includes at least one column label and at least one cell below
     *                                        the column label in which you specify a condition for the
     *                                        column.
     * @return    float
     *
     */
    public static function dVarp(database, field, criteria)
    {
        let field = self::fieldExtract(database, field);
        
        if (is_null(field)) {
            return null;
        }
        
        // Return
        return call_user_func("\ZExcel\Calculation\Statistical::VARP", self::getFilteredColumn(database, field, criteria));
    }
}
