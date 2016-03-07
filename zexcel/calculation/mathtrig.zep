namespace ZExcel\Calculation;

class MathTrig
{
    /**
     * Private method to return an array of the factors of the input value
     * @return array
     */
    private static function factors(var value)
    {
        array factorArray = [];
        int i, startVal;
        
        let startVal = 0 + floor(sqrt(value));
        
        for i in reverse range(2, startVal) {
            if ((value % i) == 0) {
                let factorArray = array_merge(factorArray, self::factors(value / i));
                let factorArray = array_merge(factorArray, self::factors(i));
                
                if (i <= sqrt(value)) {
                    break;
                }
            }
        }
        
        if (!empty(factorArray)) {
            rsort(factorArray);
            return factorArray;
        } else {
            return [intval(value)];
        }
    }

    private static function romanCut(num, n)
    {
        return (num - (num % n)) / n;
    }

    /**
     * ATAN2
     *
     * This function calculates the arc tangent of the two variables x and y. It is similar to
     *        calculating the arc tangent of y ÷ x, except that the signs of both arguments are used
     *        to determine the quadrant of the result.
     * The arctangent is the angle from the x-axis to a line containing the origin (0, 0) and a
     *        point with coordinates (xCoordinate, yCoordinate). The angle is given in radians between
     *        -pi and pi, excluding -pi.
     *
     * Note that the Excel ATAN2() function accepts its arguments in the reverse order to the standard
     *        PHP atan2() function, so we need to reverse them here before calling the PHP atan() function.
     *
     * Excel Function:
     *        ATAN2(xCoordinate,yCoordinate)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    xCoordinate        The x-coordinate of the point.
     * @param    float    yCoordinate        The y-coordinate of the point.
     * @return    float    The inverse tangent of the specified x- and y-coordinates.
     */
    public static function aTan2(var xCoordinate = null, var yCoordinate = null)
    {
        let xCoordinate = \ZExcel\Calculation\Functions::flattenSingleValue(xCoordinate);
        let yCoordinate = \ZExcel\Calculation\Functions::flattenSingleValue(yCoordinate);

        let xCoordinate = (xCoordinate !== null) ? xCoordinate : 0.0;
        let yCoordinate = (yCoordinate !== null) ? yCoordinate : 0.0;

        if (((is_numeric(xCoordinate)) || (is_bool(xCoordinate))) &&
            ((is_numeric(yCoordinate)))  || (is_bool(yCoordinate))) {
            let xCoordinate    = (float) xCoordinate;
            let yCoordinate    = (float) yCoordinate;

            if ((xCoordinate == 0) && (yCoordinate == 0)) {
                return \ZExcel\Calculation\Functions::DiV0();
            }

            return atan2(yCoordinate, xCoordinate);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * CEILING
     *
     * Returns number rounded up, away from zero, to the nearest multiple of significance.
     *        For example, if you want to avoid using pennies in your prices and your product is
     *        priced at 4.42, use the formula =CEILING(4.42,0.05) to round prices up to the
     *        nearest nickel.
     *
     * Excel Function:
     *        CEILING(number[,significance])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    number            The number you want to round.
     * @param    float    significance    The multiple to which you want to round.
     * @return   float    Rounded Number
     */
    public static function ceiling(float number, var significance = null)
    {
        let number       = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let significance = \ZExcel\Calculation\Functions::flattenSingleValue(significance);

        if ((is_null(significance)) &&
            (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC)) {
            let significance = number / abs(number);
        }

        if ((is_numeric(number)) && (is_numeric(significance))) {
            if ((number == 0.0 ) || (significance == 0.0)) {
                return 0.0;
            } elseif (self::SiGN(number) == self::SiGN(significance)) {
                return ceil(number / significance) * significance;
            } else {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * COMBIN
     *
     * Returns the number of combinations for a given number of items. Use COMBIN to
     *        determine the total possible number of groups for a given number of items.
     *
     * Excel Function:
     *        COMBIN(numObjs,numInSet)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    int        numObjs    Number of different objects
     * @param    int        numInSet    Number of objects in each combination
     * @return   int        Number of combinations
     */
    public static function combin(var numObjs, var numInSet)
    {
        let numObjs  = \ZExcel\Calculation\Functions::flattenSingleValue(numObjs);
        let numInSet = \ZExcel\Calculation\Functions::flattenSingleValue(numInSet);

        if ((is_numeric(numObjs)) && (is_numeric(numInSet))) {
            if (numObjs < numInSet) {
                return \ZExcel\Calculation\Functions::NaN();
            } elseif (numInSet < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return round(self::FaCT(numObjs) / self::FaCT(numObjs - numInSet)) / self::FaCT(numInSet);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * EVEN
     *
     * Returns number rounded up to the nearest even integer.
     * You can use this function for processing items that come in twos. For example,
     *        a packing crate accepts rows of one or two items. The crate is full when
     *        the number of items, rounded up to the nearest two, matches the crate's
     *        capacity.
     *
     * Excel Function:
     *        EVEN(number)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    number            Number to round
     * @return   int      Rounded Number
     */
    public static function even(var number)
    {
        var significance;
        
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);

        if (is_null(number)) {
            return 0;
        } elseif (is_bool(number)) {
            let number = (int) number;
        }

        if (is_numeric(number)) {
            let significance = 2 * self::SiGN(number);
            return (int) self::CeILING(number, significance);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FACT
     *
     * Returns the factorial of a number.
     * The factorial of a number is equal to 1*2*3*...* number.
     *
     * Excel Function:
     *        FACT(factVal)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    factVal    Factorial Value
     * @return    int        Factorial
     */
    public static function fact(var factVal)
    {
        var factLoop;
        double factorial;
        
        let factVal = \ZExcel\Calculation\Functions::flattenSingleValue(factVal);

        if (is_numeric(factVal)) {
            if (factVal < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let factLoop = floor(factVal);
            
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                if (factVal > factLoop) {
                    return \ZExcel\Calculation\Functions::NaN();
                }
            }

            let factorial = 1;
            
            while (factLoop > 1) {
                let factorial = factorial * factLoop;
                let factLoop = factLoop - 1;
            }
            
            return factorial ;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FACTDOUBLE
     *
     * Returns the double factorial of a number.
     *
     * Excel Function:
     *        FACTDOUBLE(factVal)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    factVal    Factorial Value
     * @return    int        Double Factorial
     */
    public static function factDouble(var factVal)
    {
        var factLoop;
        double factorial;
        
        let factLoop = \ZExcel\Calculation\Functions::flattenSingleValue(factVal);

        if (is_numeric(factLoop)) {
            let factLoop = floor(factLoop);
            if (factVal < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let factorial = 1;
            
            while (factLoop > 1) {
                let factorial = factorial * factLoop;
                let factLoop = factLoop - 2;
            }
            
            return factorial ;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FLOOR
     *
     * Rounds number down, toward zero, to the nearest multiple of significance.
     *
     * Excel Function:
     *        FLOOR(number[,significance])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    number            Number to round
     * @param    float    significance    Significance
     * @return    float    Rounded Number
     */
    public static function floor(var number, var significance = null)
    {
        let number       = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let significance = \ZExcel\Calculation\Functions::flattenSingleValue(significance);

        if ((is_null(significance)) &&
            (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC)) {
            let significance = number / abs(number);
        }

        if ((is_numeric(number)) && (is_numeric(significance))) {
            if (significance == 0.0) {
                return \ZExcel\Calculation\Functions::DiV0();
            } elseif (number == 0.0) {
                return 0.0;
            } elseif (self::SiGN(number) == self::SiGN(significance)) {
                return floor(number / significance) * significance;
            } else {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }

        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * GCD
     *
     * Returns the greatest common divisor of a series of numbers.
     * The greatest common divisor is the largest integer that divides both
     *        number1 and number2 without a remainder.
     *
     * Excel Function:
     *        GCD(number1[,number2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed    arg,...        Data values
     * @return    integer                    Greatest Common Divisor
     */
    public static function gcd()
    {
        var i, myFactors, myCountedFactors, allValuesCount, mergedArray, mergedArrayValues,
            highestPowerTest, testKey, testValue, mergedKey, mergedValue, key, value, keys,
            returnValue = 1, allValuesFactors = [];
        
        // Loop through arguments
        for value in \ZExcel\Calculation\Functions::flattenArray(func_get_args()) {
            if (!is_numeric(value)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } elseif (value == 0) {
                continue;
            } elseif (value < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            let myFactors = self::factors(value);
            let myCountedFactors = array_count_values(myFactors);
            let allValuesFactors[] = myCountedFactors;
        }
        
        let allValuesCount = count(allValuesFactors);
        
        if (allValuesCount == 0) {
            return 0;
        }

        let mergedArray = allValuesFactors[0];
        
        for i in range(1, allValuesCount - 1) {
            let mergedArray = array_intersect_key(mergedArray, allValuesFactors[i]);
        }
        
        let mergedArrayValues = count(mergedArray);
        
        if (mergedArrayValues == 0) {
            return returnValue;
        } elseif (mergedArrayValues > 1) {
            for mergedKey, mergedValue in mergedArray {
                for highestPowerTest in allValuesFactors {
                    for testKey, testValue in highestPowerTest {
                        if ((testKey == mergedKey) && (testValue < mergedValue)) {
                            let mergedArray[mergedKey] = testValue;
                            let mergedValue = testValue;
                        }
                    }
                }
            }

            let returnValue = 1;
            
            for key, value in mergedArray {
                let returnValue = returnValue * pow(key, value);
            }
            
            return returnValue;
        } else {
            let keys = array_keys(mergedArray);
            let key = keys[0];
            let value = mergedArray[key];
            
            for testValue in allValuesFactors {
                for mergedKey, mergedValue in testValue {
                    if ((mergedKey == key) && (mergedValue < value)) {
                        let value = mergedValue;
                    }
                }
            }
            
            return pow(key, value);
        }
    }


    /**
     * INT (rename Intt for zephir reserved word)
     *
     * Casts a floating point value to an integer
     *
     * Excel Function:
     *        INT(number)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    number            Number to cast to an integer
     * @return    integer    Integer value
     */
    public static function intt(number)
    {
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);

        if (is_null(number)) {
            return 0;
        } elseif (is_bool(number)) {
            return (int) number;
        }
        if (is_numeric(number)) {
            return (int) floor(number);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * LCM
     *
     * Returns the lowest common multiplier of a series of numbers
     * The least common multiple is the smallest positive integer that is a multiple
     * of all integer arguments number1, number2, and so on. Use LCM to add fractions
     * with different denominators.
     *
     * Excel Function:
     *        LCM(number1[,number2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed    arg,...        Data values
     * @return    int        Lowest Common Multiplier
     */
    public static function lcm()
    {
        var value, myFactors, myCountedFactor, myCountedFactors,
            myCountedPower, myPoweredFactors, myPoweredValue, allPoweredFactor,
            myPoweredFactor, returnValue = 1, allPoweredFactors = [];
        
        // Loop through arguments
        for value in \ZExcel\Calculation\Functions::flattenArray(func_get_args()) {
            if (!is_numeric(value)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            
            if (value == 0) {
                return 0;
            } elseif (value < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let myFactors = self::factors(floor(value));
            let myCountedFactors = array_count_values(myFactors);
            let myPoweredFactors = [];
            
            for myCountedFactor, myCountedPower in myCountedFactors {
                let myPoweredFactors[myCountedFactor] = pow(myCountedFactor, myCountedPower);
            }
            
            for myPoweredValue, myPoweredFactor in myPoweredFactors {
                if (array_key_exists(myPoweredValue, allPoweredFactors)) {
                    if (allPoweredFactors[myPoweredValue] < myPoweredFactor) {
                        let allPoweredFactors[myPoweredValue] = myPoweredFactor;
                    }
                } else {
                    let allPoweredFactors[myPoweredValue] = myPoweredFactor;
                }
            }
        }
        
        for allPoweredFactor in allPoweredFactors {
            let returnValue = returnValue * (int) allPoweredFactor;
        }
        
        return returnValue;
    }


    /**
     * LOG_BASE
     *
     * Returns the logarithm of a number to a specified base. The default base is 10.
     *
     * Excel Function:
     *        LOG(number[,base])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    float    number        The positive real number for which you want the logarithm
     * @param    float    base        The base of the logarithm. If base is omitted, it is assumed to be 10.
     * @return    float
     */
    public static function Log_base(var number = null, var base = 10)
    {
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let base   = (is_null(base)) ? 10 : (float) \ZExcel\Calculation\Functions::flattenSingleValue(base);

        if ((!is_numeric(base)) || (!is_numeric(number))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        if ((base <= 0) || (number <= 0)) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        return log(number, base);
    }


    /**
     * MDETERM
     *
     * Returns the matrix determinant of an array.
     *
     * Excel Function:
     *        MDETERM(array)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    array    matrixValues    A matrix of values
     * @return    float
     */
    public static function mDeterm(var matrixValues)
    {
        var matrixData = [], row = 0, column = 0, maxColumn = 0, matrix, matrixRow, matrixCell, ex;
        
        if (!is_array(matrixValues)) {
            let matrixValues = [
                [matrixValues]
            ];
        }

        for matrixRow in matrixValues {
            if (!is_array(matrixRow)) {
                let matrixRow = [matrixRow];
            }
            
            let column = 0;
            
            for matrixCell in matrixRow {
                if ((is_string(matrixCell)) || (matrixCell === null)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
                let matrixData[column][row] = matrixCell;
                let column = column + 1;
            }
            
            if (column > maxColumn) {
                let maxColumn = column;
            }
            
            let row = row + 1;
        }
        
        if (row != maxColumn) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        try {
            let matrix = new \ZExcel\Shared\JAMA\Matrix();
            call_user_func([matrix, "initialize"], matrixData);
            
            return matrix->det();
        } catch \ZExcel\Exception, ex {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
    }


    /**
     * MINVERSE
     *
     * Returns the inverse matrix for the matrix stored in an array.
     *
     * Excel Function:
     *        MINVERSE(array)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    array    matrixValues    A matrix of values
     * @return    array
     */
    public static function minverse(var matrixValues)
    {
        var matrixData = [], row = 0, column = 0, maxColumn = 0,
            matrixRow, matrixCell, matrix, ex;
        
        if (!is_array(matrixValues)) {
            let matrixValues = [
                [matrixValues]
            ];
        }

        
        for matrixRow in matrixValues {
            if (!is_array(matrixRow)) {
                let matrixRow = [matrixRow];
            }
            
            let column = 0;
            
            for matrixCell in matrixRow {
                if ((is_string(matrixCell)) || (matrixCell === null)) {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
                let matrixData[column][row] = matrixCell;
                let column = column + 1;
            }
            
            if (column > maxColumn) {
                let maxColumn = column;
            }
            let row = row + 1;
        }
        if (row != maxColumn) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }

        try {
            let matrix = new \ZExcel\Shared\JAMA\Matrix();
            call_user_func([matrix, "initialize"], matrixData);
            
            return matrix->inverse()->getArray();
        } catch \ZExcel\Exception, ex {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
    }


    /**
     * MMULT
     *
     * @param    array    matrixData1    A matrix of values
     * @param    array    matrixData2    A matrix of values
     * @return    array
     */
    public static function mmult(var matrixData1, var matrixData2)
    {
        var rowA, columnA, rowB, columnB, matrixRow, matrixCell, matrixA, matrixB, ex;
        array matrixAData = [], matrixBData = [];
        
        if (!is_array(matrixData1)) {
            let matrixData1 = [[matrixData1]];
        }
        if (!is_array(matrixData2)) {
            let matrixData2 = [[matrixData2]];
        }

        try {
            let rowA = 0;
            for matrixRow in matrixData1 {
                if (!is_array(matrixRow)) {
                    let matrixRow = [matrixRow];
                }
                
                let columnA = 0;
                
                for matrixCell in matrixRow {
                    if ((!is_numeric(matrixCell)) || (matrixCell === null)) {
                        return \ZExcel\Calculation\Functions::VaLUE();
                    }
                    
                    let matrixAData[rowA][columnA] = matrixCell;
                    let columnA = columnA + 1;
                }
                let rowA = rowA + 1;
            }
            
            let rowB = 0;
            let matrixA = new \ZExcel\Shared\JAMA\Matrix();
            call_user_func([matrixA, "initialize"], matrixAData);
            
            for matrixRow in matrixData2 {
                if (!is_array(matrixRow)) {
                    let matrixRow = [matrixRow];
                }
                
                let columnB = 0;
                
                for matrixCell in matrixRow {
                    if ((!is_numeric(matrixCell)) || (matrixCell === null)) {
                        return \ZExcel\Calculation\Functions::VaLUE();
                    }
                    
                    let matrixBData[rowB][columnB] = matrixCell;
                    let columnB = columnB + 1;
                }
                let rowB = rowB + 1;
            }
            
            let matrixB = new \ZExcel\Shared\JAMA\Matrix();
            call_user_func([matrixB, "initialize"], matrixBData);

            if (columnA != rowB) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
            
            let ex = call_user_func([matrixA, "times"], matrixB);

            return ex->getArray();
        } catch \ZExcel\Exception, ex {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
    }


    /**
     * MOD
     *
     * @param    int        a        Dividend
     * @param    int        b        Divisor
     * @return    int        Remainder
     */
    public static function mod(var a = 1, var b = 1) -> int
    {
        let a = \ZExcel\Calculation\Functions::flattenSingleValue(a);
        let b = \ZExcel\Calculation\Functions::flattenSingleValue(b);

        if (b == 0.0) {
            return \ZExcel\Calculation\Functions::DiV0();
        } elseif ((a < 0.0) && (b > 0.0)) {
            return b - fmod(abs(a), b);
        } elseif ((a > 0.0) && (b < 0.0)) {
            return b + fmod(a, abs(b));
        }

        return fmod(a, b);
    }


    /**
     * MROUND
     *
     * Rounds a number to the nearest multiple of a specified value
     *
     * @param    float    number            Number to round
     * @param    int        multiple        Multiple to which you want to round number
     * @return    float    Rounded Number
     */
    public static function mRound(var number, var multiple)
    {
        var multiplier;
        
        let number   = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let multiple = \ZExcel\Calculation\Functions::flattenSingleValue(multiple);

        if ((is_numeric(number)) && (is_numeric(multiple))) {
            if (multiple == 0) {
                return 0;
            }
            if ((self::SiGN(number)) == (self::SiGN(multiple))) {
                let multiplier = 1 / multiple;
                return round((double) number * multiplier) / multiplier;
            }
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * MULTINOMIAL
     *
     * Returns the ratio of the factorial of a sum of values to the product of factorials.
     *
     * @param    array of mixed        Data Series
     * @return    float
     */
    public static function multinomial()
    {
        var summer = 0, divisor = 1, aArgs, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        // Loop through arguments
        for arg in aArgs {
            // Is it a numeric value?
            if (is_numeric(arg)) {
                if (arg < 1) {
                    return \ZExcel\Calculation\Functions::NaN();
                }
                let summer = summer + floor(arg);
                let divisor = divisor * self::FaCT(arg);
            } else {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }

        // Return
        if (summer > 0) {
            let summer = self::FaCT(summer);
            return summer / divisor;
        }
        
        return 0;
    }


    /**
     * ODD
     *
     * Returns number rounded up to the nearest odd integer.
     *
     * @param    float    number            Number to round
     * @return    int        Rounded Number
     */
    public static function odd(var number)
    {
        var result, significance;
        
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);

        if (is_null(number)) {
            return 1;
        } elseif (is_bool(number)) {
            return 1;
        } elseif (is_numeric(number)) {
            let significance = self::SiGN(number);
            if (significance == 0) {
                return 1;
            }

            let result = self::CeILING(number, significance);
            
            if (result == self::eVEN(result)) {
                let result = result + significance;
            }

            return (int) result;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * POWER
     *
     * Computes x raised to the power y.
     *
     * @param    float        x
     * @param    float        y
     * @return    float
     */
    public static function power(var x = 0, var y = 2)
    {
        var result;
        
        let x = \ZExcel\Calculation\Functions::flattenSingleValue(x);
        let y = \ZExcel\Calculation\Functions::flattenSingleValue(y);

        // Validate parameters
        if (x == 0.0 && y == 0.0) {
            return \ZExcel\Calculation\Functions::NaN();
        } elseif (x == 0.0 && y < 0.0) {
            return \ZExcel\Calculation\Functions::DiV0();
        }

        // Return
        let result = pow(x, y);
        
        return (!is_nan(result) && !is_infinite(result)) ? result : \ZExcel\Calculation\Functions::NaN();
    }


    /**
     * PRODUCT
     *
     * PRODUCT returns the product of all the values and cells referenced in the argument list.
     *
     * Excel Function:
     *        PRODUCT(value1[,value2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function product()
    {
        var returnValue = null, aArgs, arg;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        // Loop through arguments
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                if (is_null(returnValue)) {
                    let returnValue = arg;
                } else {
                    let returnValue = returnValue * arg;
                }
            }
        }

        // Return
        if (is_null(returnValue)) {
            return 0;
        }
        
        return returnValue;
    }


    /**
     * QUOTIENT
     *
     * QUOTIENT function returns the integer portion of a division. Numerator is the divided number
     *        and denominator is the divisor.
     *
     * Excel Function:
     *        QUOTIENT(value1[,value2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function quotient()
    {
        var returnValue = null, aArgs, arg;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        // Loop through arguments
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                if (is_null(returnValue)) {
                    let returnValue = (arg == 0) ? 0 : arg;
                } else {
                    if ((returnValue == 0) || (arg == 0)) {
                        let returnValue = 0;
                    } else {
                        let returnValue = returnValue / arg;
                    }
                }
            }
        }

        // Return
        return intval(returnValue);
    }


    /**
     * RAND
     *
     * @param    int        min    Minimal value
     * @param    int        max    Maximal value
     * @return    int        Random number
     */
    public static function rand(min = 0, max = 0)
    {
        let min = \ZExcel\Calculation\Functions::flattenSingleValue(min);
        let max = \ZExcel\Calculation\Functions::flattenSingleValue(max);

        if (min == 0 && max == 0) {
            return (mt_rand(0, 10000000)) / 10000000;
        }

        return mt_rand(min, max);
    }


    public static function roman(var aValue, var style = 0) -> string
    {
        var mill, cent, tens, ones, roman, m, c, t;
        
        let aValue = \ZExcel\Calculation\Functions::flattenSingleValue(aValue);
        let style = ((is_null(style)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(style));
        
        if ((!is_numeric(aValue)) || (aValue < 0) || (aValue >= 4000)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let aValue = (int) aValue;
        
        if (aValue == 0) {
            return "";
        }

        let mill = ["", "M", "MM", "MMM", "MMMM", "MMMMM"];
        let cent = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"];
        let tens = ["", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"];
        let ones = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"];

        let roman = "";
        
        while (aValue > 5999) {
            let roman  = roman . "M";
            let aValue = aValue - 1000;
        }
        
        let m = self::romanCut(aValue, 1000);
        let aValue = aValue % 1000;
        let c = self::romanCut(aValue, 100);
        let aValue = aValue % 100;
        let t = self::romanCut(aValue, 10);
        let aValue = aValue % 10;

        return roman . mill[m] . cent[c] . tens[t] . ones[aValue];
    }


    /**
     * ROUNDUP
     *
     * Rounds a number up to a specified number of decimal places
     *
     * @param    float    number            Number to round
     * @param    int        digits            Number of digits to which you want to round number
     * @return    float    Rounded Number
     */
    public static function roundUp(var number, var digits)
    {
        var significance;
        
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let digits = \ZExcel\Calculation\Functions::flattenSingleValue(digits);

        if ((is_numeric(number)) && (is_numeric(digits))) {
            let significance = pow(10, (int) digits);
            if (number < 0.0) {
                return floor(number * significance) / significance;
            } else {
                return ceil(number * significance) / significance;
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * ROUNDDOWN
     *
     * Rounds a number down to a specified number of decimal places
     *
     * @param    float    number            Number to round
     * @param    int        digits            Number of digits to which you want to round number
     * @return    float    Rounded Number
     */
    public static function roundDown(var number, var digits)
    {
        var significance;
        
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);
        let digits = \ZExcel\Calculation\Functions::flattenSingleValue(digits);

        if ((is_numeric(number)) && (is_numeric(digits))) {
            let significance = pow(10, (int) digits);
            if (number < 0.0) {
                return ceil(number * significance) / significance;
            } else {
                return floor(number * significance) / significance;
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SERIESSUM
     *
     * Returns the sum of a power series
     *
     * @param    float            x    Input value to the power series
     * @param    float            n    Initial power to which you want to raise x
     * @param    float            m    Step by which to increase n for each term in the series
     * @param    array of mixed        Data Series
     * @return    float
     */
    public static function seriesSum()
    {
        var returnValue = 0, aArgs, x, n, m, i, arg;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        let x = array_shift(aArgs);
        let n = array_shift(aArgs);
        let m = array_shift(aArgs);

        if ((is_numeric(x)) && (is_numeric(n)) && (is_numeric(m))) {
            // Calculate
            let i = 0;
            for arg in aArgs {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let returnValue = returnValue + (arg * pow(x, n + (m * i)));
                    let i = i + 1;
                } else {
                    return \ZExcel\Calculation\Functions::VaLUE();
                }
            }
            
            return returnValue;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SIGN
     *
     * Determines the sign of a number. Returns 1 if the number is positive, zero (0)
     *        if the number is 0, and -1 if the number is negative.
     *
     * @param    float    number            Number to round
     * @return    int        sign value
     */
    public static function SiGN(var number)
    {
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);

        if (is_bool(number)) {
            return (int) number;
        }
        if (is_numeric(number)) {
            if (number == 0.0) {
                return 0;
            }
            
            return number / abs(number);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SQRTPI
     *
     * Returns the square root of (number * pi).
     *
     * @param    float    number        Number
     * @return    float    Square Root of Number * Pi
     */
    public static function sqrTPI(var number)
    {
        let number = \ZExcel\Calculation\Functions::flattenSingleValue(number);

        if (is_numeric(number)) {
            if (number < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return sqrt(number * M_PI) ;
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SUBTOTAL
     *
     * Returns a subtotal in a list or database.
     *
     * @param    int        the number 1 to 11 that specifies which function to
     *                    use in calculating subtotals within a list.
     * @param    array of mixed        Data Series
     * @return    float
     */
    public static function subTotal()
    {
        var aArgs, subtotal;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let subtotal = array_shift(aArgs);

        if ((is_numeric(subtotal)) && (!is_string(subtotal))) {
            switch (subtotal) {
                case 1:
                    return call_user_func("\ZExcel\Calculation\Statistical::AVERAGE", aArgs);
                case 2:
                    return call_user_func("\ZExcel\Calculation\Statistical::COUNT", aArgs);
                case 3:
                    return call_user_func("\ZExcel\Calculation\Statistical::COUNTA", aArgs);
                case 4:
                    return call_user_func("\ZExcel\Calculation\Statistical::MAX", aArgs);
                case 5:
                    return call_user_func("\ZExcel\Calculation\Statistical::MIN", aArgs);
                case 6:
                    return call_user_func("self::PRODUCT", aArgs);
                case 7:
                    return call_user_func("\ZExcel\Calculation\Statistical::STDEV", aArgs);
                case 8:
                    return call_user_func("\ZExcel\Calculation\Statistical::STDEVP", aArgs);
                case 9:
                    return call_user_func("self::SUM", aArgs);
                case 10:
                    return call_user_func("\ZExcel\Calculation\Statistical::VARFunc", aArgs);
                case 11:
                    return call_user_func("\ZExcel\Calculation\Statistical::VARP", aArgs);
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * SUM
     *
     * SUM computes the sum of all the values and cells referenced in the argument list.
     *
     * Excel Function:
     *        SUM(value1[,value2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function sum() -> float
    {
        var aArgs, arg;
        float returnValue = 0;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Loop through the arguments
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let returnValue = returnValue + (float) arg;
            }
        }

        return returnValue;
    }


    /**
     * SUMIF
     *
     * Counts the number of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        SUMIF(value1[,value2[, ...]],condition)
     *
     * @access   public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @param    string       condition      The criteria that defines which cells will be summed.
     * @return   float
     */
    public static function sumIf(var aArgs, var condition, var sumArgs = []) -> float
    {
        var aArgs, sumArgs, key, arg, testCondition;
        float returnValue = 0;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(aArgs);
        let sumArgs = \ZExcel\Calculation\Functions::flattenArray(sumArgs);
        
        if (empty(sumArgs)) {
            let sumArgs = aArgs;
        }
        
        let condition = \ZExcel\Calculation\Functions::ifCondition(condition);
        
        // Loop through arguments
        for key, arg in aArgs {
            if (!is_numeric(arg)) {
                let arg = str_replace("\"", "\"\"", arg);
                let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
            }
            
            let testCondition = "=" . arg . condition;
            
            if (\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                // Is it a value within our criteria
                let returnValue = returnValue + (float) sumArgs[key];
            }
        }

        return returnValue;
    }

    /**
     *  SUMIFS
     *
     *  Counts the number of cells that contain numbers within the list of arguments
     *
     *  Excel Function:
     *      SUMIFS(value1[,value2[, ...]],condition)
     *
     *  @access public
     *  @category Mathematical and Trigonometric Functions
     *  @param  mixed       arg,...        Data values
     *  @param  string      condition      The criteria that defines which cells will be summed.
     *  @return float
     */
    public static function sumIfs() {
        var arrayList, sumArgs, aArgsArray, conditions, condition, testCondition, index, aArgs, key, arg;
        boolean wildcard;
    
        let arrayList = func_get_args();

        let sumArgs = \ZExcel\Calculation\Functions::flattenArray(array_shift(arrayList));

        let aArgsArray = [];
        let conditions = [];
        while (count(arrayList) > 0) {
            let aArgsArray[] = \ZExcel\Calculation\Functions::flattenArray(array_shift(arrayList));
            let conditions[] = \ZExcel\Calculation\Functions::ifCondition(array_shift(arrayList));
        }

        // Loop through each set of arguments and conditions
        for index, condition in conditions {
            let aArgs = aArgsArray[index];
            let wildcard = false;
            if ((strpos(condition, "*") !== false) || (strpos(condition, "?") !== false)) {
                // * and ? are wildcard characters.
                // Use ~* and ~? for literal star and question mark
                // Code logic doesn"t yet handle escaping
                let condition = trim(ltrim(condition, "=<>"), "\"");
                let wildcard = true;
            }
            // Loop through arguments
            for key, arg in aArgs {
                if (wildcard) {
                    if (!fnmatch(condition, arg, FNM_CASEFOLD)) {
                        // Is it a value within our criteria
                        let sumArgs[key] = 0.0;
                    }
                } else {
                    if (!is_numeric(arg)) {
                        let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
                    }
                    
                    let testCondition = "=" . arg . condition;
                    
                    if (!\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                        // Is it a value within our criteria
                        let sumArgs[key] = 0.0;
                    }
                }
            }
        }

        // Return
        return array_sum(sumArgs);
    }

    /**
     * SUMPRODUCT
     *
     * Excel Function:
     *        SUMPRODUCT(value1[,value2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function sumProduct()
    {
        var arrayList, wrkArray, wrkCellCount, i, matrixData, array2, count, val;
        
        let arrayList = func_get_args();
        let wrkArray = \ZExcel\Calculation\Functions::flattenArray(array_shift(arrayList));
        let wrkCellCount = count(wrkArray);

        for i in range(0, wrkCellCount - 1) {
            if ((!is_numeric(wrkArray[i])) || (is_string(wrkArray[i]))) {
                let wrkArray[i] = 0;
            }
        }

        for matrixData in arrayList {
            let array2 = \ZExcel\Calculation\Functions::flattenArray(matrixData);
            let count = count(array2);
            if (wrkCellCount != count) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }

            for i, val in array2 {
                if ((!is_numeric(val)) || (is_string(val))) {
                    let val = 0;
                }
                let wrkArray[i] = wrkArray[i] * val;
            }
        }

        return array_sum(wrkArray);
    }


    /**
     * SUMSQ
     *
     * SUMSQ returns the sum of the squares of the arguments
     *
     * Excel Function:
     *        SUMSQ(value1[,value2[, ...]])
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function sumSQ() -> float
    {
        var aArgs, arg;
        float returnValue = 0;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Loop through arguments
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let returnValue = returnValue + (arg * arg);
            }
        }

        return returnValue;
    }


    /**
     * SUMX2MY2
     *
     * @param    mixed[]    matrixData1    Matrix #1
     * @param    mixed[]    matrixData2    Matrix #2
     * @return    float
     */
    public static function SumX2MY2(var matrixData1, var matrixData2) -> float
    {
        var array1, array2, count;
        float result;
        int i;
        
        let array1 = \ZExcel\Calculation\Functions::flattenArray(matrixData1);
        let array2 = \ZExcel\Calculation\Functions::flattenArray(matrixData2);
        let count = min(count(array1), count(array2)) - 1;

        let result = 0;
        for i in range(0, count) {
            if (((is_numeric(array1[i])) && (!is_string(array1[i]))) && ((is_numeric(array2[i])) && (!is_string(array2[i])))) {
                let result = result + (float) (array1[i] * array1[i]) - (array2[i] * array2[i]);
            }
        }

        return result;
    }


    /**
     * SUMX2PY2
     *
     * @param    mixed[]    matrixData1    Matrix #1
     * @param    mixed[]    matrixData2    Matrix #2
     * @return    float
     */
    public static function SumX2PY2(var matrixData1, var matrixData2) -> float
    {
        var array1, array2, count;
        float result;
        int i;
        
        let array1 = \ZExcel\Calculation\Functions::flattenArray(matrixData1);
        let array2 = \ZExcel\Calculation\Functions::flattenArray(matrixData2);
        let count = min(count(array1), count(array2)) - 1;

        let result = 0;
        for i in range(0,count) {
            if (((is_numeric(array1[i])) && (!is_string(array1[i]))) && ((is_numeric(array2[i])) && (!is_string(array2[i])))) {
                let result = result + (float) (array1[i] * array1[i]) + (array2[i] * array2[i]);
            }
        }

        return result;
    }


    /**
     * SUMXMY2
     *
     * @param    mixed[]    matrixData1    Matrix #1
     * @param    mixed[]    matrixData2    Matrix #2
     * @return    float
     */
    public static function SumXMY2(var matrixData1, var matrixData2) -> float
    {
        var array1, array2, count;
        float result;
        int i;
        
        let array1 = \ZExcel\Calculation\Functions::flattenArray(matrixData1);
        let array2 = \ZExcel\Calculation\Functions::flattenArray(matrixData2);
        let count = min(count(array1), count(array2)) - 1;

        let result = 0;
        for i in range(0, count) {
            if (((is_numeric(array1[i])) && (!is_string(array1[i]))) && ((is_numeric(array2[i])) && (!is_string(array2[i])))) {
                let result = result + (float) (array1[i] - array2[i]) * (array1[i] - array2[i]);
            }
        }

        return result;
    }


    /**
     * TRUNC
     *
     * Truncates value to the number of fractional digits by number_digits.
     *
     * @param    float        value
     * @param    int          digits
     * @return   float        Truncated value
     */
    public static function trunc(var value = 0, var digits = 0)
    {
        var adjust;
        
        let value    = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let digits   = \ZExcel\Calculation\Functions::flattenSingleValue(digits);

        // Validate parameters
        if ((!is_numeric(value)) || (!is_numeric(digits))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        let digits = floor(digits);

        // Truncate
        let adjust = pow(10, digits);

        if ((digits > 0) && (rtrim(intval((abs(value) - abs(intval(value))) * adjust), strval(0)) < adjust/10)) {
            return value;
        }

        return (intval(value * adjust)) / adjust;
    }
}
