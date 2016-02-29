namespace ZExcel\Shared\JAMA;

class Matrix 
{

    const POLYMORPHIC_ARGUMENT_EXCEPTION = -1;
    const ARGUMENT_TYPE_EXCEPTION = -2;
    const ARGUMENT_BOUNDS_EXCEPTION = -3;
    const MATRIX_DIMENSION_EXCEPTION = -4;
    const PRECISION_LOSS_EXCEPTION = -5;
    const MATRIX_SPD_EXCEPTION = -6;
    const MATRIX_SINGULAR_EXCEPTION = -7;
    const MATRIX_RANK_EXCEPTION = -8;
    const ARRAY_LENGTH_EXCEPTION = -9;
    const ROW_LENGTH_EXCEPTION = -10;
    
    private static jamaLang = "EN";
    
    private static errors = [
        "EN": [
            "-1": "Invalid argument pattern for polymorphic function.",
            "-2": "Invalid argument type.",
            "-3": "Invalid argument range.",
            "-4": "Matrix dimensions are not equal.",
            "-5": "Significant precision loss detected.",
            "-6": "Can only perform operation on symmetric positive definite matrix.",
            "-7": "Can only perform operation on singular matrix.",
            "-8": "Can only perform operation on full-rank matrix.",
            "-9": "Array length must be a multiple of m.",
            "-10": "All rows must have the same length."
        ],
        "FR": [
            "-1": "Modèle inadmissible d'argument pour la fonction polymorphe.",
            "-2": "Type inadmissible d'argument.",
            "-3": "Gamme inadmissible d'argument.",
            "-4": "Les dimensions de Matrix ne sont pas égales.",
            "-5": "Perte significative de précision détectée.",
            "-6": "Perte significative de précision détectée."
        ],
        "DE": [
            "-1": "Unzulässiges Argumentmuster für polymorphe Funktion.",
            "-2": "Unzulässige Argumentart.",
            "-3": "Unzulässige Argumentstrecke.",
            "-4": "Matrixmaße sind nicht gleich.",
            "-5": "Bedeutender Präzision Verlust ermittelte.",
            "-6": "Bedeutender Präzision Verlust ermittelte."
        ]
    ];
    
    /**
     *    Matrix storage
     *
     *    @var array
     *    @access public
     */
    public a = [];

    /**
     *    Matrix row dimension
     *
     *    @var int
     *    @access private
     */
    private m;

    /**
     *    Matrix column dimension
     *
     *    @var int
     *    @access private
     */
    private n;

    /**
     *    Polymorphic constructor
     *
     *    As PHP has no support for polymorphic constructors, we hack our own sort of polymorphism using func_num_args, func_get_arg, and gettype. In essence, we"re just implementing a simple RTTI filter and calling the appropriate constructor.
     */
    public function __construct()
    {
    }
    
    public function initialize()
    {
        var args, match;
        int i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                // Rectangular matrix - m x n initialized from 2D array
                case "array":
                    let this->m = count(args[0]);
                    let this->n = count(args[0][0]);
                    let this->a = args[0];
                    break;
                // Square matrix - n x n
                case "integer":
                    let this->m = args[0];
                    let this->n = args[0];
                    let this->a = array_fill(0, this->m, array_fill(0, this->n, 0));
                    break;
                // Rectangular matrix - m x n
                case "integer,integer":
                    let this->m = args[0];
                    let this->n = args[1];
                    let this->a = array_fill(0, this->m, array_fill(0, this->n, 0));
                    break;
                // Rectangular matrix - m x n initialized from packed array
                case "array,integer":
                    let this->m = args[1];
                    
                    if (this->m != 0) {
                        let this->n = count(args[0]) / this->m;
                    } else {
                        let this->n = 0;
                    }
                    
                    if ((this->m * this->n) == count(args[0])) {
                        for i in range(0, this->m - 1) {
                            for j in range(0, this->n - 1) {
                                let this->a[i][j] = args[0][i + j * this->m];
                            }
                        }
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARRAY_LENGTH_EXCEPTION));
                    }
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
                    break;
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
    }

    /**
     *    getArray
     *
     *    @return array Matrix array
     */
    public function getArray()
    {
        return this->a;
    }

    /**
     *    getRowDimension
     *
     *    @return int Row dimension
     */
    public function getRowDimension()
    {
        return this->m;
    }

    /**
     *    getColumnDimension
     *
     *    @return int Column dimension
     */
    public function getColumnDimension()
    {
        return this->n;
    }

    /**
     *    get
     *
     *    Get the i,j-th element of the matrix.
     *    @param int i Row position
     *    @param int j Column position
     *    @return mixed Element (int/float/double)
     */
    public function get(var i = null, var j = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    getMatrix
     *
     *    Get a submatrix
     *    @param int i0 Initial row index
     *    @param int iF Final row index
     *    @param int j0 Initial column index
     *    @param int jF Final column index
     *    @return Matrix Submatrix
     */
    public function getMatrix()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    checkMatrixDimensions
     *
     *    Is matrix B the same size?
     *    @param Matrix B Matrix B
     *    @return boolean
     */
    public function checkMatrixDimensions(var b = null)
    {
        throw new \Exception("Not implemented yet!");
    }    //    function checkMatrixDimensions()

    /**
     *    set
     *
     *    Set the i,j-th element of the matrix.
     *    @param int i Row position
     *    @param int j Column position
     *    @param mixed c Int/float/double value
     *    @return mixed Element (int/float/double)
     */
    public function set(var i = null, var j = null, var c = null)
    {
        // Optimized set version just has this
        throw new \Exception("Not implemented yet!");
    }    //    function set()

    /**
     *    identity
     *
     *    Generate an identity matrix.
     *    @param int m Row dimension
     *    @param int n Column dimension
     *    @return Matrix Identity matrix
     */
    public function identity(var m = null, var n = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    diagonal
     *
     *    Generate a diagonal matrix
     *    @param int m Row dimension
     *    @param int n Column dimension
     *    @param mixed c Diagonal value
     *    @return Matrix Diagonal matrix
     */
    public function diagonal(var m = null, var n = null, var c = 1)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    getMatrixByRow
     *
     *    Get a submatrix by row index/range
     *    @param int i0 Initial row index
     *    @param int iF Final row index
     *    @return Matrix Submatrix
     */
    public function getMatrixByRow(var i0 = null, var iFF = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    getMatrixByCol
     *
     *    Get a submatrix by column index/range
     *    @param int i0 Initial column index
     *    @param int iF Final column index
     *    @return Matrix Submatrix
     */
    public function getMatrixByCol(var j0 = null, var jF = null)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    transpose
     *
     *    Tranpose matrix
     *    @return Matrix Transposed matrix
     */
    public function transpose()
    {
        throw new \Exception("Not implemented yet!");
    }    //    function transpose()

    /**
     *    trace
     *
     *    Sum of diagonal elements
     *    @return float Sum of diagonal elements
     */
    public function trace()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    uminus
     *
     *    Unary minus matrix -A
     *    @return Matrix Unary minus matrix
     */
    public function uminus()
    {
    }

    /**
     *    plus
     *
     *    A + B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function plus()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    plusEquals
     *
     *    A = A + B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function plusEquals()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    minus
     *
     *    A - B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function minus()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    minusEquals
     *
     *    A = A - B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function minusEquals()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    arrayTimes
     *
     *    Element-by-element multiplication
     *    Cij = Aij * Bij
     *    @param mixed B Matrix/Array
     *    @return Matrix Matrix Cij
     */
    public function arrayTimes()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    arrayTimesEquals
     *
     *    Element-by-element multiplication
     *    Aij = Aij * Bij
     *    @param mixed B Matrix/Array
     *    @return Matrix Matrix Aij
     */
    public function arrayTimesEquals()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    arrayRightDivide
     *
     *    Element-by-element right division
     *    A / B
     *    @param Matrix B Matrix B
     *    @return Matrix Division result
     */
    public function arrayRightDivide()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    arrayRightDivideEquals
     *
     *    Element-by-element right division
     *    Aij = Aij / Bij
     *    @param mixed B Matrix/Array
     *    @return Matrix Matrix Aij
     */
    public function arrayRightDivideEquals()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    arrayLeftDivide
     *
     *    Element-by-element Left division
     *    A / B
     *    @param Matrix B Matrix B
     *    @return Matrix Division result
     */
    public function arrayLeftDivide()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    arrayLeftDivideEquals
     *
     *    Element-by-element Left division
     *    Aij = Aij / Bij
     *    @param mixed B Matrix/Array
     *    @return Matrix Matrix Aij
     */
    public function arrayLeftDivideEquals()
    {
        throw new \Exception("Not implemented yet!");
    }


    /**
     *    times
     *
     *    Matrix multiplication
     *    @param mixed n Matrix/Array/Scalar
     *    @return Matrix Product
     */
    public function times()
    {
        var args, match, b, c;
        array Bcolj, Arowi;
        int i, j, k;
        double s;
        
        if (func_num_args() > 0) {
            let args  = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let b = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    
                    if (this->n == b->m) {
                        let c = new \ZExcel\Shared\JAMA\Matrix();
                        call_user_func([c, "initialize"], this->m, b->n);
                        
                        var Bcolj = [];
                        var Arowi = [];
                        
                        for j in range(0, b->n - 1) {
                            for k in range(0, this->n - 1) {
                                let Bcolj[k] = b->a[k][j];
                            }
                            
                            for i in range(0, this->m - 1) {
                                let Arowi = this->a[i];
                                let s = 0;
                                
                                for k in range(0, this->n - 1) {
                                    let s = s + (Arowi[k] * Bcolj[k]);
                                }
                                
                                let c->a[i][j] = s;
                            }
                        }
                        
                        return c;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::MATRIX_DIMENSION_EXCEPTION));
                    }
                    break;
                case "array":
                    let b = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([b, "initialize"], args[0]);
                    
                    if (this->n == b->m) {
                        let c = new \ZExcel\Shared\JAMA\Matrix();
                        call_user_func([c, "initialize"], this->m, b->n);
                        
                        for i in range(0, c->m - 1) {
                            for j in range(0, c->n - 1) {
                                let s = 0;
                                
                                for k in range(0, c->n - 1) {
                                    let s = s + (this->a[i][k] * b->a[k][j]);
                                }
                                
                                let c->a[i][j] = s;
                            }
                        }
                        
                        return c;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::MATRIX_DIMENSION_EXCEPTION));
                    }
                    break;
                case "integer":
                    let c = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([c, "initialize"], this->a);
                    
                    for i in range(0, c->m - 1) {
                        for j in range(0, c->n - 1) {
                            let c->a[i][j] *= args[0];
                        }
                    }
                    
                    return c;
                    break;
                case "double":
                    let c = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([c, "initialize"], this->m, this->n);
                    
                    for i in range(0, c->m - 1) {
                        for j in range(0, c->n - 1) {
                            let c->a[i][j] = args[0] * this->a[i][j];
                        }
                    }
                    
                    return c;
                    break;
                case "float":
                    let c = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([c, "initialize"], this->a);
                    
                    for i in range(0, c->m - 1) {
                        for j in range(0, c->n - 1) {
                            let c->a[i][j] *= args[0];
                        }
                    }
                    
                    return c;
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
                    break;
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
    }

    /**
     *    power
     *
     *    A = A ^ B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function power()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    concat
     *
     *    A = A & B
     *    @param mixed B Matrix/Array
     *    @return Matrix Sum
     */
    public function concat()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Solve A*X = B.
     *
     *    @param Matrix B Right hand side
     *    @return Matrix ... Solution if A is square, least squares solution otherwise
     */
    public function solve(b)
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Matrix inverse or pseudoinverse.
     *
     *    @return Matrix ... Inverse(A) if A is square, pseudoinverse otherwise.
     */
    public function inverse()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    det
     *
     *    Calculate determinant
     *    @return float Determinant
     */
    public function det()
    {
        throw new \Exception("Not implemented yet!");
    }
    
    /**
     *    Custom error handler
     *    @param int num Error number
     */
    public static function JAMAError(var errorNumber = null)
    {
        if (errorNumber != null) {
            if (isset(self::errors[self::jamaLang][errorNumber])) {
                return self::errors[self::jamaLang][errorNumber];
            } elseif (isset(self::errors["EN"][errorNumber])) {
                return self::errors["EN"][errorNumber];
            }
        }

        return ("Invalid argument to JAMAError()");
    }
}
