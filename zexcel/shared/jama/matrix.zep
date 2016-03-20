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
            "-1": "Modèle inadmissible d argument pour la fonction polymorphe.",
            "-2": "Type inadmissible d argument.",
            "-3": "Gamme inadmissible d argument.",
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
    public function get(int i, int j)
    {
        return this->a[i][j];
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
        var args, match, i, j, i0, j0, m, n, r, iFF, jF, rl, cl;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                // A(i0...; j0...)
                case "integer,integer":
                    let i0 = args[0];
                    let j0 = args[1];
                    
                    if (i0 >= 0) {
                        let m = this->m - i0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if (j0 >= 0) {
                        let n = this->n - j0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m, n);
                    
                    for i in range(i0, this->m - 1) {
                        for j in range(j0, this->n - 1) {
                            r->set(i, j, this->a[i][j]);
                        }
                    }
                    
                    return r;
                // A(i0...iF; j0...jF)
                case "integer,integer,integer,integer":
                    let i0 = args[0];
                    let iFF = args[1];
                    let j0 = args[2];
                    let jF = args[3];
                    
                    if ((iFF > i0) && (this->m >= iFF) && (i0 >= 0)) {
                        let m = iFF - i0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if ((jF > j0) && (this->n >= jF) && (j0 >= 0)) {
                        let n = jF - j0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m + 1, n + 1);
                    
                    for i in range(i0, iFF) {
                        for j in range(j0, jF) {
                            r->set(i - i0, j - j0, this->a[i][j]);
                        }
                    }
                    
                    return r;
                //r = array of row indices; C = array of column indices
                case "array,array":
                    let rl = args[0];
                    let cl = args[1];
                    
                    if (count(rl) > 0) {
                        let m = count(rl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if (count(cl) > 0) {
                        let n = count(cl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m, n);
                    
                    for i in range(0, m - 1) {
                        for j in range(0, n - 1) {
                            r->set(i - i0, j - j0, this->a[rl[i]][cl[j]]);
                        }
                    }
                    
                    return r;
                //rl = array of row indices; cl = array of column indices
                case "array,array":
                    let rl = args[0];
                    let cl = args[1];
                    
                    if (count(rl) > 0) {
                        let m = count(rl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if (count(cl) > 0) {
                        let n = count(cl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m, n);
                    
                    for i in range(0, m - 1) {
                        for j in range(0, n - 1) {
                            r->set(i, j, this->a[rl[i]][cl[j]]);
                        }
                    }
                    
                    return r;
                //A(i0...iFF); cl = array of column indices
                case "integer,integer,array":
                    let i0 = args[0];
                    let iFF = args[1];
                    let cl = args[2];
                    
                    if ((iFF > i0) && (this->m >= iFF) && (i0 >= 0)) {
                        let m = iFF - i0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if (count(cl) > 0) {
                        let n = count(cl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m, n);
                    
                    for i in range(i0, iFF - 1) {
                        for j in range(0, n - 1) {
                            r->set(i - i0, j, this->a[rl[i]][j]);
                        }
                    }
                    
                    return r;
                //rl = array of row indices
                case "array,integer,integer":
                    let rl = args[0];
                    let j0 = args[1];
                    let jF = args[2];
                    
                    if (count(rl) > 0) {
                        let m = count(rl);
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    if ((jF >= j0) && (this->n >= jF) && (j0 >= 0)) {
                        let n = jF - j0;
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_BOUNDS_EXCEPTION));
                    }
                    
                    let r = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([r, "initialize"], m, n + 1);
                    
                    for i in range(0, m - 1) {
                        for j in range(j0, jF) {
                            r->set(i, j - j0, this->a[rl[i]][j]);
                        }
                    }
                    return r;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        if (b instanceof \ZExcel\Shared\JAMA\Matrix) {
            if ((this->m == b->getRowDimension()) && (this->n == b->getColumnDimension())) {
                return true;
            } else {
                throw new \ZExcel\Calculation\Exception(self::MATRIX_DIMENSION_EXCEPTION);
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::ARGUMENT_TYPE_EXCEPTION);
        }
    }

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
        let this->a[i][j] = c;
    }

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
        return this->diagonal(m, n, 1);
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
        var r;
        int i;
        
        let r = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([r, "initialize"], m, n);
        
        for i in range(0, m - 1) {
            r->set(i, i, c);
        }
        
        return r;
    }

    /**
     *    getMatrixByRow
     *
     *    Get a submatrix by row index/range
     *    @param int i0 Initial row index
     *    @param int iFF Final row index
     *    @return Matrix Submatrix
     */
    public function getMatrixByRow(var i0 = null, var iFF = null)
    {
        if (is_int(i0)) {
            if (is_int(iFF)) {
                return call_user_func([this, "getMatrix"], i0, 0, iFF + 1, this->n);
            } else {
                return call_user_func([this, "getMatrix"], i0, 0, i0 + 1, this->n);
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
        }
    }

    /**
     *    getMatrixByCol
     *
     *    Get a submatrix by column index/range
     *    @param int j0 Initial column index
     *    @param int jF Final column index
     *    @return Matrix Submatrix
     */
    public function getMatrixByCol(var j0 = null, var jF = null)
    {
        if (is_int(j0)) {
            if (is_int(jF)) {
                return call_user_func([this, "getMatrix"], 0, j0, this->m, jF + 1);
            } else {
                return call_user_func([this, "getMatrix"], 0, j0, this->m, j0 + 1);
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
        }
    }

    /**
     *    transpose
     *
     *    Tranpose matrix
     *    @return Matrix Transposed matrix
     */
    public function transpose()
    {
        var r;
        int i, j;
        
        let r = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([r, "initialize"], this->n, this->m);
        
        for i in range(0, this->m - 1) {
            for j in range(0, this->n - 1) {
                r->set(j, i, this->a[i][j]);
            }
        }
        
        return r;
    }    //    function transpose()

    /**
     *    trace
     *
     *    Sum of diagonal elements
     *    @return float Sum of diagonal elements
     */
    public function trace()
    {
        double s;
        int i, n;
        
        let n = 0 + min(this->m, this->n);
        
        for i in range(0, n - 1) {
            let s = s + this->a[i][i];
        }
        
        return s;
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
        var i, j, m, args, match;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch(match) {
                case "object":
                        if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                            let m = args[0];
                        } else {
                            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                        }
                        break;
                case "array":
                        let m = new \ZExcel\Shared\JAMA\Matrix();
                        call_user_func([m, "initialize"], args[0]);
                        break;
                default:
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    m->set(i, j, m->get(i, j) + this->a[i][j]);
                }
            }
            
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j, validValues, value;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let validValues = true;
                    let value = m->get(i, j);
                    
                    if ((is_string(this->a[i][j])) && (strlen(this->a[i][j]) > 0) && (!is_numeric(this->a[i][j]))) {
                        let this->a[i][j] = trim(this->a[i][j], "\"");
                        let this->a[i][j] = \ZExcel\Shared\Stringg::convertToNumberIfFraction(this->a[i][j]);
                        let validValues = validValues && (this->a[i][j] !== false);
                    }
                    
                    if ((is_string(value)) && (strlen(value) > 0) && (!is_numeric(value))) {
                        let value = trim(value, "\"");
                        let value = \ZExcel\Shared\Stringg::convertToNumberIfFraction(value);
                        let validValues = validValues && (value !== false);
                    }
                    
                    if (validValues) {
                        let this->a[i][j] = this->a[i][j] + value;
                    } else {
                        let this->a[i][j] = \ZExcel\Calculation\Functions::NaN();
                    }
                }
            }
            
            return this;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared_JAMA_Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initilliaze"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    m->set(i, j, m->get(i, j) - this->a[i][j]);
                }
            }
            
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j, validValues, value;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared_JAMA_Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let validValues = true;
                    let value = m->get(i, j);
                    
                    if ((is_string(this->a[i][j])) && (strlen(this->a[i][j]) > 0) && (!is_numeric(this->a[i][j]))) {
                        let this->a[i][j] = trim(this->a[i][j], "\"");
                        let this->a[i][j] = \ZExcel\Shared\Stringg::convertToNumberIfFraction(this->a[i][j]);
                        let validValues = validValues && (this->a[i][j] !== false);
                    }
                    
                    if ((is_string(value)) && (strlen(value) > 0) && (!is_numeric(value))) {
                        let value = trim(value, "\"");
                        let value = \ZExcel\Shared\Stringg::convertToNumberIfFraction(value);
                        let validValues = validValues && (value !== false);
                    }
                    
                    if (validValues) {
                        let this->a[i][j] = this->a[i][j] - value;
                    } else {
                        let this->a[i][j] = \ZExcel\Calculation\Functions::NaN();
                    }
                }
            }
            
            return this;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared_JAMA_Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m + 1) {
                for j in range(0, this->n - 1) {
                    m->set(i, j, m->get(i, j) * this->a[i][j]);
                }
            
            }
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j, validValues, value;
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared_JAMA_Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let validValues = true;
                    let value = m->get(i, j);
                    
                    if ((is_string(this->a[i][j])) && (strlen(this->a[i][j]) > 0) && (!is_numeric(this->a[i][j]))) {
                        let this->a[i][j] = trim(this->a[i][j], "\"");
                        let this->a[i][j] = \ZExcel\Shared\Stringg::convertToNumberIfFraction(this->a[i][j]);
                        let validValues = validValues && (this->a[i][j] !== false);
                    }
                    
                    if ((is_string(value)) && (strlen(value) > 0) && (!is_numeric(value))) {
                        let value = trim(value, "\"");
                        let value = \ZExcel\Shared\Stringg::convertToNumberIfFraction(value);
                        let validValues = validValues && (value !== false);
                    }
                    
                    if (validValues) {
                        let this->a[i][j] = this->a[i][j] * value;
                    } else {
                        let this->a[i][j] = \ZExcel\Calculation\Functions::NaN();
                    }
                }
            }
            
            return this;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j, validValues, value;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared_JAMA_Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let validValues = true;
                    let value = m->get(i, j);
                    
                    if ((is_string(this->a[i][j])) && (strlen(this->a[i][j]) > 0) && (!is_numeric(this->a[i][j]))) {
                        let this->a[i][j] = trim(this->a[i][j], "\"");
                        let this->a[i][j] = \ZExcel\Shared\Stringg::convertToNumberIfFraction(this->a[i][j]);
                        let validValues = validValues && (this->a[i][j] !== false);
                    }
                    
                    if ((is_string(value)) && (strlen(value) > 0) && (!is_numeric(value))) {
                        let value = trim(value, "\"");
                        let value = \ZExcel\Shared\Stringg::convertToNumberIfFraction(value);
                        let validValues = validValues && (value !== false);
                    }
                    
                    if (validValues) {
                        if (value == 0) {
                            //    Trap for Divide by Zero error
                            m->set(i, j, "#DIV/0!");
                        } else {
                            m->set(i, j, this->a[i][j] / value);
                        }
                    } else {
                        m->set(i, j, \ZExcel\Calculation\Functions::NaN());
                    }
                }
            }
            
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let this->a[i][j] = this->a[i][j] / m->get(i, j);
                }
            }
            
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    m->set(i, j, m->get(i, j) / this->a[i][j]);
                }
            }
            
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let this->a[i][j] = m->get(i, j) / this->a[i][j];
                }
            }
            return m;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
                case "double":
                    let c = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([c, "initialize"], this->m, this->n);
                    
                    for i in range(0, c->m - 1) {
                        for j in range(0, c->n - 1) {
                            let c->a[i][j] = args[0] * this->a[i][j];
                        }
                    }
                    
                    return c;
                case "float":
                    let c = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([c, "initialize"], this->a);
                    
                    for i in range(0, c->m - 1) {
                        for j in range(0, c->n - 1) {
                            let c->a[i][j] *= args[0];
                        }
                    }
                    
                    return c;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
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
        var args, match, m, i, j, validValues, value;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                    break;
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let validValues = true;
                    let value = m->get(i, j);
                    
                    if ((is_string(this->a[i][j])) && (strlen(this->a[i][j]) > 0) && (!is_numeric(this->a[i][j]))) {
                        let this->a[i][j] = trim(this->a[i][j], "\"");
                        let this->a[i][j] = \ZExcel\Shared\Stringg::convertToNumberIfFraction(this->a[i][j]);
                        let validValues = validValues && (this->a[i][j] !== false);
                    }
                    
                    if ((is_string(value)) && (strlen(value) > 0) && (!is_numeric(value))) {
                        let value = trim(value, "\"");
                        let value = \ZExcel\Shared\Stringg::convertToNumberIfFraction(value);
                        let validValues = validValues && (value !== false);
                    }
                    
                    if (validValues) {
                        let this->a[i][j] = pow(this->a[i][j], value);
                    } else {
                        let this->a[i][j] = \ZExcel\Calculation\Functions::NaN();
                    }
                }
            }
            
            return this;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
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
        var args, match, m, i, j;
        
        if (func_num_args() > 0) {
            let args = func_get_args();
            let match = implode(",", array_map("gettype", args));

            switch (match) {
                case "object":
                    if (args[0] instanceof \ZExcel\Shared\JAMA\Matrix) {
                        let m = args[0];
                    } else {
                        throw new \ZExcel\Calculation\Exception(self::JAMAError(self::ARGUMENT_TYPE_EXCEPTION));
                    }
                case "array":
                    let m = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([m, "initialize"], args[0]);
                    break;
                default:
                    throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
            }
            
            this->checkMatrixDimensions(m);
            
            for i in range(0, this->m - 1) {
                for j in range(0, this->n - 1) {
                    let this->a[i][j] = trim(this->a[i][j], "\"") . trim(m->get(i, j), "\"");
                }
            }
            
            return this;
        } else {
            throw new \ZExcel\Calculation\Exception(self::JAMAError(self::POLYMORPHIC_ARGUMENT_EXCEPTION));
        }
    }

    /**
     *    Solve A*X = B.
     *
     *    @param Matrix B Right hand side
     *    @return Matrix ... Solution if A is square, least squares solution otherwise
     */
    public function solve(var b)
    {
        var comp;
        
        if (this->m == this->n) {
            let comp = new \ZExcel\Shared\JAMA\LudeComposition(this);
            return comp->solve(b);
        } else {
            let comp = new \ZExcel\Shared\JAMA\QRDecomposition(this);
            return comp->solve(b);
        }
    }

    /**
     *    Matrix inverse or pseudoinverse.
     *
     *    @return Matrix ... Inverse(A) if A is square, pseudoinverse otherwise.
     */
    public function inverse()
    {
        return this->solve(this->identity(this->m, this->m));
    }

    /**
     *    det
     *
     *    Calculate determinant
     *    @return float Determinant
     */
    public function det()
    {
        var l;
        
        let l = new \ZExcel\Shared\JAMA\ludeComposition(this);
        
        return l->det();
    }
    
    public static function hypo(double a, double b) {
        double r;
        
        if (abs(a) > abs(b)) {
            let r = b / a;
            let r = abs(a) * sqrt(1 + r * r);
        } else {
            if (b != 0) {
                let r = a / b;
                let r = abs(b) * sqrt(1 + r * r);
            } else {
                let r = 0.0;
            }
        }
        
        return r;
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
            } else {
                if (isset(self::errors["EN"][errorNumber])) {
                    return self::errors["EN"][errorNumber];
                }
            }
        }

        return ("Invalid argument to JAMAError()");
    }
}
