namespace ZExcel\Shared\JAMA;

class Matrix 
{
    const POLYMORPHIC_ARGUMENT_EXCEPTION = "Invalid argument pattern for polymorphic function.";
    const ARGUMENT_TYPE_EXCEPTION        = "Invalid argument type.";
    const ARGUMENT_BOUNDS_EXCEPTION      = "Invalid argument range.";
    const MATRIX_DIMENSION_EXCEPTION     = "Matrix dimensions are not equal.";
    const ARRAY_LENGTH_EXCEPTION         = "Array length must be a multiple of m.";

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
     *    As PHP has no support for polymorphic constructors, we hack our own sort of polymorphism using func_num_args, func_get_arg, and gettype. In essence, we're just implementing a simple RTTI filter and calling the appropriate constructor.
     */
    public function __construct(var args)
    {
        if (!is_array(args)) {
            let args = [args];
        }
        
        throw new \Exception("Not implemented yet!");
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
     *    @param int $i Row position
     *    @param int $j Column position
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
     *    @param int $i0 Initial row index
     *    @param int $iF Final row index
     *    @param int $j0 Initial column index
     *    @param int $jF Final column index
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
     *    @param Matrix $B Matrix B
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
     *    @param int $i Row position
     *    @param int $j Column position
     *    @param mixed $c Int/float/double value
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
     *    @param int $m Row dimension
     *    @param int $n Column dimension
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
     *    @param int $m Row dimension
     *    @param int $n Column dimension
     *    @param mixed $c Diagonal value
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
     *    @param int $i0 Initial row index
     *    @param int $iF Final row index
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
     *    @param int $i0 Initial column index
     *    @param int $iF Final column index
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
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
     *    @param Matrix $B Matrix B
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
     *    @param mixed $B Matrix/Array
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
     *    @param Matrix $B Matrix B
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
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $n Matrix/Array/Scalar
     *    @return Matrix Product
     */
    public function times()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    power
     *
     *    A = A ^ B
     *    @param mixed $B Matrix/Array
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
     *    @param mixed $B Matrix/Array
     *    @return Matrix Sum
     */
    public function concat()
    {
        throw new \Exception("Not implemented yet!");
    }

    /**
     *    Solve A*X = B.
     *
     *    @param Matrix $B Right hand side
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
}
