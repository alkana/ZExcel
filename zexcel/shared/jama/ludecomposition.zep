namespace ZExcel\Shared\JaMa;

class ludeComposition
{
    const MATRIX_SINGULAR_EXCEPTION = "Can only perform operation on singular matrix.";
    
    const MATRIX_SQUARE_EXCEPTION   = "Mismatched Row dimension";

    /**
     *    Decomposition storage
     *    @var array
     */
    private lu = [];

    /**
     *    Row dimension.
     *    @var int
     */
    private m;

    /**
     *    Column dimension.
     *    @var int
     */
    private n;

    /**
     *    Pivot sign.
     *    @var int
     */
    private pivsign;

    /**
     *    Internal storage of pivot vector.
     *    @var array
     */
    private piv = [];


    /**
     *    lu Decomposition constructor.
     *
     *    @param a Rectangular matrix
     *    @return Structure to access l, u and piv.
     */
    public function __construct(var a)
    {
        var i, j, k, p, s, t, lurowi, lucolj, kmax;
        
        if (a instanceof \ZExcel\Shared\JaMa\Matrix) {
            // use a "left-looking", dot-product, Crout/Doolittle algorithm.
            let this->lu = a->getarray();
            let this->m  = a->getRowDimension();
            let this->n  = a->getColumnDimension();
            
            for i in range(0, this->m - 1) {
                let this->piv[i] = i;
            }
            
            let this->pivsign = 1;
            let lurowi = [];
            let lucolj = [];

            // Outer loop.
            for j in range(0, this->n - 1) {
                // Make a copy of the j-th column to localize references.
                for i in range(0, this->m - 1) {
                    let lucolj[i] = this->lu[i][j]; // @FIXME set by reference in PHPExcel
                }
                // apply previous transformations.
                for i in range(0, this->m - 1) {
                    let lurowi = this->lu[i];
                    // Most of the time is spent in the following dot product.
                    let kmax = min(i,j);
                    let s = 0.0;
                    
                    for k in range(0, kmax - 1) {
                        let s = s + (lurowi[k] * lucolj[k]);
                    }
                    
                    let lucolj[i] = (double) lucolj[i] - s;
                    let this->lu[i][j] = lucolj[i]; // back reference
                    let lurowi[j] = lucolj[i];
                }
                
                // Find pivot and exchange if necessary.
                let p = j;
                
                for i in range(j + 1, this->m - 1) {
                    if (abs(lucolj[i]) > abs(lucolj[p])) {
                        let p = i;
                    }
                }
                
                if (p != j) {
                    for k in range(0, this->n - 1) {
                        let t = this->lu[p][k];
                        let this->lu[p][k] = this->lu[j][k];
                        let this->lu[j][k] = t;
                    }
                    
                    let k = this->piv[p];
                    let this->piv[p] = this->piv[j];
                    let this->piv[j] = k;
                    let this->pivsign = this->pivsign * -1;
                }
                
                // Compute multipliers.
                if ((j < this->m) && (this->lu[j][j] != 0.0)) {
                    for i in range(j + 1, this->m - 1) {
                        let this->lu[i][j] = this->lu[i][j] * (1 / this->lu[j][j]);
                    }
                }
            }
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JaMa\Matrix::JaMaError(\ZExcel\Shared\JaMa\Matrix::ARGUMENT_TYPE_EXCEPTION));
        }
    }    //    function __construct()


    /**
     *    Get lower triangular factor.
     *
     *    @return array lower triangular factor
     */
    public function getL()
    {
        var i, j, l, jm;
        
        for i in range(0, this->m - 1) {
            for j in range(0, this->n - 1) {
                if (i > j) {
                    let l[i][j] = this->lu[i][j];
                } elseif (i == j) {
                    let l[i][j] = 1.0;
                } else {
                    let l[i][j] = 0.0;
                }
            }
        }
        
        let jm = new \ZExcel\Shared\JaMa\Matrix();
        call_user_func([jm, "initialize"], l);
        
        return jm;
    }    //    function getL()


    /**
     *    Get upper triangular factor.
     *
     *    @return array upper triangular factor
     */
    public function getU()
    {
        var i, j, u, jm;
        
        for i in range(0, this->n - 1) {
            for j in range(0, this->n - 1) {
                if (i <= j) {
                    let u[i][j] = this->lu[i][j];
                } else {
                    let u[i][j] = 0.0;
                }
            }
        }
        
        let jm = new \ZExcel\Shared\JaMa\Matrix();
        call_user_func([jm, "initialize"], u);
        
        return jm;
    }    //    function getU()


    /**
     *    Return pivot permutation vector.
     *
     *    @return array Pivot vector
     */
    public function getPivot()
    {
        return this->piv;
    }    //    function getPivot()


    /**
     *    alias for getPivot
     *
     *    @see getPivot
     */
    public function getDoublePivot()
    {
        return this->getPivot();
    }    //    function getDoublePivot()


    /**
     *    Is the matrix nonsingular?
     *
     *    @return true if u, and hence a, is nonsingular.
     */
    public function isNonsingular() -> boolean
    {
        var j;
        
        for j in range(0, this->n - 1) {
            if (this->lu[j][j] == 0) {
                return false;
            }
        }
        
        return true;
    }    //    function isNonsingular()


    /**
     *    Count determinants
     *
     *    @return array d matrix deterninat
     */
    public function det()
    {
        var d, j;
        
        if (this->m == this->n) {
            let d = (double) this->pivsign;
            
            for j in range(0, this->n - 1) {
                let d = d * (double) this->lu[j][j];
            }
            
            return d;
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JaMa\Matrix::JaMaError(\ZExcel\Shared\JaMa\Matrix::MATRIX_DIMENSION_EXCEPTION));
        }
    }    //    function det()


    /**
     *    Solve a*x = b
     *
     *    @param  b  a Matrix with as many rows as a and any number of columns.
     *    @return  x so that l*U*x = b(piv,:)
     *    @\ZExcel\Calculation\Exception  IllegalargumentException Matrix row dimensions must agree.
     *    @\ZExcel\Calculation\Exception  RuntimeException  Matrix is singular.
     */
    public function solve(b)
    {
        var nx, x, k, i, j;
        
        if (b->getRowDimension() == this->m) {
            if (this->isNonsingular()) {
                // Copy right hand side with pivoting
                let nx = b->getColumnDimension();
                let x  = b->getMatrix(this->piv, 0, nx - 1);
                
                // Solve l*Y = b(piv,:)
                for k in range(0, this->n - 1) {
                    for i in range(k + 1, this->n - 1) {
                        for j in range(0, nx - 1) {
                            let x->a[i][j] = x->a[i][j] - (x->a[k][j] * this->lu[i][k]);
                        }
                    }
                }
                
                // Solve u*x = Y;
                for k in reverse range(0, this->n - 1) {
                    for j in range(0, nx - 1) {
                        let x->a[k][j] = x->a[k][j] / this->lu[k][k];
                    }
                    
                    for i in range(0, k - 1) {
                        for j in range(0, nx - 1) {
                            let x->a[i][j] = x->a[i][j] - (x->a[k][j] * this->lu[i][k]);
                        }
                    }
                }
                
                return x;
            } else {
                throw new \ZExcel\Calculation\Exception(self::MATRIX_SINGULAR_EXCEPTION);
            }
        } else {
            throw new \ZExcel\Calculation\Exception(self::MATRIX_SQUARE_EXCEPTION);
        }
    }
}
