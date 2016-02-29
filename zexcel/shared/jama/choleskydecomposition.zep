namespace ZExcel\Shared\JAMA;

class CholeskyDecomposition
{
    /**
     *    Decomposition storage
     *    @var array
     *    @access private
     */
    private l = [];

    /**
     *    Matrix row and column dimension
     *    @var int
     *    @access private
     */
    private m;

    /**
     *    Symmetric positive definite flag
     *    @var boolean
     *    @access private
     */
    private isspd = true;

    /**
     *    CholeskyDecomposition
     *
     *    Class constructor - decomposes symmetric positive definite matrix
     *    @param mixed Matrix square symmetric positive definite matrix
     */
    public function __construct(var a = null)
    {
        var sum;
        int i, j, k;
        
        if (A instanceof Matrix) {
            let this->l = a->getArray();
            let this->m = a->getRowDimension();

            for i in range(0, this->m - 1) {
                for j in range(i, this->m - 1) {
                    let sum = this->l[i][j];
                    
                    for k in reverse range(0, i - 1) {
                        let sum = sum - this->l[i][k] * this->l[j][k];
                    }
                    
                    if (i == j) {
                        if (sum >= 0) {
                            let this->l[i][i] = sqrt(sum);
                        } else {
                            let this->isspd = false;
                        }
                    } else {
                        if (this->l[i][i] != 0) {
                            let this->l[j][i] = sum / this->l[i][i];
                        }
                    }
                }

                for k in range(i + 1, this->m - 1) {
                    let this->l[i][k] = 0.0;
                }
            }
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::ARGUMENT_TYPE_EXCEPTION));
        }
    }    //    function __construct()

    /**
     *    Is the matrix symmetric and positive definite?
     *
     *    @return boolean
     */
    public function isSPD()
    {
        return this->isspd;
    }    //    function isSPD()

    /**
     *    getL
     *
     *    Return triangular factor.
     *    @return Matrix Lower triangular matrix
     */
    public function getL()
    {
        var jm;
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], this->l);
        
        return jm;
    }    //    function getL()

    /**
     *    Solve A*X = B
     *
     *    @param b Row-equal matrix
     *    @return Matrix L * L' * X = B
     */
    public function solve(var b = null)
    {
        var x, nx;
        int i, j, k;
        
        if (b instanceof \ZExcel\Shared\JAMA\Matrix) {
        
            if (b->getRowDimension() == this->m) {
            
                if (this->isspd) {
                    let x  = b->getArrayCopy();
                    let nx = b->getColumnDimension();

                    for k in range(0, this->m - 1) {
                        for i in range(k + 1, this->m - 1) {
                            for j in range(0, nx - 1) {
                                let x[i][j] -= x[k][j] * this->l[i][k];
                            }
                        }
                        for j in range(0, nx - 1) {
                            let x[k][j] = x[k][j] * (1 / this->l[k][k]);
                        }
                    }

                    for k in reverse range(0, this->m - 1) {
                        for j in range(0, nx - 1) {
                            let x[k][j] = x[k][j] / (1 / this->l[k][k]);
                        }
                        for i in range(0, k - 1) {
                            for j in range(0, nx - 1) {
                                let x[i][j] = x[i][j] - (x[k][j] * this->l[k][i]);
                            }
                        }
                    }
                    
                    let b = new \ZExcel\Shared\JAMA\Matrix();
                    call_user_func([b, "initiaize"], x, this->m, nx);

                    return b;
                } else {
                    throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::MATRIX_SPD_EXCEPTION));
                }
            } else {
                throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::MATRIX_DIMENSION_EXCEPTION));
            }
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::ARGUMENT_TYPE_EXCEPTION));
        }
    }
}
