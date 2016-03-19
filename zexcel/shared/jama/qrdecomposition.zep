namespace ZExcel\Shared\JAMA;

class QRDecomposition 
{
    const MatrixRankException = "Can only perform operation on full-rank matrix.";

    /**
     *    Array for internal storage of decomposition.
     *    @var array
     */
    private qr = [];

    /**
     *    Row dimension.
     *    @var integer
     */
    private m;

    /**
    *    Column dimension.
    *    @var integer
    */
    private n;

    /**
     *    Array for internal storage of diagonal of R.
     *    @var  array
     */
    private rdiag = [];


    /**
     *    QR Decomposition computed by Householder reflections.
     *
     *    @param matrix A Rectangular matrix
     *    @return Structure to access R and the Householder vectors and compute Q.
     */
    public function __construct(a)
    {
        var nrm, i, j, k, s;
        
        if(a instanceof \ZExcel\Shared\JAMA\Matrix) {
            // Initialize.
            let this->qr = a->getArrayCopy();
            let this->m  = a->getRowDimension();
            let this->n  = a->getColumnDimension();
            
            // Main loop.
             for k in range(0, this->n - 1) {
                // Compute 2-norm of k-th column without under/overflow.
                let nrm = 0.0;
                 for i in range(k, this->m - 1) {
                    let nrm = \ZExcel\Shared\JAMA\Matrix::hypo(nrm, this->qr[i][k]);
                }
                if (nrm != 0.0) {
                    // Form k-th Householder vector.
                    if (this->qr[k][k] < 0) {
                        let nrm = -nrm;
                    }
                    
                     for i in range(k, this->m - 1) {
                        let this->qr[i][k] = this->qr[i][k] / nrm;
                    }
                    
                    let this->qr[k][k] = this->qr[k][k] + 1.0;
                    
                    // Apply transformation to remaining columns.
                    for j in range(k + 1, this->n - 1) {
                        let s = 0.0;
                         for i in range(k, this->m - 1) {
                            let s = s + (this->qr[i][k] * this->qr[i][j]);
                        }
                        
                        let s = -s / this->qr[k][k];
                        
                         for i in range(k, this->m - 1) {
                            let this->qr[i][j] = this->qr[i][j] + (s * this->qr[i][k]);
                        }
                    }
                }
                
                let this->rdiag[k] = -nrm;
            }
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::ARGUMENT_TYPE_EXCEPTION));
        }
    }    //    function __construct()


    /**
     *    Is the matrix full rank?
     *
     *    @return boolean true if R, and hence A, has full rank, else false.
     */
    public function isFullRank() -> boolean
    {
        var j;
        
        for j in range(0, this->n - 1) {
            if (this->rdiag[j] == 0) {
                return false;
            }
        }
        
        return true;
    }    //    function isFullRank()


    /**
     *    Return the Householder vectors
     *
     *    @return Matrix Lower trapezoidal matrix whose columns define the reflections
     */
    public function getH()
    {
        var i, j, h, jm;
        
        let h = [];
        for i in range(0, this->m - 1) {
            let h[i] = [];
            
            for j in range(0, this->n - 1) {
                if (i >= j) {
                    let h[i][j] = this->qr[i][j];
                } else {
                    let h[i][j] = 0.0;
                }
            }
        }
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], h);
        
        return jm;
    }    //    function getH()


    /**
     *    Return the upper triangular factor
     *
     *    @return Matrix upper triangular factor
     */
    public function getR()
    {
        var i, j, r, jm;
        
        for i in range(0, this->n - 1) {
            let r[i] = [];
            
            for j in range(0, this->n - 1) {
                if (i < j) {
                    let r[i][j] = this->qr[i][j];
                } else {
                    if (i == j) {
                        let r[i][j] = this->rdiag[i];
                    } else {
                        let r[i][j] = 0.0;
                    }
                }
            }
        }
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], r);
        
        return jm;
    }    //    function getR()


    /**
     *    Generate and return the (economy-sized) orthogonal factor
     *
     *    @return Matrix orthogonal factor
     */
    public function getQ()
    {
        var k, q, j, i, s, jm;
        
        for k in reverse range(0, this->n - 1) {
            for i in range(0, this->m - 1) {
                let q[i][k] = 0.0;
            }
            
            let q[k][k] = 1.0;
            
            for j in range(k, this->n - 1) {
                if (this->qr[k][k] != 0) {
                    let s = 0.0;
                    
                    for i in range(k, this->m - 1) {
                        let s = s + (this->qr[i][k] * q[i][j]);
                    }
                    
                    let s = -s / this->qr[k][k];
                    
                     for i in range(k, this->m - 1) {
                        let q[i][j] = q[i][j] + (s * this->qr[i][k]);
                    }
                }
            }
        }
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], q);
        
        return jm;
    }    //    function getQ()


    /**
     *    Least squares solution of A*X = B
     *
     *    @param Matrix B A Matrix with as many rows as A and any number of columns.
     *    @return Matrix Matrix that minimizes the two norm of q*R*X-B.
     */
    public function solve(b)
    {
        var nx, x, i, j, k, s, jm;
        
        if (b->getRowDimension() == this->m) {
            if (this->isFullRank()) {
                // Copy right hand side
                let nx = b->getColumnDimension();
                let x  = b->getArrayCopy();
                
                // Compute Y = transpose(Q)*B
                 for k in range(0, this->n - 1) {
                     for j in range(0, nx - 1) {
                        let s = 0.0;
                         for i in range(k, this->m - 1) {
                            let s = s + (this->qr[i][k] * x[i][j]);
                        }
                        
                        let s = -s/this->qr[k][k];
                        
                         for i in range(k, this->m - 1) {
                            let x[i][j] = x[i][j] + (s * this->qr[i][k]);
                        }
                    }
                }
                
                // Solve R*X = Y;
                for k in reverse range(0, this->n - 1) {
                    for j in range(0, nx - 1) {
                        let x[k][j] = x[k][j] / this->rdiag[k];
                    }
                    
                     for i in range(0, k - 1) {
                         for j in range(0, nx - 1) {
                            let x[i][j] = x[i][j] - (x[k][j]* this->qr[i][k]);
                        }
                    }
                }
                
                let x = new \ZExcel\Shared\JAMA\Matrix();
                call_user_func([jm, "initialize"], x);
                
                let nx = call_user_func([x, "getMatrix"], 0, this->n - 1, 0, nx);
                
                return nx;
            } else {
                throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::MATRIX_RANK_EXCEPTION));
            }
        } else {
            throw new \ZExcel\Calculation\Exception(\ZExcel\Shared\JAMA\Matrix::JAMAError(\ZExcel\Shared\JAMA\Matrix::MATRIX_DIMENSION_EXCEPTION));
        }
    }
}
