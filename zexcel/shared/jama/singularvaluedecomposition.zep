namespace ZExcel\Shared\JAMA;

class SingularValueDecomposition 
{
    /**
     *    Internal storage of U.
     *    @var array
     */
    private u = [];

    /**
     *    Internal storage of V.
     *    @var array
     */
    private v = [];

    /**
     *    Internal storage of singular values.
     *    @var array
     */
    private s = [];

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
     *    Construct the singular value decomposition
     *
     *    Derived from LINPACK code.
     *
     *    @param a Rectangular matrix
     *    @return Structure to access U, S and V.
     */
    public function __construct(var Arg)
    {
        var a, nu, e, work, wantu, wantv, nct, nrt,i, j, k,
            p, pp, t, iter, eps, kase, ks, f, cs, sn, scale,
            sp, spm1, epm1, sk, ek, b, c, shift, g;
        
        // Initialize.
        let a = Arg->getArrayCopy();
        let this->m = Arg->getRowDimension();
        let this->n = Arg->getColumnDimension();
        let nu      = min(this->m, this->n);
        let e       = [];
        let work    = [];
        let wantu   = true;
        let wantv   = true;
        let nct = min(this->m - 1, this->n);
        let nrt = max(0, min(this->n - 2, this->m));

        // Reduce A to bidiagonal form, storing the diagonal elements
        // in s and the super-diagonal elements in e.
        for k in range(0, max(nct, nrt) - 1) {

            if (k < nct) {
                // Compute the transformation for the k-th column and
                // place the k-th diagonal in s[k].
                // Compute 2-norm of k-th column without under/overflow.
                let this->s[k] = 0;
                
                for i in range(k, this->m - 1) {
                    let this->s[k] = \ZExcel\Shared\JAMA\Matrix::hypo(this->s[k], a[i][k]);
                }
                
                if (this->s[k] != 0.0) {
                    if (a[k][k] < 0.0) {
                        let this->s[k] = -this->s[k];
                    }
                    
                    for i in range(k, this->m - 1) {
                        let a[i][k] = a[i][k] / this->s[k];
                    }
                    
                    let a[k][k] = a[k][k] + 1.0;
                }
                
                let this->s[k] = -this->s[k];
            }

            for j in range(k + 1, this->n - 1) {
                if ((k < nct) & (this->s[k] != 0.0)) {
                    // Apply the transformation.
                    let t = 0;
                    
                    for i in range(k, this->m - 1) {
                        let t = t + (a[i][k] * a[i][j]);
                    }
                    
                    let t = -t / a[k][k];
                    
                    for i in range(k, this->m - 1) {
                        let a[i][j] = a[i][j] + (t * a[i][k]);
                    }
                    // Place the k-th row of A into e for the
                    // subsequent calculation of the row transformation.
                    let e[j] = a[k][j];
                }
            }

            if (wantu && (k < nct)) {
                // Place the transformation in U for subsequent back
                // multiplication.
                for i in range(k, this->m - 1) {
                    let this->u[i][k] = a[i][k];
                }
            }

            if (k < nrt) {
                // Compute the k-th row transformation and place the
                // k-th super-diagonal in e[k].
                // Compute 2-norm without under/overflow.
                let e[k] = 0;
                
                for i in range(k + 1, this->n - 1) {
                    let e[k] = \ZExcel\Shared\JAMA\Matrix::hypo(e[k], e[i]);
                }
                
                if (e[k] != 0.0) {
                    if (e[k + 1] < 0.0) {
                        let e[k] = -e[k];
                    }
                    
                    for i in range(k + 1, this->n - 1) {
                        let e[i] = e[i] / e[k];
                    }
                    
                    let e[k + 1] = e[k + 1] + 1.0;
                }
                
                let e[k] = -e[k];
                
                if ((k + 1 < this->m) && (e[k] != 0.0)) {
                    // Apply the transformation.
                    for i in range(k + 1, this->m - 1) {
                        let work[i] = 0.0;
                    }
                    
                    for j in range(k + 1, this->n - 1) {
                        for i in range(k + 1, this->m - 1) {
                            let work[i] = work[i] + (e[j] * a[i][j]);
                        }
                    }
                    
                    for j in range(k + 1, this->n - 1) {
                        let t = -e[j] / e[k + 1];
                        
                        for i in range(k + 1, this->m - 1) {
                            let a[i][j] = a[i][j] + (t * work[i]);
                        }
                    }
                }
                if (wantv) {
                    // Place the transformation in V for subsequent
                    // back multiplication.
                    for i in range(k + 1, this->n - 1) {
                        let this->v[i][k] = e[i];
                    }
                }
            }
        }

        // Set up the final bidiagonal matrix or order p.
        let p = min(this->n, this->m + 1);
        
        if (nct < this->n) {
            let this->s[nct] = a[nct][nct];
        }
        
        if (this->m < p) {
            let this->s[p - 1] = 0.0;
        }
        
        if (nrt + 1 < p) {
            let e[nrt] = a[nrt][p - 1];
        }
        
        let e[p - 1] = 0.0;
        // If required, generate U.
        if (wantu) {
            for j in range(nct, nu - 1) {
                for i in range(0, this->m - 1) {
                    let this->u[i][j] = 0.0;
                }
                
                let this->u[j][j] = 1.0;
            }
            
            for k in reverse range(0, nct - 1) {
                if (this->s[k] != 0.0) {
                    for j in range(k + 1, nu - 1) {
                        let t = 0;
                        
                        for i in range(k, this->m - 1) {
                            let t = t + (this->u[i][k] * this->u[i][j]);
                        }
                        
                        let t = -t / this->u[k][k];
                        
                        for i in range(k, this->m - 1) {
                            let this->u[i][j] = this->u[i][j] + (t * this->u[i][k]);
                        }
                    }
                    
                    for i in range(k, this->m - 1) {
                        let this->u[i][k] = -this->u[i][k];
                    }
                    
                    let this->u[k][k] = 1.0 + this->u[k][k];
                    
                    for i in range(0, k - 2) {
                        let this->u[i][k] = 0.0;
                    }
                } else {
                    for i in range(0, this->m - 1) {
                        let this->u[i][k] = 0.0;
                    }
                    
                    let this->u[k][k] = 1.0;
                }
            }
        }

        // If required, generate V.
        if (wantv) {
            for k in reverse range(0, this->n - 1) {
                if ((k < nrt) && (e[k] != 0.0)) {
                    for j in range(k + 1, nu - 1) {
                        let t = 0;
                        
                        for i in range(k + 1, this->n - 1) {
                            let t = t + (this->v[i][k]* this->v[i][j]);
                        }
                        
                        let t = -t / this->v[k + 1][k];
                        
                        for i in range(k + 1, this->n - 1) {
                            let this->v[i][j] = this->v[i][j] + (t * this->v[i][k]);
                        }
                    }
                }
                for i in range(0, this->n - 1) {
                    let this->v[i][k] = 0.0;
                }
                
                let this->v[k][k] = 1.0;
            }
        }

        // Main iteration loop for the singular values.
        let pp   = p - 1;
        let iter = 0;
        let eps  = pow(2.0, -52.0);

        while (p > 0) {
            // Here is where a test for too many iterations would go.
            // This section of the program inspects for negligible
            // elements in the s and e arrays.  On completion the
            // variables kase and k are set as follows:
            // kase = 1  if s(p) and e[k - 1] are negligible and k<p
            // kase = 2  if s(k) is negligible and k<p
            // kase = 3  if e[k - 1] is negligible, k<p, and
            //           s(k), ..., s(p) are not negligible (qr step).
            // kase = 4  if e(p - 1) is negligible (convergence).
            for k in reverse range(-1, p - 2) {
                if (k == -1) {
                    break;
                }
                
                if (abs(e[k]) <= eps * (abs(this->s[k]) + abs(this->s[k + 1]))) {
                    let e[k] = 0.0;
                    break;
                }
            }
            if (k == p - 2) {
                let kase = 4;
            } else {
                for ks in reverse range(k, p - 1) {
                    if (ks == k) {
                        break;
                    }
                    
                    let t = (ks != p ? abs(e[ks]) : 0) + (ks != k + 1 ? abs(e[ks - 1]) : 0);
                    
                    if (abs(this->s[ks]) <= eps * t)  {
                        let this->s[ks] = 0.0;
                        break;
                    }
                }
                if (ks == k) {
                    let kase = 3;
                } else {
                    if (ks == p - 1) {
                        let kase = 1;
                    } else {
                        let kase = 2;
                        let k = ks;
                    }
                }
            }
            
            let k = k + 1;

            // Perform the task indicated by kase.
            switch (kase) {
                // Deflate negligible s(p).
                case 1:
                        let f = e[p - 2];
                        let e[p - 2] = 0.0;
                        
                        for j in reverse range(k, p - 2) {
                            let t  = \ZExcel\Shared\JAMA\Matrix::hypo(this->s[j],f);
                            let cs = this->s[j] / t;
                            let sn = f / t;
                            let this->s[j] = t;
                            
                            if (j != k) {
                                let f = -sn * e[j - 1];
                                let e[j - 1] = cs * e[j - 1];
                            }
                            
                            if (wantv) {
                                for i in range(0, this->n - 1) {
                                    let t = cs * this->v[i][j] + sn * this->v[i][p - 1];
                                    let this->v[i][p - 1] = -sn * this->v[i][j] + cs * this->v[i][p - 1];
                                    let this->v[i][j] = t;
                                }
                            }
                        }
                        break;
                // Split at negligible s(k).
                case 2:
                        let f = e[k - 1];
                        let e[k - 1] = 0.0;
                        
                        for j in range(k, p - 1) {
                            let t = \ZExcel\Shared\JAMA\Matrix::hypo(this->s[j], f);
                            let cs = this->s[j] / t;
                            let sn = f / t;
                            let this->s[j] = t;
                            let f = -sn * e[j];
                            let e[j] = cs * e[j];
                            
                            if (wantu) {
                                for i in range(0, this->m - 1) {
                                    let t = cs * this->u[i][j] + sn * this->u[i][k - 1];
                                    let this->u[i][k - 1] = -sn * this->u[i][j] + cs * this->u[i][k - 1];
                                    let this->u[i][j] = t;
                                }
                            }
                        }
                        break;
                // Perform one qr step.
                case 3:
                        // Calculate the shift.
                        let scale = max(max(max(max(
                                    abs(this->s[p - 1]),abs(this->s[p - 2])),abs(e[p - 2])),
                                    abs(this->s[k])), abs(e[k]));
                        let sp   = this->s[p - 1] / scale;
                        let spm1 = this->s[p - 2] / scale;
                        let epm1 = e[p - 2] / scale;
                        let sk   = this->s[k] / scale;
                        let ek   = e[k] / scale;
                        let b    = ((spm1 + sp) * (spm1 - sp) + epm1 * epm1) / 2.0;
                        let c    = (sp * epm1) * (sp * epm1);
                        let shift = 0.0;
                        
                        if ((b != 0.0) || (c != 0.0)) {
                            let shift = sqrt(b * b + c);
                            if (b < 0.0) {
                                let shift = -shift;
                            }
                            let shift = c / (b + shift);
                        }
                        
                        let f = (sk + sp) * (sk - sp) + shift;
                        let g = sk * ek;
                        
                        // Chase zeros.
                        for j in range(k, p - 2) {
                            let t  = \ZExcel\Shared\JAMA\Matrix::hypo(f,g);
                            let cs = f/t;
                            let sn = g/t;
                            
                            if (j != k) {
                                let e[j - 1] = t;
                            }
                            
                            let f = cs * this->s[j] + sn * e[j];
                            let e[j] = cs * e[j] - sn * this->s[j];
                            let g = sn * this->s[j + 1];
                            let this->s[j + 1] = cs * this->s[j + 1];
                            
                            if (wantv) {
                                for i in range(0, this->n - 1) {
                                    let t = cs * this->v[i][j] + sn * this->v[i][j + 1];
                                    let this->v[i][j + 1] = -sn * this->v[i][j] + cs * this->v[i][j + 1];
                                    let this->v[i][j] = t;
                                }
                            }
                            
                            let t = \ZExcel\Shared\JAMA\Matrix::hypo(f,g);
                            let cs = f/t;
                            let sn = g/t;
                            let this->s[j] = t;
                            let f = cs * e[j] + sn * this->s[j + 1];
                            let this->s[j + 1] = -sn * e[j] + cs * this->s[j + 1];
                            let g = sn * e[j + 1];
                            let e[j + 1] = cs * e[j + 1];
                            
                            if (wantu && (j < this->m - 1)) {
                                for i in range(0, this->m - 1) {
                                    let t = cs * this->u[i][j] + sn * this->u[i][j + 1];
                                    let this->u[i][j + 1] = -sn * this->u[i][j] + cs * this->u[i][j + 1];
                                    let this->u[i][j] = t;
                                }
                            }
                        }
                        
                        let e[p - 2] = f;
                        let iter = iter + 1;
                        break;
                // Convergence.
                case 4:
                        // Make the singular values positive.
                        if (this->s[k] <= 0.0) {
                            let this->s[k] = (this->s[k] < 0.0 ? -this->s[k] : 0.0);
                            
                            if (wantv) {
                                for i in range(0, pp) {
                                    let this->v[i][k] = -this->v[i][k];
                                }
                            }
                        }
                        // Order the singular values.
                        while (k < pp) {
                            if (this->s[k] >= this->s[k + 1]) {
                                break;
                            }
                            
                            let t = this->s[k];
                            let this->s[k] = this->s[k + 1];
                            let this->s[k + 1] = t;
                            
                            if (wantv && (k < this->n - 1)) {
                                for i in range(0, this->n - 1) {
                                    let t = this->v[i][k + 1];
                                    let this->v[i][k + 1] = this->v[i][k];
                                    let this->v[i][k] = t;
                                }
                            }
                            
                            if (wantu && (k < this->m - 1)) {
                                for i in range(0, this->m - 1) {
                                    let t = this->u[i][k + 1];
                                    let this->u[i][k + 1] = this->u[i][k];
                                    let this->u[i][k] = t;
                                }
                            }
                            
                            let k = k + 1;
                        }
                        
                        let iter = 0;
                        let p = p - 1;
                        break;
            } // end switch
        } // end while

    } // end constructor


    /**
     *    Return the left singular vectors
     *
     *    @access public
     *    @return U
     */
    public function getU()
    {
        var jm;
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], this->u, this->m, min(this->m + 1, this->n));
        
        return jm;
    }


    /**
     *    Return the right singular vectors
     *
     *    @access public
     *    @return V
     */
    public function getV()
    {
        var jm;
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], this->v);
        
        return jm;
    }


    /**
     *    Return the one-dimensional array of singular values
     *
     *    @access public
     *    @return diagonal of S.
     */
    public function getSingularValues() {
        return this->s;
    }


    /**
     *    Return the diagonal matrix of singular values
     *
     *    @access public
     *    @return S
     */
    public function getS() {
        var i, j, jm, s;
        
        let s = [];
        for i in range(0, this->n - 1) {
            let s[i] = [];
            for j in range(0, this->n - 1) {
                let s[i][j] = 0.0;
            }
            
            let s[i][i] = this->s[i];
        }
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], s);
        
        return jm;
    }


    /**
     *    Two norm
     *
     *    @access public
     *    @return max(S)
     */
    public function norm2() {
        return this->s[0];
    }


    /**
     *    Two norm condition number
     *
     *    @access public
     *    @return max(S)/min(S)
     */
    public function cond() {
        return this->s[0] / this->s[min(this->m, this->n) - 1];
    }


    /**
     *    Effective numerical matrix rank
     *
     *    @access public
     *    @return Number of nonnegligible singular values.
     */
    public function rank() {
        var eps, tol, r, i;
        
        let eps = pow(2.0, -52.0);
        let tol = max(this->m, this->n) * this->s[0] * eps;
        let r = 0;
        
        for i in range(0, count(this->s) - 1) {
            if (this->s[i] > tol) {
                let r = r + 1;
            }
        }
        
        return r;
    }
}
