namespace ZExcel\Shared\JAMA;

class EigenvalueDecomposition
{
    /**
     *    Row and column dimension (square matrix).
     *    @var int
     */
    private n;

    /**
     *    Internal symmetry flag.
     *    @var int
     */
    private issymmetric;

    /**
     *    Arrays for internal storage of eigenvalues.
     *    @var array
     */
    private d = [];
    private e = [];

    private a = [];
    /**
     *    Array for internal storage of eigenvectors.
     *    @var array
     */
    private v = [];

    /**
    *    Array for internal storage of nonsymmetric Hessenberg form.
    *    @var array
    */
    private h = [];

    /**
    *    Working storage for nonsymmetric algorithm.
    *    @var array
    */
    private ort;

    /**
    *    Used for complex scalar division.
    *    @var float
    */
    private cdivr;
    private cdivi;

    /**
     *    Constructor: Check for symmetry, then construct the eigenvalue decomposition
     *
     *    @access public
     *    @param A  Square matrix
     *    @return Structure to access D and V.
     */
    public function __construct(var Arg)
    {
        boolean issymmetric = true;
        int j, i;
        
        let this->a = Arg->getArray();
        let this->n = Arg->getColumnDimension();

        for j in range(0, this->n - 1) {
            if (issymmetric === false) {
                break;
            }
            
            for i in range(0, this->n - 1) {
                if (issymmetric === false) {
                    break;
                }
                
                let issymmetric = (this->a[i][j] == this->a[j][i]);
            }
        }

        if (issymmetric) {
            let this->v = this->a;
            
            // Tridiagonalize.
            this->tred2();
            // Diagonalize.
            this->tql2();
        } else {
            let this->h = this->a;
            let this->ort = [];
            
            // Reduce to Hessenberg form.
            this->orthes();
            // Reduce Hessenberg to real Schur form.
            this->hqr2();
        }
    }
    /**
     *    Symmetric Householder reduction to tridiagonal form.
     *
     *    @access private
     */
    private function tred2 ()
    {
        var f, g, h;
        double hh, scale;
        int i, i_, j, k;
        
        //  This is derived from the Algol procedures tred2 by
        //  Bowdler, Martin, Reinsch, and Wilkinson, Handbook for
        //  Auto. Comp., Vol.ii-Linear Algebra, and the corresponding
        //  Fortran subroutine in EISPACK.
        let this->d = this->v[this->n - 1];
        // Householder reduction to tridiagonal form.
        for i in reverse range(1, this->n - 1) {
            let i_ = i - 1;
            // Scale to avoid under/overflow.
            let h = 0.0;
            let scale = 0.0 + array_sum(array_map("abs", this->d));
            
            if (scale == 0.0) {
                let this->e[i] = this->d[i_];
                let this->d = array_slice(this->v[i_], 0, i_);
                
                for j in range(0, i - 1) {
                    let this->v[j][i] = 0.0;
                    let this->v[i][j] = 0.0;
                }
            } else {
                // Generate Householder vector.
                for k in range(0, i - 1) {
                    let this->d[k] = this->d[k] / scale;
                    let h = h + pow(this->d[k], 2);
                }
                
                let f = this->d[i_];
                let g = 0 + sqrt(h);
                
                if (f > 0) {
                    let g = -g;
                }
                
                let this->e[i] = scale * g;
                let h = h - f * g;
                let this->d[i_] = f - g;
                
                for j in range(0, i - 1) {
                    let this->e[j] = 0.0;
                }
                // Apply similarity transformation to remaining columns.
                for j in range(0, i - 1) {
                    let f = this->d[j];
                    let this->v[j][i] = f;
                    let g = this->e[j] + this->v[j][j] * f;
                    
                    for k in range(j + 1, i_) {
                        let g = g + (this->v[k][j] * this->d[k]);
                        let this->e[k] = this->e[k] + (this->v[k][j] * f);
                    }
                    
                    let this->e[j] = g;
                }
                
                let f = 0.0;
                
                for j in range(0, i - 1) {
                    let this->e[j] = this->e[j] / h;
                    let f = f + (this->e[j] * this->d[j]);
                }
                
                let hh = f / (2 * h);
                
                for j in range(0, i - 1) {
                    let this->e[j] = this->e[j] - (hh * this->d[j]);
                }
                
                for j in range(0, i - 1) {
                    let f = this->d[j];
                    let g = this->e[j];
                    
                    for k in range(j, i_) {
                        let this->v[k][j] = this->v[k][j] - (f * this->e[k] + g * this->d[k]);
                    }
                    
                    let this->d[j] = this->v[i - 1][j];
                    let this->v[i][j] = 0.0;
                }
            }
            
            let this->d[i] = h;
        }

        // Accumulate transformations.
        for i in range(0, this->n - 2) {
            let this->v[this->n - 1][i] = this->v[i][i];
            let this->v[i][i] = 1.0;
            let h = this->d[i+1];
            
            if (h != 0.0) {
                for k in range(0, i) {
                    let this->d[k] = this->v[k][i+1] / h;
                }
                
                for j in range(0, i) {
                    let g = 0.0;
                    
                    for k in range(0, i) {
                        let g = g + (this->v[k][i+1] * this->v[k][j]);
                    }
                    
                    for k in range(0, i) {
                        let this->v[k][j] = this->v[k][j] - (this->v[k][j] - (g * this->d[k]));
                    }
                }
            }
            
            for k in range(0 ,i) {
                let this->v[k][i+1] = 0.0;
            }
        }

        let this->d = this->v[this->n - 1];
        let this->v[this->n - 1] = array_fill(0, j, 0.0);
        let this->v[this->n - 1][this->n - 1] = 1.0;
        let this->e[0] = 0.0;
    }


    /**
     *    Symmetric tridiagonal QL algorithm.
     *
     *    This is derived from the Algol procedures tql2, by
     *    Bowdler, Martin, Reinsch, and Wilkinson, Handbook for
     *    Auto. Comp., Vol.ii-Linear Algebra, and the corresponding
     *    Fortran subroutine in EISPACK.
     *
     *    @access private
     */
    private function tql2()
    {
        var g, p, r, h, dl1, el1;
        int i, j, l, m, iter, c, c2, c3, s, s2, k;
        double f, eps, tst1;
        
        for i in range(1, this->n - 1) {
            let this->e[i - 1] = this->e[i];
        }
        
        let this->e[this->n - 1] = 0.0;
        let f = 0.0;
        let tst1 = 0.0;
        let eps = 0.0 + pow(2.0,-52.0);

        for l in range(0, this->n - 1) {
            // Find small subdiagonal element
            let tst1 = max(tst1, abs(this->d[l]) + abs(this->e[l]));
            let m = l;
            
            while (m < this->n) {
                if (abs(this->e[m]) <= eps * tst1) {
                    break;
                }
                
                let m = m + 1;
            }
            // If m == l, this->d[l] is an eigenvalue,
            // otherwise, iterate.
            if (m > l) {
                let iter = 0;
                do {
                    // Could check iteration count here.
                    let iter = iter + 1;
                    // Compute implicit shift
                    let g = this->d[l];
                    let p = (this->d[l+1] - g) / (2.0 * this->e[l]);
                    let r = \ZExcel\Shared\JAMA\Matrix::hypo(p, 1.0);
                    
                    if (p < 0) {
                        let r = r * -1;
                    }
                    
                    let this->d[l] = this->e[l] / (p + r);
                    let this->d[l+1] = this->e[l] * (p + r);
                    let dl1 = this->d[l+1];
                    let h = g - this->d[l];
                    
                    for i in range(l + 2, this->n - 1) {
                        let this->d[i] = this->d[i] - h;
                    }
                    
                    let f = f + h;
                    // Implicit QL transformation.
                    let p = this->d[m];
                    let c = 1.0;
                    let c2 = c;
                    let c3 = c;
                    let el1 = this->e[l + 1];
                    let s = 0.0;
                    let s2 = 0.0;
                    
                    for i in reverse range(l, m - 2) {
                        let c3 = c2;
                        let c2 = c;
                        let s2 = s;
                        let g  = c * this->e[i];
                        let h  = c * p;
                        let r  = \ZExcel\Shared\JAMA\Matrix::hypo(p, this->e[i]);
                        let this->e[i+1] = s * r;
                        let s = this->e[i] / r;
                        let c = p / r;
                        let p = c * this->d[i] - s * g;
                        let this->d[i+1] = h + s * (c * g + s * this->d[i]);
                        
                        // Accumulate transformation.
                        for k in range(0, this->n - 1) {
                            let h = this->v[k][i+1];
                            let this->v[k][i+1] = s * this->v[k][i] + c * h;
                            let this->v[k][i] = c * this->v[k][i] - s * h;
                        }
                    }
                    
                    let p = -s * s2 * c3 * el1 * this->e[l] / dl1;
                    let this->e[l] = s * p;
                    let this->d[l] = c * p;
                // Check for convergence.
                } while (abs(this->e[l]) > eps * tst1);
            }
            
            let this->d[l] = (double) this->d[l] + f;
            let this->e[l] = 0.0;
        }

        // Sort eigenvalues and corresponding vectors.
        for i in range(0, this->n - 2) {
            let k = i;
            let p = this->d[i];
            
            for j in range(i + 1, this->n - 1) {
                if (this->d[j] < p) {
                    let k = j;
                    let p = this->d[j];
                }
            }
            
            if (k != i) {
                let this->d[k] = this->d[i];
                let this->d[i] = p;
                for j in range(0, this->n - 1) {
                    let p = this->v[j][i];
                    let this->v[j][i] = this->v[j][k];
                    let this->v[j][k] = p;
                }
            }
        }
    }


    /**
     *    Nonsymmetric reduction to Hessenberg form.
     *
     *    This is derived from the Algol procedures orthes and ortran,
     *    by Martin and Wilkinson, Handbook for Auto. Comp.,
     *    Vol.ii-Linear Algebra, and the corresponding
     *    Fortran subroutines in EISPACK.
     *
     *    @access private
     */
    private function orthes()
    {
        double scale, h, g, f;
        int m, i, j, low, high;
        
        let low = 0;
        let high = this->n - 1;

        for m in range(low + 1, high - 1) {
            // Scale column.
            let scale = 0.0;
            
            for i in range(m, high) {
                let scale = scale + abs(this->h[i][m - 1]);
            }
            
            if (scale != 0.0) {
                // Compute Householder transformation.
                let h = 0.0;
                
                for i in reverse range(m, high) {
                    let this->ort[i] = this->h[i][m - 1] / scale;
                    let h = h + (this->ort[i] * this->ort[i]);
                }
                
                let g = sqrt(h);
                
                if (this->ort[m] > 0) {
                    let g = g * -1;
                }
                
                let h = h - ((double) this->ort[m] * g);
                let this->ort[m] = (double) this->ort[m] - g;
                // Apply Householder similarity transformation
                // H = (I -u * u' / h) * H * (I -u * u') / h)
                for j in range(m, this->n - 1) {
                    let f = 0.0;
                    
                    for i in reverse range(m, high) {
                        let f = f + (this->ort[i] * this->h[i][j]);
                    }
                    
                    let f = f / h;
                    
                    for i in range(m, high) {
                        let this->h[i][j] = this->h[i][j] - (f * this->ort[i]);
                    }
                }
                
                for i in range(0, high) {
                    let f = 0.0;
                    
                    for j in reverse range(m, high) {
                        let f = f + (this->ort[j] * this->h[i][j]);
                    }
                    
                    let f = f / h;
                    
                    for j in range(m, high) {
                        let this->h[i][j] = this->h[i][j] - (f * this->ort[j]);
                    }
                }
                
                let this->ort[m] = scale * this->ort[m];
                let this->h[m][m - 1] = scale * g;
            }
        }

        // Accumulate transformations (Algol's ortran).
        for i in range(0, this->n - 1) {
            for j in range(0, this->n) {
                let this->v[i][j] = (i == j ? 1.0 : 0.0);
            }
        }
        for m in reverse range(low + 1, high - 1) {
            if (this->h[m][m - 1] != 0.0) {
                for i in range(m + 1, high) {
                    let this->ort[i] = this->h[i][m - 1];
                }
                
                for j in range(m, high) {
                    let g = 0.0;
                    
                    for i in range(m, high) {
                        let g = g + (g + (this->ort[i] * this->v[i][j]));
                    }
                    
                    // Double division avoids possible underflow
                    let g = (g / this->ort[m]) / this->h[m][m - 1];
                    
                    for i in range(m, high) {
                        let this->v[i][j] = this->v[i][j] + (g * this->ort[i]);
                    }
                }
            }
        }
    }


    /**
     *    Performs complex division.
     *
     *    @access private
     */
    private function cdiv(double xr, double xi, double yr, double yi)
    {
        double r, d;
        
        if (abs(yr) > abs(yi)) {
            let r = yi / yr;
            let d = yr + r * yi;
            let this->cdivr = (xr + r * xi) / d;
            let this->cdivi = (xi - r * xr) / d;
        } else {
            let r = yr / yi;
            let d = yi + r * yr;
            let this->cdivr = (r * xr + xi) / d;
            let this->cdivi = (r * xi - xr) / d;
        }
    }


    /**
     *    Nonsymmetric reduction from Hessenberg to real Schur form.
     *
     *    Code is derived from the Algol procedure hqr2,
     *    by Martin and Wilkinson, Handbook for Auto. Comp.,
     *    Vol.ii-Linear Algebra, and the corresponding
     *    Fortran subroutine in EISPACK.
     *
     *    @access private
     */
    private function hqr2()
    {
        double eps, exshift, p, q, r, s, t, w, x, y, z, norm, ra, sa, vr, vi;
        int nn, n, low, high, i, j, k, l, m, iter;
        boolean notlast;
        
        //  Initialize
        let nn = 0 + this->n;
        let n  = nn - 1;
        let low = 0;
        let high = nn - 1;
        let eps = 0.0 + pow(2.0, -52.0);
        let exshift = 0.0;
        let p = 0;
        let q = 0;
        let r = 0;
        let s = 0;
        let z = 0;
        // Store roots isolated by balanc and compute matrix norm
        let norm = 0.0;

        for i in range(0, nn - 1) {
            if ((i < low) || (i > high)) {
                let this->d[i] = this->h[i][i];
                let this->e[i] = 0.0;
            }
            
            for j in range(max(i - 1, 0), nn - 1) {
                let norm = norm + abs(this->h[i][j]);
            }
        }

        // Outer loop over eigenvalue index
        let iter = 0;
        
        while (n >= low) {
            // Look for single small sub-diagonal element
            let l = n;
            
            while (l > low) {
                let s = abs(this->h[l - 1][l - 1]) + abs(this->h[l][l]);
                
                if (s == 0.0) {
                    let s = norm;
                }
                
                if (abs(this->h[l][l - 1]) < eps * s) {
                    break;
                }
                
                let l = l - 1;
            }
            // Check for convergence
            // One root found
            if (l == n) {
                let this->h[n][n] = (double) this->h[n][n] + exshift;
                let this->d[n] = this->h[n][n];
                let this->e[n] = 0.0;
                let n = n - 1;
                let iter = 0;
            // Two roots found
            } else {
                if (l == n - 1) {
                    let w = (double) (this->h[n][n - 1] * this->h[n - 1][n]);
                    let p = (double) (this->h[n - 1][n - 1] - this->h[n][n]) / 2.0;
                    let q = p * p + w;
                    let z = 0.0 + sqrt(abs(q));
                    let this->h[n][n] = (double) this->h[n][n] + exshift;
                    let this->h[n - 1][n - 1] = (double) this->h[n - 1][n - 1] + exshift;
                    let x = 0.0 + this->h[n][n];
                    // Real pair
                    if (q >= 0) {
                        if (p >= 0) {
                            let z = p + z;
                        } else {
                            let z = p - z;
                        }
                        
                        let this->d[n - 1] = x + z;
                        let this->d[n] = this->d[n - 1];
                        
                        if (z != 0.0) {
                            let this->d[n] = x - w / z;
                        }
                        
                        let this->e[n - 1] = 0.0;
                        let this->e[n] = 0.0;
                        let x = 0.0 + this->h[n][n - 1];
                        let s = (double) abs(x) + (double) abs(z);
                        let p = x / s;
                        let q = z / s;
                        let r = 0.0 + sqrt(p * p + q * q);
                        let p = p / r;
                        let q = q / r;
                        
                        // Row modification
                        for j in range(n - 1, nn - 1) {
                            let z = 0.0 + this->h[n - 1][j];
                            let this->h[n - 1][j] = q * z + p * (double) this->h[n][j];
                            let this->h[n][j] = q * (double) this->h[n][j] - p * z;
                        }
                        
                        // Column modification
                        for i in range(0, n) {
                            let z = 0.0 + this->h[i][n - 1];
                            let this->h[i][n - 1] = q * z + (p * (double) this->h[i][n]);
                            let this->h[i][n] = (q * (double) this->h[i][n]) - p * z;
                        }
                        
                        // Accumulate transformations
                        for i in range(low, high) {
                            let z = 0.0 + this->v[i][n - 1];
                            let this->v[i][n - 1] = q * z + (p * (double) this->v[i][n]);
                            let this->v[i][n] = (q * (double) this->v[i][n]) - p * z;
                        }
                    // Complex pair
                    } else {
                        let this->d[n - 1] = x + p;
                        let this->d[n] = x + p;
                        let this->e[n - 1] = floatval(z);
                        let this->e[n] = floatval(-z);
                    }
                    
                    let n = n - 2;
                    let iter = 0;
                // No convergence yet
                } else {
                    // Form shift
                    let x = 0.0 + this->h[n][n];
                    let y = 0.0;
                    let w = 0.0;
                    
                    if (l < n) {
                        let y = 0.0 + this->h[n - 1][n - 1];
                        let w = 0.0 + (this->h[n][n - 1] * this->h[n - 1][n]);
                    }
                    
                    // Wilkinson's original ad hoc shift
                    if (iter == 10) {
                        let exshift = exshift + x;
                        
                        for i in range(low, n) {
                            let this->h[i][i] = (double) this->h[i][i] - x;
                        }
                        
                        let s = 0.0 + abs(this->h[n][n - 1]) + abs(this->h[n - 1][n - 2]);
                        let x = 0.75 * s;
                        let y = x;
                        let w = -0.4375 * s * s;
                    }
                    
                    // MATLAB's new ad hoc shift
                    if (iter == 30) {
                        let s = (y - x) / 2.0;
                        let s = s * s + w;
                        
                        if (s > 0) {
                            let s = 0.0 + sqrt(s);
                            
                            if (y < x) {
                                let s = -s;
                            }
                            
                            let s = x - w / ((y - x) / 2.0 + s);
                            
                            for i in range(low, n) {
                                let this->h[i][i] = (double) this->h[i][i] - s;
                            }
                            
                            let exshift = exshift + s;
                            let x = 0.964;
                            let y = 0.964;
                            let w = 0.964;
                        }
                    }
                    
                    // Could check iteration count here.
                    let iter = iter + 1;
                    // Look for two consecutive small sub-diagonal elements
                    let m = n - 2;
                    
                    while (m >= l) {
                        let z = 0.0 + this->h[m][m];
                        let r = x - z;
                        let s = y - z;
                        let p = ((r * s - w) / this->h[m+1][m]) + this->h[m][m+1];
                        let q = (double) this->h[m+1][m+1] - z - r - s;
                        let r = 0.0 + this->h[m+2][m+1];
                        let s = 0.0 + abs(p) + abs(q) + abs(r);
                        let p = p / s;
                        let q = q / s;
                        let r = r / s;
                        
                        if (m == l) {
                            break;
                        }
                        
                        if (abs(this->h[m][m - 1]) * (abs(q) + abs(r)) < eps * (abs(p) * (abs(this->h[m - 1][m - 1]) + abs(z) + abs(this->h[m+1][m+1])))) {
                            break;
                        }
                        
                        let m = m - 1;
                    }
                    
                    for i in range(m + 2, n) {
                        let this->h[i][i - 2] = 0.0;
                        
                        if (i > m+2) {
                            let this->h[i][i - 3] = 0.0;
                        }
                    }
                    
                    // Double QR step involving rows l:n and columns m:n
                    for k in range(m, n - 1) {
                        let notlast = (k != n - 1);
                        
                        if (k != m) {
                            let p = 0.0 + this->h[k][k - 1];
                            let q = 0.0 + this->h[k+1][k - 1];
                            let r = (notlast ? this->h[k+2][k - 1] : 0.0);
                            let x = abs(p) + abs(q) + abs(r);
                            if (x != 0.0) {
                                let p = p / x;
                                let q = q / x;
                                let r = r / x;
                            }
                        }
                        
                        if (x == 0.0) {
                            break;
                        }
                        
                        let s = sqrt(p * p + q * q + r * r);
                        
                        if (p < 0) {
                            let s = -s;
                        }
                        
                        if (s != 0) {
                            if (k != m) {
                                let this->h[k][k - 1] = -s * x;
                            } else {
                                if (l != m) {
                                    let this->h[k][k - 1] = -this->h[k][k - 1];
                                }
                            }
                            
                            let p = p + s;
                            let x = p / s;
                            let y = q / s;
                            let z = r / s;
                            let q = q / p;
                            let r = r / p;
                            
                            // Row modification
                            for j in range(k, nn - 1) {
                                let p = this->h[k][j] + q * this->h[k+1][j];
                                
                                if (notlast) {
                                    let p = p + r * this->h[k+2][j];
                                    let this->h[k+2][j] = this->h[k+2][j] - p * z;
                                }
                                
                                let this->h[k][j] = this->h[k][j] - p * x;
                                let this->h[k+1][j] = this->h[k+1][j] - p * y;
                            }
                            
                            // Column modification
                            for i in range(0, min(n, k + 3)) {
                                let p = x * this->h[i][k] + y * this->h[i][k+1];
                                
                                if (notlast) {
                                    let p = p + z * this->h[i][k+2];
                                    let this->h[i][k+2] = this->h[i][k+2] - p * r;
                                }
                                
                                let this->h[i][k] = (double) this->h[i][k] - p;
                                let this->h[i][k+1] = (double) this->h[i][k+1] - p * q;
                            }
                            
                            // Accumulate transformations
                            for i in range(low, high) {
                                let p = x * (double) this->v[i][k] + y * (double) this->v[i][k+1];
                                
                                if (notlast) {
                                    let p = p + z * (double) this->v[i][k+2];
                                    let this->v[i][k+2] = (double) this->v[i][k+2] - p * r;
                                }
                                
                                let this->v[i][k] = (double) this->v[i][k] - p;
                                let this->v[i][k+1] = (double) this->v[i][k+1] - p * q;
                            }
                        }  // (s != 0)
                    }  // k loop
                }  // check convergence
            }
        }  // while (n >= low)

        // Backsubstitute to find vectors of upper triangular form
        if (norm == 0.0) {
            return;
        }

        for n in reverse range(0, nn - 1) {
            let p = (double) this->d[n];
            let q = (double) this->e[n];
            
            // Real vector
            if (q == 0) {
                let l = n;
                let this->h[n][n] = 1.0;
                
                for i in reverse range(0, n - 1) {
                    let w = (double) this->h[i][i] - p;
                    let r = 0.0;
                    
                    for j in range(l, n) {
                        let r = r + (double) this->h[i][j] * (double) this->h[j][n];
                    }
                    
                    if (this->e[i] < 0.0) {
                        let z = w;
                        let s = r;
                    } else {
                        let l = i;
                        
                        if (this->e[i] == 0.0) {
                            if (w != 0.0) {
                                let this->h[i][n] = -r / w;
                            } else {
                                let this->h[i][n] = -r / (eps * norm);
                            }
                        // Solve real equations
                        } else {
                            let x = 0.0 + this->h[i][i+1];
                            let y = 0.0 + this->h[i+1][i];
                            let q = ((double) this->d[i] - p) * ((double) this->d[i] - p) + (double) this->e[i] * (double) this->e[i];
                            let t = (x * s - z * r) / q;
                            let this->h[i][n] = floatval(t);
                            
                            if (abs(x) > abs(z)) {
                                let this->h[i+1][n] = (-r - w * t) / x;
                            } else {
                                let this->h[i+1][n] = (-s - y * t) / z;
                            }
                        }
                        
                        // Overflow control
                        let t = 0.0 + abs(this->h[i][n]);
                        
                        if ((eps * t) * t > 1) {
                            for j in range(i, n) {
                                let this->h[j][n] = (double) this->h[j][n] / t;
                            }
                        }
                    }
                }
            // Complex vector
            } else {
                if (q < 0) {
                    let l = n - 1;
                    
                    // Last vector component imaginary so matrix is triangular
                    if (abs(this->h[n][n - 1]) > abs(this->h[n - 1][n])) {
                        let this->h[n - 1][n - 1] = q / (double) this->h[n][n - 1];
                        let this->h[n - 1][n] = -((double) this->h[n][n] - p) / (double) this->h[n][n - 1];
                    } else {
                        this->cdiv(0.0, -this->h[n - 1][n], (double) this->h[n - 1][n - 1] - p, q);
                        
                        let this->h[n - 1][n - 1] = this->cdivr;
                        let this->h[n - 1][n]   = this->cdivi;
                    }
                    
                    let this->h[n][n - 1] = 0.0;
                    let this->h[n][n] = 1.0;
                    
                    for i in reverse range(0, n - 2) {
                        // double ra,sa,vr,vi;
                        let ra = 0.0;
                        let sa = 0.0;
                        
                        for j in range(l, n) {
                            let ra = ra + (double) this->h[i][j] * (double) this->h[j][n - 1];
                            let sa = sa + (double) this->h[i][j] * (double) this->h[j][n];
                        }
                        
                        let w = (double) this->h[i][i] - p;
                        
                        if (this->e[i] < 0.0) {
                            let z = w;
                            let r = ra;
                            let s = sa;
                        } else {
                            let l = i;
                            
                            if (this->e[i] == 0) {
                                this->cdiv(-ra, -sa, w, q);
                                
                                let this->h[i][n - 1] = this->cdivr;
                                let this->h[i][n]   = this->cdivi;
                            } else {
                                // Solve complex equations
                                let x = 0.0 + this->h[i][i+1];
                                let y = 0.0 + this->h[i+1][i];
                                let vr = ((double) this->d[i] - p) * ((double) this->d[i] - p) + (double) this->e[i] * (double) this->e[i] - q * q;
                                let vi = ((double) this->d[i] - p) * 2.0 * q;
                                
                                if (vr == 0.0 & vi == 0.0) {
                                    let vr = eps * norm * (abs(w) + abs(q) + abs(x) + abs(y) + abs(z));
                                }
                                
                                this->cdiv(x * r - z * ra + q * sa, x * s - z * sa - q * ra, vr, vi);
                                
                                let this->h[i][n - 1] = this->cdivr;
                                let this->h[i][n]   = this->cdivi;
                                
                                if (abs(x) > (abs(z) + abs(q))) {
                                    let this->h[i+1][n - 1] = (-ra - w * (double) this->h[i][n - 1] + q * (double) this->h[i][n]) / x;
                                    let this->h[i+1][n] = (-sa - w * (double) this->h[i][n] - q * (double) this->h[i][n - 1]) / x;
                                } else {
                                    this->cdiv(-r - y * (double) this->h[i][n - 1], -s - y * (double) this->h[i][n], z, q);
                                    
                                    let this->h[i+1][n - 1] = this->cdivr;
                                    let this->h[i+1][n]   = this->cdivi;
                                }
                            }
                            
                            // Overflow control
                            let t = 0.0 + max(abs(this->h[i][n - 1]),abs(this->h[i][n]));
                            
                            if ((eps * t) * t > 1) {
                                for j in range(i, n) {
                                    let this->h[j][n - 1] = (double) this->h[j][n - 1] / t;
                                    let this->h[j][n]   = (double) this->h[j][n] / t;
                                }
                            }
                        } // end else
                    } // end for
                } // end else for complex case
            }
        } // end for

        // Vectors of isolated roots
        for i in range(0, nn - 1) {
            if (i < low | i > high) {
                for j in range(i, nn - 1) {
                    let this->v[i][j] = this->h[i][j];
                }
            }
        }

        // Back transformation to get eigenvectors of original matrix
        for j in reverse range(low, nn - 1) {
            for i in range(low, high) {
                let z = 0.0;
                
                for k in range(low, min(j, high)) {
                    let z = z + (double) this->v[i][k] * (double) this->h[k][j];
                }
                
                let this->v[i][j] = z;
            }
        }
    } // end hqr2

    /**
     *    Return the eigenvector matrix
     *
     *    @access public
     *    @return V
     */
    public function getV()
    {
        var jm;
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initialize"], this->v, this->n, this->n);
        
        return jm;
    }

    /**
     *    Return the real parts of the eigenvalues
     *
     *    @access public
     *    @return real(diag(D))
     */
    public function getRealEigenvalues() {
        return this->d;
    }

    /**
     *    Return the imaginary parts of the eigenvalues
     *
     *    @access public
     *    @return imag(diag(D))
     */
    public function getImagEigenvalues()
    {
        return this->e;
    }


    /**
     *    Return the block diagonal eigenvalue matrix
     *
     *    @access public
     *    @return D
     */
    public function getD()
    {
        var jm;
        array d = [];
        int i, o;
        
        for i in range(0, this->n - 1) {
            let d[i] = array_fill(0, this->n, 0.0);
            let d[i][i] = this->d[i];
            
            if (this->e[i] == 0) {
                continue;
            }
            
            if (this->e[i] > 0) {
                let o = i + 1;
            } else {
                let o = i - 1;
            }

            let d[i][o] = this->e[i];
        }
        
        let jm = new \ZExcel\Shared\JAMA\Matrix();
        call_user_func([jm, "initiaze"], d);
        
        return jm;
    }
}
