namespace ZExcel\Calculation;

class Statistical
{
    const LOG_GAMMA_X_MAX_VALUE = "2.55e305";
    
    const XMININ = "2.23e-308";
    
    const EPS = "2.22e-16";
    
    const SQRT2PI = 2.5066282746310005024157652848110452530069867406099;

    // Function cache for logGamma
    private static logGammaCacheResult = 0.0;
    private static logGammaCacheX      = 0.0;

    // Function cache for logBeta function
    private static logBetaCacheP      = 0.0;
    private static logBetaCacheQ      = 0.0;
    private static logBetaCacheResult = 0.0;
    
    
    // Log Gamma related constants
    private static lg_d1 = -0.5772156649015328605195174;
    private static lg_d2 = 0.4227843350984671393993777;
    private static lg_d4 = 1.791759469228055000094023;

    private static lg_p1 = [
        4.945235359296727046734888,
        201.8112620856775083915565,
        2290.838373831346393026739,
        11319.67205903380828685045,
        28557.24635671635335736389,
        38484.96228443793359990269,
        26377.48787624195437963534,
        7225.813979700288197698961
    ];
    private static lg_p2 = [
        4.974607845568932035012064,
        542.4138599891070494101986,
        15506.93864978364947665077,
        184793.2904445632425417223,
        1088204.76946882876749847,
        3338152.967987029735917223,
        5106661.678927352456275255,
        3074109.054850539556250927
    ];
    private static lg_p4 = [
        14745.02166059939948905062,
        2426813.369486704502836312,
        121475557.4045093227939592,
        2663432449.630976949898078,
        29403789566.34553899906876,
        170266573776.5398868392998,
        492612579337.743088758812,
        560625185622.3951465078242
    ];
    private static lg_q1 = [
        67.48212550303777196073036,
        1113.332393857199323513008,
        7738.757056935398733233834,
        27639.87074403340708898585,
        54993.10206226157329794414,
        61611.22180066002127833352,
        36351.27591501940507276287,
        8785.536302431013170870835
    ];
    private static lg_q2 = [
        183.0328399370592604055942,
        7765.049321445005871323047,
        133190.3827966074194402448,
        1136705.821321969608938755,
        5267964.117437946917577538,
        13467014.54311101692290052,
        17827365.30353274213975932,
        9533095.591844353613395747
    ];
    private static lg_q4 = [
        2690.530175870899333379843,
        639388.5654300092398984238,
        41355999.30241388052042842,
        1120872109.61614794137657,
        14886137286.78813811542398,
        101680358627.2438228077304,
        341747634550.7377132798597,
        446315818741.9713286462081
    ];
    private static lg_c  = [
        -0.001910444077728,
        8.4171387781295,
        -0.0005952379913043012,
        0.000793650793500350248,
        -0.002777777777777681622553,
        0.08333333333333333331554247,
        0.0057083835261
    ];

    // Rough estimate of the fourth root of logGamma_xBig
    private static lg_frtbig = "2.25e76";
    private static pnt68     = 0.6796875;
    
    private static array1;
    private static array2;
    
    private static function checkTrendArrays()
    {
        var key, value;
        
        if (!is_array(self::array1)) {
            let self::array1 = [self::array1];
        }
        if (!is_array(self::array2)) {
            let self::array2 = [self::array2];
        }

        let self::array1 = \ZExcel\Calculation\Functions::flattenArray(self::array1);
        let self::array2 = \ZExcel\Calculation\Functions::flattenArray(self::array2);
        
        for key, value in self::array1 {
            if ((is_bool(value)) || (is_string(value)) || (is_null(value))) {
                unset(self::array1[key]);
                unset(self::array2[key]);
            }
        }
        
        for key, value in self::array2 {
            if ((is_bool(value)) || (is_string(value)) || (is_null(value))) {
                unset(self::array1[key]);
                unset(self::array2[key]);
            }
        }
        
        let self::array1 = array_merge(self::array1, []);
        let self::array2 = array_merge(self::array2, []);

        return true;
    }


    /**
     * Beta function.
     *
     * @author Jaco van Kooten
     *
     * @param p require p>0
     * @param q require q>0
     * @return 0 if p<=0, q<=0 or p+q>2.55E305 to avoid errors and over/underflow
     */
    private static function beta(p, q)
    {
        if (p <= 0.0 || q <= 0.0 || (p + q) > floatval(self::LOG_GAMMA_X_MAX_VALUE)) {
            return 0.0;
        } else {
            return exp(self::logBeta(p, q));
        }
    }


    /**
     * Incomplete beta function
     *
     * @author Jaco van Kooten
     * @author Paul Meagher
     *
     * The computation is based on formulas from Numerical Recipes, Chapter 6.4 (W.H. Press et al, 1992).
     * @param x require 0<=x<=1
     * @param p require p>0
     * @param q require q>0
     * @return 0 if x<0, p<=0, q<=0 or p+q>2.55E305 and 1 if x>1 to avoid errors and over/underflow
     */
    private static function incompleteBeta(x, p, q)
    {
        var beta_gam;
        
        if (x <= 0.0) {
            return 0.0;
        } else {
            if (x >= 1.0) {
                return 1.0;
            } else {
                if ((p <= 0.0) || (q <= 0.0) || ((p + q) > floatval(self::LOG_GAMMA_X_MAX_VALUE))) {
                    return 0.0;
                }
            }
        }
        
        let beta_gam = exp((0 - self::logBeta(p, q)) + p * log(x) + q * log(1.0 - x));
        
        if (x < (p + 1.0) / (p + q + 2.0)) {
            return beta_gam * self::betaFraction(x, p, q) / p;
        } else {
            return 1.0 - (beta_gam * self::betaFraction(1 - x, q, p) / q);
        }
    }

    /**
     * The natural logarithm of the beta function.
     *
     * @param p require p>0
     * @param q require q>0
     * @return 0 if p<=0, q<=0 or p+q>2.55E305 to avoid errors and over/underflow
     * @author Jaco van Kooten
     */
    private static function logBeta(p, q)
    {
        if (p != self::logBetaCacheP || q != self::logBetaCacheQ) {
            let self::logBetaCacheP = p;
            let self::logBetaCacheQ = q;
            
            if ((p <= 0.0) || (q <= 0.0) || ((p + q) > floatval(self::LOG_GAMMA_X_MAX_VALUE))) {
                let self::logBetaCacheResult = 0.0;
            } else {
                let self::logBetaCacheResult = self::logGamma(p) + self::logGamma(q) - self::logGamma(p + q);
            }
        }
        return self::logBetaCacheResult;
    }


    /**
     * Evaluates of continued fraction part of incomplete beta function.
     * Based on an idea from Numerical Recipes (W.H. Press et al, 1992).
     * @author Jaco van Kooten
     */
    private static function betaFraction(x, p, q)
    {
        var c, sum_pq, p_plus, p_minus, h, frac, m, m2, delta, d;
        
        let c = 1.0;
        let sum_pq = p + q;
        let p_plus = p + 1.0;
        let p_minus = p - 1.0;
        let h = 1.0 - sum_pq * x / p_plus;
        
        if (abs(h) < floatval(self::XMININ)) {
            let h = floatval(self::XMININ);
        }
        
        let h = 1.0 / h;
        let frac = h;
        let m = 1;
        let delta = 0.0;
        
        while (m <= \ZExcel\Calculation\Functions::MAX_ITERATIONS && abs(delta - 1.0) > floatval(\ZExcel\Calculation\Functions::PRECISION)) {
            let m2 = 2 * m;
            // even index for d
            let d = m * (q - m) * x / ( (p_minus + m2) * (p + m2));
            let h = 1.0 + d * h;
            
            if (abs(h) < floatval(self::XMININ)) {
                let h = floatval(self::XMININ);
            }
            
            let h = 1.0 / h;
            let c = 1.0 + d / c;
            
            if (abs(c) < floatval(self::XMININ)) {
                let c = floatval(self::XMININ);
            }
            
            let frac = frac * h * c;
            // odd index for d
            let d = -(p + m) * (sum_pq + m) * x / ((p + m2) * (p_plus + m2));
            let h = 1.0 + d * h;
            
            if (abs(h) < floatval(self::XMININ)) {
                let h = floatval(self::XMININ);
            }
            
            let h = 1.0 / h;
            let c = 1.0 + d / c;
            
            if (abs(c) < floatval(self::XMININ)) {
                let c = floatval(self::XMININ);
            }
            
            let delta = h * c;
            let frac = frac * delta;
            let m = m + 1;
        }
        return frac;
    }


    /**
     * logGamma function
     *
     * @version 1.1
     * @author Jaco van Kooten
     *
     * Original author was Jaco van Kooten. Ported to PHP by Paul Meagher.
     *
     * The natural logarithm of the gamma function. <br />
     * Based on public domain NETLIB (Fortran) code by W. J. Cody and L. Stoltz <br />
     * Applied Mathematics Division <br />
     * Argonne National Laboratory <br />
     * Argonne, IL 60439 <br />
     * <p>
     * References:
     * <ol>
     * <li>W. J. Cody and K. E. Hillstrom, 'Chebyshev Approximations for the Natural
     *     Logarithm of the Gamma Function,' Math. Comp. 21, 1967, pp. 198-203.</li>
     * <li>K. E. Hillstrom, ANL/AMD Program ANLC366S, DGAMMA/DLGAMA, May, 1969.</li>
     * <li>Hart, Et. Al., Computer Approximations, Wiley and sons, New York, 1968.</li>
     * </ol>
     * </p>
     * <p>
     * From the original documentation:
     * </p>
     * <p>
     * This routine calculates the LOG(GAMMA) function for a positive real argument X.
     * Computation is based on an algorithm outlined in references 1 and 2.
     * The program uses rational functions that theoretically approximate LOG(GAMMA)
     * to at least 18 significant decimal digits. The approximation for X > 12 is from
     * reference 3, while approximations for X < 12.0 are similar to those in reference
     * 1, but are unpublished. The accuracy achieved depends on the arithmetic system,
     * the compiler, the intrinsic functions, and proper selection of the
     * machine-dependent constants.
     * </p>
     * <p>
     * Error returns: <br />
     * The program returns the value XINF for X .LE. 0.0 or when overflow would occur.
     * The computation is believed to be free of underflow and overflow.
     * </p>
     * @return MAX_VALUE for x < 0.0 or when overflow would occur, i.e. x > 2.55E305
     */
    private static function logGamma(x)
    {
        var y, ysq, res, corr, xm1, xnum, xden, xm2, xm4, i;
        
        if (x == self::logGammaCacheX) {
            return self::logGammaCacheResult;
        }
        
        let y = x;
        
        if (y > 0.0 && y <= floatval(self::LOG_GAMMA_X_MAX_VALUE)) {
            if (y <= floatval(self::EPS)) {
                let res = -log(y);
            } else {
                if (y <= 1.5) {
                    // ---------------------
                    //    floatval(self::EPS) .LT. X .LE. 1.5
                    // ---------------------
                    if (y < self::pnt68) {
                        let corr = -log(y);
                        let xm1 = y;
                    } else {
                        let corr = 0.0;
                        let xm1 = y - 1.0;
                    }
                    
                    if (y <= 0.5 || y >= self::pnt68) {
                        let xden = 1.0;
                        let xnum = 0.0;
                        for i in range(1, 7) {
                            let xnum = xnum * xm1 + self::lg_p1[i];
                            let xden = xden * xm1 + self::lg_q1[i];
                        }
                        let res = corr + xm1 * (self::lg_d1 + xm1 * (xnum / xden));
                    } else {
                        let xm2 = y - 1.0;
                        let xden = 1.0;
                        let xnum = 0.0;
                        for i in range(1, 7) {
                            let xnum = xnum * xm2 + self::lg_p2[i];
                            let xden = xden * xm2 + self::lg_q2[i];
                        }
                        let res = corr + xm2 * (self::lg_d2 + xm2 * (xnum / xden));
                    }
                } else {
                    if (y <= 4.0) {
                        // ---------------------
                        //    1.5 .LT. X .LE. 4.0
                        // ---------------------
                        let xm2 = y - 2.0;
                        let xden = 1.0;
                        let xnum = 0.0;
                        for i in range(1, 7) {
                            let xnum = xnum * xm2 + self::lg_p2[i];
                            let xden = xden * xm2 + self::lg_q2[i];
                        }
                        let res = xm2 * (self::lg_d2 + xm2 * (xnum / xden));
                    } else {
                        if (y <= 12.0) {
                            // ----------------------
                            //    4.0 .LT. X .LE. 12.0
                            // ----------------------
                            let xm4 = y - 4.0;
                            let xden = -1.0;
                            let xnum = 0.0;
                            for i in range(1, 7) {
                                let xnum = xnum * xm4 +  self::lg_p4[i];
                                let xden = xden * xm4 + self::lg_q4[i];
                            }
                            let res = self::lg_d4 + xm4 * (xnum / xden);
                        } else {
                            // ---------------------------------
                            //    Evaluate for argument .GE. 12.0
                            // ---------------------------------
                            let res = 0.0;
                            if (y <= floatval(self::lg_frtbig)) {
                                let res =   floatval(self::lg_c[6]);
                                let ysq = y * y;
                                for i in range(1, 5) {
                                    let res = res / ysq +   floatval(self::lg_c[i]);
                                }
                                let res = res / y;
                                let corr = log(y);
                                let res = res + (y * (corr - 1.0));
                            }
                        }
                    }
                }
            }
        } else {
            // --------------------------
            //    Return for bad arguments
            // --------------------------
            let res = floatval(\ZExcel\Calculation\Functions::MAX_VALUE);
        }
        // ------------------------------
        //    Final adjustments and return
        // ------------------------------
        let self::logGammaCacheX = x;
        let self::logGammaCacheResult = res;
        
        return res;
    }


    /**
     * Private implementation of the incomplete Gamma function
     */
    private static function incompleteGamma(a, x)
    {
        var n, i, divisor, summer = 0;
        
        for n in range(0, 32) {
            let divisor = a;
            for i in range(1, n) {
                let divisor = divisor * (a + i);
            }
            let summer = summer  + (pow(x, n) / divisor);
        }
        
        return pow(x, a) * exp(0-x) * summer;
    }


    /**
     * Private implementation of the Gamma function
     */
    private static function gamma(data)
    {
        var p0, p, y, x, tmp, summer, j;
        
        let p0 = 1.000000000190015;
        let p = [
            1: 76.18009172947146,
            2: -86.50532032941677,
            3: 24.01409824083091,
            4: -1.231739572450155,
            5: 0.001208650973866179,
            6: -0.000005395239384953
        ];
        
        if (data == 0.0) {
            return 0;
        }

        let y = data;
        let x = data;
        let tmp = x + 5.5;
        let tmp = tmp - (x + 0.5) * log(tmp);
        let summer = p0;
        
        for j in range(1, 6) {
            let y = y + 1;
            let summer = summer + (floatval(p[j]) / y);
        }
        
        return exp(0 - tmp + log(self::SQRT2PI * (double) summer / x));
    }


    

    /***************************************************************************
     *                                inverse_ncdf.php
     *                            -------------------
     *    begin                : Friday, January 16, 2004
     *    copyright            : (C) 2004 Michael Nickerson
     *    email                : nickersonm@yahoo.com
     *
     ***************************************************************************/
    private static function inverseNcdf(p)
    {
        var p, q, r, p_low, p_high;
        array a, b, c, d;
        
        //    Coefficients in rational approximations
        let a = [
            1: -39.69683028665376,
            2: 220.9460984245205,
            3: -275.9285104469687,
            4: 138.3577518672690,
            5: -30.66479806614716,
            6: 2.506628277459239
        ];
        let b = [
            1: -54.47609879822406,
            2: 161.5858368580409,
            3: -155.6989798598866,
            4: 66.80131188771972,
            5: -13.28068155288572
        ];
        let c = [
            1: -0.007784894002430293,
            2: -0.3223964580411365,
            3: -2.400758277161838,
            4: -2.549732539343734,
            5: 4.374664141464968,
            6: 2.938163982698783
        ];
        let d = [
            1: 0.007784695709041462,
            2: 0.3224671290700398,
            3: 2.445134137142996,
            4: 3.754408661907416
        ];
    
        //    Inverse ncdf approximation by Peter J. Acklam, implementation adapted to
        //    PHP by Michael Nickerson, using Dr. Thomas Ziegler's C implementation as
        //    a guide. http://home.online.no/~pjacklam/notes/invnorm/index.html
        //    I have not checked the accuracy of this implementation. Be aware that PHP
        //    will truncate the coeficcients to 14 digits.

        //    You have permission to use and distribute this function freely for
        //    whatever purpose you want, but please show common courtesy and give credit
        //    where credit is due.

        //    Input paramater is p - probability - where 0 < p < 1.

        //    Define lower and upper region break-points.
        let p_low = 0.02425;            // Use lower region approx. below this
        let p_high = 1 - p_low;        // Use upper region approx. above this

        if (0 < p && p < p_low) {
            //    Rational approximation for lower region.
            let q = sqrt(-2 * log(p));
            return (((((c[1] * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) * q + c[6]) /
                    ((((d[1] * q + d[2]) * q + d[3]) * q + d[4]) * q + 1);
        } else {
            if (p_low <= p && p <= p_high) {
                //    Rational approximation for central region.
                let q = p - 0.5;
                let r = q * q;
                return (((((a[1] * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * r + a[6]) * q /
                       (((((b[1] * r + b[2]) * r + b[3]) * r + b[4]) * r + b[5]) * r + 1);
            } else {
                if (p_high < p && p < 1) {
                    //    Rational approximation for upper region.
                    let q = sqrt(-2 * log(1 - p));
                    return -(((((c[1] * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) * q + c[6]) /
                             ((((d[1] * q + d[2]) * q + d[3]) * q + d[4]) * q + 1);
                }
            }
        }
        
        //    If 0 < p < 1, return a null value
        return \ZExcel\Calculation\Functions::NuLLL();
    }


    private static function inverseNcdf2(double prob)
    {
        //    Approximation of inverse standard normal CDF developed by
        //    B. Moro, "The Full Monte," Risk 8(2), Feb 1995, 57-58.

        float
            a1 = 2.50662823884,
            a2 = -18.61500062529,
            a3 = 41.39119773534,
            a4 = -25.44106049637,

            b1 = -8.4735109309,
            b2 = 23.08336743743,
            b3 = -21.06224101826,
            b4 = 3.13082909833,

            c1 = 0.337475482272615,
            c2 = 0.976169019091719,
            c3 = 0.160797971491821,
            c4 = 0.0276438810333863,
            c5 = 0.0038405729373609,
            c6 = 0.0003951896511919,
            c7 = 0.0000321767881768,
            c8 = 0.0000002888167364,
            c9 = 0.0000003960315187,
            y, z;

        let y = prob - 0.5;
        
        if (abs(y) < 0.42) {
            let z = (y * y);
            let z = y * (((a4 * z + a3) * z + a2) * z + a1) / ((((b4 * z + b3) * z + b2) * z + b1) * z + 1);
        } else {
            if (y > 0) {
                let z = 0.0 + log(-log(1 - prob));
            } else {
                let z = 0.0 + log(-log(prob));
            }
            
            let z = c1 + z * (c2 + z * (c3 + z * (c4 + z * (c5 + z * (c6 + z * (c7 + z * (c8 + z * c9)))))));
            
            if (y < 0) {
                let z = -z;
            }
        }
        
        return z;
    }    //    function inverseNcdf2()


    private static function inverseNcdf3(p)
    {
        //    ALGORITHM AS241 APPL. STATIST. (1988) VOL. 37, NO. 3.
        //    Produces the normal deviate Z corresponding to a given lower
        //    tail area of P; Z is accurate to about 1 part in 10**16.
        //
        //    This is a PHP version of the original FORTRAN code that can
        //    be found at http://lib.stat.cmu.edu/apstat/
        int   split2 = 5;
        float split1 = 0.425,
              const1 = 0.180625,
              const2 = 1.6;
        //    coefficients for p close to 0.5
              
        var a0 = 3.387132872796366608,
            a1 = 133.14166789178437745,
            a2 = 1971.5909503065514427,
            a3 = 13731.693765509461125,
            a4 = 45921.953931549871457,
            a5 = 67265.770927008700853,
            a6 = 33430.575583588128105,
            a7 = 2509.0809287301226727,

            b1 = 42.313330701600911252,
            b2 = 687.1870074920579083,
            b3 = 5394.1960214247511077,
            b4 = 21213.794301586595867,
            b5 = 39307.895800092710610,
            b6 = 28729.085735721942674,
            b7 = 5226.495278852854561,

        //    coefficients for p not close to 0, 0.5 or 1.
            c0 = 1.42343711074968357734,
            c1 = 4.6303378461565452959,
            c2 = 5.7694972214606914055,
            c3 = 3.64784832476320460504,
            c4 = 1.27045825245236838258,
            c5 = 0.24178072517745061177,
            c6 = 0.0227238449892691845833,
            c7 = 0.00077454501427834140764,

            d1 = 2.05319162663775882187,
            d2 = 1.6763848301838038494,
            d3 = 0.68976733498510000455,
            d4 = 0.14810397642748007459,
            d5 = 0.0151986665636164571966,
            d6 = 0.0005475938084995344946,
            d7 = 0.00000000105075007164441684324,

        //    coefficients for p near 0 or 1.
            e0 = 6.6579046435011037772,
            e1 = 5.4637849111641143699,
            e2 = 1.7848265399172913358,
            e3 = 0.29656057182850489123,
            e4 = 0.026532189526576123093,
            e5 = 0.0012426609473880784386,
            e6 = 0.0000271155556874348757815,
            e7 = 0.000000201033439929228813265,

            f1 = 0.59983220655588793769,
            f2 = 0.13692988092273580531,
            f3 = 0.0148753612908506148525,
            f4 = 0.0007868691311456132591,
            f5 = 0.000018463183175100546818,
            f6 = 0.00000014215117583164458887,
            f7 = 0.00000000000000204426310338993978564,
            q, r, z;

        let q = p - 0.5;

        //    computation for p close to 0.5
        if (abs(q) <= split1) {
            let r = const1 - q * q;
            let z =  q * ((((((((double) a7 * r + (double) a6) * r +(double)  a5) * r + (double) a4) * r + (double) a3) * r + (double) a2) * r + (double) a1) * r + (double) a0) /
                      ((((((((double) b7 * r + (double) b6) * r + (double) b5) * r + (double) b4) * r + (double) b3) * r + (double) b2) * r + (double) b1) * r + 1);
        } else {
            if (q < 0) {
                let r = p;
            } else {
                let r = 1 - p;
            }
            let r = pow(-log(r), 2);

            //    computation for p not close to 0, 0.5 or 1.
            if (r <= split2) {
                let r = r - (double) const2;
                let z = ((((((((double) c7 * r + (double) c6) * r + (double) c5) * r + (double) c4) * r + (double) c3) * r + (double) c2) * r + (double) c1) * r + c0) /
                     ((((((((double) d7 * r + (double) d6) * r + (double) d5) * r + (double) d4) * r + (double) d3) * r + (double) d2) * r + (double) d1) * r + 1);
            } else {
            //    computation for p near 0 or 1.
                let r = r - (double) split2;
                let z = ((((((((double) e7 * r + (double) e6) * r + (double) e5) * r + (double) e4) * r + (double) e3) * r + (double) e2) * r + (double) e1) * r + (double) e0) /
                     ((((((((double) f7 * r + (double) f6) * r + (double) f5) * r + (double) f4) * r + (double) f3) * r + (double) f2) * r + f1) * r + 1);
            }
            if (q < 0) {
                let z = -z;
            }
        }
        
        return z;
    }


    /**
     * AVEDEV
     *
     * Returns the average of the absolute deviations of data points from their mean.
     * AVEDEV is a measure of the variability in a data set.
     *
     * Excel Function:
     *        AVEDEV(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function aveDev()
    {
        var aArgs, returnValue, aMean, aCount, k, arg;

        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aMean = call_user_func("self::AVeRAGE", aArgs);
        if (aMean != \ZExcel\Calculation\Functions::DiV0()) {
            let aCount = 0;
            for k, arg in aArgs {
                if ((is_bool(arg)) &&
                    ((!\ZExcel\Calculation\Functions::isCellValue(k)) || (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                    let arg = (int) arg;
                }
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    if (is_null(returnValue)) {
                        let returnValue = abs(arg - aMean);
                    } else {
                        let returnValue = returnValue + abs(arg - aMean);
                    }
                    let aCount = aCount + 1;
                }
            }

            // Return
            if (aCount == 0) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
            return returnValue / aCount;
        }
        return \ZExcel\Calculation\Functions::NaN();
    }


    /**
     * AVERAGE
     *
     * Returns the average (arithmetic mean) of the arguments
     *
     * Excel Function:
     *        AVERAGE(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function average()
    {
        var k, arg, returnValue = 0, aCount = 0;

        // Loop through arguments
        for k, arg in \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args()) {
            if ((is_bool(arg)) &&
                ((!\ZExcel\Calculation\Functions::isCellValue(k)) || (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                let arg = (int) arg;
            }
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                if (is_null(returnValue)) {
                    let returnValue = arg;
                } else {
                    let returnValue = returnValue + arg;
                }
                
                let aCount = aCount + 1;
            }
        }

        // Return
        if (aCount > 0) {
            return returnValue / aCount;
        } else {
            return \ZExcel\Calculation\Functions::DiV0();
        }
    }


    /**
     * AVERAGEA
     *
     * Returns the average of its arguments, including numbers, text, and logical values
     *
     * Excel Function:
     *        AVERAGEA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function averageA()
    {
        var k, arg, returnValue = null, aCount = 0;
        
        // Loop through arguments
        for k, arg in \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args()) {
            if ((is_bool(arg)) &&
                (!\ZExcel\Calculation\Functions::isMatrixValue(k))) {
            } else {
                if ((is_numeric(arg)) || (is_bool(arg)) || ((is_string(arg) && (arg != "")))) {
                    if (is_bool(arg)) {
                        let arg = (int) arg;
                    } else {
                        if (is_string(arg)) {
                            let arg = 0;
                        }
                    }
                    
                    if (is_null(returnValue)) {
                        let returnValue = arg;
                    } else {
                        let returnValue = returnValue + arg;
                    }
                    
                    let aCount = aCount + 1;
                }
            }
        }

        if (aCount > 0) {
            return returnValue / aCount;
        } else {
            return \ZExcel\Calculation\Functions::DiV0();
        }
    }


    /**
     * AVERAGEIF
     *
     * Returns the average value from a range of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        AVERAGEIF(value1[,value2[, ...]],condition)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @param    string        condition        The criteria that defines which cells will be checked.
     * @param    mixed[]        averageArgs    Data values
     * @return    float
     */
    public static function averageIf(aArgs, condition, averageArgs = [])
    {
        var aCount, arg, testCondition, returnValue = 0;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(aArgs);
        let averageArgs = \ZExcel\Calculation\Functions::flattenArray(averageArgs);
        
        if (empty(averageArgs)) {
            let averageArgs = aArgs;
        }
        
        let condition = \ZExcel\Calculation\Functions::ifCondition(condition);
        // Loop through arguments
        let aCount = 0;
        for arg in aArgs {
            if (!is_numeric(arg)) {
                let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
            }
            
            let testCondition = "=".arg.condition;
            
            if (\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                if ((is_null(returnValue)) || (arg > returnValue)) {
                    let returnValue = returnValue + arg;
                    let aCount = aCount + 1;
                }
            }
        }

        if (aCount > 0) {
            return returnValue / aCount;
        }
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * BETADIST
     *
     * Returns the beta distribution.
     *
     * @param    float        value            Value at which you want to evaluate the distribution
     * @param    float        alpha            Parameter to the distribution
     * @param    float        beta            Parameter to the distribution
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function betaDist(value, alpha, beta, rMin = 0, rMax = 1)
    {
        var tmp, rMin, rMax;
        
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let alpha = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);
        let beta  = \ZExcel\Calculation\Functions::flattenSingleValue(beta);
        let rMin  = \ZExcel\Calculation\Functions::flattenSingleValue(rMin);
        let rMax  = \ZExcel\Calculation\Functions::flattenSingleValue(rMax);

        if ((is_numeric(value)) && (is_numeric(alpha)) && (is_numeric(beta)) && (is_numeric(rMin)) && (is_numeric(rMax))) {
            if ((value < rMin) || (value > rMax) || (alpha <= 0) || (beta <= 0) || (rMin == rMax)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            if (rMin > rMax) {
                let tmp = rMin;
                let rMin = rMax;
                let rMax = tmp;
            }
            
            let value = value - rMin;
            let value = value / (rMax - rMin);
            
            return self::incompleteBeta(value, alpha, beta);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * BETAINV
     *
     * Returns the inverse of the beta distribution.
     *
     * @param    float        probability    Probability at which you want to evaluate the distribution
     * @param    float        alpha            Parameter to the distribution
     * @param    float        beta            Parameter to the distribution
     * @param    float        rMin            Minimum value
     * @param    float        rMax            Maximum value
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function betaInv(probability, alpha, beta, rMin = 0, rMax = 1)
    {
        var tmp, rMin, rMax, a, b, i, guess, result;
        
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let alpha       = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);
        let beta        = \ZExcel\Calculation\Functions::flattenSingleValue(beta);
        let rMin        = \ZExcel\Calculation\Functions::flattenSingleValue(rMin);
        let rMax        = \ZExcel\Calculation\Functions::flattenSingleValue(rMax);

        if ((is_numeric(probability)) && (is_numeric(alpha)) && (is_numeric(beta)) && (is_numeric(rMin)) && (is_numeric(rMax))) {
            if ((alpha <= 0) || (beta <= 0) || (rMin == rMax) || (probability <= 0) || (probability > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if (rMin > rMax) {
                let tmp = rMin;
                let rMin = rMax;
                let rMax = tmp;
            }
            
            let a = 0;
            let b = 2;

            let i = 0;
            while (((b - a) > floatval(\ZExcel\Calculation\Functions::PRECISION)) && (i < \ZExcel\Calculation\Functions::MAX_ITERATIONS)) {
                let i = i + 1;
                
                let guess = (a + b) / 2;
                let result = self::BeTADIST(guess, alpha, beta);
                if ((result == probability) || (result == 0)) {
                    let b = a;
                } else {
                    if (result > probability) {
                        let b = guess;
                    } else {
                        let a = guess;
                    }
                }
            }
            if (i == \ZExcel\Calculation\Functions::MAX_ITERATIONS) {
                return \ZExcel\Calculation\Functions::Na();
            }
            return round(rMin + guess * (rMax - rMin), 12);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * BINOMDIST
     *
     * Returns the individual term binomial distribution probability. Use BINOMDIST in problems with
     *        a fixed number of tests or trials, when the outcomes of any trial are only success or failure,
     *        when trials are independent, and when the probability of success is constant throughout the
     *        experiment. For example, BINOMDIST can calculate the probability that two of the next three
     *        babies born are male.
     *
     * @param    float        value            Number of successes in trials
     * @param    float        trials            Number of trials
     * @param    float        probability    Probability of success on each trial
     * @param    boolean        cumulative
     * @return    float
     *
     * @todo    Cumulative distribution function
     *
     */
    public static function binomDist(value, trials, probability, cumulative)
    {
        var summer, i;
        
        let value       = floor(\ZExcel\Calculation\Functions::flattenSingleValue(value));
        let trials      = floor(\ZExcel\Calculation\Functions::flattenSingleValue(trials));
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);

        if ((is_numeric(value)) && (is_numeric(trials)) && (is_numeric(probability))) {
            if ((value < 0) || (value > trials)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            if ((probability < 0) || (probability > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    let summer = 0;
                    for i in range(0, value) {
                        let summer = summer + \ZExcel\Calculation\MathTrig::CoMBIN(trials, i) * pow(probability, i) * pow(1 - probability, trials - i);
                    }
                    return summer;
                } else {
                    return \ZExcel\Calculation\MathTrig::CoMBIN(trials, value) * pow(probability, value) * pow(1 - probability, trials - value) ;
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * CHIDIST
     *
     * Returns the one-tailed probability of the chi-squared distribution.
     *
     * @param    float        value            Value for the function
     * @param    float        degrees        degrees of freedom
     * @return    float
     */
    public static function chiDist(value, degrees)
    {
        let value   = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let degrees = floor(\ZExcel\Calculation\Functions::flattenSingleValue(degrees));

        if ((is_numeric(value)) && (is_numeric(degrees))) {
            if (degrees < 1) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if (value < 0) {
                if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                    return 1;
                }
                return \ZExcel\Calculation\Functions::NaN();
            }
            return 1 - (self::incompleteGamma(degrees/2, value/2) / self::gamma(degrees/2));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * CHIINV
     *
     * Returns the one-tailed probability of the chi-squared distribution.
     *
     * @param    float        probability    Probability for the function
     * @param    float        degrees        degrees of freedom
     * @return    float
     */
    public static function chiInv(probability, degrees)
    {
        var x, xLo, xHi, xNew, dx, i, result, error;
        
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let degrees     = floor(\ZExcel\Calculation\Functions::flattenSingleValue(degrees));

        if ((is_numeric(probability)) && (is_numeric(degrees))) {
            let xLo = 100;
            let xHi = 0;

            let x = 1;
            let xNew = 1;
            let dx = 1;
            let i = 0;

            while ((abs(dx) > floatval(\ZExcel\Calculation\Functions::PRECISION)) && (i < \ZExcel\Calculation\Functions::MAX_ITERATIONS)) {
                let i = i + 1;
                
                // Apply Newton-Raphson step
                let result = self::CHiDIST(x, degrees);
                let error = result - probability;
                
                if (error == 0.0) {
                    let dx = 0;
                } else {
                    if (error < 0.0) {
                        let xLo = x;
                    } else {
                        let xHi = x;
                    }
                }
                
                // Avoid division by zero
                if (result != 0.0) {
                    let dx = error / result;
                    let xNew = x - dx;
                }
                // If the NR fails to converge (which for example may be the
                // case if the initial guess is too rough) we apply a bisection
                // step to determine a more narrow interval around the root.
                if ((xNew < xLo) || (xNew > xHi) || (result == 0.0)) {
                    let xNew = (xLo + xHi) / 2;
                    let dx = xNew - x;
                }
                let x = xNew;
            }
            
            if (i == \ZExcel\Calculation\Functions::MAX_ITERATIONS) {
                return \ZExcel\Calculation\Functions::Na();
            }
            
            return round(x, 12);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * CONFIDENCE
     *
     * Returns the confidence interval for a population mean
     *
     * @param    float        alpha
     * @param    float        stdDev        Standard Deviation
     * @param    float        size
     * @return    float
     *
     */
    public static function confidence(alpha, stdDev, size)
    {
        let alpha  = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);
        let stdDev = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);
        let size   = floor(\ZExcel\Calculation\Functions::flattenSingleValue(size));

        if ((is_numeric(alpha)) && (is_numeric(stdDev)) && (is_numeric(size))) {
            if ((alpha <= 0) || (alpha >= 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((stdDev <= 0) || (size < 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return self::NoRMSINV(1 - alpha / 2) * stdDev / sqrt(size);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * CORREL
     *
     * Returns covariance, the average of the products of deviations for each data point pair.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function correl(yValues, xValues = null)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        if ((is_null(xValues)) || (!is_array(yValues)) || (!is_array(xValues))) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getCorrelation();
    }


    /**
     * COUNT
     *
     * Counts the number of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        COUNT(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    int
     */
    public static function count()
    {
        var aArgs, k, arg, returnValue = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        
        for k, arg in aArgs {
            if ((is_bool(arg)) &&
                ((!\ZExcel\Calculation\Functions::isCellValue(k)) || (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                let arg = (int) arg;
            }
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let returnValue = returnValue + 1;
            }
        }

        return returnValue;
    }


    /**
     * COUNTA
     *
     * Counts the number of cells that are not empty within the list of arguments
     *
     * Excel Function:
     *        COUNTA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    int
     */
    public static function countA()
    {
        var aArgs, arg, returnValue = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        for arg in aArgs {
            // Is it a numeric, boolean or string value?
            if ((is_numeric(arg)) || (is_bool(arg)) || ((is_string(arg) && (arg != "")))) {
                let returnValue = returnValue + 1;
            }
        }

        return returnValue;
    }


    /**
     * COUNTBLANK
     *
     * Counts the number of empty cells within the list of arguments
     *
     * Excel Function:
     *        COUNTBLANK(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    int
     */
    public static function countBlank()
    {
        var aArgs, arg, returnValue = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        for arg in aArgs {
            // Is it a blank cell?
            if ((is_null(arg)) || ((is_string(arg)) && (arg == ""))) {
                let returnValue = returnValue + 1;
            }
        }

        return returnValue;
    }


    /**
     * COUNTIF
     *
     * Counts the number of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        COUNTIF(value1[,value2[, ...]],condition)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    string        condition        The criteria that defines which cells will be counted.
     * @return    int
     */
    public static function countIf(aArgs, condition)
    {
        var arg, testCondition, returnValue = 0;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(aArgs);
        let condition = \ZExcel\Calculation\Functions::ifCondition(condition);
        // Loop through arguments
        for arg in aArgs {
            if (!is_numeric(arg)) {
                let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
            }
            
            let testCondition = "=" . arg.condition;
            
            if (\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                // Is it a value within our criteria
                let returnValue = returnValue + 1;
            }
        }

        return returnValue;
    }


    /**
     * COVAR
     *
     * Returns covariance, the average of the products of deviations for each data point pair.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function covar(yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getCovariance();
    }


    /**
     * CRITBINOM
     *
     * Returns the smallest value for which the cumulative binomial distribution is greater
     *        than or equal to a criterion value
     *
     * See http://support.microsoft.com/kb/828117/ for details of the algorithm used
     *
     * @param    float        trials            number of Bernoulli trials
     * @param    float        probability    probability of a success on each trial
     * @param    float        alpha            criterion value
     * @return    int
     *
     * @todo    Warning. This implementation differs from the algorithm detailed on the MS
     *            web site in that CumPGuessMinus1 = CumPGuess - 1 rather than CumPGuess - PGuess
     *            This eliminates a potential endless loop error, but may have an adverse affect on the
     *            accuracy of the function (although all my tests have so far returned correct results).
     *
     */
    public static function critBinom(trials, probability, alpha)
    {
        var t, Guess, trialsApprox, TotalUnscaledProbability, UnscaledPGuess,
            UnscaledCumPGuess, EssentiallyZero, m, PreviousValue, Done, k,
            CurrentValue, PGuess, CumPGuess, CumPGuessMinus1, PGuessPlus1,
            PGuessMinus1;
        
        let trials      = floor(\ZExcel\Calculation\Functions::flattenSingleValue(trials));
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let alpha       = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);

        if ((is_numeric(trials)) && (is_numeric(probability)) && (is_numeric(alpha))) {
            if (trials < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            } else {
                if ((probability < 0) || (probability > 1)) {
                    return \ZExcel\Calculation\Functions::NaN();
                } else {
                    if ((alpha < 0) || (alpha > 1)) {
                        return \ZExcel\Calculation\Functions::NaN();
                    } else {
                        if (alpha <= 0.5) {
                            let t = sqrt(log(1 / (alpha * alpha)));
                            let trialsApprox = 0 - (t + (2.515517 + 0.802853 * t + 0.010328 * t * t) / (1 + 1.432788 * t + 0.189269 * t * t + 0.001308 * t * t * t));
                        } else {
                            let t = sqrt(log(1 / pow(1 - alpha, 2)));
                            let trialsApprox = t - (2.515517 + 0.802853 * t + 0.010328 * t * t) / (1 + 1.432788 * t + 0.189269 * t * t + 0.001308 * t * t * t);
                        }
                    }
                }
            }
            
            let Guess = floor(trials * probability + trialsApprox * sqrt(trials * probability * (1 - probability)));
            
            if (Guess < 0) {
                let Guess = 0;
            } else {
                if (Guess > trials) {
                    let Guess = trials;
                }
            }

            let TotalUnscaledProbability = 0.0;
            let UnscaledPGuess = 0.0;
            let UnscaledCumPGuess = 0.0;
            let EssentiallyZero = 0.00000000001;

            let m = floor(trials * probability);
            let TotalUnscaledProbability = TotalUnscaledProbability + 1;
            
            if (m == Guess) {
                let UnscaledPGuess = UnscaledPGuess + 1;
            }
            if (m <= Guess) {
                let UnscaledCumPGuess = UnscaledCumPGuess + 1;
            }

            let PreviousValue = 1;
            let Done = false;
            let k = m + 1;
            while ((!Done) && (k <= trials)) {
                let CurrentValue = PreviousValue * (trials - k + 1) * probability / (k * (1 - probability));
                let TotalUnscaledProbability = TotalUnscaledProbability + CurrentValue;
                
                if (k == Guess) {
                    let UnscaledPGuess = UnscaledPGuess + CurrentValue;
                }
                
                if (k <= Guess) {
                    let UnscaledCumPGuess = UnscaledCumPGuess + CurrentValue;
                }
                
                if (CurrentValue <= EssentiallyZero) {
                    let Done = true;
                }
                
                let PreviousValue = CurrentValue;
                let k = k + 1;
            }

            let PreviousValue = 1;
            let Done = false;
            let k = m - 1;
            while ((!Done) && (k >= 0)) {
                let CurrentValue = PreviousValue * k + 1 * (1 - probability) / ((trials - k) * probability);
                let TotalUnscaledProbability = TotalUnscaledProbability + CurrentValue;
                
                if (k == Guess) {
                    let UnscaledPGuess = UnscaledPGuess + CurrentValue;
                }
                
                if (k <= Guess) {
                    let UnscaledCumPGuess = UnscaledCumPGuess + CurrentValue;
                }
                
                if (CurrentValue <= EssentiallyZero) {
                    let Done = true;
                }
                
                let PreviousValue = CurrentValue;
                let k = k - 1;
            }

            let PGuess = UnscaledPGuess / TotalUnscaledProbability;
            let CumPGuess = UnscaledCumPGuess / TotalUnscaledProbability;

//            CumPGuessMinus1 = CumPGuess - PGuess;
            let CumPGuessMinus1 = CumPGuess - 1;

            while (true) {
                if ((CumPGuessMinus1 < alpha) && (CumPGuess >= alpha)) {
                    return Guess;
                } else {
                    if ((CumPGuessMinus1 < alpha) && (CumPGuess < alpha)) {
                        let PGuessPlus1 = PGuess * (trials - Guess) * probability / Guess / (1 - probability);
                        let CumPGuessMinus1 = CumPGuess;
                        let CumPGuess = CumPGuess + PGuessPlus1;
                        let PGuess = PGuessPlus1;
                        let Guess = Guess + 1;
                    } else {
                        if ((CumPGuessMinus1 >= alpha) && (CumPGuess >= alpha)) {
                            let PGuessMinus1 = PGuess * Guess * (1 - probability) / (trials - Guess + 1) / probability;
                            let CumPGuess = CumPGuessMinus1;
                            let CumPGuessMinus1 = CumPGuessMinus1 - PGuess;
                            let PGuess = PGuessMinus1;
                            let Guess = Guess - 1;
                        }
                    }
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * DEVSQ
     *
     * Returns the sum of squares of deviations of data points from their sample mean.
     *
     * Excel Function:
     *        DEVSQ(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function devsq()
    {
        var returnValue = null, aArgs, aMean, k, arg, aCount;
        
        // Return value
        let returnValue = null;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());

        let aMean = call_user_func("self::AVeRAGE", aArgs);
        if (aMean != \ZExcel\Calculation\Functions::DiV0()) {
            let aCount = -1;
            for k, arg in aArgs {
                // Is it a numeric value?
                if ((is_bool(arg)) &&
                    ((!\ZExcel\Calculation\Functions::isCellValue(k)) ||
                    (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                    let arg = (int) arg;
                }
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    if (is_null(returnValue)) {
                        let returnValue = pow((arg - aMean), 2);
                    } else {
                        let returnValue = returnValue + pow((arg - aMean), 2);
                    }
                    let aCount = aCount + 1;
                }
            }

            // Return
            if (is_null(returnValue)) {
                return \ZExcel\Calculation\Functions::NaN();
            } else {
                return returnValue;
            }
        }
        
        return \ZExcel\Calculation\Functions::Na();
    }


    /**
     * EXPONDIST
     *
     *    Returns the exponential distribution. Use EXPONDIST to model the time between events,
     *        such as how long an automated bank teller takes to deliver cash. For example, you can
     *        use EXPONDIST to determine the probability that the process takes at most 1 minute.
     *
     * @param    float        value            Value of the function
     * @param    float        lambda            The parameter value
     * @param    boolean        cumulative
     * @return    float
     */
    public static function exponDist(value, lambda, cumulative)
    {
        let value    = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let lambda    = \ZExcel\Calculation\Functions::flattenSingleValue(lambda);
        let cumulative    = \ZExcel\Calculation\Functions::flattenSingleValue(cumulative);

        if ((is_numeric(value)) && (is_numeric(lambda))) {
            if ((value < 0) || (lambda < 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    return 1 - exp(0-value*lambda);
                } else {
                    return lambda * exp(0-value*lambda);
                }
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FISHER
     *
     * Returns the Fisher transformation at x. This transformation produces a function that
     *        is normally distributed rather than skewed. Use this function to perform hypothesis
     *        testing on the correlation coefficient.
     *
     * @param    float        value
     * @return    float
     */
    public static function fisher(value)
    {
        let value    = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        if (is_numeric(value)) {
            if ((value <= -1) || (value >= 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return 0.5 * log((1+value)/(1-value));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FISHERINV
     *
     * Returns the inverse of the Fisher transformation. Use this transformation when
     *        analyzing correlations between ranges or arrays of data. If y = FISHER(x), then
     *        FISHERINV(y) = x.
     *
     * @param    float        value
     * @return    float
     */
    public static function fisherInv(value)
    {
        let value    = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        if (is_numeric(value)) {
            return (exp(2 * value) - 1) / (exp(2 * value) + 1);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * FORECAST
     *
     * Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value.
     *
     * @param    float                Value of X for which we want to find Y
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function forecast(xValue, yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let xValue = \ZExcel\Calculation\Functions::flattenSingleValue(xValue);
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!is_numeric(xValue)) {
            return \ZExcel\Calculation\Functions::VaLUE();
        } else {
            if (!self::checkTrendArrays()) {
                return \ZExcel\Calculation\Functions::VaLUE();
            }
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getValueOfYForX(xValue);
    }


    /**
     * GAMMADIST
     *
     * Returns the gamma distribution.
     *
     * @param    float        value            Value at which you want to evaluate the distribution
     * @param    float        a                Parameter to the distribution
     * @param    float        b                Parameter to the distribution
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function gammaDist(value, a, b, cumulative)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let a     = \ZExcel\Calculation\Functions::flattenSingleValue(a);
        let b     = \ZExcel\Calculation\Functions::flattenSingleValue(b);

        if ((is_numeric(value)) && (is_numeric(a)) && (is_numeric(b))) {
            if ((value < 0) || (a <= 0) || (b <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    return self::incompleteGamma(a, value / b) / self::gamma(a);
                } else {
                    return (1 / (pow(b, a) * self::gamma(a))) * pow(value, a - 1) * exp(0 - (value / b));
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * GAMMAINV
     *
     * Returns the inverse of the beta distribution.
     *
     * @param    float        probability    Probability at which you want to evaluate the distribution
     * @param    float        alpha            Parameter to the distribution
     * @param    float        beta            Parameter to the distribution
     * @return    float
     *
     */
    public static function gammaInv(probability, alpha, beta)
    {
        var xLo, xHi, x, xNew, dx, i, error, pdf;
        
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let alpha       = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);
        let beta        = \ZExcel\Calculation\Functions::flattenSingleValue(beta);

        if ((is_numeric(probability)) && (is_numeric(alpha)) && (is_numeric(beta))) {
            if ((alpha <= 0) || (beta <= 0) || (probability < 0) || (probability > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }

            let xLo = 0;
            let xHi = alpha * beta * 5;

            let x = 1;
            let xNew = 1;
            let error = 0;
            let pdf = 0;
            let dx = 1024;
            let i = 0;

            while ((abs(dx) > floatval(\ZExcel\Calculation\Functions::PRECISION)) && (i < \ZExcel\Calculation\Functions::MAX_ITERATIONS)) {
                let i = i + 1;
                
                // Apply Newton-Raphson step
                let error = self::GaMMADIST(x, alpha, beta, true) - probability;
                if (error < 0.0) {
                    let xLo = x;
                } else {
                    let xHi = x;
                }
                let pdf = self::GaMMADIST(x, alpha, beta, false);
                // Avoid division by zero
                if (pdf != 0.0) {
                    let dx = error / pdf;
                    let xNew = x - dx;
                }
                // If the NR fails to converge (which for example may be the
                // case if the initial guess is too rough) we apply a bisection
                // step to determine a more narrow interval around the root.
                if ((xNew < xLo) || (xNew > xHi) || (pdf == 0.0)) {
                    let xNew = (xLo + xHi) / 2;
                    let dx = xNew - x;
                }
                let x = xNew;
            }
            if (i == \ZExcel\Calculation\Functions::MAX_ITERATIONS) {
                return \ZExcel\Calculation\Functions::Na();
            }
            return x;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * GAMMALN
     *
     * Returns the natural logarithm of the gamma function.
     *
     * @param    float        value
     * @return    float
     */
    public static function gammaLn(value)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        if (is_numeric(value)) {
            if (value <= 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return log(self::gamma(value));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * GEOMEAN
     *
     * Returns the geometric mean of an array or range of positive data. For example, you
     *        can use GEOMEAN to calculate average growth rate given compound interest with
     *        variable rates.
     *
     * Excel Function:
     *        GEOMEAN(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function geoman()
    {
        var aArgs, aMean, aCount;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        let aMean = call_user_func("\ZExcel\Calculation\MathTrig::PRoDUCT", aArgs);
        
        if (is_numeric(aMean) && (aMean > 0)) {
            let aCount = call_user_func("self::CoUNT", aArgs) ;
            if (call_user_func("self::MiN", aArgs) > 0) {
                return pow(aMean, (1 / aCount));
            }
        }
        return \ZExcel\Calculation\Functions::NaN();
    }


    /**
     * GROWTH
     *
     * Returns values along a predicted emponential trend
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @param    array of mixed        Values of X for which we want to find Y
     * @param    boolean                A logical value specifying whether to force the intersect to equal 0.
     * @return    array of float
     */
    public static function growth(yValues, xValues = [], newValues = [], constt = true)
    {
        var bestFitExponential, xValue, returnArray;
        
        let yValues = \ZExcel\Calculation\Functions::flattenArray(yValues);
        let xValues = \ZExcel\Calculation\Functions::flattenArray(xValues);
        let newValues = \ZExcel\Calculation\Functions::flattenArray(newValues);
        let constt = (is_null(constt)) ? true : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(constt);

        let bestFitExponential = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_EXPONENTIAL, yValues, xValues, constt);
        
        if (empty(newValues)) {
            let newValues = bestFitExponential->getXValues();
        }

        let returnArray = [];
        
        for xValue in newValues {
            let returnArray[0][] = bestFitExponential->getValueOfYForX(xValue);
        }

        return returnArray;
    }


    /**
     * HARMEAN
     *
     * Returns the harmonic mean of a data set. The harmonic mean is the reciprocal of the
     *        arithmetic mean of reciprocals.
     *
     * Excel Function:
     *        HARMEAN(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function harMean()
    {
        var returnValue = null, aArgs, aCount, arg;
        
        // Return value
        let returnValue = \ZExcel\Calculation\Functions::Na();

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        if (call_user_func("self::MiN", aArgs) < 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }
        
        let aCount = 0;
        
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                if (arg <= 0) {
                    return \ZExcel\Calculation\Functions::NaN();
                }
                if (is_null(returnValue)) {
                    let returnValue = (1 / arg);
                } else {
                    let returnValue = returnValue + (1 / arg);
                }
                let aCount = aCount + 1;
            }
        }

        // Return
        if (aCount > 0) {
            return 1 / (returnValue / aCount);
        } else {
            return returnValue;
        }
    }


    /**
     * HYPGEOMDIST
     *
     * Returns the hypergeometric distribution. HYPGEOMDIST returns the probability of a given number of
     * sample successes, given the sample size, population successes, and population size.
     *
     * @param    float        sampleSuccesses        Number of successes in the sample
     * @param    float        sampleNumber            Size of the sample
     * @param    float        populationSuccesses    Number of successes in the population
     * @param    float        populationNumber        Population size
     * @return    float
     *
     */
    public static function hypGeomDist(sampleSuccesses, sampleNumber, populationSuccesses, populationNumber)
    {
        let sampleSuccesses     = floor(\ZExcel\Calculation\Functions::flattenSingleValue(sampleSuccesses));
        let sampleNumber        = floor(\ZExcel\Calculation\Functions::flattenSingleValue(sampleNumber));
        let populationSuccesses = floor(\ZExcel\Calculation\Functions::flattenSingleValue(populationSuccesses));
        let populationNumber    = floor(\ZExcel\Calculation\Functions::flattenSingleValue(populationNumber));

        if ((is_numeric(sampleSuccesses)) && (is_numeric(sampleNumber)) && (is_numeric(populationSuccesses)) && (is_numeric(populationNumber))) {
            if ((sampleSuccesses < 0) || (sampleSuccesses > sampleNumber) || (sampleSuccesses > populationSuccesses)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((sampleNumber <= 0) || (sampleNumber > populationNumber)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((populationSuccesses <= 0) || (populationSuccesses > populationNumber)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return \ZExcel\Calculation\MathTrig::CoMBIN(populationSuccesses, sampleSuccesses) *
                   \ZExcel\Calculation\MathTrig::CoMBIN(populationNumber - populationSuccesses, sampleNumber - sampleSuccesses) /
                   \ZExcel\Calculation\MathTrig::CoMBIN(populationNumber, sampleNumber);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * INTERCEPT
     *
     * Calculates the point at which a line will intersect the y-axis by using existing x-values and y-values.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function intercept(yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getIntersect();
    }


    /**
     * KURT
     *
     * Returns the kurtosis of a data set. Kurtosis characterizes the relative peakedness
     * or flatness of a distribution compared with the normal distribution. Positive
     * kurtosis indicates a relatively peaked distribution. Negative kurtosis indicates a
     * relatively flat distribution.
     *
     * @param    array    Data Series
     * @return    float
     */
    public static function kurt()
    {
        var aArgs, mean, stdDev, count, summer, k, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let mean = call_user_func("self::AVeRAGE", aArgs);
        let stdDev = call_user_func("self::STDeV", aArgs);

        if (stdDev > 0) {
            let count = 0;
            let summer = 0;
            // Loop through arguments
            for k, arg in aArgs {
                if ((is_bool(arg)) &&
                    (!\ZExcel\Calculation\Functions::isMatrixValue(k))) {
                } else {
                    // Is it a numeric value?
                    if ((is_numeric(arg)) && (!is_string(arg))) {
                        let summer = summer + pow(((arg - mean) / stdDev), 4);
                        let count = count + 1;
                    }
                }
            }

            // Return
            if (count > 3) {
                return summer * (count * (count + 1) / ((count - 1) * (count - 2) * (count - 3))) - (3 * pow(count - 1, 2) / ((count - 2) * (count - 3)));
            }
        }
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * LARGE
     *
     * Returns the nth largest value in a data set. You can use this function to
     *        select a value based on its relative standing.
     *
     * Excel Function:
     *        LARGE(value1[,value2[, ...]],entry)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    int            entry            Position (ordered from the largest) in the array or range of data to return
     * @return    float
     *
     */
    public static function large()
    {
        var aArgs, entry, mArgs, arg, count;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let entry = floor(array_pop(aArgs));

        if ((is_numeric(entry)) && (!is_string(entry))) {
            let mArgs = [];
            
            for arg in aArgs {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let mArgs[] = arg;
                }
            }
            
            let count = call_user_func("self::CoUNT", mArgs);
            let entry = floor(entry - 1);
            
            if ((entry < 0) || (entry >= count) || (count == 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            rsort(mArgs);
            
            return mArgs[entry];
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * LINEST
     *
     * Calculates the statistics for a line by using the "least squares" method to calculate a straight line that best fits your data,
     *        and then returns an array that describes the line.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @param    boolean                A logical value specifying whether to force the intersect to equal 0.
     * @param    boolean                A logical value specifying whether to return additional regression statistics.
     * @return    array
     */
    public static function linEst(yValues, xValues = null, constt = true, stats = false)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let constt = (is_null(constt)) ? true : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(constt);
        let stats = (is_null(stats)) ? false : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(stats);
        
        if (is_null(xValues)) {
            let xValues = range(1, count(\ZExcel\Calculation\Functions::flattenArray(yValues)));
        }

        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return 0;
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues, constt);
        
        if (stats) {
            return [
                [
                    bestFitLinear->getSlope(),
                    bestFitLinear->getSlopeSE(),
                    bestFitLinear->getGoodnessOfFit(),
                    bestFitLinear->getF(),
                    bestFitLinear->getSSRegression()
                ],
                [
                    bestFitLinear->getIntersect(),
                    bestFitLinear->getIntersectSE(),
                    bestFitLinear->getStdevOfResiduals(),
                    bestFitLinear->getDFResiduals(),
                    bestFitLinear->getSSResiduals()
                ]
            ];
        } else {
            return [
                bestFitLinear->getSlope(),
                bestFitLinear->getIntersect()
            ];
        }
    }


    /**
     * LOGEST
     *
     * Calculates an exponential curve that best fits the X and Y data series,
     *        and then returns an array that describes the line.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @param    boolean                A logical value specifying whether to force the intersect to equal 0.
     * @param    boolean                A logical value specifying whether to return additional regression statistics.
     * @return    array
     */
    public static function logEst(yValues, xValues = null, constt = true, stats = false)
    {
        var yValueCount, xValueCount, bestFitExponential, value;
        
        let constt = (is_null(constt)) ? true : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(constt);
        let stats = (is_null(stats)) ? false : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(stats);
        
        if (is_null(xValues)) {
            let xValues = range(1, count(\ZExcel\Calculation\Functions::flattenArray(yValues)));
        }

        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        for value in yValues {
            if (value <= 0.0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
        }


        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return 1;
            }
        }

        let bestFitExponential = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_EXPONENTIAL, yValues, xValues, constt);
        
        if (stats) {
            return [
                [
                    bestFitExponential->getSlope(),
                    bestFitExponential->getSlopeSE(),
                    bestFitExponential->getGoodnessOfFit(),
                    bestFitExponential->getF(),
                    bestFitExponential->getSSRegression()
                ],
                [
                    bestFitExponential->getIntersect(),
                    bestFitExponential->getIntersectSE(),
                    bestFitExponential->getStdevOfResiduals(),
                    bestFitExponential->getDFResiduals(),
                    bestFitExponential->getSSResiduals()
                ]
            ];
        } else {
            return [
                bestFitExponential->getSlope(),
                bestFitExponential->getIntersect()
            ];
        }
    }


    /**
     * LOGINV
     *
     * Returns the inverse of the normal cumulative distribution
     *
     * @param    float        probability
     * @param    float        mean
     * @param    float        stdDev
     * @return    float
     *
     * @todo    Try implementing P J Acklam's refinement algorithm for greater
     *            accuracy if I can get my head round the mathematics
     *            (as described at) http://home.online.no/~pjacklam/notes/invnorm/
     */
    public static function logInv(probability, mean, stdDev)
    {
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let mean        = \ZExcel\Calculation\Functions::flattenSingleValue(mean);
        let stdDev      = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);

        if ((is_numeric(probability)) && (is_numeric(mean)) && (is_numeric(stdDev))) {
            if ((probability < 0) || (probability > 1) || (stdDev <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return exp(mean + stdDev * self::NoRMSINV(probability));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * LOGNORMDIST
     *
     * Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed
     * with parameters mean and standard_dev.
     *
     * @param    float        value
     * @param    float        mean
     * @param    float        stdDev
     * @return    float
     */
    public static function logNormDist(value, mean, stdDev)
    {
        let value  = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let mean   = \ZExcel\Calculation\Functions::flattenSingleValue(mean);
        let stdDev = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);

        if ((is_numeric(value)) && (is_numeric(mean)) && (is_numeric(stdDev))) {
            if ((value <= 0) || (stdDev <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return self::NoRMSDIST((log(value) - mean) / stdDev);
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * MAX
     *
     * MAX returns the value of the element of the values passed that has the highest value,
     *        with negative numbers considered smaller than positive numbers.
     *
     * Excel Function:
     *        MAX(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function max()
    {
        var returnValue = null, aArgs, arg;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        for arg in aArgs {
            // Is it a numeric value?
            if (is_numeric(arg) && !is_string(arg)) {
                if (is_null(returnValue) || arg > returnValue) {
                    let returnValue = arg;
                }
            }
        }
        
        if (is_null(returnValue)) {
            return 0;
        }
        
        return returnValue;
    }


    /**
     * MAXA
     *
     * Returns the greatest value in a list of arguments, including numbers, text, and logical values
     *
     * Excel Function:
     *        MAXA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function maxA()
    {
        var returnValue = null, aArgs, arg;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) || (is_bool(arg)) || ((is_string(arg) && (arg != "")))) {
                if (is_bool(arg)) {
                    let arg = (int) arg;
                } else {
                    if (is_string(arg)) {
                        let arg = 0;
                    }
                }
                
                if ((is_null(returnValue)) || (arg > returnValue)) {
                    let returnValue = arg;
                }
            }
        }

        if (is_null(returnValue)) {
            return 0;
        }
        
        return returnValue;
    }


    /**
     * MAXIF
     *
     * Counts the maximum value within a range of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        MAXIF(value1[,value2[, ...]],condition)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @param    string        condition        The criteria that defines which cells will be checked.
     * @return    float
     */
    public static function maxIf(aArgs, condition, sumArgs = [])
    {
        var returnValue = null, testCondition, arg;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(aArgs);
        let sumArgs = \ZExcel\Calculation\Functions::flattenArray(sumArgs);
        
        if (empty(sumArgs)) {
            let sumArgs = aArgs;
        }
        
        let condition = \ZExcel\Calculation\Functions::ifCondition(condition);
        
        // Loop through arguments
        for arg in aArgs {
            if (!is_numeric(arg)) {
                let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
            }
            
            let testCondition = "=" . arg . condition;
            
            if (\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                if ((is_null(returnValue)) || (arg > returnValue)) {
                    let returnValue = arg;
                }
            }
        }

        return returnValue;
    }

    /**
     * MEDIAN
     *
     * Returns the median of the given numbers. The median is the number in the middle of a set of numbers.
     *
     * Excel Function:
     *        MEDIAN(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function median()
    {
        var returnValue = null, mArgs, aArgs, arg, mValueCount;
        
        let returnValue = \ZExcel\Calculation\Functions::NaN();

        let mArgs = [];
        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let mArgs[] = arg;
            }
        }

        let mValueCount = count(mArgs);
        
        if (mValueCount > 0) {
            sort(mArgs, SORT_NUMERIC);
            let mValueCount = mValueCount / 2;
            
            if (mValueCount == floor(mValueCount)) {
                let mValueCount = mValueCount - 1;
                let returnValue = (mArgs[mValueCount] + mArgs[mValueCount]) / 2;
            } else {
                let mValueCount = floor(mValueCount);
                let returnValue = mArgs[mValueCount];
            }
        }

        return returnValue;
    }


    /**
     * MIN
     *
     * MIN returns the value of the element of the values passed that has the smallest value,
     *        with negative numbers considered smaller than positive numbers.
     *
     * Excel Function:
     *        MIN(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function min()
    {
        var returnValue = null, aArgs, arg;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                if ((is_null(returnValue)) || (arg < returnValue)) {
                    let returnValue = arg;
                }
            }
        }

        if (is_null(returnValue)) {
            return 0;
        }
        
        return returnValue;
    }


    /**
     * MINA
     *
     * Returns the smallest value in a list of arguments, including numbers, text, and logical values
     *
     * Excel Function:
     *        MINA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function minA()
    {
        var returnValue = null, aArgs, arg;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) || (is_bool(arg)) || ((is_string(arg) && (arg != "")))) {
                if (is_bool(arg)) {
                    let arg = (int) arg;
                } else {
                    if (is_string(arg)) {
                        let arg = 0;
                    }
                }
                
                if ((is_null(returnValue)) || (arg < returnValue)) {
                    let returnValue = arg;
                }
            }
        }

        if (is_null(returnValue)) {
            return 0;
        }
        
        return returnValue;
    }


    /**
     * MINIF
     *
     * Returns the minimum value within a range of cells that contain numbers within the list of arguments
     *
     * Excel Function:
     *        MINIF(value1[,value2[, ...]],condition)
     *
     * @access    public
     * @category Mathematical and Trigonometric Functions
     * @param    mixed        arg,...        Data values
     * @param    string        condition        The criteria that defines which cells will be checked.
     * @return    float
     */
    public static function minIf(aArgs, condition, sumArgs = [])
    {
        var returnValue = null, testCondition, arg;

        let aArgs = \ZExcel\Calculation\Functions::flattenArray(aArgs);
        let sumArgs = \ZExcel\Calculation\Functions::flattenArray(sumArgs);
        
        if (empty(sumArgs)) {
            let sumArgs = aArgs;
        }
        
        let condition = \ZExcel\Calculation\Functions::ifCondition(condition);

        // Loop through arguments
        for arg in aArgs {
            if (!is_numeric(arg)) {
                let arg = \ZExcel\Calculation::wrapResult(strtoupper(arg));
            }
            
            let testCondition = "=" . arg . condition;
            
            if (\ZExcel\Calculation::getInstance()->_calculateFormulaValue(testCondition)) {
                if ((is_null(returnValue)) || (arg < returnValue)) {
                    let returnValue = arg;
                }
            }
        }

        return returnValue;
    }


    //
    //    Special variant of array_count_values that isn't limited to strings and integers,
    //        but can work with floating point numbers as values
    //
    private static function modeCalc(data)
    {
        var datum, key, value, found;
        array frequencyArray, frequencyList, valueList;
        
        for datum in data {
            let found = false;
            for key, value in frequencyArray {
                if ((string) value["value"] == (string) datum) {
                    let frequencyArray[key]["frequency"] = frequencyArray[key]["frequency"] + 1;
                    let found = true;
                    break;
                }
            }
            if (!found) {
                let frequencyArray[] = [
                    "value"    : datum,
                    "frequency": 1
                ];
            }
        }

        for key, value in frequencyArray {
            let frequencyList[key] = value["frequency"];
            let valueList[key] = value["value"];
        }
        
        array_multisort(frequencyList, SORT_DESC, valueList, SORT_ASC, SORT_NUMERIC, frequencyArray);

        if (frequencyArray[0]["frequency"] == 1) {
            return \ZExcel\Calculation\Functions::Na();
        }
        return frequencyArray[0]["value"];
    }


    /**
     * MODE
     *
     * Returns the most frequently occurring, or repetitive, value in an array or range of data
     *
     * Excel Function:
     *        MODE(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function mode()
    {
        var returnValue = null, aArgs, arg;
        array mArgs;
        
        let returnValue = \ZExcel\Calculation\Functions::Na();

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        let mArgs = [];
        for arg in aArgs {
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let mArgs[] = arg;
            }
        }

        if (!empty(mArgs)) {
            return self::modeCalc(mArgs);
        }

        return returnValue;
    }


    /**
     * NEGBINOMDIST
     *
     * Returns the negative binomial distribution. NEGBINOMDIST returns the probability that
     *        there will be number_f failures before the number_s-th success, when the constant
     *        probability of a success is probability_s. This function is similar to the binomial
     *        distribution, except that the number of successes is fixed, and the number of trials is
     *        variable. Like the binomial, trials are assumed to be independent.
     *
     * @param    float        failures        Number of Failures
     * @param    float        successes        Threshold number of Successes
     * @param    float        probability    Probability of success on each trial
     * @return    float
     *
     */
    public static function negBinNomDist(failures, successes, probability)
    {
        let failures    = floor(\ZExcel\Calculation\Functions::flattenSingleValue(failures));
        let successes   = floor(\ZExcel\Calculation\Functions::flattenSingleValue(successes));
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);

        if ((is_numeric(failures)) && (is_numeric(successes)) && (is_numeric(probability))) {
            if ((failures < 0) || (successes < 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            } else {
                if ((probability < 0) || (probability > 1)) {
                    return \ZExcel\Calculation\Functions::NaN();
                }
            }
            
            if (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_GNUMERIC) {
                if ((failures + successes - 1) <= 0) {
                    return \ZExcel\Calculation\Functions::NaN();
                }
            }
            
            return (\ZExcel\Calculation\MathTrig::CoMBIN(failures + successes - 1, successes - 1)) * (pow(probability, successes)) * (pow(1 - probability, failures));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * NORMDIST
     *
     * Returns the normal distribution for the specified mean and standard deviation. This
     * function has a very wide range of applications in statistics, including hypothesis
     * testing.
     *
     * @param    float        value
     * @param    float        mean        Mean Value
     * @param    float        stdDev        Standard Deviation
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function normDist(value, mean, stdDev, cumulative)
    {
        let value  = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let mean   = \ZExcel\Calculation\Functions::flattenSingleValue(mean);
        let stdDev = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);

        if ((is_numeric(value)) && (is_numeric(mean)) && (is_numeric(stdDev))) {
            if (stdDev < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    return 0.5 * (1 + \ZExcel\Calculation\Engineering::erfVal((value - mean) / (stdDev * sqrt(2))));
                } else {
                    return (1 / (self::SQRT2PI * stdDev)) * exp(0 - (pow(value - mean, 2) / (2 * (stdDev * stdDev))));
                }
            }
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * NORMINV
     *
     * Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.
     *
     * @param    float        value
     * @param    float        mean        Mean Value
     * @param    float        stdDev        Standard Deviation
     * @return    float
     *
     */
    public static function normInv(probability, mean, stdDev)
    {
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let mean        = \ZExcel\Calculation\Functions::flattenSingleValue(mean);
        let stdDev      = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);

        if ((is_numeric(probability)) && (is_numeric(mean)) && (is_numeric(stdDev))) {
            if ((probability < 0) || (probability > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if (stdDev < 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return (self::inverseNcdf(probability) * stdDev) + mean;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * NORMSDIST
     *
     * Returns the standard normal cumulative distribution function. The distribution has
     * a mean of 0 (zero) and a standard deviation of one. Use this function in place of a
     * table of standard normal curve areas.
     *
     * @param    float        value
     * @return    float
     */
    public static function normsDist(value)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);

        return self::NoRMDIST(value, 0, 1, true);
    }


    /**
     * NORMSINV
     *
     * Returns the inverse of the standard normal cumulative distribution
     *
     * @param    float        value
     * @return    float
     */
    public static function normSinv(value)
    {
        return self::NoRMINV(value, 0, 1);
    }


    /**
     * PERCENTILE
     *
     * Returns the nth percentile of values in a range..
     *
     * Excel Function:
     *        PERCENTILE(value1[,value2[, ...]],entry)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    float        entry            Percentile value in the range 0..1, inclusive.
     * @return    float
     */
    public static function percentTile()
    {
        var aArgs, arg, entry, mValueCount, count, index, iBase, iNext, iProportion;
        array mArgs;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let entry = array_pop(aArgs);

        if ((is_numeric(entry)) && (!is_string(entry))) {
            if ((entry < 0) || (entry > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let mArgs = [];
            
            for arg in aArgs {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let mArgs[] = arg;
                }
            }
            
            let mValueCount = count(mArgs);
            
            if (mValueCount > 0) {
                sort(mArgs);
                let count = call_user_func("self::CoUNT", mArgs);
                let index = entry * (count - 1);
                let iBase = floor(index);
                if (index == iBase) {
                    return mArgs[index];
                } else {
                    let iNext = iBase + 1;
                    let iProportion = index - iBase;
                    return mArgs[iBase] + ((mArgs[iNext] - mArgs[iBase]) * iProportion) ;
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * PERCENTRANK
     *
     * Returns the rank of a value in a data set as a percentage of the data set.
     *
     * @param    array of number        An array of, or a reference to, a list of numbers.
     * @param    number                The number whose rank you want to find.
     * @param    number                The number of significant digits for the returned percentage value.
     * @return    float
     */
    public static function percentRank(valueSet, value, significance = 3)
    {
        var key, valueEntry, valueCount, pos, testValue, valueAdjustor;
        
        let valueSet     = \ZExcel\Calculation\Functions::flattenArray(valueSet);
        let value        = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let significance = (is_null(significance)) ? 3 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(significance);

        for key, valueEntry in valueSet {
            if (!is_numeric(valueEntry)) {
                unset(valueSet[key]);
            }
        }
        
        sort(valueSet, SORT_NUMERIC);
        let valueCount = count(valueSet);
        
        if (valueCount == 0) {
            return \ZExcel\Calculation\Functions::NaN();
        }

        let valueAdjustor = valueCount - 1;
        
        if ((value < valueSet[0]) || (value > valueSet[valueAdjustor])) {
            return \ZExcel\Calculation\Functions::Na();
        }

        let pos = array_search(value, valueSet);
        
        if (pos === false) {
            let pos = 0;
            let testValue = valueSet[0];
            while (testValue < value) {
                let pos = pos + 1;
                let testValue = valueSet[pos];
            }
            let pos = pos - 1;
            let pos = pos + (((value - valueSet[pos]) / (testValue - valueSet[pos])));
        }

        return round(pos / valueAdjustor, significance);
    }


    /**
     * PERMUT
     *
     * Returns the number of permutations for a given number of objects that can be
     *        selected from number objects. A permutation is any set or subset of objects or
     *        events where internal order is significant. Permutations are different from
     *        combinations, for which the internal order is not significant. Use this function
     *        for lottery-style probability calculations.
     *
     * @param    int        numObjs    Number of different objects
     * @param    int        numInSet    Number of objects in each permutation
     * @return    int        Number of permutations
     */
    public static function permut(numObjs, numInSet)
    {
        let numObjs  = \ZExcel\Calculation\Functions::flattenSingleValue(numObjs);
        let numInSet = \ZExcel\Calculation\Functions::flattenSingleValue(numInSet);

        if ((is_numeric(numObjs)) && (is_numeric(numInSet))) {
            let numInSet = floor(numInSet);
            if (numObjs < numInSet) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return round(\ZExcel\Calculation\MathTrig::FaCT(numObjs) / \ZExcel\Calculation\MathTrig::FaCT(numObjs - numInSet));
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * POISSON
     *
     * Returns the Poisson distribution. A common application of the Poisson distribution
     * is predicting the number of events over a specific time, such as the number of
     * cars arriving at a toll plaza in 1 minute.
     *
     * @param    float        value
     * @param    float        mean        Mean Value
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function poisson(value, mean, cumulative)
    {
        var summer, i;
        
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let mean  = \ZExcel\Calculation\Functions::flattenSingleValue(mean);

        if ((is_numeric(value)) && (is_numeric(mean))) {
            if ((value < 0) || (mean <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    let summer = 0;
                    for i in range(0, floor(value)) {
                        let summer = summer + pow(mean, i) / \ZExcel\Calculation\MathTrig::FaCT(i);
                    }
                    return exp(0-mean) * summer;
                } else {
                    return (exp(0-mean) * pow(mean, value)) / \ZExcel\Calculation\MathTrig::FaCT(value);
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * QUARTILE
     *
     * Returns the quartile of a data set.
     *
     * Excel Function:
     *        QUARTILE(value1[,value2[, ...]],entry)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    int            entry            Quartile value in the range 1..3, inclusive.
     * @return    float
     */
    public static function quartile()
    {
        var aArgs, entry;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let entry = floor(array_pop(aArgs));

        if ((is_numeric(entry)) && (!is_string(entry))) {
            let entry = entry / 4;
            if ((entry < 0) || (entry > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            return call_user_func("self::PeRCENTILE", aArgs, entry);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * RANK
     *
     * Returns the rank of a number in a list of numbers.
     *
     * @param    number                The number whose rank you want to find.
     * @param    array of number        An array of, or a reference to, a list of numbers.
     * @param    mixed                Order to sort the values in the value set
     * @return    float
     */
    public static function rank(value, valueSet, order = 0)
    {
        var key, valueEntry, pos;
        
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let valueSet = \ZExcel\Calculation\Functions::flattenArray(valueSet);
        let order = (is_null(order)) ? 0 : (int) \ZExcel\Calculation\Functions::flattenSingleValue(order);

        for key, valueEntry in valueSet {
            if (!is_numeric(valueEntry)) {
                unset(valueSet[key]);
            }
        }

        if (order == 0) {
            rsort(valueSet, SORT_NUMERIC);
        } else {
            sort(valueSet, SORT_NUMERIC);
        }
        
        let pos = array_search(value, valueSet);
        
        if (pos === false) {
            return \ZExcel\Calculation\Functions::Na();
        }
        
        let pos = pos + 1;

        return pos;
    }


    /**
     * RSQ
     *
     * Returns the square of the Pearson product moment correlation coefficient through data points in known_y's and known_x's.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function rsq(yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getGoodnessOfFit();
    }


    /**
     * SKEW
     *
     * Returns the skewness of a distribution. Skewness characterizes the degree of asymmetry
     * of a distribution around its mean. Positive skewness indicates a distribution with an
     * asymmetric tail extending toward more positive values. Negative skewness indicates a
     * distribution with an asymmetric tail extending toward more negative values.
     *
     * @param    array    Data Series
     * @return    float
     */
    public static function skew()
    {
        var aArgs, mean, k, arg, count, summer, stdDev;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let mean = call_user_func("self::AVeRAGE", aArgs);
        let stdDev = call_user_func("self::STDeV", aArgs);

        let count = 0;
        let summer = 0;
        
        // Loop through arguments
        for k, arg in aArgs {
            if ((is_bool(arg)) &&
                (!\ZExcel\Calculation\Functions::isMatrixValue(k))) {
            } else {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let summer = summer + pow(((arg - mean) / stdDev), 3);
                    let count = count + 1;
                }
            }
        }

        if (count > 2) {
            return summer * (count / ((count - 1) * (count - 2)));
        }
        
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * SLOPE
     *
     * Returns the slope of the linear regression line through data points in known_y's and known_x's.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function slope(yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getSlope();
    }


    /**
     * SMALL
     *
     * Returns the nth smallest value in a data set. You can use this function to
     *        select a value based on its relative standing.
     *
     * Excel Function:
     *        SMALL(value1[,value2[, ...]],entry)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    int            entry            Position (ordered from the smallest) in the array or range of data to return
     * @return    float
     */
    public static function small()
    {
        var aArgs, arg, count, entry;
        array mArgs;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let entry = array_pop(aArgs);

        if ((is_numeric(entry)) && (!is_string(entry))) {
            let mArgs = [];
            
            for arg in aArgs {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let mArgs[] = arg;
                }
            }
            
            let count = call_user_func("self::CoUNT", mArgs);
            let entry = floor(entry - 1);
            
            if ((entry < 0) || (entry >= count) || (count == 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            sort(mArgs);
            
            return mArgs[entry];
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * STANDARDIZE
     *
     * Returns a normalized value from a distribution characterized by mean and standard_dev.
     *
     * @param    float    value        Value to normalize
     * @param    float    mean        Mean Value
     * @param    float    stdDev        Standard Deviation
     * @return    float    Standardized value
     */
    public static function standardize(value, mean, stdDev)
    {
        let value  = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let mean   = \ZExcel\Calculation\Functions::flattenSingleValue(mean);
        let stdDev = \ZExcel\Calculation\Functions::flattenSingleValue(stdDev);

        if ((is_numeric(value)) && (is_numeric(mean)) && (is_numeric(stdDev))) {
            if (stdDev <= 0) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            return (value - mean) / stdDev ;
        }
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * STDEV
     *
     * Estimates standard deviation based on a sample. The standard deviation is a measure of how
     *        widely values are dispersed from the average value (the mean).
     *
     * Excel Function:
     *        STDEV(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function stdev()
    {
        var aArgs, returnValue, aMean, aCount, k, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());

        let aMean = call_user_func("self::AVeRAGE", aArgs);
        if (!is_null(aMean)) {
            let aCount = -1;
            for k, arg in aArgs {
                if ((is_bool(arg)) &&
                    ((!\ZExcel\Calculation\Functions::isCellValue(k)) || (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                    let arg = (int) arg;
                }
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    if (is_null(returnValue)) {
                        let returnValue = pow((arg - aMean), 2);
                    } else {
                        let returnValue = returnValue + pow((arg - aMean), 2);
                    }
                    let aCount = aCount + 1;
                }
            }

            // Return
            if ((aCount > 0) && (returnValue >= 0)) {
                return sqrt(returnValue / aCount);
            }
        }
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * STDEVA
     *
     * Estimates standard deviation based on a sample, including numbers, text, and logical values
     *
     * Excel Function:
     *        STDEVA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function stdeva()
    {
        var aArgs, returnValue, aMean, aCount, k, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aMean = call_user_func("self::AVeRAGE", aArgs);
        
        if (!is_null(aMean)) {
            let aCount = -1;
            for k, arg in aArgs {
                if ((is_bool(arg)) && (!\ZExcel\Calculation\Functions::isMatrixValue(k))) {
                } else {
                    // Is it a numeric value?
                    if (is_numeric(arg) || is_bool(arg) || (is_string(arg) && arg != "")) {
                        if (is_bool(arg)) {
                            let arg = (int) arg;
                        } else {
                            if (is_string(arg)) {
                                let arg = 0;
                            }
                        }
                        
                        if (is_null(returnValue)) {
                            let returnValue = pow((arg - aMean), 2);
                        } else {
                            let returnValue = returnValue + pow((arg - aMean), 2);
                        }
                        
                        let aCount = aCount + 1;
                    }
                }
            }

            if (aCount > 0 && returnValue >= 0) {
                return sqrt(returnValue / aCount);
            }
        }
        
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * STDEVP
     *
     * Calculates standard deviation based on the entire population
     *
     * Excel Function:
     *        STDEVP(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function stdevp()
    {
        var aArgs, returnValue, aMean, aCount, k, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aMean = call_user_func("self::AVeRAGE", aArgs);
        
        if (!is_null(aMean)) {
            let aCount = 0;
            for k, arg in aArgs {
                if ((is_bool(arg)) &&
                    ((!\ZExcel\Calculation\Functions::isCellValue(k)) || (\ZExcel\Calculation\Functions::getCompatibilityMode() == \ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE))) {
                    let arg = (int) arg;
                }
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    if (is_null(returnValue)) {
                        let returnValue = pow((arg - aMean), 2);
                    } else {
                        let returnValue = returnValue + pow((arg - aMean), 2);
                    }
                    let aCount = aCount + 1;
                }
            }

            if ((aCount > 0) && (returnValue >= 0)) {
                return sqrt(returnValue / aCount);
            }
        }
        
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * STDEVPA
     *
     * Calculates standard deviation based on the entire population, including numbers, text, and logical values
     *
     * Excel Function:
     *        STDEVPA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function stdevpa()
    {
        var aArgs, returnValue, aMean, aCount, k, arg;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aMean = call_user_func("self::AVeRAGE", aArgs);
        
        if (!is_null(aMean)) {
            let aCount = 0;
            for k, arg in aArgs {
                if ((is_bool(arg)) &&
                    (!\ZExcel\Calculation\Functions::isMatrixValue(k))) {
                } else {
                    // Is it a numeric value?
                    if (is_numeric(arg) || is_bool(arg) || (is_string(arg) && arg != "")) {
                        if (is_bool(arg)) {
                            let arg = (int) arg;
                        } else {
                            if (is_string(arg)) {
                                let arg = 0;
                            }
                        }
                        
                        if (is_null(returnValue)) {
                            let returnValue = pow((arg - aMean), 2);
                        } else {
                            let returnValue = returnValue + pow((arg - aMean), 2);
                        }
                        
                        let aCount = aCount + 1;
                    }
                }
            }

            if ((aCount > 0) && (returnValue >= 0)) {
                return sqrt(returnValue / aCount);
            }
        }
        
        return \ZExcel\Calculation\Functions::DiV0();
    }


    /**
     * STEYX
     *
     * Returns the standard error of the predicted y-value for each x in the regression.
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @return    float
     */
    public static function sTeyx(yValues, xValues)
    {
        var yValueCount, xValueCount, bestFitLinear;
        
        let self::array1 = yValues;
        let self::array2 = xValues;
        
        if (!self::checkTrendArrays()) {
            return \ZExcel\Calculation\Functions::VaLUE();
        }
        
        let yValues = self::array1;
        let xValues = self::array2;
        
        let yValueCount = count(yValues);
        let xValueCount = count(xValues);

        if ((yValueCount == 0) || (yValueCount != xValueCount)) {
            return \ZExcel\Calculation\Functions::Na();
        } else {
            if (yValueCount == 1) {
                return \ZExcel\Calculation\Functions::DiV0();
            }
        }
        
        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues);
        
        return bestFitLinear->getStdevOfResiduals();
    }


    /**
     * TDIST
     *
     * Returns the probability of Student's T distribution.
     *
     * @param    float        value            Value for the function
     * @param    float        degrees        degrees of freedom
     * @param    float        tails            number of tails (1 or 2)
     * @return    float
     */
    public static function tDist(value, degrees, tails)
    {
        var tterm, ttheta, tc, ts, tsum, ti, tValue;
        
        let value        = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let degrees    = floor(\ZExcel\Calculation\Functions::flattenSingleValue(degrees));
        let tails        = floor(\ZExcel\Calculation\Functions::flattenSingleValue(tails));

        if ((is_numeric(value)) && (is_numeric(degrees)) && (is_numeric(tails))) {
            if ((value < 0) || (degrees < 1) || (tails < 1) || (tails > 2)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            //    tdist, which finds the probability that corresponds to a given value
            //    of t with k degrees of freedom. This algorithm is translated from a
            //    pascal function on p81 of "Statistical Computing in Pascal" by D
            //    Cooke, A H Craven & G M Clark (1985: Edward Arnold (Pubs.) Ltd:
            //    London). The above Pascal algorithm is itself a translation of the
            //    fortran algoritm "AS 3" by B E Cooper of the Atlas Computer
            //    Laboratory as reported in (among other places) "Applied Statistics
            //    Algorithms", editied by P Griffiths and I D Hill (1985; Ellis
            //    Horwood Ltd.; W. Sussex, England).
            let tterm = degrees;
            let ttheta = atan2(value, sqrt(tterm));
            let tc = cos(ttheta);
            let ts = sin(ttheta);
            let tsum = 0;

            if ((degrees % 2) == 1) {
                let ti = 3;
                let tterm = tc;
            } else {
                let ti = 2;
                let tterm = 1;
            }

            let tsum = tterm;
            
            while (ti < degrees) {
                let tterm = tterm * tc * tc * (ti - 1) / ti;
                let tsum = tsum + tterm;
                let ti = ti + 2;
            }
            
            let tsum = tsum * ts;
            
            if ((degrees % 2) == 1) {
                let tsum = \ZExcel\Calculation\Functions::M_2DIVPI * (tsum + ttheta);
            }
            
            let tValue = 0.5 * (1 + tsum);
            
            if (tails == 1) {
                return 1 - abs(tValue);
            } else {
                return 1 - abs((1 - tValue) - tValue);
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * TINV
     *
     * Returns the one-tailed probability of the chi-squared distribution.
     *
     * @param    float        probability    Probability for the function
     * @param    float        degrees        degrees of freedom
     * @return    float
     */
    public static function tinv(probability, degrees)
    {
        var xLo, xHi, x, xNew, dx, i, result, error;
        
        let probability = \ZExcel\Calculation\Functions::flattenSingleValue(probability);
        let degrees     = floor(\ZExcel\Calculation\Functions::flattenSingleValue(degrees));

        if ((is_numeric(probability)) && (is_numeric(degrees))) {
            let xLo = 100;
            let xHi = 0;

            let x = 1;
            let xNew = 1;
            let dx = 1;
            let i = 0;

            while ((abs(dx) > floatval(\ZExcel\Calculation\Functions::PRECISION)) && (i < \ZExcel\Calculation\Functions::MAX_ITERATIONS)) {
                let i = i + 1;
                
                // Apply Newton-Raphson step
                let result = call_user_func("self::TDiST", x, degrees, 2);
                let error = result - probability;
                
                if (error == 0.0) {
                    let dx = 0;
                } else {
                    if (error < 0.0) {
                        let xLo = x;
                    } else {
                        let xHi = x;
                    }
                }
                
                // Avoid division by zero
                if (result != 0.0) {
                    let dx = error / result;
                    let xNew = x - dx;
                }
                // If the NR fails to converge (which for example may be the
                // case if the initial guess is too rough) we apply a bisection
                // step to determine a more narrow interval around the root.
                if ((xNew < xLo) || (xNew > xHi) || (result == 0.0)) {
                    let xNew = (xLo + xHi) / 2;
                    let dx = xNew - x;
                }
                
                let x = xNew;
            }
            if (i == \ZExcel\Calculation\Functions::MAX_ITERATIONS) {
                return \ZExcel\Calculation\Functions::Na();
            }
            
            return round(x, 12);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * TREND
     *
     * Returns values along a linear trend
     *
     * @param    array of mixed        Data Series Y
     * @param    array of mixed        Data Series X
     * @param    array of mixed        Values of X for which we want to find Y
     * @param    boolean                A logical value specifying whether to force the intersect to equal 0.
     * @return    array of float
     */
    public static function trend(yValues, xValues = [], newValues = [], constt = true)
    {
        var bestFitLinear, xValue;
        array returnArray;
        
        let yValues = \ZExcel\Calculation\Functions::flattenArray(yValues);
        let xValues = \ZExcel\Calculation\Functions::flattenArray(xValues);
        let newValues = \ZExcel\Calculation\Functions::flattenArray(newValues);
        let constt = (is_null(constt)) ? true : (boolean) \ZExcel\Calculation\Functions::flattenSingleValue(constt);

        let bestFitLinear = \ZExcel\Shared\TrendClass::calculate(\ZExcel\Shared\TrendClass::TREND_LINEAR, yValues, xValues, constt);
        
        if (empty(newValues)) {
            let newValues = bestFitLinear->getXValues();
        }

        let returnArray = [];
        
        for xValue in newValues {
            let returnArray[0][] = bestFitLinear->getValueOfYForX(xValue);
        }

        return returnArray;
    }


    /**
     * TRIMMEAN
     *
     * Returns the mean of the interior of a data set. TRIMMEAN calculates the mean
     *        taken by excluding a percentage of data points from the top and bottom tails
     *        of a data set.
     *
     * Excel Function:
     *        TRIMEAN(value1[,value2[, ...]], discard)
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @param    float        discard        Percentage to discard
     * @return    float
     */
    public static function trimMean()
    {
        var aArgs, percent, arg, discard, i;
        array mArgs;
        
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());

        // Calculate
        let percent = array_pop(aArgs);

        if ((is_numeric(percent)) && (!is_string(percent))) {
            if ((percent < 0) || (percent > 1)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            
            let mArgs = [];
            
            for arg in aArgs {
                // Is it a numeric value?
                if ((is_numeric(arg)) && (!is_string(arg))) {
                    let mArgs[] = arg;
                }
            }
            
            let discard = floor(call_user_func("self::CoUNT", mArgs) * percent / 2);
            sort(mArgs);
            
            for i in range(0, discard - 1) {
                array_pop(mArgs);
                array_shift(mArgs);
            }
            
            return call_user_func("self::AVeRAGE", aArgs);
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * VARFunc
     *
     * Estimates variance based on a sample.
     *
     * Excel Function:
     *        VAR(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function varFunc()
    {
        var returnValue = null, summerA, summerB, aArgs, aCount, arg;
        
        let returnValue = \ZExcel\Calculation\Functions::DiV0();

        let summerA = 0;
        let summerB = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        let aCount = 0;
        
        for arg in aArgs {
            if (is_bool(arg)) {
                let arg = (int) arg;
            }
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let summerA = summerA + (arg * arg);
                let summerB = summerB + arg;
                let aCount = aCount + 1;
            }
        }

        if (aCount > 1) {
            let summerA = summerA * aCount;
            let summerB = summerB * summerB;
            let returnValue = (summerA - summerB) / (aCount * (aCount - 1));
        }
        
        return returnValue;
    }


    /**
     * VARA
     *
     * Estimates variance based on a sample, including numbers, text, and logical values
     *
     * Excel Function:
     *        VARA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function vara()
    {
        var returnValue = null, summerA, summerB, aArgs, aCount, k, arg;
        
        let returnValue = \ZExcel\Calculation\Functions::DiV0();

        let summerA = 0;
        let summerB = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aCount = 0;
        
        for k, arg in aArgs {
            if (is_string(arg) && \ZExcel\Calculation\Functions::isValue(k)) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } else {
                if (is_string(arg) && !\ZExcel\Calculation\Functions::isMatrixValue(k)) {
                } else {
                    // Is it a numeric value?
                    if (is_numeric(arg) || is_bool(arg) || (is_string(arg) && arg != "")) {
                        if (is_bool(arg)) {
                            let arg = (int) arg;
                        } else {
                            if (is_string(arg)) {
                                let arg = 0;
                            }
                        }
                        
                        let summerA = summerA + (arg * arg);
                        let summerB = summerB + arg;
                        let aCount = aCount + 1;
                    }
                }
            }
        }

        if (aCount > 1) {
            let summerA = summerA * aCount;
            let summerB = summerB * summerB;
            let returnValue = (summerA - summerB) / (aCount * (aCount - 1));
        }
        return returnValue;
    }


    /**
     * VARP
     *
     * Calculates variance based on the entire population
     *
     * Excel Function:
     *        VARP(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function varp()
    {
        var returnValue = null, summerA, summerB, aArgs, aCount, arg;
        
        // Return value
        let returnValue = \ZExcel\Calculation\Functions::DiV0();

        let summerA = 0;
        let summerB = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArray(func_get_args());
        let aCount = 0;
        for arg in aArgs {
            if (is_bool(arg)) {
                let arg = (int) arg;
            }
            // Is it a numeric value?
            if ((is_numeric(arg)) && (!is_string(arg))) {
                let summerA = summerA + (arg * arg);
                let summerB = summerB + arg;
                let aCount = aCount + 1;
            }
        }

        if (aCount > 0) {
            let summerA = summerA * aCount;
            let summerB = summerB * summerB;
            let returnValue = (summerA - summerB) / (aCount * aCount);
        }
        return returnValue;
    }


    /**
     * VARPA
     *
     * Calculates variance based on the entire population, including numbers, text, and logical values
     *
     * Excel Function:
     *        VARPA(value1[,value2[, ...]])
     *
     * @access    public
     * @category Statistical Functions
     * @param    mixed        arg,...        Data values
     * @return    float
     */
    public static function varpa()
    {
        var returnValue = null, summerA, summerB, aArgs, aCount, k, arg;
        
        let returnValue = \ZExcel\Calculation\Functions::DiV0();

        let summerA = 0;
        let summerB = 0;

        // Loop through arguments
        let aArgs = \ZExcel\Calculation\Functions::flattenArrayIndexed(func_get_args());
        let aCount = 0;
        
        for k, arg in aArgs {
            if ((is_string(arg)) &&
                (\ZExcel\Calculation\Functions::isValue(k))) {
                return \ZExcel\Calculation\Functions::VaLUE();
            } else {
                if (!is_string(arg) || \ZExcel\Calculation\Functions::isMatrixValue(k)) {
                    // Is it a numeric value?
                    if (is_numeric(arg) || is_bool(arg) || (is_string(arg) && arg != "")) {
                        if (is_bool(arg)) {
                            let arg = (int) arg;
                        } else {
                            if (is_string(arg)) {
                                let arg = 0;
                            }
                        }
                        
                        let summerA = summerA + (arg * arg);
                        let summerB = summerB + arg;
                        let aCount = aCount + 1;
                    }
                }
            }
        }

        if (aCount > 0) {
            let summerA = summerA * aCount;
            let summerB = summerB * summerB;
            let returnValue = (summerA - summerB) / (aCount * aCount);
        }
        
        return returnValue;
    }


    /**
     * WEIBULL
     *
     * Returns the Weibull distribution. Use this distribution in reliability
     * analysis, such as calculating a device's mean time to failure.
     *
     * @param    float        value
     * @param    float        alpha        Alpha Parameter
     * @param    float        beta        Beta Parameter
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function weibull(value, alpha, beta, cumulative)
    {
        let value = \ZExcel\Calculation\Functions::flattenSingleValue(value);
        let alpha = \ZExcel\Calculation\Functions::flattenSingleValue(alpha);
        let beta  = \ZExcel\Calculation\Functions::flattenSingleValue(beta);

        if ((is_numeric(value)) && (is_numeric(alpha)) && (is_numeric(beta))) {
            if ((value < 0) || (alpha <= 0) || (beta <= 0)) {
                return \ZExcel\Calculation\Functions::NaN();
            }
            if ((is_numeric(cumulative)) || (is_bool(cumulative))) {
                if (cumulative) {
                    return 1 - exp(0 - pow(value / beta, alpha));
                } else {
                    return (alpha / pow(beta, alpha)) * pow(value, alpha - 1) * exp(0 - pow(value / beta, alpha));
                }
            }
        }
        
        return \ZExcel\Calculation\Functions::VaLUE();
    }


    /**
     * ZTEST
     *
     * Returns the Weibull distribution. Use this distribution in reliability
     * analysis, such as calculating a device's mean time to failure.
     *
     * @param    float        dataSet
     * @param    float        m0        Alpha Parameter
     * @param    float        sigma    Beta Parameter
     * @param    boolean        cumulative
     * @return    float
     *
     */
    public static function ztest(dataSet, m0, sigma = null)
    {
        var n;
        
        let dataSet = \ZExcel\Calculation\Functions::flattenArrayIndexed(dataSet);
        let m0      = \ZExcel\Calculation\Functions::flattenSingleValue(m0);
        let sigma   = \ZExcel\Calculation\Functions::flattenSingleValue(sigma);

        if (is_null(sigma)) {
            let sigma = call_user_func("self::STDeV", dataSet);
        }
        
        let n = count(dataSet);

        return 1 - self::NoRMSDIST((call_user_func("self::AVeRAGE", dataSet) - m0) / (sigma / sqrt(n)));
    }
}
