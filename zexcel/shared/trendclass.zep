namespace ZExcel\Shared;

class TrendClass
{
    const TREND_LINEAR            = "Linear";
    const TREND_LOGARITHMIC       = "Logarithmic";
    const TREND_EXPONENTIAL       = "Exponential";
    const TREND_POWER             = "Power";
    const TREND_POLYNOMIAL_2      = "Polynomial_2";
    const TREND_POLYNOMIAL_3      = "Polynomial_3";
    const TREND_POLYNOMIAL_4      = "Polynomial_4";
    const TREND_POLYNOMIAL_5      = "Polynomial_5";
    const TREND_POLYNOMIAL_6      = "Polynomial_6";
    const TREND_BEST_FIT          = "Bestfit";
    const TREND_BEST_FIT_NO_POLY  = "Bestfit_no_Polynomials";

    /**
     * Names of the best-fit trend analysis methods
     *
     * @var string[]
     **/
    private static trendTypes = [
        self::TREND_LINEAR,
        self::TREND_LOGARITHMIC,
        self::TREND_EXPONENTIAL,
        self::TREND_POWER
    ];

    /**
     * Names of the best-fit trend polynomial orders
     *
     * @var string[]
     **/
    private static trendTypePolynomialOrders = [
        self::TREND_POLYNOMIAL_2,
        self::TREND_POLYNOMIAL_3,
        self::TREND_POLYNOMIAL_4,
        self::TREND_POLYNOMIAL_5,
        self::TREND_POLYNOMIAL_6
    ];

    /**
     * Cached results for each method when trying to identify which provides the best fit
     *
     * @var PHPExcel_Best_Fit[]
     **/
    private static trendCache = [];


    public static function calculate(var trendType = self::TREND_BEST_FIT, var yValues, var xValues = [], var constt = true)
    {
        var nY, nX, key, className, order, trendMethod, bestFit, bestFitValue, bestFitType;
        
        //    Calculate number of points in each dataset
        let nY = count(yValues);
        let nX = count(xValues);

        //    Define X Values if necessary
        if (nX == 0) {
            let xValues = range(1, nY);
        } else {
            if (nY != nX) {
                //    Ensure both arrays of points are the same size
                trigger_error("trend(): Number of elements in coordinate arrays do not match.", E_USER_ERROR);
            }
        }
        
        let key = md5(trendType . constt . serialize(yValues) . serialize(xValues));
        
        //    Determine which trend method has been requested
        switch (trendType) {
            //    Instantiate and return the class for the requested trend method
            case self::TREND_LINEAR:
            case self::TREND_LOGARITHMIC:
            case self::TREND_EXPONENTIAL:
            case self::TREND_POWER:
                if (!isset(self::trendCache[key])) {
                    let className = "\\ZExcel\\Shared\\Trend\\" . trendType . "BestFit";
                    let self::trendCache[key] = new {className}(yValues, xValues, constt);
                }
                return self::trendCache[key];
            case self::TREND_POLYNOMIAL_2:
            case self::TREND_POLYNOMIAL_3:
            case self::TREND_POLYNOMIAL_4:
            case self::TREND_POLYNOMIAL_5:
            case self::TREND_POLYNOMIAL_6:
                if (!isset(self::trendCache[key])) {
                    let order = substr(trendType, -1);
                    let self::trendCache[key] = new \ZExcel\Shared\Trend\PolynomialBestFit(order, yValues, xValues, constt);
                }
                return self::trendCache[key];
            case self::TREND_BEST_FIT:
            case self::TREND_BEST_FIT_NO_POLY:
                //    If the request is to determine the best fit regression, then we test each trend line in turn
                //    Start by generating an instance of each available trend method
                for trendMethod in self::trendTypes {
                    let className = "\\ZExcel\\Shared\\Trend\\" . trendMethod . "BestFit";
                    let bestFit[trendMethod] = new {className}(yValues, xValues, constt);
                    let bestFitValue[trendMethod] = bestFit[trendMethod]->getGoodnessOfFit();
                }
                if (trendType != self::TREND_BEST_FIT_NO_POLY) {
                    for trendMethod in self::trendTypePolynomialOrders {
                        let order = substr(trendMethod, -1);
                        let bestFit[trendMethod] = new \ZExcel\Shared\Trend\PolynomialBestFit(order, yValues, xValues, constt);
                        if (bestFit[trendMethod]->getError()) {
                            unset(bestFit[trendMethod]);
                        } else {
                            let bestFitValue[trendMethod] = bestFit[trendMethod]->getGoodnessOfFit();
                        }
                    }
                }
                //    Determine which of our trend lines is the best fit, and then we return the instance of that trend class
                arsort(bestFitValue);
                
                let bestFitType = key(bestFitValue);
                
                return bestFit[bestFitType];
            default:
                return false;
        }
    }
}
