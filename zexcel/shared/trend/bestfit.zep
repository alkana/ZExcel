namespace ZExcel\Shared\Trend;

/**
 * PHPExcel_Best_Fit
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Shared_Trend
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class BestFit
{
    /**
     * Indicator flag for a calculation error
     *
     * @var    boolean
     **/
    protected error = false;

    /**
     * Algorithm type to use for best-fit
     *
     * @var    string
     **/
    protected bestFitType = "undetermined";

    /**
     * Number of entries in the sets of x- and y-value arrays
     *
     * @var    int
     **/
    protected valueCount = 0;

    /**
     * X-value dataseries of values
     *
     * @var    float[]
     **/
    protected xValues = [];

    /**
     * Y-value dataseries of values
     *
     * @var    float[]
     **/
    protected yValues = [];

    /**
     * Flag indicating whether values should be adjusted to Y=0
     *
     * @var    boolean
     **/
    protected adjustToZero = false;

    /**
     * Y-value series of best-fit values
     *
     * @var    float[]
     **/
    protected yBestFitValues = [];

    protected goodnessOfFit = 1;

    protected stdevOfResiduals = 0;

    protected covariance = 0;

    protected correlation = 0;

    protected SSRegression = 0;

    protected SSResiduals = 0;

    protected DFResiduals = 0;

    protected f = 0;

    protected slope = 0;

    protected slopeSE = 0;

    protected intersect = 0;

    protected intersectSE = 0;

    protected xOffset = 0;

    protected yOffset = 0;

    /**
     * Define the regression
     *
     * @param    float[]        yValues    The set of Y-values for this regression
     * @param    float[]        xValues    The set of X-values for this regression
     * @param    boolean        constt
     */
    public function __construct(var yValues, var xValues = [], boolean constt = true)
    {
        var nY, nX;
        
        //    Calculate number of points
        let nY = count(yValues);
        let nX = count(xValues);

        //    Define X Values if necessary
        if (nX == 0) {
            let xValues = range(1, nY);
        } else {
            if (nY != nX) {
                //    Ensure both arrays of points are the same size
                let this->error = true;
            } else {
                let this->valueCount = nY;
                let this->xValues = xValues;
                let this->yValues = yValues;
            }
        }
    }

    public function getError() -> boolean
    {
        return this->error;
    }


    public function getBestFitType() -> string
    {
        return this->bestFitType;
    }

    /**
     * Return the Y-Value for a specified value of X
     *
     * @param     float        xValue            X-Value
     * @return     float                        Y-Value
     */
    public function getValueOfYForX(float xValue) -> boolean
    {
        return false;
    }

    /**
     * Return the X-Value for a specified value of Y
     *
     * @param     float        yValue            Y-Value
     * @return     float                        X-Value
     */
    public function getValueOfXForY(float yValue) -> boolean
    {
        return false;
    }

    /**
     * Return the original set of X-Values
     *
     * @return     float[]                X-Values
     */
    public function getXValues() -> array
    {
        return this->xValues;
    }

    /**
     * Return the Equation of the best-fit line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     */
    public function getEquation(int dp = 0) -> boolean
    {
        return false;
    }

    /**
     * Return the Slope of the line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     */
    public function getSlope(int dp = 0)
    {
        if (dp != 0) {
            return round(this->slope, dp);
        }
        
        return this->slope;
    }

    /**
     * Return the standard error of the Slope
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     */
    public function getSlopeSE(int dp = 0)
    {
        if (dp != 0) {
            return round(this->slopeSE, dp);
        }
        return this->slopeSE;
    }

    /**
     * Return the Value of X where it intersects Y = 0
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     */
    public function getIntersect(int dp = 0)
    {
        if (dp != 0) {
            return round(this->intersect, dp);
        }
        return this->intersect;
    }

    /**
     * Return the standard error of the Intersect
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     */
    public function getIntersectSE(int dp = 0)
    {
        if (dp != 0) {
            return round(this->intersectSE, dp);
        }
        return this->intersectSE;
    }

    /**
     * Return the goodness of fit for this regression
     *
     * @param     int        dp        Number of places of decimal precision to return
     * @return     float
     */
    public function getGoodnessOfFit(int dp = 0)
    {
        if (dp != 0) {
            return round(this->goodnessOfFit, dp);
        }
        return this->goodnessOfFit;
    }

    public function getGoodnessOfFitPercent(int dp = 0)
    {
        if (dp != 0) {
            return round(this->goodnessOfFit * 100, dp);
        }
        return this->goodnessOfFit * 100;
    }

    /**
     * Return the standard deviation of the residuals for this regression
     *
     * @param     int        dp        Number of places of decimal precision to return
     * @return     float
     */
    public function getStdevOfResiduals(int dp = 0)
    {
        if (dp != 0) {
            return round(this->stdevOfResiduals, dp);
        }
        return this->stdevOfResiduals;
    }

    public function getSSRegression(int dp = 0)
    {
        if (dp != 0) {
            return round(this->SSRegression, dp);
        }
        return this->SSRegression;
    }

    public function getSSResiduals(int dp = 0)
    {
        if (dp != 0) {
            return round(this->SSResiduals, dp);
        }
        return this->SSResiduals;
    }

    public function getDFResiduals(int dp = 0)
    {
        if (dp != 0) {
            return round(this->DFResiduals, dp);
        }
        return this->DFResiduals;
    }

    public function getF(int dp = 0)
    {
        if (dp != 0) {
            return round(this->f, dp);
        }
        return this->f;
    }

    public function getCovariance(int dp = 0)
    {
        if (dp != 0) {
            return round(this->covariance, dp);
        }
        return this->covariance;
    }

    public function getCorrelation(int dp = 0)
    {
        if (dp != 0) {
            return round(this->correlation, dp);
        }
        return this->correlation;
    }

    public function getYBestFitValues()
    {
        return this->yBestFitValues;
    }

    protected function calculateGoodnessOfFit(var sumX, var sumY, var sumX2, var sumY2, var sumXY, var meanX, var meanY, var constt)
    {
        var xKey, xValue, bestFitY;
        float SSres = 0.0, SScov =  0.0, SStot =  0.0, SSsex = 0.0;
        
        for xKey, xValue in this->xValues {
            let this->yBestFitValues[xKey] = this->getValueOfYForX(xValue);
            let bestFitY = this->yBestFitValues[xKey];

            let SSres = SSres + (this->yValues[xKey] - bestFitY) * (this->yValues[xKey] - bestFitY);
            if (constt) {
                let SStot = SStot + (this->yValues[xKey] - meanY) * (this->yValues[xKey] - meanY);
            } else {
                let SStot = SStot + this->yValues[xKey] * this->yValues[xKey];
            }
            let SScov = SScov + (this->xValues[xKey] - meanX) * (this->yValues[xKey] - meanY);
            if (constt) {
                let SSsex = SSsex + (this->xValues[xKey] - meanX) * (this->xValues[xKey] - meanX);
            } else {
                let SSsex = SSsex + this->xValues[xKey] * this->xValues[xKey];
            }
        }

        let this->SSResiduals = SSres;
        let this->DFResiduals = this->valueCount - 1 - constt;

        if (this->DFResiduals == 0.0) {
            let this->stdevOfResiduals = 0.0;
        } else {
            let this->stdevOfResiduals = sqrt(SSres / this->DFResiduals);
        }
        if ((SStot == 0.0) || (SSres == SStot)) {
            let this->goodnessOfFit = 1;
        } else {
            let this->goodnessOfFit = 1 - (SSres / SStot);
        }

        let this->SSRegression = this->goodnessOfFit * (double) SStot;
        let this->covariance = SScov / this->valueCount;
        let this->correlation = (this->valueCount * sumXY - sumX * sumY) / sqrt((this->valueCount * sumX2 - pow(sumX, 2)) * (this->valueCount * sumY2 - pow(sumY, 2)));
        let this->slopeSE = this->stdevOfResiduals / sqrt(SSsex);
        let this->intersectSE = this->stdevOfResiduals * sqrt(1 / (this->valueCount - (sumX * sumX) / sumX2));
        if (this->SSResiduals != 0.0) {
            if (this->DFResiduals == 0.0) {
                let this->f = 0.0;
            } else {
                let this->f = this->SSRegression / (this->SSResiduals / this->DFResiduals);
            }
        } else {
            if (this->DFResiduals == 0.0) {
                let this->f = 0.0;
            } else {
                let this->f = this->SSRegression / this->DFResiduals;
            }
        }
    }

    protected function leastSquareFit(var yValues, var xValues, var constt)
    {
        var  i, x_sum, y_sum;
        float meanX, meanY, mBase = 0.0, mDivisor = 0.0, xx_sum = 0.0, xy_sum = 0.0, yy_sum = 0.0;
        
        // calculate sums
        let x_sum = array_sum(xValues);
        let y_sum = array_sum(yValues);
        let meanX = x_sum / this->valueCount;
        let meanY = y_sum / this->valueCount;
        
        for i in range(0, this->valueCount - 1) {
            let xy_sum = xy_sum + (xValues[i] * yValues[i]);
            let xx_sum = xx_sum + (xValues[i] * xValues[i]);
            let yy_sum = yy_sum + (yValues[i] * yValues[i]);

            if (constt) {
                let mBase = mBase + ((double) xValues[i] - meanX) * ((double) yValues[i] - meanY);
                let mDivisor = mDivisor + ((double) xValues[i] - meanX) * ((double) xValues[i] - meanX);
            } else {
                let mBase = mBase + ((double) xValues[i] * (double) yValues[i]);
                let mDivisor = mDivisor + ((double) xValues[i] * (double) xValues[i]);
            }
        }

        // calculate slope
//        this->slope = ((this->valueCount * xy_sum) - (x_sum * y_sum)) / ((this->valueCount * xx_sum) - (x_sum * x_sum));
        let this->slope = mBase / mDivisor;

        // calculate intersect
//        this->intersect = (y_sum - (this->slope * x_sum)) / this->valueCount;
        if (constt) {
            let this->intersect = meanY - ((double) this->slope * meanX);
        } else {
            let this->intersect = 0;
        }

        this->calculateGoodnessOfFit(x_sum, y_sum, xx_sum, yy_sum, xy_sum, meanX, meanY, constt);
    }
}
