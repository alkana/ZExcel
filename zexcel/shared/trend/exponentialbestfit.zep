namespace ZExcel\Shared\Trend;

/**
 * PHPExcel_Exponential_Best_Fit
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
class ExponentialBestFit extends \ZExcel\Shared\Trend\BestFit
{
    /**
     * Algorithm type to use for best-fit
     * (Name of this trend class)
     *
     * @var    string
     **/
    protected bestFitType = "exponential";

    /**
     * Define the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param    float[]        yValues    The set of Y-values for this regression
     * @param    float[]        xValues    The set of X-values for this regression
     * @param    boolean        const
     */
    public function __construct(yValues, xValues = [], constt = true)
    {
        if (parent::__construct(yValues, xValues) !== false) {
            this->exponentialRegression(yValues, xValues, constt);
        }
    }
    
    /**
     * Return the Y-Value for a specified value of X
     *
     * @param     float        xValue            X-Value
     * @return     float                        Y-Value
     **/
    public function getValueOfYForX(xValue)
    {
        return this->getIntersect() * pow(this->getSlope(), (xValue - this->xOffset));
    }

    /**
     * Return the X-Value for a specified value of Y
     *
     * @param     float        yValue            Y-Value
     * @return     float                        X-Value
     **/
    public function getValueOfXForY(yValue)
    {
        return log((yValue + this->yOffset) / this->getIntersect()) / log(this->getSlope());
    }

    /**
     * Return the Equation of the best-fit line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     **/
    public function getEquation(int dp = 0)
    {
        var slope, intersect;
        
        let slope = this->getSlope(dp);
        let intersect = this->getIntersect(dp);

        return "Y = " . intersect . " * " . slope . "^X";
    }

    /**
     * Return the Slope of the line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     **/
    public function getSlope(dp = 0)
    {
        if (dp != 0) {
            return round(exp(this->slope), dp);
        }
        return exp(this->slope);
    }

    /**
     * Return the Value of X where it intersects Y = 0
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     **/
    public function getIntersect(dp = 0)
    {
        if (dp != 0) {
            return round(exp(this->intersect), dp);
        }
        return exp(this->intersect);
    }

    /**
     * Execute the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param     float[]    yValues    The set of Y-values for this regression
     * @param     float[]    xValues    The set of X-values for this regression
     * @param     boolean    const
     */
    private function exponentialRegression(var yValues, var xValues, constt)
    {
        var k, value;
        
        for k, value in yValues {
            if (value < 0.0) {
                let yValues[k] = 0 - log(abs(value));
            } else {
                if (value > 0.0) {
                    let yValues[k] = log(value);
                }
            }
        }

        this->leastSquareFit(yValues, xValues, constt);
    }
}
