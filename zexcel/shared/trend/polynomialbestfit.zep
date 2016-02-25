namespace ZExcel\Shared\Trend;

/**
 * PHPExcel_Polynomial_Best_Fit
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
class PolynomialBestFit extends \ZExcel\Shared\Trend\BestFit
{
    /**
     * Algorithm type to use for best-fit
     * (Name of this trend class)
     *
     * @var    string
     **/
    protected bestFitType = "polynomial";

    /**
     * Polynomial order
     *
     * @protected
     * @var    int
     **/
    protected order = 0;

    /**
     * Define the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param    int            order        Order of Polynomial for this regression
     * @param    float[]        yValues    The set of Y-values for this regression
     * @param    float[]        xValues    The set of X-values for this regression
     * @param    boolean        const
     */
    public function __construct(var order, var yValues, var xValues = [], var constt = true)
    {
        if (parent::__construct(yValues, xValues) !== false) {
            if (order < this->valueCount) {
                let this->bestFitType .= this->bestFitType . "_" . order;
                let this->order = order;
                
                this->polynomialRegression(order, yValues, xValues, constt);
                
                if ((this->getGoodnessOfFit() < 0.0) || (this->getGoodnessOfFit() > 1.0)) {
                    let this->_error = true;
                }
            } else {
                let this->_error = true;
            }
        }
    }


    /**
     * Return the order of this polynomial
     *
     * @return     int
     **/
    public function getOrder()
    {
        return this->order;
    }


    /**
     * Return the Y-Value for a specified value of X
     *
     * @param     float        xValue            X-Value
     * @return     float                        Y-Value
     **/
    public function getValueOfYForX(xValue)
    {
        var slope, retVal, key, value;
        
        let retVal = this->getIntersect();
        let slope = this->getSlope();
        
        for key, value in slope {
            if (value != 0.0) {
                let retVal = retVal + (value * pow(xValue, key + 1));
            }
        }
        
        return retVal;
    }


    /**
     * Return the X-Value for a specified value of Y
     *
     * @param     float        yValue            Y-Value
     * @return     float                        X-Value
     **/
    public function getValueOfXForY(float yValue) -> float
    {
        return (yValue - this->getIntersect()) / this->getSlope();
    }


    /**
     * Return the Equation of the best-fit line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     **/
    public function getEquation(int dp = 0) -> string
    {
        var slope, intersect, equation, key, value;
        
        let slope = this->getSlope(dp);
        let intersect = this->getIntersect(dp);

        let equation = "Y = " . intersect;
        
        for key, value in slope {
            if (value != 0.0) {
                let equation = equation . " + " . value . " * X";
                if (key > 0) {
                    let equation = "^" . (key + 1);
                }
            }
        }
        
        return equation;
    }


    /**
     * Return the Slope of the line
     *
     * @param     int        dp        Number of places of decimal precision to display
     * @return     string
     **/
    public function getSlope(int dp = 0)
    {
        var coefficient;
        array coefficients;
        
        if (dp != 0) {
            let coefficients = [];
            for coefficient in this->slope {
                let coefficients[] = round(coefficient, dp);
            }
            return coefficients;
        }
        
        return this->slope;
    }


    public function getCoefficients(dp = 0)
    {
        return array_merge([this->getIntersect(dp)], this->getSlope(dp));
    }


    /**
     * Execute the regression and calculate the goodness of fit for a set of X and Y data values
     *
     * @param    int            order        Order of Polynomial for this regression
     * @param    float[]        yValues    The set of Y-values for this regression
     * @param    float[]        xValues    The set of X-values for this regression
     * @param    boolean        const
     */
    private function polynomialRegression(order, yValues, xValues, constt)
    {
        var x_sum, y_sum, yy_sum, xy_sum, xx_sum, a, b, c, i, j, r, matrixA, matrixB, xKey, xValue;
        array coefficients;
        
        // calculate sums
        let x_sum = array_sum(xValues);
        let y_sum = array_sum(yValues);
        let xx_sum = 0;
        let xy_sum = 0;
        let yy_sum = 0;
        
        for i in range(0, this->valueCount - 1) {
            let xy_sum = xy_sum + (xValues[i] * yValues[i]);
            let xx_sum = xx_sum + (xValues[i] * xValues[i]);
            let yy_sum = yy_sum + (yValues[i] * yValues[i]);
        }
        /*
         *    This routine uses logic from the PHP port of polyfit version 0.1
         *    written by Michael Bommarito and Paul Meagher
         *
         *    The function fits a polynomial function of order order through
         *    a series of x-y data points using least squares.
         *
         */
        let a = [];
         
        for i in range(i, this->valueCount - 1) {
            let a[i] = [];
            for j in range(0, order) {
                let a[i][j] = pow(xValues[i], j);
            }
        }
        
        let b = [];
        
        for i in range(i, this->valueCount - 1) {
            let b[i] = [yValues[i]];
        }
        
        let matrixA = new \ZExcel\Shared\JAMA\Matrix(a);
        let matrixB = new \ZExcel\Shared\JAMA\Matrix(b);
        let c = matrixA->solve(matrixB);

        let coefficients = [];
        
        for i in range(0, c->m - 1) {
            let r = c->get(i, 0);
            
            if (abs(r) <= pow(10, -9)) {
                let r = 0;
            }
            
            let coefficients[] = r;
        }

        let this->intersect = array_shift(coefficients);
        let this->slope = coefficients;

        call_user_func("this->calculateGoodnessOfFit", x_sum, y_sum, xx_sum, yy_sum, xy_sum);
        
        for xKey, xValue in this->xValues {
            let this->yBestFitValues[xKey] = this->getValueOfYForX(xValue);
        }
    }
}
