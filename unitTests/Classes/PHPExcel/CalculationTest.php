<?php

require_once 'testDataFileIterator.php';

class CalculationTest extends PHPUnit_Framework_TestCase
{

    /**
     * @dataProvider providerBinaryComparisonOperation
     */
    public function testBinaryComparisonOperation($formula, $expectedResultExcel, $expectedResultOpenOffice)
    {
        \ZExcel\Calculation\Functions::setCompatibilityMode(\ZExcel\Calculation\Functions::COMPATIBILITY_EXCEL);
        $resultExcel = \ZExcel\Calculation::getInstance()->_calculateFormulaValue($formula);
        $this->assertEquals($expectedResultExcel, $resultExcel, 'should be Excel compatible');

        \ZExcel\Calculation\Functions::setCompatibilityMode(\ZExcel\Calculation\Functions::COMPATIBILITY_OPENOFFICE);
        $resultOpenOffice = \ZExcel\Calculation::getInstance()->_calculateFormulaValue($formula);
        $this->assertEquals($expectedResultOpenOffice, $resultOpenOffice, 'should be OpenOffice compatible');
    }

    public function providerBinaryComparisonOperation()
    {
        return new testDataFileIterator('rawTestData/CalculationBinaryComparisonOperation.data');
    }
}
