<?php


require_once 'testDataFileIterator.php';

class NumberFormatTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        

        \ZExcel\Shared\Stringg::setDecimalSeparator('.');
        \ZExcel\Shared\Stringg::setThousandsSeparator(',');
    }

    /**
     * @dataProvider providerNumberFormat
     */
    public function testFormatValueWithMask()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Style_NumberFormat','toFormattedString'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerNumberFormat()
    {
        return new testDataFileIterator('rawTestData/Style/NumberFormat.data');
    }
}
