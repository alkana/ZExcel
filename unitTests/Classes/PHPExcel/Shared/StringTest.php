<?php


require_once 'testDataFileIterator.php';

class StringTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testGetIsMbStringEnabled()
    {
        $result = call_user_func(array('\ZExcel\Shared\Stringg','getIsMbstringEnabled'));
        $this->assertTrue($result);
    }

    public function testGetIsIconvEnabled()
    {
        $result = call_user_func(array('\ZExcel\Shared\Stringg','getIsIconvEnabled'));
        $this->assertTrue($result);
    }

    public function testGetDecimalSeparator()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['decimal_point'])) ? $localeconv['decimal_point'] : ',';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','getDecimalSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetDecimalSeparator()
    {
        $expectedResult = ',';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','setDecimalSeparator'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared\Stringg','getDecimalSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testGetThousandsSeparator()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['thousands_sep'])) ? $localeconv['thousands_sep'] : ',';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','getThousandsSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetThousandsSeparator()
    {
        $expectedResult = ' ';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','setThousandsSeparator'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared\Stringg','getThousandsSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testGetCurrencyCode()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['currency_symbol'])) ? $localeconv['currency_symbol'] : '$';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','getCurrencyCode'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetCurrencyCode()
    {
        $expectedResult = '£';
        $result = call_user_func(array('\ZExcel\Shared\Stringg','setCurrencyCode'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared\Stringg','getCurrencyCode'));
        $this->assertEquals($expectedResult, $result);
    }
}
