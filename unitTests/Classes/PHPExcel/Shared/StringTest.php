<?php


require_once 'testDataFileIterator.php';

class StringTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testGetIsMbStringEnabled()
    {
        $result = call_user_func(array('\ZExcel\Shared_String','getIsMbstringEnabled'));
        $this->assertTrue($result);
    }

    public function testGetIsIconvEnabled()
    {
        $result = call_user_func(array('\ZExcel\Shared_String','getIsIconvEnabled'));
        $this->assertTrue($result);
    }

    public function testGetDecimalSeparator()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['decimal_point'])) ? $localeconv['decimal_point'] : ',';
        $result = call_user_func(array('\ZExcel\Shared_String','getDecimalSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetDecimalSeparator()
    {
        $expectedResult = ',';
        $result = call_user_func(array('\ZExcel\Shared_String','setDecimalSeparator'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared_String','getDecimalSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testGetThousandsSeparator()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['thousands_sep'])) ? $localeconv['thousands_sep'] : ',';
        $result = call_user_func(array('\ZExcel\Shared_String','getThousandsSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetThousandsSeparator()
    {
        $expectedResult = ' ';
        $result = call_user_func(array('\ZExcel\Shared_String','setThousandsSeparator'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared_String','getThousandsSeparator'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testGetCurrencyCode()
    {
        $localeconv = localeconv();

        $expectedResult = (!empty($localeconv['currency_symbol'])) ? $localeconv['currency_symbol'] : '$';
        $result = call_user_func(array('\ZExcel\Shared_String','getCurrencyCode'));
        $this->assertEquals($expectedResult, $result);
    }

    public function testSetCurrencyCode()
    {
        $expectedResult = '£';
        $result = call_user_func(array('\ZExcel\Shared_String','setCurrencyCode'), $expectedResult);

        $result = call_user_func(array('\ZExcel\Shared_String','getCurrencyCode'));
        $this->assertEquals($expectedResult, $result);
    }
}
