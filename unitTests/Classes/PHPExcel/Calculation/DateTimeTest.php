<?php


require_once 'testDataFileIterator.php';

class DateTimeTest extends PHPUnit_Framework_TestCase
{


    /**
     * @dataProvider providerDATE
     */
    public function testDATE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DATE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDATE()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/DATE.data');
    }

    public function testDATEtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::DATE(2012, 1, 31);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(1327968000, $result, null, 1E-8);
    }

    public function testDATEtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::DATE(2012, 1, 31);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('d-M-Y'), '31-Jan-2012');
    }

    public function testDATEwith1904Calendar()
    {
        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_MAC_1904);
        $result = \ZExcel\Calculation\DateTime::DATE(1918, 11, 11);
        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_WINDOWS_1900);
        $this->assertEquals($result, 5428);
    }

    public function testDATEwith1904CalendarError()
    {
        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_MAC_1904);
        $result = \ZExcel\Calculation\DateTime::DATE(1901, 1, 31);
        \ZExcel\Shared\Date::setExcelCalendar(\ZExcel\Shared\Date::CALENDAR_WINDOWS_1900);
        $this->assertEquals($result, '#NUM!');
    }

    /**
     * @dataProvider providerDATEVALUE
     */
    public function testDATEVALUE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DATEVALUE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDATEVALUE()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/DATEVALUE.data');
    }

    public function testDATEVALUEtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::DATEVALUE('2012-1-31');
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(1327968000, $result, null, 1E-8);
    }

    public function testDATEVALUEtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::DATEVALUE('2012-1-31');
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('d-M-Y'), '31-Jan-2012');
    }

    /**
     * @dataProvider providerYEAR
     */
    public function testYEAR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'YEAR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerYEAR()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/YEAR.data');
    }

    /**
     * @dataProvider providerMONTH
     */
    public function testMONTH()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'MONTHOFYEAR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerMONTH()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/MONTH.data');
    }

    /**
     * @dataProvider providerWEEKNUM
     */
    public function testWEEKNUM()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'WEEKOFYEAR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerWEEKNUM()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/WEEKNUM.data');
    }

    /**
     * @dataProvider providerWEEKDAY
     */
    public function testWEEKDAY()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DAYOFWEEK'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerWEEKDAY()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/WEEKDAY.data');
    }

    /**
     * @dataProvider providerDAY
     */
    public function testDAY()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DAYOFMONTH'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDAY()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/DAY.data');
    }

    /**
     * @dataProvider providerTIME
     */
    public function testTIME()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'TIME'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerTIME()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/TIME.data');
    }

    public function testTIMEtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::TIME(7, 30, 20);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(27020, $result, null, 1E-8);
    }

    public function testTIMEtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::TIME(7, 30, 20);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('H:i:s'), '07:30:20');
    }

    /**
     * @dataProvider providerTIMEVALUE
     */
    public function testTIMEVALUE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'TIMEVALUE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerTIMEVALUE()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/TIMEVALUE.data');
    }

    public function testTIMEVALUEtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::TIMEVALUE('7:30:20');
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(23420, $result, null, 1E-8);
    }

    public function testTIMEVALUEtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::TIMEVALUE('7:30:20');
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('H:i:s'), '07:30:20');
    }

    /**
     * @dataProvider providerHOUR
     */
    public function testHOUR()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'HOUROFDAY'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerHOUR()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/HOUR.data');
    }

    /**
     * @dataProvider providerMINUTE
     */
    public function testMINUTE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'MINUTEOFHOUR'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerMINUTE()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/MINUTE.data');
    }

    /**
     * @dataProvider providerSECOND
     */
    public function testSECOND()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'SECONDOFMINUTE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerSECOND()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/SECOND.data');
    }

    /**
     * @dataProvider providerNETWORKDAYS
     */
    public function testNETWORKDAYS()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'NETWORKDAYS'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerNETWORKDAYS()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/NETWORKDAYS.data');
    }

    /**
     * @dataProvider providerWORKDAY
     */
    public function testWORKDAY()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'WORKDAY'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerWORKDAY()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/WORKDAY.data');
    }

    /**
     * @dataProvider providerEDATE
     */
    public function testEDATE()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'EDATE'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerEDATE()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/EDATE.data');
    }

    public function testEDATEtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::EDATE('2012-1-26', -1);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(1324857600, $result, null, 1E-8);
    }

    public function testEDATEtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::EDATE('2012-1-26', -1);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('d-M-Y'), '26-Dec-2011');
    }

    /**
     * @dataProvider providerEOMONTH
     */
    public function testEOMONTH()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'EOMONTH'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerEOMONTH()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/EOMONTH.data');
    }

    public function testEOMONTHtoPHP()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_NUMERIC);
        $result = \ZExcel\Calculation\DateTime::EOMONTH('2012-1-26', -1);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        $this->assertEquals(1325289600, $result, null, 1E-8);
    }

    public function testEOMONTHtoPHPObject()
    {
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_PHP_OBJECT);
        $result = \ZExcel\Calculation\DateTime::EOMONTH('2012-1-26', -1);
        \ZExcel\Calculation\Functions::setReturnDateType(\ZExcel\Calculation\Functions::RETURNDATE_EXCEL);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, 'DateTime'));
        //    ... with the correct value
        $this->assertEquals($result->format('d-M-Y'), '31-Dec-2011');
    }

    /**
     * @dataProvider providerDATEDIF
     */
    public function testDATEDIF()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DATEDIF'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDATEDIF()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/DATEDIF.data');
    }

    /**
     * @dataProvider providerDAYS360
     */
    public function testDAYS360()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'DAYS360'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerDAYS360()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/DAYS360.data');
    }

    /**
     * @dataProvider providerYEARFRAC
     */
    public function testYEARFRAC()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Calculation\DateTime', 'YEARFRAC'), $args);
        $this->assertEquals($expectedResult, $result, null, 1E-8);
    }

    public function providerYEARFRAC()
    {
        return new testDataFileIterator('rawTestData/Calculation/DateTime/YEARFRAC.data');
    }
}
