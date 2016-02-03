<?php


class DataTypeTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testGetErrorCodes()
    {
        $result = call_user_func(array('\ZExcel\Cell_DataType','getErrorCodes'));
        $this->assertInternalType('array', $result);
        $this->assertGreaterThan(0, count($result));
        $this->assertArrayHasKey('#NULL!', $result);
    }
}
