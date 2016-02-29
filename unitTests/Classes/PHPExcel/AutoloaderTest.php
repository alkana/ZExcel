<?php


class AutoloaderTest extends PHPUnit_Framework_TestCase
{

    public function testAutoloaderNonPHPExcelClass()
    {
        $className = 'InvalidClass';

        $result = \ZExcel\Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloaderInvalidPHPExcelClass()
    {
        $className = '\ZExcel\Invalid\Class';

        $result = \ZExcel\Autoloader::Load($className);
        //    Must return a boolean...
        $this->assertTrue(is_bool($result));
        //    ... indicating failure
        $this->assertFalse($result);
    }

    public function testAutoloadValidPHPExcelClass()
    {
        $className = '\ZExcel\IOFactory';

        $result = \ZExcel\Autoloader::Load($className);
        //    Check that class has been loaded
        $this->assertTrue(class_exists($className));
    }

    public function testAutoloadInstantiateSuccess()
    {
        $result = new \ZExcel\Calculation\Functionn(1, 2, 3);
        //    Must return an object...
        $this->assertTrue(is_object($result));
        //    ... of the correct type
        $this->assertTrue(is_a($result, '\ZExcel\Calculation\Functionn'));
    }
}
