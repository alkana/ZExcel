<?php


class LayoutTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testSetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new \ZExcel\Chart_Layout;

        $result = $testInstance->setLayoutTarget($LayoutTargetValue);
        $this->assertTrue($result instanceof \ZExcel\Chart_Layout);
    }

    public function testGetLayoutTarget()
    {
        $LayoutTargetValue = 'String';

        $testInstance = new \ZExcel\Chart_Layout;
        $setValue = $testInstance->setLayoutTarget($LayoutTargetValue);

        $result = $testInstance->getLayoutTarget();
        $this->assertEquals($LayoutTargetValue, $result);
    }
}
