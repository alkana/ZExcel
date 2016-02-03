<?php


class LegendTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testSetPosition()
    {
        $positionValues = array(
            \ZExcel\Chart_Legend::POSITION_RIGHT,
            \ZExcel\Chart_Legend::POSITION_LEFT,
            \ZExcel\Chart_Legend::POSITION_TOP,
            \ZExcel\Chart_Legend::POSITION_BOTTOM,
            \ZExcel\Chart_Legend::POSITION_TOPRIGHT,
        );

        $testInstance = new \ZExcel\Chart_Legend;

        foreach ($positionValues as $positionValue) {
            $result = $testInstance->setPosition($positionValue);
            $this->assertTrue($result);
        }
    }

    public function testSetInvalidPositionReturnsFalse()
    {
        $testInstance = new \ZExcel\Chart_Legend;

        $result = $testInstance->setPosition('BottomLeft');
        $this->assertFalse($result);
        //    Ensure that value is unchanged
        $result = $testInstance->getPosition();
        $this->assertEquals(\ZExcel\Chart_Legend::POSITION_RIGHT, $result);
    }

    public function testGetPosition()
    {
        $PositionValue = \ZExcel\Chart_Legend::POSITION_BOTTOM;

        $testInstance = new \ZExcel\Chart_Legend;
        $setValue = $testInstance->setPosition($PositionValue);

        $result = $testInstance->getPosition();
        $this->assertEquals($PositionValue, $result);
    }

    public function testSetPositionXL()
    {
        $positionValues = array(
            \ZExcel\Chart_Legend::xlLegendPositionBottom,
            \ZExcel\Chart_Legend::xlLegendPositionCorner,
            \ZExcel\Chart_Legend::xlLegendPositionCustom,
            \ZExcel\Chart_Legend::xlLegendPositionLeft,
            \ZExcel\Chart_Legend::xlLegendPositionRight,
            \ZExcel\Chart_Legend::xlLegendPositionTop,
        );

        $testInstance = new \ZExcel\Chart_Legend;

        foreach ($positionValues as $positionValue) {
            $result = $testInstance->setPositionXL($positionValue);
            $this->assertTrue($result);
        }
    }

    public function testSetInvalidXLPositionReturnsFalse()
    {
        $testInstance = new \ZExcel\Chart_Legend;

        $result = $testInstance->setPositionXL(999);
        $this->assertFalse($result);
        //    Ensure that value is unchanged
        $result = $testInstance->getPositionXL();
        $this->assertEquals(\ZExcel\Chart_Legend::xlLegendPositionRight, $result);
    }

    public function testGetPositionXL()
    {
        $PositionValue = \ZExcel\Chart_Legend::xlLegendPositionCorner;

        $testInstance = new \ZExcel\Chart_Legend;
        $setValue = $testInstance->setPositionXL($PositionValue);

        $result = $testInstance->getPositionXL();
        $this->assertEquals($PositionValue, $result);
    }

    public function testSetOverlay()
    {
        $overlayValues = array(
            true,
            false,
        );

        $testInstance = new \ZExcel\Chart_Legend;

        foreach ($overlayValues as $overlayValue) {
            $result = $testInstance->setOverlay($overlayValue);
            $this->assertTrue($result);
        }
    }

    public function testSetInvalidOverlayReturnsFalse()
    {
        $testInstance = new \ZExcel\Chart_Legend;

        $result = $testInstance->setOverlay('INVALID');
        $this->assertFalse($result);

        $result = $testInstance->getOverlay();
        $this->assertFalse($result);
    }

    public function testGetOverlay()
    {
        $OverlayValue = true;

        $testInstance = new \ZExcel\Chart_Legend;
        $setValue = $testInstance->setOverlay($OverlayValue);

        $result = $testInstance->getOverlay();
        $this->assertEquals($OverlayValue, $result);
    }
}
