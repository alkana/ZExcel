<?php


class HyperlinkTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testGetUrl()
    {
        $urlValue = 'http://www.phpexcel.net';

        $testInstance = new \ZExcel\Cell\Hyperlink($urlValue);

        $result = $testInstance->getUrl();
        $this->assertEquals($urlValue, $result);
    }

    public function testSetUrl()
    {
        $initialUrlValue = 'http://www.phpexcel.net';
        $newUrlValue = 'http://github.com/PHPOffice/PHPExcel';

        $testInstance = new \ZExcel\Cell\Hyperlink($initialUrlValue);
        $result = $testInstance->setUrl($newUrlValue);
        $this->assertTrue($result instanceof \ZExcel\Cell\Hyperlink);

        $result = $testInstance->getUrl();
        $this->assertEquals($newUrlValue, $result);
    }

    public function testGetTooltip()
    {
        $tooltipValue = 'PHPExcel Web Site';

        $testInstance = new \ZExcel\Cell\Hyperlink(null, $tooltipValue);

        $result = $testInstance->getTooltip();
        $this->assertEquals($tooltipValue, $result);
    }

    public function testSetTooltip()
    {
        $initialTooltipValue = 'PHPExcel Web Site';
        $newTooltipValue = 'PHPExcel Repository on Github';

        $testInstance = new \ZExcel\Cell\Hyperlink(null, $initialTooltipValue);
        $result = $testInstance->setTooltip($newTooltipValue);
        $this->assertTrue($result instanceof \ZExcel\Cell\Hyperlink);

        $result = $testInstance->getTooltip();
        $this->assertEquals($newTooltipValue, $result);
    }

    public function testIsInternal()
    {
        $initialUrlValue = 'http://www.phpexcel.net';
        $newUrlValue = 'sheet://Worksheet1!A1';

        $testInstance = new \ZExcel\Cell\Hyperlink($initialUrlValue);
        $result = $testInstance->isInternal();
        $this->assertFalse($result);

        $testInstance->setUrl($newUrlValue);
        $result = $testInstance->isInternal();
        $this->assertTrue($result);
    }

    public function testGetHashCode()
    {
        $urlValue = 'http://www.phpexcel.net';
        $tooltipValue = 'PHPExcel Web Site';
        $initialExpectedHash = 'd84d713aed1dbbc8a7c5af183d6c7dbb';

        $testInstance = new \ZExcel\Cell\Hyperlink($urlValue, $tooltipValue);

        $result = $testInstance->getHashCode();
        $this->assertEquals($initialExpectedHash, $result);
    }
}
