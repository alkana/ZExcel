<?php
class ReaderTest extends PHPUnit_Framework_TestCase
{
    public function testExtensionExists()
	{
		$this->assertTrue(extension_loaded('zexcel'));
	}
	
    public function testLoadXlsx()
    {
    	$objZExcel = ZExcel\IOFactory::load(__DIR__ . "/documents/testLoadXlsx.xlsx");
		
		$this->assertSame('ZExcel\ZExcel', get_class($objZExcel));
		$this->assertSame(2, $objZExcel->getSheetCount());
		
		$names = $objZExcel->getSheetNames();
		$this->assertCount(2, $names);
		$this->assertSame("Feuille1", $names[0]);
		$this->assertSame("Feuille2", $names[1]);
    }
}
