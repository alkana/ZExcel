<?php
class ReaderTest extends PHPUnit_Framework_TestCase
{
    public function testExtensionExists()
	{
		$this->assertTrue(extension_loaded('zexcel'));
	}

	public function testCreateWriter()
	{
		$reader = ZExcel\IOFactory::createReader('excel2007');
		
		echo "\nreader : " . get_class($reader) . "\n\n";
		
		$writer = ZExcel\IOFactory::createWriter(new ZExcel\ZExcel(), 'excel2007');
		
		echo "\nwriter : " . get_class($writer) . "\n\n";
	}
	
    public function testLoadXlsx()
    {
		try {
			$objPHPExcel = ZExcel\IOFactory::load(__DIR__ . "/documents/testLoadXlsx.xlsx");
			echo "\nloader : " . get_class($objPHPExcel) . "\n\n";
		} catch (Exception $ex) {
			echo "message : " . $ex->getMessage() . "\n\n";
			echo $ex->getTraceAsString();
		}
    }
}
