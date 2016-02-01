<?php
class ReaderTest extends PHPUnit_Framework_TestCase
{
    public function testExtensionExists()
    {
        $this->assertTrue(extension_loaded('zexcel'));
    }
    
    public function testZExcelLoad()
    {
        // TEST Load files with simple/multi sheet(s)
        try {
            $objZExcel = ZExcel\IOFactory::load(__DIR__ . "/documents/simple.xlsx");
            
            $this->assertEquals('ZExcel\ZExcel', get_class($objZExcel));
            $this->assertEquals("1", $objZExcel->getSheetCount());
            
            $names = $objZExcel->getSheetNames();
            $this->assertCount(1, $names);
            $this->assertEquals("Feuille1", $names[0]);
            
            // TEST Multi Sheets
            $objZExcel = ZExcel\IOFactory::load(__DIR__ . "/documents/multi.xlsx");
            
            $this->assertEquals('ZExcel\ZExcel', get_class($objZExcel));
            $this->assertEquals(2, $objZExcel->getSheetCount());
            
            $names = $objZExcel->getSheetNames();
            $this->assertCount(2, $names);
            $this->assertEquals("Feuille1", $names[0]);
            $this->assertEquals("Feuille2", $names[1]);

            // Test active sheet
            $sheet = $objZExcel->getActiveSheet();
            $this->assertEquals("Feuille1", $sheet->getTitle());
            $this->assertEquals("B1", $sheet->getActiveCell());
        } catch (\Exception $ex) {
            echo $ex->getMessage() . "\n";
            echo $ex->getTraceAsString();
        }
    }
}
