<?php

class CellCollectionTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }


    public function testCacheLastCell()
    {
        $methods = \ZExcel\CachedObjectStorageFactory::getCacheStorageMethods();
        foreach ($methods as $method) {
            \ZExcel\CachedObjectStorageFactory::initialize($method);
            $workbook = new \ZExcel\ZExcel();
            $cells = array('A1', 'A2');
            $worksheet = $workbook->getActiveSheet();
            $worksheet->setCellValue('A1', 1);
            $worksheet->setCellValue('A2', 2);
            $this->assertEquals($cells, $worksheet->getCellCollection(), "Cache method \"$method\".");
            \ZExcel\CachedObjectStorageFactory::finalize();
        }
    }
}
