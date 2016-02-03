<?php

class WorksheetRowTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRow;

    public function setUp()
    {
        
        
        $this->mockWorksheet = $this->getMockBuilder('\ZExcel\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();
        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
    }


    public function testInstantiateRowDefault()
    {
        $row = new \ZExcel\Worksheet_Row($this->mockWorksheet);
        $this->assertInstanceOf('\ZExcel\Worksheet_Row', $row);
        $rowIndex = $row->getRowIndex();
        $this->assertEquals(1, $rowIndex);
    }

    public function testInstantiateRowSpecified()
    {
        $row = new \ZExcel\Worksheet_Row($this->mockWorksheet, 5);
        $this->assertInstanceOf('\ZExcel\Worksheet_Row', $row);
        $rowIndex = $row->getRowIndex();
        $this->assertEquals(5, $rowIndex);
    }

    public function testGetCellIterator()
    {
        $row = new \ZExcel\Worksheet_Row($this->mockWorksheet);
        $cellIterator = $row->getCellIterator();
        $this->assertInstanceOf('\ZExcel\Worksheet_RowCellIterator', $cellIterator);
    }
}
