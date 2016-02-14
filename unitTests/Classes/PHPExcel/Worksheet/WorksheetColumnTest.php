<?php

class WorksheetColumnTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumn;

    public function setUp()
    {
        
        
        $this->mockWorksheet = $this->getMockBuilder('\ZExcel\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();
        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
    }


    public function testInstantiateColumnDefault()
    {
        $column = new \ZExcel\Worksheet\Column($this->mockWorksheet);
        $this->assertInstanceOf('\ZExcel\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('A', $columnIndex);
    }

    public function testInstantiateColumnSpecified()
    {
        $column = new \ZExcel\Worksheet\Column($this->mockWorksheet, 'E');
        $this->assertInstanceOf('\ZExcel\Worksheet\Column', $column);
        $columnIndex = $column->getColumnIndex();
        $this->assertEquals('E', $columnIndex);
    }

    public function testGetCellIterator()
    {
        $column = new \ZExcel\Worksheet\Column($this->mockWorksheet);
        $cellIterator = $column->getCellIterator();
        $this->assertInstanceOf('\ZExcel\Worksheet\ColumnCellIterator', $cellIterator);
    }
}
