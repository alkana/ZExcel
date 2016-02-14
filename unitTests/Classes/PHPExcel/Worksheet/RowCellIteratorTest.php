<?php

class RowCellIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRowCell;

    public function setUp()
    {
        
        
        $this->mockCell = $this->getMockBuilder('\ZExcel\Cell')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\ZExcel\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
        $this->mockWorksheet->expects($this->any())
                 ->method('getCellByColumnAndRow')
                 ->will($this->returnValue($this->mockCell));
    }


    public function testIteratorFullRange()
    {
        $iterator = new \ZExcel\Worksheet\RowCellIterator($this->mockWorksheet, 1, 'A');
        $RowCellIndexResult = 'A';
        $this->assertEquals($RowCellIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $RowCell) {
            $this->assertEquals($RowCellIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Cell', $RowCell);
        }
    }

    public function testIteratorStartEndRange()
    {
        $iterator = new \ZExcel\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $RowCellIndexResult = 'B';
        $this->assertEquals($RowCellIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $RowCell) {
            $this->assertEquals($RowCellIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Cell', $RowCell);
        }
    }

    public function testIteratorSeekAndPrev()
    {
        $ranges = range('A', 'E');
        $iterator = new \ZExcel\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $RowCellIndexResult = 'D';
        $iterator->seek('D');
        $this->assertEquals($RowCellIndexResult, $iterator->key());

        for ($i = 1; $i < array_search($RowCellIndexResult, $ranges); $i++) {
            $iterator->prev();
            $expectedResult = $ranges[array_search($RowCellIndexResult, $ranges) - $i];
            $this->assertEquals($expectedResult, $iterator->key());
        }
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $iterator->seek(1);
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet\RowCellIterator($this->mockWorksheet, 2, 'B', 'D');
        $iterator->prev();
    }
}
