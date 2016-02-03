<?php

class ColumnIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockColumn;

    public function setUp()
    {
        
        
        $this->mockColumn = $this->getMockBuilder('\ZExcel\Worksheet_Column')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\ZExcel\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestColumn')
                 ->will($this->returnValue('E'));
        $this->mockWorksheet->expects($this->any())
                 ->method('current')
                 ->will($this->returnValue($this->mockColumn));
    }


    public function testIteratorFullRange()
    {
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet);
        $columnIndexResult = 'A';
        $this->assertEquals($columnIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $column) {
            $this->assertEquals($columnIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Worksheet_Column', $column);
        }
    }

    public function testIteratorStartEndRange()
    {
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet, 'B', 'D');
        $columnIndexResult = 'B';
        $this->assertEquals($columnIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $column) {
            $this->assertEquals($columnIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Worksheet_Column', $column);
        }
    }

    public function testIteratorSeekAndPrev()
    {
        $ranges = range('A', 'E');
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet, 'B', 'D');
        $columnIndexResult = 'D';
        $iterator->seek('D');
        $this->assertEquals($columnIndexResult, $iterator->key());

        for ($i = 1; $i < array_search($columnIndexResult, $ranges); $i++) {
            $iterator->prev();
            $expectedResult = $ranges[array_search($columnIndexResult, $ranges) - $i];
            $this->assertEquals($expectedResult, $iterator->key());
        }
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testStartOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet, 'IA', 'IV');
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet, 'B', 'D');
        $iterator->seek('A');
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet_ColumnIterator($this->mockWorksheet, 'B', 'D');
        $iterator->prev();
    }
}
