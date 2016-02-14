<?php

class RowIteratorTest extends PHPUnit_Framework_TestCase
{
    public $mockWorksheet;
    public $mockRow;

    public function setUp()
    {
        $this->mockRow = $this->getMockBuilder('\ZExcel\Worksheet\Row')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet = $this->getMockBuilder('\ZExcel\Worksheet')
            ->disableOriginalConstructor()
            ->getMock();

        $this->mockWorksheet->expects($this->any())
                 ->method('getHighestRow')
                 ->will($this->returnValue(5));
        $this->mockWorksheet->expects($this->any())
                 ->method('current')
                 ->will($this->returnValue($this->mockRow));
    }


    public function testIteratorFullRange()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet);
        $rowIndexResult = 1;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Worksheet\Row', $row);
        }
    }

    public function testIteratorStartEndRange()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 2;
        $this->assertEquals($rowIndexResult, $iterator->key());
        
        foreach ($iterator as $key => $row) {
            $this->assertEquals($rowIndexResult++, $key);
            $this->assertInstanceOf('\ZExcel\Worksheet\Row', $row);
        }
    }

    public function testIteratorSeekAndPrev()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $rowIndexResult = 4;
        $iterator->seek($rowIndexResult);
        $this->assertEquals($rowIndexResult, $iterator->key());

        for ($i = 1; $i < $rowIndexResult-1; $i++) {
            $iterator->prev();
            $this->assertEquals($rowIndexResult - $i, $iterator->key());
        }
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testStartOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet, 256, 512);
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testSeekOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $iterator->seek(1);
    }

    /**
     * @expectedException \ZExcel\Exception
     */
    public function testPrevOutOfRange()
    {
        $iterator = new \ZExcel\Worksheet\RowIterator($this->mockWorksheet, 2, 4);
        $iterator->prev();
    }
}
