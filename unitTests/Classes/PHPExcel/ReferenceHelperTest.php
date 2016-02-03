<?php


class ReferenceHelperTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testColumnSort()
    {
        $columnBase = $columnExpectedResult = array(
            'A','B','Z',
            'AA','AB','AZ',
            'BA','BB','BZ',
            'ZA','ZB','ZZ',
            'AAA','AAB','AAZ',
            'ABA','ABB','ABZ',
            'AZA','AZB','AZZ',
            'BAA','BAB','BAZ',
            'BBA','BBB','BBZ',
            'BZA','BZB','BZZ'
        );
        shuffle($columnBase);
        usort($columnBase, array('\ZExcel\ReferenceHelper','columnSort'));
        foreach ($columnBase as $key => $value) {
            $this->assertEquals($columnExpectedResult[$key], $value);
        }
    }

    public function testColumnReverseSort()
    {
        $columnBase = $columnExpectedResult = array(
            'A','B','Z',
            'AA','AB','AZ',
            'BA','BB','BZ',
            'ZA','ZB','ZZ',
            'AAA','AAB','AAZ',
            'ABA','ABB','ABZ',
            'AZA','AZB','AZZ',
            'BAA','BAB','BAZ',
            'BBA','BBB','BBZ',
            'BZA','BZB','BZZ'
        );
        shuffle($columnBase);
        $columnExpectedResult = array_reverse($columnExpectedResult);
        usort($columnBase, array('\ZExcel\ReferenceHelper','columnReverseSort'));
        foreach ($columnBase as $key => $value) {
            $this->assertEquals($columnExpectedResult[$key], $value);
        }
    }
}
