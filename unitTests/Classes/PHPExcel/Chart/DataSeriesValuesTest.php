<?php


class DataSeriesValuesTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    public function testSetDataType()
    {
        $dataTypeValues = array(
            'Number',
            'String'
        );

        $testInstance = new \ZExcel\Chart_DataSeriesValues;

        foreach ($dataTypeValues as $dataTypeValue) {
            $result = $testInstance->setDataType($dataTypeValue);
            $this->assertTrue($result instanceof \ZExcel\Chart_DataSeriesValues);
        }
    }

    public function testSetInvalidDataTypeThrowsException()
    {
        $testInstance = new \ZExcel\Chart_DataSeriesValues;

        try {
            $result = $testInstance->setDataType('BOOLEAN');
        } catch (Exception $e) {
            $this->assertEquals($e->getMessage(), 'Invalid datatype for chart data series values');
            return;
        }
        $this->fail('An expected exception has not been raised.');
    }

    public function testGetDataType()
    {
        $dataTypeValue = 'String';

        $testInstance = new \ZExcel\Chart_DataSeriesValues;
        $setValue = $testInstance->setDataType($dataTypeValue);

        $result = $testInstance->getDataType();
        $this->assertEquals($dataTypeValue, $result);
    }
}
