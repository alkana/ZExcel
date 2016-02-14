<?php

class AdvancedValueBinderTest extends PHPUnit_Framework_TestCase
{
    public function provider()
    {
        if (!class_exists('\ZExcel\Style_NumberFormat')) {
            $this->setUp();
        }
        $currencyUSD = \ZExcel\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE;
        $currencyEURO = str_replace('$', '€', \ZExcel\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

        return array(
            array('10%', 0.1, \ZExcel\Style\NumberFormat::FORMAT_PERCENTAGE_00, ',', '.', '$'),
            array('$10.11', 10.11, $currencyUSD, ',', '.', '$'),
            array('$1,010.12', 1010.12, $currencyUSD, ',', '.', '$'),
            array('$20,20', 20.2, $currencyUSD, '.', ',', '$'),
            array('$2.020,20', 2020.2, $currencyUSD, '.', ',', '$'),
            array('€2.020,20', 2020.2, $currencyEURO, '.', ',', '€'),
            array('€ 2.020,20', 2020.2, $currencyEURO, '.', ',', '€'),
            array('€2,020.22', 2020.22, $currencyEURO, ',', '.', '€'),
        );
    }

    /**
     * @dataProvider provider
     */
    public function testCurrency($value, $valueBinded, $format, $thousandsSeparator, $decimalSeparator, $currencyCode)
    {
        $sheet = $this->getMock(
            '\ZExcel\Worksheet',
            array('getStyle', 'getNumberFormat', 'setFormatCode','getCellCacheController')
        );
        $cache = $this->getMockBuilder('\ZExcel\CachedObjectStorage\Memory')
            ->disableOriginalConstructor()
            ->getMock();
        $cache->expects($this->any())
                 ->method('getParent')
                 ->will($this->returnValue($sheet));

        $sheet->expects($this->once())
                 ->method('getStyle')
                 ->will($this->returnSelf());
        $sheet->expects($this->once())
                 ->method('getNumberFormat')
                 ->will($this->returnSelf());
        $sheet->expects($this->once())
                 ->method('setFormatCode')
                 ->with($format)
                 ->will($this->returnSelf());
        $sheet->expects($this->any())
                 ->method('getCellCacheController')
                 ->will($this->returnValue($cache));

        \ZExcel\Shared\Stringg::setCurrencyCode($currencyCode);
        \ZExcel\Shared\Stringg::setDecimalSeparator($decimalSeparator);
        \ZExcel\Shared\Stringg::setThousandsSeparator($thousandsSeparator);

        $cell = new \ZExcel\Cell(null, \ZExcel\Cell\DataType::TYPE_STRING, $sheet);

        $binder = new \ZExcel\Cell\AdvancedValueBinder();
        $binder->bindValue($cell, $value);
        $this->assertEquals($valueBinded, $cell->getValue());
    }
}
