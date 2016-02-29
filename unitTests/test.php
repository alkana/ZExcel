<?php

require __DIR__ . '/../../PHPExcel.php';

ob_start();

echo \ZExcel\Calculation\MathTrig::COMBIN(100, 3) . "\n";
echo PHPExcel_Calculation_MathTrig::COMBIN(100, 3) . "\n";
echo "\n";