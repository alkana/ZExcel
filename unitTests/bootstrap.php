<?php
/**
 * $Id: bootstrap.php 2892 2011-08-14 15:11:50Z markbaker@phpexcel.net $
 *
 * @copyright   Copyright (C) 2011-2014 PHPExcel. All rights reserved.
 * @package     PHPExcel
 * @subpackage  PHPExcel Unit Tests
 * @author      Mark Baker
 */

chdir(dirname(__FILE__));

setlocale(LC_ALL, 'en_US.utf8');

// PHP 5.3 Compat
date_default_timezone_set('Europe/London');

/**
 * @todo Sort out xdebug in vagrant so that this works in all sandboxes
 * For now, it is safer to test for it rather then remove it.
 */
echo "ZExcel tests beginning\n";

if (extension_loaded('xdebug')) {
    echo "Xdebug extension loaded and running\n";
    xdebug_enable();
} else {
    echo 'Xdebug not found, you should run the following at the command line: echo "zend_extension=/usr/lib64/php/modules/xdebug.so" > /etc/php.d/xdebug.ini' . "\n";
}
