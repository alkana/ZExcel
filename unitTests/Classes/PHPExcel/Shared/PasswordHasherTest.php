<?php


require_once 'testDataFileIterator.php';

class PasswordHasherTest extends PHPUnit_Framework_TestCase
{

    public function setUp()
    {
        
    }

    /**
     * @dataProvider providerHashPassword
     */
    public function testHashPassword()
    {
        $args = func_get_args();
        $expectedResult = array_pop($args);
        $result = call_user_func_array(array('\ZExcel\Shared_PasswordHasher','hashPassword'), $args);
        $this->assertEquals($expectedResult, $result);
    }

    public function providerHashPassword()
    {
        return new testDataFileIterator('rawTestData/Shared/PasswordHashes.data');
    }
}
