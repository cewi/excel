<?php
namespace App\Test\TestCase\Controller\Component;

use Cewi\Excel\Controller\Component\ImportComponent;
use Cake\Controller\ComponentRegistry;
use Cake\TestSuite\TestCase;

/**
 * App\Controller\Component\ImportComponent Test Case
 */
class ImportComponentTest extends TestCase
{
    
    /**
     * Path for files with test-data
     * 
     * @var string
     */
    protected $path = ROOT . DS. 'vendor'. DS . 'Cewi' . DS . 'Excel' . DS . 'tests' . DS . 'Files'. DS; 
    
    /**
     * setUp method
     *
     * @return void
     */
    public function setUp()
    {
        parent::setUp();
        $registry = new ComponentRegistry();
        $this->Import = new ImportComponent($registry);
    }

    /**
     * test prepareEntityData() function
     * @TODO make it working ;-)
     */
    public function testPrepareEntityData(){
        $result = $this->Import->prepareEntityData($this->path.'articles.ods');
        //dd($result);
        $this->assertTrue(true);
    }

        /**
     * tearDown method
     *
     * @return void
     */
    public function tearDown()
    {
        unset($this->Import);

        parent::tearDown();
    }

    /**
     * Test initial setup
     *
     * @return void
     */
    public function testInitialization()
    {
        $this->assertInstanceOf('\Cewi\Excel\Controller\Component\ImportComponent', $this->Import);
    }
}
