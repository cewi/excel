<?php
namespace App\Test\TestCase\View\Helper;

use Cewi\Excel\View\Helper\ExcelHelper;
use Cewi\Excel\View;
use Cake\TestSuite\TestCase;
use Cake\ORM\TableRegistry;

/**
 * App\View\Helper\ExcelHelper Test Case
 */
class ExcelHelperTest extends TestCase
{

    public $fixtures = ['core.articles'];

    /**
     * setUp method
     *
     * @return void
     */
    public function setUp()
    {
        parent::setUp();
        $excelView = new View\ExcelView();
        $this->Excel = new ExcelHelper($excelView);
        $this->Articles = TableRegistry::get('Articles');
    }

    /**
     * tearDown method
     *
     * @return void
     */
    public function tearDown()
    {
        unset($this->Excel);
        unset($this->Articles);
        unset($this->excelView);
        parent::tearDown();
    }

    /**
     * Test initial setup
     *
     * @return void
     */
    public function testInitialization()
    {
        // is the correct Object loaded?
        $this->assertInstanceOf('Cewi\Excel\View\Helper\ExcelHelper', $this->Excel);
    }

    public function testAddTable()
    {
        $query = $this->Articles->find();
        $this->Excel->addTable($query, 'Articles');
    }

}
