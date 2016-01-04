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

    /**
     * test method which converts entities to a flat array
     */
    public function testPrepareEntityData()
    {
        $entity = $this->Articles->get(1);

        $result = $this->Excel->prepareEntityData($entity);

        $this->assertEquals(count($result), 2);

        $headers = $result[0];
        $properties = $result[1];

        $this->assertEquals(count($headers), count($properties));

        $this->assertEquals($headers, ['id', 'author_id', 'title', 'body', 'published']);

        $this->assertEquals($properties, [1, 1, 'First Article', 'First Article Body', 'Y']);
    }

    /**
     * test method which converts a collection to a flat array
     */
    public function testPrepareCollectionData()
    {
        $articles = $this->Articles->find('all');

        $data = collection($articles->toArray());

        $result = $this->Excel->prepareCollectionData($data);

        $this->assertEquals(count($result), 4);

        $headers = $result[0];
        $first = $result[1];
        $last = $result[3];

        $this->assertEquals($headers, ['id', 'author_id', 'title', 'body', 'published']);

        $this->assertEquals($first, [1, 1, 'First Article', 'First Article Body', 'Y']);

        $this->assertEquals($last, [3, 1, 'Third Article', 'Third Article Body', 'Y']);
    }

}
