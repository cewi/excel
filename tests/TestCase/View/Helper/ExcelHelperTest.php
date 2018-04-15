<?php

namespace Cewi\Excel\Test\TestCase\View\Helper;

use Cewi\Excel\View\Helper\ExcelHelper;
use Cewi\Excel\View;
use Cake\TestSuite\TestCase;
use Cake\ORM\TableRegistry;
use Cake\I18n\FrozenTime;

/**
 * App\View\Helper\ExcelHelper Test Case
 */
class ExcelHelperTest extends TestCase
{

    public $fixtures = ['plugin.Cewi/Excel.articles'];

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

        //dd($entity);

        $result = $this->Excel->prepareEntityData($entity);

        $this->assertEquals(count($result), 2);

        $headers = $result[0];
        $properties = $result[1];

        $this->assertEquals(count($headers), count($properties));

        $this->assertEquals($headers, ['id', 'name', 'body', 'published', 'created', 'modified', 'number']);

        $this->assertEquals($properties, [
            1,
            'First Article',
            'First Article Body',
            1,
            new FrozenTime('2007-03-18T10:39:23+00:00'),
            new FrozenTime('2007-03-18T10:41:31+00:00'),
            1.4
        ]);
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

        $this->assertEquals($headers, ['id', 'name', 'body', 'published', 'created', 'modified', 'number']);

        $this->assertEquals($first, [
            1,
            'First Article',
            'First Article Body',
            1,
            new FrozenTime('2007-03-18T10:39:23'),
            new FrozenTime('2007-03-18T10:41:31'),
            1.4
        ]);

        $this->assertEquals($last, [
            3,
            'Third Article',
            '000123',
            0,
            new FrozenTime('2007-03-18 10:43:23'),
            new FrozenTime('2007-03-18 10:45:31'),
            20000000
        ]);
    }

}
