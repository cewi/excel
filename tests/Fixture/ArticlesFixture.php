<?php

namespace Cewi\Excel\Test\Fixture;

use Cake\TestSuite\Fixture\TestFixture;

/**
 * ArticlesFixture
 *
 * configuration and data taken from https://book.cakephp.org/3.0/en/development/testing.html#fixtures
 * with slight addon
 *
 */
class ArticlesFixture extends TestFixture
{

    /**
     * Fields
     *
     * @var array
     */
    // @codingStandardsIgnoreStart
    public $fields = [
        'id' => ['type' => 'integer'],
        'name' => ['type' => 'string', 'length' => 255, 'null' => false],
        'body' => 'text',
        'published' => ['type' => 'integer', 'default' => '0', 'null' => false],
        'created' => 'datetime',
        'modified' => 'datetime',
        'number' => 'float',
        '_constraints' => [
            'primary' => ['type' => 'primary', 'columns' => ['id']]
        ]
    ];
    // @codingStandardsIgnoreEnd
    public $records = [
        [
            'name' => 'First Article',
            'body' => 'First Article Body',
            'published' => true,
            'created' => '2007-03-18 10:39:23',
            'modified' => '2007-03-18 10:41:31',
            'number' => 1.4
        ],
        [
            'name' => 'Second Article',
            'body' => 'Second Article Body',
            'published' => true,
            'created' => '2007-03-18 10:41:23',
            'modified' => '2007-03-18 10:43:31',
            'number' => -1
        ],
        [
            'name' => 'Third Article',
            'body' => '000123',
            'published' => false,
            'created' => '2007-03-18 10:43:23',
            'modified' => '2007-03-18 10:45:31',
            'number' => 20000000
        ]
    ];

}
