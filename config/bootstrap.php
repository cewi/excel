<?php

use Cake\Event\EventManager;
use Cake\Datasource\ConnectionManager;

/**
 * use sqlite for temporary Data
 * borrowed idea from debug_kit
 */
ConnectionManager::config('excel', [
    'className' => 'Cake\Database\Connection',
    'driver' => 'Cake\Database\Driver\Sqlite',
    'database' => TMP . 'excel.sqlite',
    'encoding' => 'utf8',
    'cacheMetadata' => true,
    'quoteIdentifiers' => false,
]);


/**
 * load and prepare RequestHandler in all Controllers
 */
EventManager::instance()
        ->attach(
                function (Cake\Event\Event $event) {
            $controller = $event->subject();
            if ($controller->components()->has('RequestHandler')) {
                $controller->RequestHandler->viewClassMap('xlsx', 'Cewi/Excel.Excel');
            }
        }, 'Controller.initialize'
);

