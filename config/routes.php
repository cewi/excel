<?php

use Cake\Routing\Router;

Router::extensions(['xlsx']);


Router::plugin('Cewi/Excel', null, function($routes){
    $routes->connect('/:controller/:action');
});
