# Cewi/Excel plugin for CakePHP 

The plugin is based on the work of [dakota]
(https://github.com/dakota/CakeExcel) and uses [PHPExcel](https://phpexcel.codeplex.com/) for the excel-related functionality. 

## Installation

You can install this plugin into your CakePHP application using [composer](http://getcomposer.org).

The recommended way to install composer packages is:

add 

    "repositories": [
             {
                "type": "vcs",
                "url": "https://github.com/cewi/excel"
            }
        ] 
        
 to your composer.json because this package is not on packagist. Then in your console:

```
composer require Cewi/Excel:dev-master
```

should fetch the plugin.

You can create Excel Workbooks from views. This works like in [dakotas](https://github.com/dakota/CakeExcel) plugin. Look there for docs. Additions:

## 1. ExcelHelper
Has a Method 'addWorksheet' which can take either an Entity, a ResultSet, a Collection or an Array and creates a worksheet from the data. If a Entity is provided, the Column-Headers will be the properties. In case of a ResultSet, a Collection or an Array of Entities, the Column-Headers are taken from the properties of the first Entity.

Example (assumed you have an article model and controller with the usual index-action) 

First include the helper in ArticlesController:

    public $helpers = ['Cewi/Excel.Excel'];

add a Folder 'xlsx' in Template/Articles and create the file 'index.ctp' in this Folder. Include this snippet of code to get an excel-file with a single worksheet called 
'Articles':    
    
    <?= $this->Excel->addWorksheet($articles, 'Articles');?>
    
create the link to generate the file somewhere in your app: 

    <?= $this->Html->link(__('Excel'), ['controller' => 'Articles', 'action' => 'index', '_ext'=>'xlsx']); ?>

done.

## 2. Import-Component

Takes a excel workbook, extracts a single worksheet with data (e.g. generated with the helper) and generates an array with data ready for building entities. 

Include the Import-Component in the controller:

     public function initialize()
     {
        parent::initialize();
        $this->loadComponent('Cewi/Excel.Import');
     }    

than you can use the method

     prepareEntityData($file = null, array $options = [])

	e.g.	

     $data = $this->Import->prepareEntityData(TMP . $this->request->data('file.name'));

and you'll get an array with data like you would get from the form-helper. You then can generate and save entities in the Controller:

     $entities = $table->newEntities($data);
     foreach ($entities as $entitiy) {
           $table->save($entitiy, ['checkExisting' => false])
     }

if your table is not empty and you don't want to replace records in the database, set `'append'=>true` in the $options array:

    $data = $this->Import->prepareEntityData($file, ['append'=> true]);

If there are more than one worksheets in the file you can supply the name of the Worksheet to use in the $options array, e.g.: `'worksheet'=>'Articles'`.
