# Cewi/Excel plugin for CakePHP 

The plugin is based on the work of [dakota]
(https://github.com/dakota/CakeExcel) and uses [PHPExcel](https://github.com/PHPOffice/PHPExcel) for the excel-related functionality. 

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

Load the Plugin in your bootstrap.php as ususal:

```
	Plugin::load('Cewi/Excel', ['bootstrap' => true, 'routes'=>true]);
```

RequestHandler Component is configured by the Plugin's bootstrap file. If not you could do this this in your controller's initialize method, e.g.:

```
	public function initialize()
		{
        		parent::initialize();
        		$this->loadComponent('RequestHandler', [
            			'viewClassMap' => ['xlsx' => 'Cewi/Excel.Excel']
        ]);
        	}
```
Be careful: RequestHandlerComponent is already loaded in your AppController by default. Adapt the settings to your needs.


You need to set up parsing for the xlsx extension. Add the following to your config/routes.php file before any route or scope definition:

```
Router::extensions('xlsx');
```
or you can add within a scope:

```
$routes->setExtensions(['xlsx']);
```
(Setting this in the plugin's config/routes.php file is currently broken. So you do have to provide the code in the application's config/routes.php file)


You further have to provide a layout for the generated Excel-Files. Add a folder xlsx in src/Template/Layout/ subdirectory and within that folder a file named default.ctp with this minimum content:
```  
<?= $this->fetch('content') ?>
```  

You can create Excel Workbooks from views. This works like in [dakotas](https://github.com/dakota/CakeExcel) plugin. Look there for docs. Additions:

## 1. ExcelHelper
Has a Method 'addworksheet' which takes a ResultSet, a Entity, a Collection of Entities or an Array of Data and creates a worksheet from the data. Properties of the Entities, or the keys of the first record in the array are set as column-headers in first row of the generated worksheet. Be careful if you use non-standard column-types. The Helper actually works only with strings, numbers and dates. 

Register xlsx-Extension in config/routes.php file before the routes that should be affected:

```
    Router::extensions(['xlsx']);
```

Example (assumed you have an article model and controller with the usual index-action) 

Include the helper in ArticlesController:

```
   public $helpers = ['Cewi/Excel.Excel'];
```

add a Folder 'xlsx' in Template/Articles and create the file 'index.ctp' in this Folder. Include this snippet of code to get an excel-file with a single worksheet called 
'Articles':    
    
```    
    $this->Excel->addWorksheet($articles, 'Articles');
```    
    
create the link to generate the file somewhere in your app: 

```
    <?= $this->Html->link(__('Excel'), ['controller' => 'Articles', 'action' => 'index', '_ext'=>'xlsx']); ?>
```

done.

## 2. ImportComponent

Takes a excel workbook, extracts a single worksheet with data (e.g. generated with the helper) and generates an array with data ready for building entities. First row must contain names of properties/database columns.

Include the Import-Component in the controller:

     public function initialize()
     {
        parent::initialize();
        $this->loadComponent('Cewi/Excel.Import');
     }    

than you can use the method

     prepareEntityData($file = null, array $options = [])
     
E.g. if you've uploaded a file:

     move_uploaded_file($this->gerRequest()->getData('file.tmp_name'), TMP . DS . $this->getRequest()->getData('file.name'));
     $data = $this->Import->prepareEntityData(TMP . $this->getRequest()->getData('file.name'));

and you'll get an array with data like you would get from the form-helper. You then can generate and save entities in the Controller:

     $entities = $table->newEntities($data);
     foreach ($entities as $entity) {
           $table->save($entity, ['checkExisting' => false])
     }

if your table is not empty and you don't want to replace records in the database, set `'append'=>true` in the $options array:

    $data = $this->Import->prepareEntityData($file, ['append'=> true]);

If there are more than one worksheets in the file you can supply the name or index of the Worksheet to use in the $options array, e.g.: 
 
	$data = $this->Import->prepareEntityData($file, ['worksheet'=> 0]);
	
or
	
	$data = $this->Import->prepareEntityData($file, ['worksheet'=> 'Articles']);

