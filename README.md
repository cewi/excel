# Cewi/Excel plugin for CakePHP

! Not  Ready For Production !  

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
composer require Cewi/Excel
```

should fetch the plugin.

You can cerate Excel Workbooks from views. This works is in [dakotas](https://github.com/dakota/CakeExcel) plugin. Look there for docs. Additions:

1. ExcelHelper: Takes a Query-Object and creates a worksheet. Properties are column-headers in first row.

Example (assumed you have an article model and controller with the usual index-action) 

include the helper in the ArticleController:

    public $helpers = ['Cewi/Excel.Excel'];

add a Folder 'xlsx' in Template/Articles and create the file 'index.ctp' in this Folder:
    
    <?php
    $this->Excel->Metadata($this->name);
    $data = $this->Excel->prepareData($articles);
    $this->Excel->addData($data);
    
create the link somwhere in your app: 

    <?= $this->Html->link(__('Excel'), ['controller' => 'Articles', 'action' => 'index', '_ext'=>'xlsx']); ?>

done.
