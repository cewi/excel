<?php

namespace Cewi\Excel\Controller\Component;

use Cake\Controller\Component;
use Cake\Controller\ComponentRegistry;
use Cake\ORM\Exception\MissingTableClassException;

/**
 * Import component
 *
 * Convert a worksheet with Database-Records in rows to an array which can
 * be used to build entities
 */
class ImportComponent extends Component
{

    /**
     * reads a worksheet from excel-file and converts every row in an array
     * which can be used to build entities
     *
     * @param string $file name of Excel-File with full path. Must be of a readable Filetype
     * @param Table $table TableObject which will be used to build entities. Needed to identify worksheet to load
     * @param array $options Override Worksheet name, set append Mode
     * @return array Array like normally provided by request->data
     * @throws MissingTableClassException
     */
    public function prepareEntityData($file = null, array $options = [])
    {

        /**  load and configure PHPExcelReader  * */
        \PHPExcel_Cell::setValueBinder(new \PHPExcel_Cell_AdvancedValueBinder());
        $fileType = \PHPExcel_IOFactory::identify($file);
        $PhpExcelReader = \PHPExcel_IOFactory::createReader($fileType);
        $PhpExcelReader->setReadDataOnly(true);

        /** identify worksheets in file and advise the Reader which WorkSheet we want to load  * */
        $worksheets = $PhpExcelReader->listWorksheetNames($file);

        /** if tableName is not set, use name of current $controller */
        if(!isset($options['tableName']) || empty($options['tableName'])){
            $options['tableName'] = $this->_registry->getController()->name;
        }

        if (isset($options['worksheet'])) {
            $worksheet = $options['worksheet']; //desired Worksheet was provided as option
        } elseif (count($worksheets) === 1) {
            $worksheet = $worksheets[0]; //take the only worksheet in file
        } elseif (in_array($tableName, $worksheets)) {
            $worksheet = $tableName; //take the worksheet with the name of the table
        } else {
            throw new MissingTableClassException(__('No proper named worksheet found'));
        }
        /** load the sheet and convert data to an array */
        $PhpExcelReader->setLoadSheetsOnly($worksheet);
        $PhpExcel = $PhpExcelReader->load($file);
        $data = $PhpExcel->getSheet(0)->toArray();

        /** convert data for building entities */
        $result = [];
        $properties = array_shift($data); //Column names in first row are the properties names
        foreach ($data as $row) {
            $record = array_combine($properties, $row);
            if (isset($record['modified'])) {
                unset($record['modified']);
            }
            if (isset($options['type']) && $options['type'] == 'append' && isset($record['id'])){
                unset($record['id']);
            }
            $result[] = $record;
        }
        $this->log('Worksheet' . $worksheet . ' contained ' . count($result) . ' records', 'debug');
        return $result;
    }

}
