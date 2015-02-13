<?php

namespace Cewi\Excel\Controller\Component;

use Cake\Controller\Component;
use Cake\Controller\ComponentRegistry;
use Cake\ORM\Exception\MissingTableClassException;
use Cake\ORM\Table;
use Cake\Database\Schema;
use Cake\Utility\Text;

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
     * @param array $options Override Worksheet name
     * @return array Array like normally provided by request->data
     * @throws MissingTableClassException
     */
    public function prepareData($file = null, $table = null, array $options = [])
    {

        /**  load and configure PHPExcelReader  * */
        \PHPExcel_Cell::setValueBinder(new \PHPExcel_Cell_AdvancedValueBinder());
        $fileType = \PHPExcel_IOFactory::identify($file);
        $PhpExcelReader = \PHPExcel_IOFactory::createReader($fileType);
        $PhpExcelReader->setReadDataOnly(true);

        /** identify worksheets in file and advise the Reader which WorkSheet we want to load  * */
        $worksheets = $PhpExcelReader->listWorksheetNames($file);

        if (isset($options['worksheet'])) {
            $worksheet = $options['worksheet']; //desired Worksheet was provided as option
        } elseif (count($worksheets) === 1) {
            $worksheet = $worksheets[0]; //take the only worksheet in file
        } elseif (in_array($table->alias(), $worksheets)) {
            $worksheet = $table->alias(); //take the worksheet with the name of the table
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
            $result[] = $record;
        }
        $this->log('Worksheet' . $worksheet . ' contained ' . count($result) . ' records', 'debug');
        return $result;
    }

    /**
     * Truncate Table and then import records
     * value of primary key will be kept
     *
     * @param string $file filename with full path
     * @param Table $table
     * @param array $options Override Worksheet name
     * @return int Number of imported records
     */
    public function replace($file = null, $table = null, array $options = [])
    {
        $result = ['successful' => 0, 'failed' => 0];

        //truncate Table
        $table->deleteAll([$table->primaryKey() . ' >' => 0]);

        //prepare Data
        $data = $this->prepareData($file, $table, $options);

        //create entitites
        $fieldList = $table->schema()->columns();
        $entities = $table->newEntities($data, ['fieldList' => $fieldList]);
        $this->log(count($entities) . ' ready to insert', 'debug');

        //save data        
        foreach ($entities as $entity) {
            if ($table->save($entity, ['checkExisting' => false])) {
                $result['successful'] ++;
            } else {
                $this->_logErrors($entity);
                $result['failed'] ++;
            }
        }
        return $result;
    }

    /**
     * adds records to table
     *
     * @param string $file filename with full path
     * @param type $table
     * @param array $options Override Worksheet name
     * @return int Number of imported records
     */
    public function append($file = null, $table = null, array $options = [])
    {
        $result = ['successful' => 0, 'failed' => 0];
        //prepare Data
        $data = $this->prepareData($file, $table, $options);
        $entities = $table->newEntities($data);
        //save data
        foreach ($entities as $entity) {
            if ($table->save($entity)) {
                $result['successful'] ++;
            } else {
                $this->_logErrors($entity);
                $result['failed'] ++;
            }
        }
        return $result;
    }

    /**
     * Update tabel with records
     *
     * @param string $file
     * @param \Cake\ORM\Table $table
     * @param array $options Override Worksheet name
     * @return array Result
     */
    public function update($file = null, $table = null, array $options = [])
    {
        $result = ['replaced' => 0, 'added' => 0];
        //prepare Data
        $data = $this->prepareData($file, $table, $options);
        //save data
        foreach ($data as $record) {
            $entity = $table->findOrCreate($record);
            if ($entity->isNew()) {
                $result['added'] ++;
            } else {
                $result['replaced'];
            }
        }
        return $result;
    }

    /**
     * Log errors
     *
     * @param \Cake\ORM\Table\Entity $entity
     * @return void
     */
    protected function _logErrors($entity = null)
    {
        $fields = $entity->visibleProperties();
        $displayfield = $fields[1];
        $this->log($entity->$displayfield . ' could not be saved: ', 'error');
        foreach ($entity->errors() as $error) {
            foreach ($error as $message) {
                $this->log($message . ' is not correct!', 'error');
            }
        }
        return;
    }

}
