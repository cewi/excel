<?php

namespace Cewi\Excel\Controller\Component;

use Cake\Controller\Component;
use Cake\Controller\ComponentRegistry;
use Cake\Validation\Validator;
use Cake\ORM\Exception\MissingTableClassException;
use Cake\ORM\Table;

/**
 * Import component
 *
 * Convert a worksheet with Database-Records in rows to an array which can
 * be used to build entities
 */
class ImportComponent extends Component
{

    /**
     * validate if uplaoded File is a valid Excel-File
     *
     * @param array $fileArray // as privded from Form-Upload
     * @return array
     */
    public function validate(array $fileArray = [])
    {
        $validator = new Validator();
        $validator
                ->requirePresence('name')
                ->notEmpty('name', __('You must submit a file'))
                ->add('type', 'validValue', [
                    'rule' => ['inList', [
                            'application/vnd.ms-excel',
                            'application/msexcel',
                            'application/x-msexcel',
                            'application/x-ms-excel',
                            'application/x-excel',
                            'application/x-dos_ms_excel',
                            'application/xls',
                            'application/x-xls',
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']],
                    'message' => __('File has wrong mime-Type')
        ]);
        return $validator->errors($fileArray);
    }

    /**
     * reads a worksheet from excel-file and converts every row in an array
     * which can be used to build entities
     *
     * @param string $file name of Excel-File with full path. Must be of a readable Filetype
     * @param Table $table TableObject which will be used to build entities. Needed to identify worksheet to load
     * @param array $options Override Worksheet name and remove prmary keys to add records instead of updating
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
        $properties = array_shift($data); //Column names in first row name the properties
        foreach ($data as $row) {
            $record = array_combine($properties, $row);
            if (isset($record['modified'])) {
                unset($record['modified']);
            }
            $result[] = $record;
        }

        return $result;
    }

    public function overwrite($file = null, $table = null)
    {
        $data = $this->prepareData($file, $table);
        $table->deleteAll([$table->primaryKey() . ' >' => 0]);
        foreach ($data as $record) {
            $table->findOrCreate($record);
        }
        return count($data);
    }

    public function add($file = null, $table = null)
    {
        $data = $this->prepareData($file, $table);
        $imported = 0;
        $entities = $table->newEntities($data);
        foreach ($entities as $entity) {
            if ($table->save($entity)) {
                $imported++;
            }
        }
        return $imported;
    }

}
