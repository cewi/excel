<?php

namespace Cewi\Excel\Controller\Component;

use Cake\Controller\Component;
use Cake\Controller\ComponentRegistry;
use Cake\ORM\Exception\MissingTableClassException;

/**
 * The MIT License
 *
 * Copyright 2015 cewi <c.wichmann@gmx.de>.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

/**
 * CakePHP Import Component
 *
 * @author cewi <c.wichmann@gmx.de>
 */
class ImportComponent extends Component
{

    /**
     * reads a file PHPExcel can understand and converts a contained worksheet into an array
     * which can be used to build entities. If the File contains more than one worksheet and it is not named like the Controller
     * you have to provide the name of the workshhet to load in the options array.
     * If you set $options['append'] to true, the primary key will be deleted.
     * @todo Find a way to make to handle primary keys anmed other than id.
     *
     * @param string $file name of Excel-File with full path. Must be of a readable Filetype (xls, xlsx, csv, ods)
     * @param array $options Override Worksheet name, set append Mode
     * @return array . The Array has the same structure as provided by request->data
     * @throws MissingTableClassException
     */
    public function prepareEntityData($file = null, array $options = [])
    {

        /**  load and configure PHPExcelReader  * */
        \PHPExcel_Cell::setValueBinder(new \PHPExcel_Cell_AdvancedValueBinder());
        $fileType = \PHPExcel_IOFactory::identify($file);

        $PhpExcelReader = \PHPExcel_IOFactory::createReader($fileType);
        $PhpExcelReader->setReadDataOnly(true);

        if ($fileType !== 'CSV') {  // csv-files have only one 'worksheet'

            /** identify worksheets in file * */
            $worksheets = $PhpExcelReader->listWorksheetNames($file);

            $worksheetToLoad = null;

            if (count($worksheets) === 1) {
                $worksheetToLoad = $worksheets[0];  //first option: if there is only one worksheet, use it
            } elseif (isset($options['worksheet'])) {
                $worksheetToLoad = $options['worksheet']; //second option: desired worksheet was provided as option
            } else {
                $worksheetToLoad = $this->_registry->getController()->name; //last option: try to load worksheet with the name of current controller
            }
            if (!in_array($worksheetToLoad, $worksheets)) {
                throw new MissingTableClassException(__('No proper named worksheet found'));
            }

            /** load the sheet and convert data to an array */
            $PhpExcelReader->setLoadSheetsOnly($worksheetToLoad);
        }

        $PhpExcel = $PhpExcelReader->load($file);
        $data = $PhpExcel->getSheet(0)->toArray();

        /** convert data for building entities */
        $result = [];
        $properties = array_shift($data); //first row columns are the properties

        foreach ($data as $row) {
            $record = array_combine($properties, $row);
            if (isset($record['modified'])) {
                unset($record['modified']);
            }
            if (isset($options['type']) && $options['type'] == 'append' && isset($record['id'])) {
                unset($record['id']);
            }
            $result[] = $record;
        }

        /** log in debug mode */
        $this->log(count($result) . ' records were extracted from File ' . $file, 'debug');

        return $result;
    }

}
