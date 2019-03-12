<?php

namespace Cewi\Excel\View\Helper;

use Cake\Collection\Collection;
use Cake\Core\Configure;
use Cake\Database\Expression\QueryExpression;
use Cake\I18n\Date;
use Cake\I18n\FrozenDate;
use Cake\I18n\FrozenTime;
use Cake\I18n\Time;
use Cake\ORM\Entity;
use Cake\ORM\Query;
use Cake\ORM\ResultSet;
use Cake\View\Helper;
use Cake\View\View;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/*
 * The MIT License
 *
 * Copyright 2018 cewi <c.wichmann@gmx.de>.
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
 * CakePHP ExcelHelper
 *
 * can add data to a workbook
 *
 * Data may be an Entity, a Query Expression, a Collection of entitites or a flat Array
 *
 * @author cewi <c.wichmann@gmx.de>
 */
class ExcelHelper extends Helper
{

    /**
     * Format in which dates are exported to excel
     * set it globally in the bootstrap file or pass it as config-Variable
     *
     * @var string
     */
    private $__dateformat = 'yyyy-MM-dd';

    /**
     * Constructor
     *
     * @param View $View
     * @param array $config
     */
    public function __construct(View $View, array $config = [])
    {
        parent::__construct($View, $config);

        if (isset($config['dateformat'])) {
            $this->__dateformat = $config['dateformat'];
        } else {
            $this->__dateformat = Configure::read('excel.dateformat');
        }
    }

    /**
     * Add new Worksheet
     *
     * @param mixed $data can be Query, Entity, Collection or flat Array
     * @param string $name
     * @return void
     */
    public function addWorksheet($data = null, $name = '')
    {
        // Add empty sheet to Workbook
        $this->addSheet($name);

        if (is_array($data)) {
            $data = $this->prepareCollectionData(collection($data));
        } elseif ($data instanceof Entity) {
            $data = $this->prepareEntityData($data);
        } elseif ($data instanceof Query) {
            $data = $this->prepareCollectionData(collection($data->toArray()));
        } elseif ($data instanceof ResultSet) {
            $data = $this->prepareCollectionData(collection($data->toArray()));
        } else {
            $data = $this->prepareCollectionData($data);
        }
        // Add the Data
        $this->addData($data);

        //auto-sizing of the columns
        $highestColumn = $this->_View->PHPSpreadsheet->getActiveSheet()->getHighestColumn();
        foreach (range('A', $highestColumn) as $column) {
            $this->_View->PHPSpreadsheet->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
        }

        return;
    }

    /**
     * Converts a Collection into a flat Array
     * properties are extracted from first item und inserted in first row
     *
     * @param mixed $collection \Cake\Collection\Collection | \Cake\ORM\Query
     * @return array
     */
    public function prepareCollectionData(Collection $collection = null)
    {
        /* Extract keys from first item */
        $first = $collection->first();
        if (is_array($first)) {
            $data = [array_keys($first)];
        } else {
            $data = [array_keys($first->toArray())];
        }

        /* Add data */
        foreach ($collection as $row) {
            if (is_array($row)) {
                $data[] = array_values($row);
            } else {
                $data[] = array_values($row->toArray());
            }
        }

        return $data;
    }

    /**
     * Converts a Entity into a flat Array
     * properties are inserted in first row
     *
     * @param Entity $entity
     * @return array
     */
    public function prepareEntityData(Entity $entity = null)
    {
        $entityArray = $entity->toArray();
        $data = [array_keys($entityArray)];
        $data[] = array_values($entityArray);

        return $data;
    }

    /**
     * Adds data to a worksheet
     *
     * @param array $data data
     * @param array $options if set row and column, data entry starts there
     * @return void
     */
    public function addData(array $data = [], array $options = [])
    {
        $rowIndex = isset($options['row']) ? $options['row'] : 1;
        foreach ($data as $row) {
            $columnIndex = isset($options['column']) ? $options['column'] : 1; // In PHPSpreadsheet Columns start with 1!
            foreach ($row as $cell) {
                $this->_addCellData($cell, $columnIndex, $rowIndex);
                $columnIndex++;
            }
            $rowIndex++;
        }

        return;
    }

    /**
     * Fills in the data in a cell.
     * respects data type
     *
     * @param mixed $cell
     * @param int $columnIndex
     * @param int $rowIndex
     * @return void
     */
    protected function _addCellData($cell = null, $columnIndex = 1, $rowIndex = 1)
    {
        if (is_array($cell)) {
            $cell = null; // adding cells of this Type is useless

            return;
        }
        if ($cell instanceof Date || $cell instanceof Time || $cell instanceof FrozenDate || $cell instanceof FrozenTime) {
            $cell = $cell->toUnixString(); // Dates must be converted in unix
            $coordinate = $this->_View->PHPSpreadsheet
                ->getActiveSheet()
                ->getCellByColumnAndRow($columnIndex, $rowIndex)
                ->getCoordinate();

            $this->_View->PHPSpreadsheet
                ->getActiveSheet()
                ->setCellValue($coordinate, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($cell))
                ->getStyle($coordinate)
                ->getNumberFormat()
                ->setFormatCode($this->__dateformat);

            return;
        }
        if ($cell instanceof QueryExpression) {
            $cell = null; //TODO find a way to get the Values and insert them into the Sheet

            return;
        }
        if (is_string($cell)) {
            $this->_View->PHPSpreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValueExplicit($cell, DataType::TYPE_STRING);

            return;
        }
        $this->_View->PHPSpreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValueExplicit($cell, DataType::TYPE_NUMERIC);

        return;
    }

    /**
     * create empty Sheet and add some Metadata
     *
     * @param string $title
     * @return void
     */
    public function addSheet($title = '')
    {
        $this->_View->PHPSpreadsheet->createSheet();
        $this->_View->currentSheetIndex++;
        $this->_View->PHPSpreadsheet->setActiveSheetIndex($this->_View->currentSheetIndex);
        $this->_View->PHPSpreadsheet->getActiveSheet()->setTitle($title);
        $this->_View->PHPSpreadsheet->getProperties()->setTitle($title);
        $this->_View->PHPSpreadsheet->getProperties()->setSubject($title . ' ' . date('d.m.Y H:i'));

        return;
    }
}
