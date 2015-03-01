<?php

namespace Cewi\Excel\View\Helper;

use Cake\View\Helper;
use Cake\View\View;
use Cake\ORM\Query;

/*
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
 * CakePHP ExcelHelper
 *
 * can add a bunch of entities as a worksheet to a workbook
 *
 * @author cewi <c.wichmann@gmx.de>
 */
class ExcelHelper extends Helper
{

    public function addTable(Query $query = null, $name = '')
    {
        $this->Metadata($name);
        $data = $this->prepareTableData($query);
        $this->addTableData($data);
        return;
    }

    /**
     * converts a query Object into a flat table
     * properties are filed in first row
     *
     * @param Query $query
     * @return array
     */
    public function prepareTableData(Query $query = null)
    {

        /* first row contains properties */
        $data = [array_keys($query->first()->toArray())];

        /** add data of an entity a row */
        foreach ($query as $entity) {
            $data[] = array_values($entity->toArray());
        }

        return $data;
    }

    /**
     * adds data to a worksheet
     *
     * @param array $array
     * @param array $options if set row and column, data entry starts there
     * @return void
     */
    public function addTableData(array $array = [], array $options = [])
    {
        $rowIndex = isset($options['row']) ? $options['row'] : 1;
        foreach ($array as $row) {
            $columnIndex = isset($options['column']) ? $options['column'] : 0;
            foreach ($row as $cell) {
                $this->_View->PhpExcel->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex)->setValue($cell);
                $columnIndex++;
            }
            $rowIndex++;
        }

        //auto-sizing of the columns
        $highestColumn = $this->_View->PhpExcel->getActiveSheet()->getHighestColumn();
        foreach (range('A', $highestColumn) as $column) {
            $this->_View->PhpExcel->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
        }

        return;
    }

    /**
     * adds Metadata to the Worksheet
     *
     * @param string $title
     * @return void
     */
    public function MetaData($title = '')
    {
        $this->_View->PhpExcel->getActiveSheet()->setTitle($title);
        $this->_View->PhpExcel->getProperties()->setTitle($title);
        $this->_View->PhpExcel->getProperties()->setSubject($title . ' ' . date('d.m.Y H:i'));
        return;
    }

}
