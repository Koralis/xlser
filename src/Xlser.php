<?php

namespace Sakalys\Xlser;

use PHPExcel;
use PHPExcel_Cell;
use ReflectionObject;

class Xlser extends PHPExcel
{
    const OUTPUT_MEMORY = 1;

    const STYLE_BOLD = 1;

    protected $currentRow = 1;
    protected $currentCol = 0;

    /** @var bool */
    protected $autoColWidth;

    public function __construct()
    {
        parent::__construct();
    }

    public static function formatCurrency($symbol, $precision, $symbolAfterValue = false, $thousandSeparator = ',', $decimalSeparator = '.')
    {
        return new CellFormat($symbol, $precision, $symbolAfterValue, $thousandSeparator, $decimalSeparator);
    }

    public function setData(array $data, $headers = [])
    {
        $this->_resetCursor();

        $this->appendRow($headers);

        foreach ($data as $rowData) {
            $this->appendRow($rowData);
        }

        return $this;
    }

    public function output($type)
    {
        $writer = \PHPExcel_IOFactory::createWriter($this, 'Excel5');

        switch ($type) {
            case self::OUTPUT_MEMORY;

                $file = tempnam(sys_get_temp_dir(), 'xlser');
                $writer->save($file);
                $contents = file_get_contents($file);
                unlink($file); //to delete an empty file that tempnam creates

                return $contents;

            default:
                throw new \Exception('Unknown output type');
        }
    }

    public function appendRow(array $rowData, CellFormat $rowFormat = null)
    {
        foreach ($rowData as $val) {
            $cellFormat = null;
            if (is_array($val)) {
                if (isset($val[1]) && $val[1]) {
                    $cellFormat = $val[1];
                    if (is_scalar($cellFormat)) {
                        $cellFormat = CellFormat::createFromScalar($cellFormat);
                    }
                }


                if (isset($val[0])) {
                    $val = $val[0];
                }
            }

            if ($rowFormat) {
                if (is_scalar($rowFormat)) {
                    $rowFormat = CellFormat::createFromScalar($rowFormat);
                }
            }

            $this->insertVal($val, $cellFormat, $rowFormat);
        }

        $this->_resetCol();
        $this->_advanceRow();
    }

    public function appendHeaderRow(array $data)
    {
        $rowFormat = new CellFormat();
        $rowFormat->setBold(true);
        $this->appendRow($data, $rowFormat);

        return $this;
    }

    private function _resetCursor()
    {
        $this->_resetRow();
        $this->_resetCol();
    }

    private function _resetCol()
    {
        $this->currentCol = 0;
    }

    private function _resetRow()
    {
        $this->currentRow = 1;
    }

    /**
     * Advances the row cursor position
     *
     * @return self
     */
    private function _advanceRow()
    {
        $this->currentRow++;

        return $this;
    }

    /**
     * Advances the colum cursor position
     *
     * @return self
     */
    private function _advanceCol()
    {
        $this->currentCol++;

        return $this;
    }

    /**
     * @return int
     */
    private function _getCurrentRow()
    {
        return $this->currentRow;
    }

    private function _getCurrentCol()
    {
        return $this->currentCol;
    }

    public function getRow()
    {
        return $this->_getCurrentRow();
    }

    public function getCol()
    {
        return $this->_getCurrentCol();
    }

    public function newRow()
    {
        return $this->_advanceRow();
    }

    public function newCol()
    {
        return $this->_advanceCol();
    }

    /**
     * @param $value
     * @param CellFormat $cellFormat
     * @param CellFormat $rowFormat
     * @return PHPExcel_Cell
     * @throws \PHPExcel_Exception
     */
    public function insertVal($value, CellFormat $cellFormat = null, CellFormat $rowFormat = null)
    {
        $sheet = $this->getActiveSheet();
        $cell = $sheet->getCellByColumnAndRow($this->_getCurrentCol(), $this->_getCurrentRow());

        if (is_callable($value)) {
            $value = $value($cell);
        }

        $cell->setValue($value);

        if ($rowFormat) {
            $rowFormat->apply($cell);
        }

        if ($cellFormat) {
            $cellFormat->apply($cell);
        }

        $this->_advanceCol();

        return $cell;
    }

    public function autoColWidth($bool)
    {
        $this->autoColWidth = $bool;
    }

}
