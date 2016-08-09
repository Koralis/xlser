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

    public function createWriter()
    {
        return new XlsWriter($this);
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

    public function appendRow(array $rowData)
    {
        foreach ($rowData as $val) {
            $format = null;
            if (is_array($val)) {
                if (isset($val[1]) && $val[1]) {
                    $format = $val[1];
                    if (is_scalar($format)) {
                        $format = CellFormat::createFromScalar($format);
                    }
                }
                if (isset($val[0])) {
                    $val = $val[0];
                }
            }

            $this->appendCell($val, $format);
        }

        $this->_resetCol();
        $this->_advanceRow();
    }

    public function appendHeaderRow(array $data)
    {
        $this->appendRow($data);

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
     * @param $format
     *
     * @throws \PHPExcel_Exception
     */
    public function appendCell($value, CellFormat $format = null)
    {
        $sheet = $this->getActiveSheet();
        $cell = $sheet->getCellByColumnAndRow($this->_getCurrentCol(), $this->_getCurrentRow());

        if (is_callable($value)) {
            $value = $value($cell);
        }

        $cell->setValue($value);

        if ($format) {
            $format->apply($cell);
        }

        $this->_advanceCol();
    }

    public function autoColWidth($bool)
    {
        $this->autoColWidth = $bool;
    }

}
