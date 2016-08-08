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

    /** @var ReflectionObject */
    protected $refl;

    public function __construct()
    {
        parent::__construct();
        $this->refl = new ReflectionObject($this);
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

    public function appendRow(array $rowData, $flags = 0)
    {
        $sheet = $this->getActiveSheet();

        foreach ($rowData as $val) {
            /** @var PHPExcel_Cell $cell */
            $cell = $sheet->setCellValueByColumnAndRow($this->_getCurrentCol(), $this->_getCurrentRow(), $val, true);

            if ($flags) {
                switch (true) {
                    case $flags & self::STYLE_BOLD:
                        $cell->getStyle()->getFont()->setBold(true);
                    default:
                        break;
                }
            }

            $this->_advanceCol();
        }

        $this->_resetCol();
        $this->_advanceRow();
    }

    public function appendHeaderRow(array $data)
    {
        $this->appendRow($data, self::STYLE_BOLD);

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
     * Advances the row cursor position and returns the last position
     *
     * @return int
     */
    private function _advanceRow()
    {
        return $this->currentRow++;
    }

    /**
     * Advances the colum cursor position and returns the last position
     *
     * @return int
     */
    private function _advanceCol()
    {
        return $this->currentCol++;
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

}
