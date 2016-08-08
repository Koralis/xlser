<?php

namespace Sakalys\Xlser;

use PHPExcel;
use ReflectionObject;
use ReflectionProperty;

class Xlser extends PHPExcel
{
    const OUTPUT_MEMORY = 1;

    const STYLE_BOLD = 1;

    protected $currentRow = 1;
    protected $currentCol = 0;

    /** @var ReflectionObject */
    protected $refl;

    /** @var [] */
    protected $data;

    /** @var [] */
    protected $headers;

    public function __construct()
    {
        parent::__construct();
        $this->refl = new ReflectionObject($this);
    }


    public function createWriter()
    {
        return new XlsWriter($this);
    }

    public function __set($name, $value)
    {
        $props = $this->refl->getProperties(ReflectionProperty::IS_PUBLIC);

        foreach ($props as $prop) {
            if ($prop->getName() == $name) {
                $this->{$name} = $value;
                return;
            }
        }

        throw new \Exception("This class doesn't accept dynamic properties");
    }

    public function setData(array $data, $headers = [])
    {
        $this->data = $data;
        $this->headers = $headers;

        $sheet = $this->getActiveSheet();

        $this->_resetCursor();

        foreach ($headers as $header) {
            $sheet->setCellValueByColumnAndRow($this->_getCurrentCol(), $this->_getCurrentRow(), $header);
            $this->_advanceCol();
        }

        $this->_advanceRow();

        foreach ($data as $rowData) {
            $this->_resetCol();
            foreach ($rowData as $columnData) {
                $sheet->setCellValueByColumnAndRow($this->_getCurrentCol(), $this->_getCurrentRow(), $columnData);
                $this->_advanceCol();
            }

            $this->_advanceRow();
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
