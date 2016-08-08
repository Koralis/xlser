<?php

namespace Sakalys\Xlser;

use PHPExcel;
use ReflectionObject;
use ReflectionProperty;

class Xlser extends PHPExcel
{
    const OUTPUT_MEMORY = 1;

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

    public function setData($data, $headers = [])
    {
        $this->data = $data;
        $this->headers = $headers;

        $row = 1;
        $col = 0;
        foreach ($headers as $header) {
            $this->getActiveSheet()->setCellValueByColumnAndRow($col++, $row, $header);
        }


        foreach ($data as $rowData) {
            $row++;
            $col = 0;
            foreach ($rowData as $columnData) {
                $this->getActiveSheet()->setCellValueByColumnAndRow($col++, $row, $columnData);
            }
        }
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

}
