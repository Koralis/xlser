<?php

namespace Sakalys\Xlser;

use PHPExcel_Worksheet;

class XlserTest extends \PHPUnit_Framework_TestCase
{
    /** @var Xlser */
    protected $man;

    /** @var PHPExcel_Worksheet */
    protected $sheet;

    public function setUp()
    {
        $this->man = new Xlser;
        $this->sheet = $this->man->getActiveSheet();
    }

    public function test_if_it_is_instantiable()
    {
        self::assertTrue(!!$this->man);
    }

    public function test_if_it_can_create_a_writer()
    {
        $writer = $this->createWriter();

        self::assertTrue(!!$writer);
    }

    public function test_if_it_accepts_data_and_headers()
    {
        $this->setData();
    }

    private function createWriter()
    {
        return $this->man->createWriter();
    }

    public function test_it_can_output_tmp_contents_as_memory()
    {
        $this->man->output(Xlser::OUTPUT_MEMORY);
    }

    private function setData()
    {
        $headers = ['header1', 'header2'];

        $data = [
            ['row2 col1', 'row2 col2'],
            ['row3 col1', 'row3 col2'],
        ];

        $this->man->setData($data, $headers);
    }

    public function test_it_writes_the_data_to_phpexcel()
    {
        $this->setData();

        $sheet = $this->man->getActiveSheet();

        self::assertEquals('header1', $sheet->getCell('A1')->getValue());
        self::assertEquals('header2', $sheet->getCell('B1')->getValue());
        self::assertEquals('row2 col1', $sheet->getCell('A2')->getValue());
        self::assertEquals('row2 col2', $sheet->getCell('B2')->getValue());
        self::assertEquals('row3 col1', $sheet->getCell('A3')->getValue());
        self::assertEquals('row3 col2', $sheet->getCell('B3')->getValue());
    }

    public function test_if_it_appends_a_row()
    {
        $this->man->appendHeaderRow(['header1', 'header2']);
        self::assertEquals('header1', $this->sheet->getCell('A1')->getValue());
        self::assertEquals('header2', $this->sheet->getCell('B1')->getValue());

        $this->man->appendRow(['data1', 'data2']);
        self::assertEquals('data1', $this->sheet->getCell('A2')->getValue());
        self::assertEquals('data2', $this->sheet->getCell('B2')->getValue());

        $val = 10;
        $this->man->appendRow([
            0,
            [0, function (\PHPExcel_Cell $c) use ($val) {
                $c->getStyle()->getFont()->setBold(true);
                $c->setValue($val);
            }]
        ]);
        self::assertEquals('0', $this->sheet->getCell('A3')->getValue());
        self::assertEquals('10', $this->sheet->getCell('B3')->getValue());
        self::assertEquals(true, $this->sheet->getCell('B3')->getStyle()->getFont()->getBold());

        $this->man->appendRow([
            function (\PHPExcel_Cell $c) use ($val) {
                $c->setValue('test');
            }
        ]);

        self::assertEquals('test', $this->sheet->getCell('A4')->getValue());
    }

    public function test_if_it_gets_the_row_and_col_props()
    {
        $this->man->getRow();
        $this->man->getCol();
    }

    public function test_if_it_moves_to_another_row()
    {
        $row = $this->man->getRow();

        $this->man
            ->newRow()
            ->newRow();

        $newRow = $this->man->getRow();

        self::assertEquals($row + 2, $newRow);
    }

    public function test_if_it_moves_to_another_col()
    {
        $col = $this->man->getCol();

        $this->man
            ->newCol()
            ->newCol();

        $newCol = $this->man->getCol();

        self::assertEquals($col + 2, $newCol);
    }
}
