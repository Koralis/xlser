<?php

namespace Koralis\Xlser;

use PHPExcel_Cell;
use PHPExcel_Worksheet;

class XlserTest extends \PHPUnit_Framework_TestCase
{
    /** @var Xlser */
    protected $man;

    /** @var PHPExcel_Worksheet */
    protected $sheet;

    /** @var \PHPUnit_Framework_MockObject_MockBuilder */
    protected $mock;

    public function setUp()
    {
        $this->mock = $this->getMockBuilder(Xlser::class);
        $this->man = new Xlser;
        $this->sheet = $this->man->getActiveSheet();
    }

    public function test_if_it_is_instantiable()
    {
        self::assertTrue(!!$this->man);
    }

    public function test_if_it_accepts_data_and_headers()
    {
        $this->setData();
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

        $this->assertCellEquals('header1', 'A1');
        $this->assertCellEquals('header2', 'B1');
        $this->assertCellEquals('row2 col1', 'A2');
        $this->assertCellEquals('row2 col2', 'B2');
        $this->assertCellEquals('row3 col1', 'A3');
        $this->assertCellEquals('row3 col2', 'B3');
    }

    public function test_if_it_appends_a_header_row()
    {
        $this->man->appendHeaderRow(['header1', 'header2']);
        $this->assertCellEquals('header1', 'A1');
        $this->assertCellEquals('header2', 'B1');

        $this->assertCellBold('A1', "The header row must be bold by default");
    }

    public function test_if_it_appends_a_row()
    {
        $this->man->appendRow(['data1', 'data2']);
        $this->assertCellEquals('data1', 'A1');
        $this->assertCellEquals('data2', 'B1');

        $boldFormat = new CellFormat();
        $boldFormat->setBold(true);

        $this->man->appendRow([
            [10, $boldFormat],
        ]);

        $this->assertCellEquals(10, 'A2');
        $this->assertCellBold('A2');
    }

    public function test_if_it_appends_a_row_with_callbacks()
    {
        /** @noinspection PhpUnusedParameterInspection */
        $this->man->appendRow([
            function (PHPExcel_Cell $c) { // The cell must be passed
                return 'test';
            }
        ]);

        $this->assertCellEquals('test');
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

    public function test_if_auto_column_width_can_be_set()
    {
        $this->man->autoColWidth(true);
    }

    public function test_if_it_is_possible_to_add_a_formatted_cell()
    {
        $cellFormat = new CellFormat();
        $cellFormat->setBold(true);
        $this->man->insertVal(123, $cellFormat);

        $assertMsg = 'It should accept a ' . CellFormat::class . ' object';
        $this->assertCellBold('A1', $assertMsg);
    }

    public function test_if_row_format_gets_applied()
    {
        $rowFormat = new CellFormat();
        $rowFormat->setBold(true);

        $this->man->insertVal(123, null, $rowFormat);

        $this->assertCellBold('A1');
    }

    /**
     * @param string $coordinate
     * @param string|null $message
     * @throws \PHPExcel_Exception
     */
    private function assertCellBold($coordinate, $message = null)
    {
        if (!$message) {
            $message = 'Cell at ' . $coordinate . ' must be bold';
        }

        $isBold = $this->sheet->getCell($coordinate)->getStyle()->getFont()->getBold();
        self::assertEquals(true, $isBold, $message);
    }

    /**
     * @param $value
     * @param $coordinate
     * @throws \PHPExcel_Exception
     */
    private function assertCellEquals($value, $coordinate = 'A1')
    {
        self::assertEquals($value, $this->sheet->getCell($coordinate)->getValue());
    }
}
