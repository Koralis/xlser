<?php

namespace Sakalys\Xlser;

class XlserTest extends \PHPUnit_Framework_TestCase
{
    /** @var Xlser */
    protected $man;

    public function setUp()
    {
        $this->man = new Xlser;
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

    public function test_if_it_does_not_accept_dynamic_properties()
    {
        $caught = false;
        try {
            $this->man->testprop_that_does_not_exist = null;
            echo 1234;
        } catch (\Exception $e) {
            $caught = true;
        }

        self::assertTrue($caught, "It should not be allowed to set non-existing properties");
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
        $this->man->appendRow(['header1', 'header2']);


    }

}
