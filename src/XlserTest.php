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
        $data = [
            ['1 1st', '1 2nd'],
            ['2 1st', '2 2nd'],
        ];
        $headers = ['1st col', '2nd col'];

        $this->man->setData($data, $headers);
    }

    public function test_it_writes_the_data_to_phpexcel()
    {
        $this->setData();

        $sheet = $this->man->getActiveSheet();

        self::assertEquals('1st col', $sheet->getCell('A1')->getValue());
        self::assertEquals('2nd col', $sheet->getCell('B1')->getValue());
        self::assertEquals('1 1st', $sheet->getCell('A2')->getValue());
        self::assertEquals('1 2nd', $sheet->getCell('B2')->getValue());
        self::assertEquals('2 1st', $sheet->getCell('A3')->getValue());
        self::assertEquals('2 2nd', $sheet->getCell('B3')->getValue());
    }
}
