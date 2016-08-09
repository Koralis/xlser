<?php

namespace Sakalys\Xlser;

class CellFormatTest extends \PHPUnit_Framework_TestCase
{
    public function test_if_sets_to_bold()
    {
        $instance= new CellFormat();

        $instance->setBold(true);
        self::assertTrue($instance->isBold());

        $instance->setBold(false);
        self::assertFalse($instance->isBold());
    }

    public function test_if_it_gets_bold_from_scalar()
    {
        $instance = CellFormat::createFromScalar(Xlser::STYLE_BOLD);

        self::assertTrue($instance->isBold());
    }

    public function test_if_applies_to_cells()
    {
        $man = new Xlser();
        $sheet = $man->getActiveSheet();

        $cell = $sheet->setCellValue('A1', 123, true);

        $cellFormat = new CellFormat();
        $cellFormat->apply($cell);
    }
}
