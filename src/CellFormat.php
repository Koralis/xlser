<?php

namespace Sakalys\Xlser;

use PHPExcel_Cell;

class CellFormat
{
    protected $_bold = false;

    /**
     * @var string
     */
    private $formatSymbol;
    /**
     * @var int
     */
    private $formatPrecision;
    /**
     * @var bool
     */
    private $formatSymbolAfterValue;
    /**
     * @var string
     */
    private $formatThousandSeparator;
    /**
     * @var string
     */
    private $formatDecimalSeparator;

    /**
     * @param string $symbol
     * @param int $precision
     * @param bool $symbolAfterValue
     * @param string $thousandSeparator
     * @param string $decimalSeparator
     * @return CellFormat
     */
/*    public static function currency($symbol, $precision = 2, $symbolAfterValue = false, $thousandSeparator = ',', $decimalSeparator = '.')
    {
        // This line here
//        $o = new self;
//
////        $numberFormat = '"$"#,##0.00_-';
//        $format = '_-';
//
//        $number = '#,##0';
//
//        if ($precision > 0) {
//            $number .= $format . str_repeat('0', $precision);
//        }
//
//
//
//        $this->formatSymbol = $symbol;
//        $this->formatPrecision = $precision;
//        $this->formatSymbolAfterValue = $symbolAfterValue;
//        $this->formatThousandSeparator = $thousandSeparator;
//        $this->formatDecimalSeparator = $decimalSeparator;
//


        return $o;
    }*/

    /**
     * @param string $string
     * @return CellFormat
     */
    public static function createFromScalar($string)
    {
        $format = new self;

        if (is_integer($string)) {
            switch (true) {
                case $string & Xlser::STYLE_BOLD:
                    $format->setBold(true);
                    break;
                default:
                    break;
            }
        }

        return $format;
    }

    public function setBold($bool)
    {
        $this->_bold = $bool;

        return $this;
    }

    /**
     * @return boolean
     */
    public function isBold()
    {
        return $this->_bold;
    }

    public function apply(PHPExcel_Cell $cell)
    {
        switch (true) {
            case $this->isBold():
                $cell->getStyle()->getFont()->setBold(true);

            default:
                break;
        }
    }
}
