<?php

namespace Koralis\Xlser;

use PHPExcel_Cell;

class CellFormat
{
    protected $_bold = false;

    protected $_textAlign = null;

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
            if ($string & Xlser::STYLE_BOLD) {
                $format->setBold(true);
            }

            if ($string & Xlser::ALIGN_RIGHT) {
                $format->setTextAlign('right');
            } elseif ($string & Xlser::ALIGN_CENTER) {
                $format->setTextAlign('center');
            } elseif ($string & Xlser::ALIGN_LEFT) {
                $format->setTextAlign('left');
            } else {
                $format->setTextAlign(null);
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

        $horizontalAlignment = $this->getTextAlign();

        $align = $cell->getStyle()->getAlignment();

        if (is_null($horizontalAlignment)) {
            $align->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_GENERAL);
        } elseif ($horizontalAlignment === -1) {
            $align->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        } elseif ($horizontalAlignment === 1) {
            $align->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        } elseif ($horizontalAlignment === 0) {
            $align->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        }
    }

    public function setTextAlign($align)
    {
        if (is_string($align)) {
            switch ($align) {
                case 'right':
                    $align = 1;
                    break;

                case 'center':
                    $align = 0;
                    break;

                case 'left':
                default:
                    // In case the string is unrecognised we set the alignment to 'left';
                    $align = -1;
            }
        }

        if (!in_array($align, [-1, 0, 1, null], true)) {
            throw new \InvalidArgumentException('Invalid alignment value');
        }

        $this->_textAlign = $align;
    }

    public function getTextAlign()
    {
        return $this->_textAlign;
    }

    public function getTextAlignText()
    {
        switch ($this->_textAlign) {
            case -1:
                return 'left';
            case 0:
                return 'center';
            case 1:
                return 'right';
            default:
                return null;
        }
    }
}
