<?php

declare(strict_types=1);
/**
 * This file is part of ILIAS, a powerful learning management system
 * published by ILIAS open source e-Learning e.V.
 *
 * ILIAS is licensed with the GPL-3.0,
 * see https://www.gnu.org/licenses/gpl-3.0.en.html
 * You should have received a copy of said license along with the
 * source code, too.
 *
 * If this is not the case or you just want to try ILIAS, you'll find
 * us at:
 * https://www.ilias.de
 * https://github.com/ILIAS-eLearning
 *
 *********************************************************************/

namespace ILIAS\Plugin\CBMChoiceQuestionExport\Model;

use ilExcel;

/**
 * Class ExcelData
 * @package ILIAS\Plugin\CBMChoiceQuestionExport\Model
 * @author Marvin Beym <mbeym@databay.de>
 */
class ExcelData
{
    /**
     * @var string
     */
    private $value;
    /**
     * @var int
     */
    private $row;
    /**
     * @var int
     */
    private $col;
    /**
     * @var bool
     */
    private $bold;

    /**
     * @var string
     */
    private $backgroundColor;
    /**
     * @var string
     */
    private $fontColor;

    /**
     * @param string $value
     * @param int $row
     * @param int $col
     * @param bool $bold
     * @param string $backgroundColor
     * @param string $fontColor
     */
    public function __construct(int $row, int $col, string $value, bool $bold = false, string $backgroundColor = "", string $fontColor = "")
    {
        $this->value = $value;
        $this->row = $row;
        $this->col = $col;
        $this->bold = $bold;
        $this->backgroundColor = $backgroundColor;
        $this->fontColor = $fontColor;
    }

    public function process(ilExcel $adapter)
    {
        $row = $this->getRow();
        $col = $this->getCol();
        $adapter->setCell($row, $col, $this->getValue());

        if ($adapter->getCoordByColumnAndRow($col, $row) === "A6") {
            $aa = "";
        }

        if ($this->getBold()) {
            $adapter->setBold($adapter->getCoordByColumnAndRow($col, $row));
        }

        $fontColor = $this->getFontColor();
        $backgroundColor = $this->getBackgroundColor();

        if ($backgroundColor) {
            $adapter->setColors($adapter->getCoordByColumnAndRow($col, $row), $backgroundColor, $fontColor);

        }
    }

    /**
     * @return string
     */
    public function getValue() : string
    {
        return $this->value;
    }

    /**
     * @param string $value
     * @return ExcelData
     */
    public function setValue(string $value) : ExcelData
    {
        $this->value = $value;
        return $this;
    }

    /**
     * @return int
     */
    public function getRow() : int
    {
        return $this->row;
    }

    /**
     * @param int $row
     * @return ExcelData
     */
    public function setRow(int $row) : ExcelData
    {
        $this->row = $row;
        return $this;
    }

    /**
     * @return int
     */
    public function getCol() : int
    {
        return $this->col;
    }

    /**
     * @param int $col
     * @return ExcelData
     */
    public function setCol(int $col) : ExcelData
    {
        $this->col = $col;
        return $this;
    }

    /**
     * @return bool|string
     */
    public function getBold()
    {
        return $this->bold;
    }

    /**
     * @param bool|string $bold
     * @return ExcelData
     */
    public function setBold($bold)
    {
        $this->bold = $bold;
        return $this;
    }

    /**
     * @return string
     */
    public function getBackgroundColor() : string
    {
        return $this->backgroundColor;
    }

    /**
     * @param string $backgroundColor
     * @return ExcelData
     */
    public function setBackgroundColor(string $backgroundColor) : ExcelData
    {
        $this->backgroundColor = $backgroundColor;
        return $this;
    }

    /**
     * @return string
     */
    public function getFontColor() : string
    {
        return $this->fontColor;
    }

    /**
     * @param string $fontColor
     * @return ExcelData
     */
    public function setFontColor(string $fontColor) : ExcelData
    {
        $this->fontColor = $fontColor;
        return $this;
    }
}
