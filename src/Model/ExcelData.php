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
    private string $value;
    private int $row;
    private int $col;
    private bool $bold;
    private string $backgroundColor;
    private string $fontColor;

    public function __construct(int $row, int $col, string $value, bool $bold = false, string $backgroundColor = "", string $fontColor = "")
    {
        $this->value = $value;
        $this->row = $row;
        $this->col = $col;
        $this->bold = $bold;
        $this->backgroundColor = $backgroundColor;
        $this->fontColor = $fontColor;
    }

    public function process(ilExcel $adapter): void
    {
        $row = $this->getRow();
        $col = $this->getCol();
        $adapter->setCell($row, $col, $this->getValue());

        if ($this->getBold()) {
            $adapter->setBold($adapter->getCoordByColumnAndRow($col, $row));
        }

        $fontColor = $this->getFontColor();
        $backgroundColor = $this->getBackgroundColor();

        if ($backgroundColor) {
            $adapter->setColors($adapter->getCoordByColumnAndRow($col, $row), $backgroundColor, $fontColor);

        }
    }

    public function getValue(): string
    {
        return $this->value;
    }

    public function setValue(string $value): ExcelData
    {
        $this->value = $value;
        return $this;
    }

    public function getRow(): int
    {
        return $this->row;
    }

    public function setRow(int $row): ExcelData
    {
        $this->row = $row;
        return $this;
    }

    public function getCol(): int
    {
        return $this->col;
    }

    public function setCol(int $col): ExcelData
    {
        $this->col = $col;
        return $this;
    }

    public function getBold(): bool
    {
        return $this->bold;
    }

    public function setBold(bool $bold): ExcelData
    {
        $this->bold = $bold;
        return $this;
    }

    public function getBackgroundColor(): string
    {
        return $this->backgroundColor;
    }

    public function setBackgroundColor(string $backgroundColor): ExcelData
    {
        $this->backgroundColor = $backgroundColor;
        return $this;
    }

    public function getFontColor(): string
    {
        return $this->fontColor;
    }

    public function setFontColor(string $fontColor): ExcelData
    {
        $this->fontColor = $fontColor;
        return $this;
    }
}
