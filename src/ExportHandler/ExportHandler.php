<?php declare(strict_types=1);
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

namespace ILIAS\Plugin\CBMChoiceQuestionExport\ExportHandler;

use assQuestion;
use ilAssExcelFormatHelper;
use ilDBInterface;
use ILIAS\DI\Container;
use ilLanguage;
use ilObjTest;
use ilTestEvaluationData;
use ilTestExportFilename;
use ilUtil;
use PhpOffice\PhpSpreadsheet\Exception;

/**
 * Class ExportHandler
 * @package ILIAS\Plugin\CBMChoiceQuestionExport\ExportHandler
 * @author Marvin Beym <mbeym@databay.de>
 */
class ExportHandler
{
    /**
     * @var ilObjTest
     */
    private $test;
    /**
     * @var ilTestExportFilename
     */
    private $filename;
    /**
     * @var Container
     */
    private $dic;
    /**
     * @var ilLanguage
     */
    private $lng;
    /**
     * @var ilDBInterface
     */
    private $db;
    /**
     * @var ilTestEvaluationData
     */
    private $data;

    public function __construct(ilObjTest $test, ilTestExportFilename $filename)
    {
        $this->test = $test;
        $this->filename = $filename;
        global $DIC;
        $this->dic = $DIC;
        $this->lng = $this->dic->language();
        $this->db = $this->dic->database();
        $this->data = $this->test->getCompleteEvaluationData(true);


    }

    /**
     * @throws Exception
     */
    public function export() : void
    {

    }
}