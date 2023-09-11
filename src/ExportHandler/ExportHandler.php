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

namespace ILIAS\Plugin\CBMChoiceQuestionExport\ExportHandler;

use assQuestion;
use CBMChoiceQuestion;
use ilAssExcelFormatHelper;
use ilCBMChoiceQuestionExportPlugin;
use ilDBInterface;
use ILIAS\DI\Container;
use ILIAS\Plugin\CBMChoiceQuestionExport\Model\ExcelData;
use ilLanguage;
use ilObjTest;
use ilPlugin;
use ilTestEvaluationData;
use ilTestEvaluationUserData;
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
    /**
     * @var ilCBMChoiceQuestionExportPlugin
     */
    private $plugin;
    /**
     * @var ilPlugin
     */
    private $cbmChoiceQuestionPlugin;

    public function __construct(ilCBMChoiceQuestionExportPlugin $plugin, ilPlugin $cbmChoiceQuestionPlugin, ilObjTest $test, ilTestExportFilename $filename)
    {
        $this->plugin = $plugin;
        $this->test = $test;
        $this->filename = $filename;
        global $DIC;
        $this->dic = $DIC;
        $this->lng = $this->dic->language();
        $this->db = $this->dic->database();
        $this->data = $this->test->getCompleteEvaluationData();
        $this->cbmChoiceQuestionPlugin = $cbmChoiceQuestionPlugin;
    }

    /**
     * @throws Exception
     */
    public function export(): void
    {
        $excelTmpFile = ilUtil::ilTempnam() . '.xlsx';

        $adapter = new ilAssExcelFormatHelper();

        /**
         * @var CBMChoiceQuestion[] $cbmQuestions
         */
        $cbmQuestions = [];
        foreach ($this->test->getQuestions() as $id) {
            $id = (int) $id;
            if (assQuestion::_getQuestionType($id) === "CBMChoiceQuestion") {
                $cbmQuestions[] = assQuestion::_instantiateQuestion($id);
            }
        }

        foreach ($this->buildOverviewResultsExcelData($cbmQuestions, $adapter, $this->data->getParticipants()) as $data) {
            $data->process($adapter);
        }

        foreach ($this->data->getParticipants() as $activeId => $userData) {
            foreach ($this->buildUserSpecificResultsExcelData($cbmQuestions, $adapter, $activeId, $userData) as $data) {
                $data->process($adapter);
            }
        }

        $adapter->writeToFile($excelTmpFile);
        //$this->deliverFile($excelTmpFile, $this->test->getTitle());

        ilUtil::makeDirParents(dirname($this->filename->getPathname('xlsx', 'cbmChoiceExport')));
        rename($excelTmpFile, $this->filename->getPathname('xlsx', 'cbmChoiceExport'));
    }

    /**
     * @param CBMChoiceQuestion[] $cbmQuestions
     * @param ilAssExcelFormatHelper $adapter
     * @param ilTestEvaluationUserData[] $participants
     * @return ExcelData[]
     */
    protected function buildOverviewResultsExcelData(array $cbmQuestions, ilAssExcelFormatHelper $adapter, array $participants): array
    {
        $row = 1;
        $worksheet_index = $adapter->addSheet($this->lng->txt("overview"));
        $adapter->setActiveSheet($worksheet_index);
        $excelData = [];

        $row += 2;

        $excelData[] = new ExcelData($row, 0, $this->lng->txt("question_title"));
        $excelData[] = new ExcelData($row, 1, $this->plugin->txt("export.averageCertainty"));
        $excelData[] = new ExcelData($row, 2, $this->plugin->txt("export.averageCorrectAnswers"));
        $row++;

        foreach ($cbmQuestions as $cbmQuestion) {
            $excelData[] = new ExcelData($row, 0, $cbmQuestion->getTitle());
            $certainCount = 0;
            $correctUserAnswersCount = 0;
            $correctAnswersCount = $cbmQuestion->getCorrectAnswerCount();

            if ($participants === []) {
                $excelData[] = new ExcelData($row, 1, "0%");
                $excelData[] = new ExcelData($row, 2, "0%");
                continue;
            }

            foreach ($participants as $activeId => $userData) {
                $solution = $cbmQuestion->mapSolution($cbmQuestion->getSolutionValues($activeId, $userData->getScoredPass()));
                $correctUserAnswersCount += $solution->getCorrectAnswerCount();
                $certainCount += (int) $solution->isCertain();
            }

            if ($correctAnswersCount > 0 && $correctUserAnswersCount > 0) {
                $correctAnswersAverage = $correctUserAnswersCount / ($correctAnswersCount * count($participants));
            } else {
                $correctAnswersAverage = 0;
            }

            if ($certainCount > 0) {
                $certainAverage = $certainCount / count($participants);
            } else {
                $certainAverage = 0;
            }

            $excelData[] = new ExcelData($row, 1, (round($certainAverage * 100, 2)) . "%");
            $excelData[] = new ExcelData($row, 2, (round($correctAnswersAverage * 100, 2)) . "%");
            $row++;
        }
        return $excelData;
    }

    /**
     * @param CBMChoiceQuestion[] $cbmQuestions
     * @param ilAssExcelFormatHelper $adapter
     * @param int $activeId
     * @param ilTestEvaluationUserData $userData
     * @return ExcelData[]
     */
    protected function buildUserSpecificResultsExcelData(array $cbmQuestions, ilAssExcelFormatHelper $adapter, int $activeId, ilTestEvaluationUserData $userData): array
    {
        $row = 1;
        $worksheet_index = $adapter->addSheet($userData->getName());
        $adapter->setActiveSheet($worksheet_index);
        $excelData = [];

        $excelData[] = new ExcelData(
            $row,
            0,
            sprintf(
                $this->lng->txt("tst_result_user_name_pass"),
                $userData->getScoredPass(),
                $userData->getName()
            )
        );

        $row += 2;

        $totalCertainCount = 0;
        $totalCorrectAnswerCount = 0;

        foreach ($cbmQuestions as $cbmQuestion) {
            $solution = $cbmQuestion->mapSolution($cbmQuestion->getSolutionValues($activeId, $userData->getScoredPass()));

            $correctAnswersCount = $cbmQuestion->getCorrectAnswerCount();
            $correctUserAnswersCount = $solution->getCorrectAnswerCount();

            if ($correctAnswersCount > 0 && $correctUserAnswersCount > 0) {
                $totalCorrectAnswerCount += $correctUserAnswersCount / $correctAnswersCount;
            }
            $totalCertainCount += (int) $solution->isCertain();
        }

        if ($totalCorrectAnswerCount > 0) {
            $correctAnswersAverage = $totalCorrectAnswerCount / count($cbmQuestions);
        } else {
            $correctAnswersAverage = 0;
        }

        if ($totalCertainCount > 0) {
            $certainAverage = $totalCertainCount / count($cbmQuestions);
        } else {
            $certainAverage = 0;
        }

        $excelData[] = new ExcelData($row, 0, $this->plugin->txt("export.averageCertainty"));
        $excelData[] = new ExcelData($row, 0, $this->plugin->txt("export.averageCorrectAnswers"));
        $excelData[] = new ExcelData($row++, 1, (round($certainAverage * 100)) . "%");
        $excelData[] = new ExcelData($row, 1, (round($correctAnswersAverage * 100)) . "%");

        $row += 2;

        foreach ($cbmQuestions as $cbmQuestion) {
            $solution = $cbmQuestion->mapSolution($cbmQuestion->getSolutionValues($activeId, $userData->getScoredPass()));
            $excelData[] = new ExcelData($row, 0, "CBMChoiceQuestion", true, EXCEL_BACKGROUND_COLOR);
            $excelData[] = new ExcelData($row, 1, $cbmQuestion->getTitle(), true, EXCEL_BACKGROUND_COLOR);
            $row += 2;
            $excelData[] = new ExcelData($row, 0, $this->lng->txt("answers"), true);
            $excelData[] = new ExcelData($row, 1, $this->lng->txt("checked"), true);
            $excelData[] = new ExcelData($row++, 2, $this->lng->txt("correct_answers"), true);

            foreach ($cbmQuestion->getAnswers() as $answer) {
                $excelData[] = new ExcelData($row, 0, $answer->getAnswerText());

                $answerAnswered = false;
                foreach ($solution->getAnswers() as $userAnswer) {
                    if ($userAnswer->getId() === $answer->getId()) {
                        $answerAnswered = true;
                    }
                }
                $excelData[] = new ExcelData($row, 1, $answerAnswered ? "X" : "");
                $excelData[] = new ExcelData($row, 2, $answer->isAnswerCorrect() ? "X" : "");
                $row++;
            }
            $row += 2;
            $excelData[] = new ExcelData($row, 0, "CBM", true);
            $excelData[] = new ExcelData($row++, 1, $this->lng->txt("checked"), true);

            foreach ($cbmQuestion->getScoringMatrix()["correct"] as $rowIndex => $data) {
                $excelData[] = new ExcelData($row, 0, $this->cbmChoiceQuestionPlugin->txt("question.cbm.$rowIndex"));
                $excelData[] = new ExcelData($row++, 1, $solution->getCbmChoice() === $rowIndex ? "X" : "");
            }
            $row++;
        }
        return $excelData;
    }


    /**
     * @param assQuestion[] $cbmQuestions
     * @return string[]
     */
    protected function getHeaders(array $cbmQuestions): array
    {
        $headers = [
            $this->lng->txt("name"),
            $this->lng->txt("login"),
        ];

        foreach ($cbmQuestions as $cbmQuestion) {
            $headers[] = "[{$cbmQuestion->getId()}] " . assQuestion::_getQuestionTitle($cbmQuestion->getId());
            $headers[] = "";
        }

        return $headers;
    }

    protected function deliverFile(string $filePath, string $title): void
    {
        $fileName = ilUtil::getASCIIFilename(preg_replace("/\s/", "_", "cbm_$title")) . ".xlsx";
        ilUtil::deliverFile($filePath, $fileName, "application/vnd.ms-excel", false, true);
        exit;
    }
}
