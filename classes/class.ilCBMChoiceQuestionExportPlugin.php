<?php

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

require_once __DIR__ . '/../vendor/autoload.php';

use ILIAS\DI\Container;
use ILIAS\Plugin\CBMChoiceQuestionExport\ExportHandler\ExportHandler;
use ILIAS\Plugin\CBMChoiceQuestionExport\Utils\UiUtil;
use PhpOffice\PhpSpreadsheet\Exception;

class ilCBMChoiceQuestionExportPlugin extends ilTestExportPlugin
{
    /** @var string */
    public const CTYPE = 'Modules';

    /** @var string */
    public const CNAME = 'Test';

    /** @var string */
    public const SLOT_ID = 'texp';

    /** @var string */
    public const PNAME = 'CBMChoiceQuestionExport';

    private static ?self $instance = null;
    public Container $dic;
    public ilSetting $settings;
    private ilLanguage $lng;
    private UiUtil $uiUtil;

    public function __construct(ilDBInterface $db, ilComponentRepositoryWrite $component_repository, string $id)
    {
        global $DIC;
        $this->dic = $DIC;
        $this->lng = $this->dic->language();
        $this->settings = new ilSetting(self::class);
        $this->uiUtil = new UiUtil($this->dic);
        parent::__construct($db, $component_repository, $id);
    }

    public function getPluginName(): string
    {
        return self::PNAME;
    }

    public static function getInstance(): self
    {
        if (self::$instance) {
            return self::$instance;
        }

        global $DIC;

        /**
         * @var ilComponentFactory $componentFactory
         */
        $componentFactory = $DIC["component.factory"];
        self::$instance = $componentFactory->getPlugin("cbmChoiceExport");
        return self::$instance;
    }

    public function redirectToHome(): void
    {
        $this->dic->ctrl()->redirectByClass("ilDashboardGUI", "show");
    }

    public function isUserAdmin(?int $userId, ?int $roleId): bool
    {
        if ($userId === null) {
            $userId = $this->dic->user->getId();
        }

        if ($roleId === null) {
            if (defined("SYSTEM_ROLE_ID")) {
                $roleId = (int) SYSTEM_ROLE_ID;
            } else {
                $roleId = 2;
            }
        }

        $roleIds = [];

        foreach ($this->dic->rbac()->review()->assignedGlobalRoles($userId) as $id) {
            $roleIds[] = (int) $id;
        }

        return in_array($roleId, $roleIds, true);
    }

    public function denyConfigIfPluginNotActive(): void
    {
        if (!$this->isActive()) {
            $this->uiUtil->sendFailure($this->txt("plugin_not_activated"), true);
            $this->dic->ctrl()->redirectByClass(ilObjComponentSettingsGUI::class, "view");
        }
    }

    protected function beforeUninstall(): bool
    {
        $this->settings->deleteAll();
        return parent::beforeUninstall();
    }

    public function assetsFolder(string $file = ''): string
    {
        return $this->getDirectory() . "/assets/$file";
    }

    public function cssFolder(string $file = ''): string
    {
        return $this->assetsFolder("css/$file");
    }

    public function jsFolder(string $file = ''): string
    {
        return $this->assetsFolder("js/$file");
    }

    public function imagesFolder(string $file = ''): string
    {
        return $this->assetsFolder("images/$file");
    }

    public function templatesFolder(string $file = ''): string
    {
        return $this->assetsFolder("templates/$file");
    }

    /**
     * @throws Exception|ilException
     */
    protected function buildExportFile(ilTestExportFilename $export_path): void
    {
        if (!$this->isCbmChoiceQuestionPluginActive()) {
            $this->uiUtil->sendFailure($this->txt("export.cbmPlugin.notActive"), true);
            return;
        }

        $cbmChoiceQuestionPlugin = $this->getCbmChoiceQuestionPluginObject();

        if (!$cbmChoiceQuestionPlugin) {
            $this->uiUtil->sendFailure($this->txt("export.cbmPlugin.notActive"), true);
            return;
        }

        $exportHandler = new ExportHandler($this, $cbmChoiceQuestionPlugin, $this->getTest(), $export_path);
        $exportHandler->export();
    }

    public function getCbmChoiceQuestionPluginObject(): ?ilPlugin
    {
        try {
            return ilCBMChoiceQuestionPlugin::getInstance();
        } catch (Throwable $ex) {
            return null;
        }
    }

    public function isCbmChoiceQuestionPluginActive(): bool
    {
        try {
            $cbmChoiceQuestionPlugin = $this->getCbmChoiceQuestionPluginObject();
            return $cbmChoiceQuestionPlugin && $cbmChoiceQuestionPlugin->isActive();
        } catch (Throwable $ex) {
            return false;
        }
    }

    protected function beforeActivation(): bool
    {
        if (!$this->isCbmChoiceQuestionPluginActive()) {
            $this->uiUtil->sendFailure($this->txt("export.cbmPlugin.notActive"), true);
            return false;
        }

        return parent::beforeActivation(); // TODO: Change the autogenerated stub
    }

    protected function getFormatIdentifier(): string
    {
        return "test";
    }

    public function getFormatLabel(): string
    {
        return $this->txt("export.format");
    }
}
