<?php
defined('TYPO3_MODE') || die('Access denied.');

\JambageCom\Div2007\Utility\FrontendUtility::init();

$eid = \TYPO3\CMS\Core\Utility\GeneralUtility::makeInstance(\JambageCom\QuotationTtProducts\Controller\ExcelController::class);

$eid->run();

