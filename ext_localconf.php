<?php
defined('TYPO3_MODE') || die('Access denied.');

define('QUOTATION_TT_PRODUCTS_EXT', 'quotation_tt_products');

call_user_func(function () {
    $extensionConfiguration = array();

    if (
        defined('TYPO3_version') &&
        version_compare(TYPO3_version, '9.0.0', '>=')
    ) {
        $extensionConfiguration = \TYPO3\CMS\Core\Utility\GeneralUtility::makeInstance(
            \TYPO3\CMS\Core\Configuration\ExtensionConfiguration::class
        )->get(QUOTATION_TT_PRODUCTS_EXT);
    } else { // before TYPO3 9
        $extensionConfiguration = unserialize($GLOBALS['TYPO3_CONF_VARS']['EXT']['extConf'][QUOTATION_TT_PRODUCTS_EXT]);
    }

    if (isset($extensionConfiguration) && is_array($extensionConfiguration)) {
        $GLOBALS['TYPO3_CONF_VARS']['EXTCONF'][QUOTATION_TT_PRODUCTS_EXT] = $extensionConfiguration;
    }

    $GLOBALS['TYPO3_CONF_VARS']['FE']['eID_include']['export_excel'] = 'EXT:' . QUOTATION_TT_PRODUCTS_EXT . '/Resources/Public/Scripts/Php/EidRunner.php';

    if (defined('TT_PRODUCTS_EXT')) {
        $GLOBALS['TYPO3_CONF_VARS']['EXTCONF'][TT_PRODUCTS_EXT]['addURLMarkers'][QUOTATION_TT_PRODUCTS_EXT] = \JambageCom\QuotationTtProducts\Hooks\TtProductsMarker::class;
    }
});


