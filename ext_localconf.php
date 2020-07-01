<?php
defined('TYPO3_MODE') || die('Access denied.');

define('QUOTATION_TT_PRODUCTS_EXT', 'quotation_tt_products');

$_EXTCONF = unserialize($_EXTCONF);    // unserializing the configuration so we can use it here:

if (isset($_EXTCONF) && is_array($_EXTCONF)) {
    $GLOBALS['TYPO3_CONF_VARS']['EXTCONF'][QUOTATION_TT_PRODUCTS_EXT] = $_EXTCONF;
}

$GLOBALS['TYPO3_CONF_VARS']['FE']['eID_include']['export_excel'] = 'EXT:' . QUOTATION_TT_PRODUCTS_EXT . '/Resources/Public/Scripts/Php/EidRunner.php';

$GLOBALS['TYPO3_CONF_VARS']['EXTCONF'][TT_PRODUCTS_EXT]['addURLMarkers'][QUOTATION_TT_PRODUCTS_EXT] = \JambageCom\QuotationTtProducts\Hooks\TtProductsMarker::class;

