<?php

namespace JambageCom\QuotationTtProducts\Hooks;


/***************************************************************
*  Copyright notice
*
*  (c) 2018 Franz Holzinger <franz@ttproducts.de>
*  All rights reserved
*
*  This script is part of the Typo3 project. The Typo3 project is
*  free software; you can redistribute it and/or modify
*  it under the terms of the GNU General Public License as published by
*  the Free Software Foundation; either version 2 of the License, or
*  (at your option) any later version.
*
*  The GNU General Public License can be found at
*  http://www.gnu.org/copyleft/gpl.html.
*  A copy is found in the textfile GPL.txt and important notices to the license
*  from the author is found in LICENSE.txt distributed with these scripts.
*
*
*  This script is distributed in the hope that it will be useful,
*  but WITHOUT ANY WARRANTY; w+ithout even the implied warranty of
*  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*  GNU General Public License for more details.
*
*  This copyright notice MUST APPEAR in all copies of the script!
***************************************************************/
/**
* Part of the quotation_tt_products extension.
*
* class with functions to control all activities
*
* @author  Franz Holzinger <franz@ttproducts.de>
* @maintainer	Franz Holzinger <franz@ttproducts.de>
* @package TYPO3
* @subpackage quotation_tt_products
*
*/

use TYPO3\CMS\Core\Utility\GeneralUtility;

use JambageCom\Div2007\Utility\FrontendUtility;

// connect with this HTML:
//         <input type="button" id="angebot_export" name="ex" value="Export als Angebot (XLS)">


class TtProductsMarker implements \TYPO3\CMS\Core\SingletonInterface {


    public function addURLMarkers (
        $pObj,
        $cObj,
        $pidNext,
        &$markerArray,
        $addQueryString,
        $excludeList,
        $singleExcludeList,
        $useBackPid = true,
        $backPid = 0,
        $target = '',
        $excludeSingleVar = true
    ) {
        $charset = 'UTF-8';
        $cnfObj = GeneralUtility::makeInstance('tx_ttproducts_config');
        $conf = $cnfObj->getConf();
        $basketPid = ($conf['PIDbasket'] ? $conf['PIDbasket'] : $GLOBALS['TSFE']->id);
        $addQueryString = array('eID' => 'export_excel');

        $url = FrontendUtility::getTypoLink_URL(
            $cObj,
            $basketPid,
            $pObj->getLinkParams(
                $singleExcludeList,
                $addQueryString,
                false,
                false,
                0
            ),
            $target,
            $conf
        );
    debug ($url, 'hook $url');
        $markerArray['###FORM_URL_QUOTATION_JSVALUE###'] = $jsUrl = GeneralUtility::quoteJSvalue($url);

        $code = '<script>
    function addListeners(){
        document.getElementById(\'angebot_export\').addEventListener(\'click\', anexport);
        function anexport() {
            this.form.action = \'https://' . $conf['domain'] . '/\' + ' . $jsUrl . ';
            this.form.submit();
        }
    }
    window.addEventListener(\'load\', addListeners); 

    </script>
';
        $JSfieldname = 'tx_quotation_tt_products-excel';
        $GLOBALS['TSFE']->additionalHeaderData[$JSfieldname] = $code;
    }
}


