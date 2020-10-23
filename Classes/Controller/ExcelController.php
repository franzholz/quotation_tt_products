<?php

namespace JambageCom\QuotationTtProducts\Controller;

/***************************************************************
*  Copyright notice
*
*  (c) 2018 Franz Holzinger (franz@ttproducts.de)
*  All rights reserved
*
*  This script is part of the TYPO3 project. The TYPO3 project is
*  free software; you can redistribute it and/or modify
*  it under the terms of the GNU General Public License as published by
*  the Free Software Foundation; either version 2 of the License or
*  (at your option) any later version.
*
*  The GNU General Public License can be found at
*  http://www.gnu.org/copyleft/gpl.html.
*  A copy is found in the textfile GPL.txt and important notices to the license
*  from the author is found in LICENSE.txt distributed with these scripts.
*
*
*  This script is distributed in the hope that it will be useful,
*  but WITHOUT ANY WARRANTY; without even the implied warranty of
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
*
*/

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

use TYPO3\CMS\Core\Utility\GeneralUtility;


class ExcelController implements \TYPO3\CMS\Core\SingletonInterface {

    public function run () {

        $load = $_POST;
        $cObj = \JambageCom\Div2007\Utility\FrontendUtility::getContentObjectRenderer();
        $conf = $GLOBALS['TSFE']->tmpl->setup['plugin.'][TT_PRODUCTS_EXT . '.'];
        $config = $GLOBALS['TYPO3_CONF_VARS']['EXTCONF'][QUOTATION_TT_PRODUCTS_EXT];
        $basketApi = GeneralUtility::makeInstance(\JambageCom\TtProducts\Api\BasketApi::class);

        $itemArray = $basketApi->readItemArray();
        $calculatedArray = $basketApi->readCalculatedArray();
        $basketExtra = $basketApi->readBasketExtra();
        if (
            empty($itemArray) ||
            empty($calculatedArray) ||
            empty($basketExtra)
        ) {
            echo('Die Daten sind nicht mehr verfügbar. Gehen Sie im Browser Fenster zurück und laden Sie die Seite neu.');
            return false;
        }

        $variants = $conf['table.']['tt_products.']['variant.'];
        $variants = array_diff($variants, array('additional'));
        $taxrate = $conf['TAXpercentage'];

        /** Spreadsheet */

        ############ Dateinamen erstellen ################
        $file = PATH_site . $config['savePath'] . 'count.txt';

        if (!file_exists($file) ) {
            throw new \RuntimeException('File "' . $file . '" not found.');
        }
    
        $fp = fopen($file, 'r');
        if (!$fp) {
            throw new \RuntimeException('File open for read failed: "' . $file . '"');
        }  
        $zahl = intval(fread($fp, 100));
        $zahl++;
        fclose($fp);

        $fp = fopen($file, 'w');
        if (!$fp) {
            throw new \RuntimeException('File open for write failed: "' . $file . '"');
        }  
        fwrite($fp, $zahl);
        fclose($fp);

        $filename = $config['savePath'] . $config['filenamePrefix'] . $zahl . '.xls';
        ###################################################

        // Create new Spreadsheet object
        $objSpreadsheet = new Spreadsheet();

        // Set properties
        $objSpreadsheet->getProperties()->setCreator($config['fileSubject']);
        $objSpreadsheet->getProperties()->setLastModifiedBy($config['fileSubject']);
        $objSpreadsheet->getProperties()->setTitle($config['fileSubject']);
        $objSpreadsheet->getProperties()->setSubject($config['fileSubject']);
        $objSpreadsheet->getProperties()->setDescription('');
        ##################################################################################################################
        $cell_array=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK','CL','CM','CN','CO','CP','CQ','CR','CS','CT','CU','CV','CW','CX','CY','CZ','DA','DB','DC','DD','DE','DF','DG','DH','DI','DJ','DK','DL','DM','DN','DO','DP','DQ','DR','DS','DT','DU','DV','DW','DX','DY','DZ','EA','EB','EC','ED','EE','EF','EG','EH','EI','EJ','EK','EL','EM','EN','EO','EP','EQ','ER','ES','ET','EU','EV','EW','EX','EY','EZ','FA','FB','FC','FD','FE','FF','FG','FH','FI','FJ','FK','FL','FM','FN','FO','FP','FQ','FR','FS','FT','FU','FV','FW','FX','FY','FZ');

        ###############################################################
        $objSpreadsheet->setActiveSheetIndex(0);
        # Formatierung
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('A')->setWidth(0);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('B')->setWidth(12);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('C')->setWidth(12);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('D')->setWidth(8);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('E')->setWidth(12);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('F')->setWidth(2);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('G')->setWidth(12);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('H')->setWidth(12);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('I')->setWidth(5);
        $objSpreadsheet->getActiveSheet()->GetColumnDimension('J')->setWidth(14);
        #Rahmen 1
        $objSpreadsheet->getActiveSheet()->getStyle('B4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('C4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('D4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('H4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('I4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J4')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);


        ############################################################################

        $objSpreadsheet->getActiveSheet()->mergeCells('C4:E4');
        $objSpreadsheet->getActiveSheet()->mergeCells('C5:E5');
        $objSpreadsheet->getActiveSheet()->mergeCells('C6:E6');
        $objSpreadsheet->getActiveSheet()->mergeCells('C7:E7');
        $objSpreadsheet->getActiveSheet()->mergeCells('C8:E8');
        $objSpreadsheet->getActiveSheet()->mergeCells('C9:E9');

        $objSpreadsheet->getActiveSheet()->mergeCells('H4:J4');
        $objSpreadsheet->getActiveSheet()->mergeCells('H5:J5');
        $objSpreadsheet->getActiveSheet()->mergeCells('H6:J6');
        $objSpreadsheet->getActiveSheet()->mergeCells('H7:J7');
        $objSpreadsheet->getActiveSheet()->mergeCells('H8:J8');
        $objSpreadsheet->getActiveSheet()->mergeCells('H9:J9');

        $objSpreadsheet->getActiveSheet()->getStyle('J1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $objSpreadsheet->getActiveSheet()->getStyle('J2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

        $datum = date('d.m.Y',time());
        $titelmitdatum = 'Angebots-Nr.: ' . $zahl . '  /  Angebotsdatum: ' . $datum . '';
        $objSpreadsheet->getActiveSheet()->SetCellValue('J1', $titelmitdatum);
        $objSpreadsheet->getActiveSheet()->SetCellValue('J2', '(Das Angebot ist 10 Wochen gültig)');

        $objSpreadsheet->getActiveSheet()->SetCellValue('B4', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C4', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B5', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C5', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B6', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C6', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B7', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C7', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B8', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C8', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B9', '');
        $objSpreadsheet->getActiveSheet()->SetCellValue('C9', '');
        ##$objSpreadsheet->getActiveSheet()->SetCellValue('E6', 'Nr.:'.$load['fa_str_nr']);
        ##$objSpreadsheet->getActiveSheet()->SetCellValue('E7', 'PLZ:'.$load['fa_plz']);
        $objSpreadsheet->getActiveSheet()->SetCellValue('G4', 'Auftraggeber');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G5', 'Name');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G6', 'Str.');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G7', 'PLZ/Ort');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G8', 'Telefon');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G9', 'E-Mail');
        ##$objSpreadsheet->getActiveSheet()->SetCellValue('J6', 'Nr.:'.$load['ag_str_nr']);
        ##$objSpreadsheet->getActiveSheet()->SetCellValue('J7', 'PLZ:'.$load['ag_plz']);
        #Rahmen
        $objSpreadsheet->getActiveSheet()->getStyle('B4')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('B5')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('B6')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('B7')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('B8')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('B9')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G4')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G5')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G6')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G7')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G8')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G9')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);

        $objSpreadsheet->getActiveSheet()->getStyle('E4')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E5')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E6')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E7')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E8')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E9')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J4')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J5')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J6')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J7')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J8')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J9')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);

        $objSpreadsheet->getActiveSheet()->getStyle('B9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('C9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('D9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('H9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('I9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J9')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);

        #Values
        #$objSpreadsheet->getActiveSheet()->SetCellValue('C5', $load['fa_name']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('C6', $load['fa_str']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('C7', $load['fa_ort']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('C8', $load['fa_tel']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('C9', $load['fa_mail']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('H5', $load['ag_name']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('H6', $load['ag_str']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('H7', $load['ag_ort']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('H8', $load['ag_tel']);
        #$objSpreadsheet->getActiveSheet()->SetCellValue('H9', $load['ag_mail']);

        # Beginn Warenkorb
        ############################################################################
        $objSpreadsheet->getActiveSheet()->getStyle('B11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('C11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('D11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('E11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('F11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('G11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('H11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('I11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        $objSpreadsheet->getActiveSheet()->getStyle('J11')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);

        $objSpreadsheet->getActiveSheet()->SetCellValue('B12', 'Warenkorb');
        $objSpreadsheet->getActiveSheet()->SetCellValue('B14', 'Artikel');
        $objSpreadsheet->getActiveSheet()->SetCellValue('G14', 'Preis (netto)');
        $objSpreadsheet->getActiveSheet()->SetCellValue('H14', 'Menge');
        $objSpreadsheet->getActiveSheet()->SetCellValue('J14', 'Summe (netto)');
        ############################################################################
        # Produkte durchlaufen
        $aktpos = 15;
        $start_formel_pos = $aktpos;
        $endpos = 21;
        $i = 0;

//         for ($i = 0; $i <= count($_POST[anartikel][name]) - 1; $i++) {
        // loop over all items in the basket indexed by sorting text
        foreach ($itemArray as $sort => $actItemArray) {
            foreach ($actItemArray as $k1 => $actItem) {
                $row = $actItem['rec'];
                if (!$row) {	// avoid bug with missing row
                    continue;
                }

                $zusatz = '';
                // edit variants
                foreach ($row as $field => $value) {
                    if (strpos($field, 'edit_') === 0) {
                        $zusatz .= ' ' . $value;
                    }
                }

                // variants
                foreach ($variants as $variant) {
                    if ($row[$variant] != '') {
                        $zusatz .= ' ' . $row[$variant];
                    }
                }

                $field = 'B' . $aktpos;
                $a_von = 'B' . $aktpos;
                $a_bis = 'E' . $aktpos;
                $a_zusammen = $a_von . ':' . $a_bis;
                $objSpreadsheet->getActiveSheet()->getStyle($a_zusammen)->getAlignment()->setWrapText(true);
                $objSpreadsheet->getActiveSheet()->mergeCells($a_zusammen);
                $objSpreadsheet->getActiveSheet()->SetCellValue($field, $row['title']);
                $objSpreadsheet->getActiveSheet()->getRowDimension($aktpos)->setRowHeight(30);


                $field='G' . $aktpos;
                $objSpreadsheet->getActiveSheet()->SetCellValue($field, $row['pricenotax']);
                $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
                $field = 'H' . $aktpos;
                $objSpreadsheet->getActiveSheet()->SetCellValue($field, $actItem['count']);
                $end = $row['pricenotax'] * $actItem['count'];
                $field = 'J' . $aktpos;
                $formel = '=G' . $aktpos . '*H' . $aktpos;
                $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
                $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
                
                if (trim($zusatz) != '') {
                    $aktpos++;
                    $field = 'B' . $aktpos;
                    $objSpreadsheet->getActiveSheet()->SetCellValue($field, '(' . $zusatz . ')');
                }
                $aktpos++;
                $endpos = $endpos + 2;
                $i++;  // neu
            }
        }


        $endpos = $endpos - 5;

        for ($i = 11; $i <= $endpos + 2; $i++) {
            $field = 'B' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
            $field = 'J' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        }

        $pos = $endpos + 2;
        for ($i = 1; $i <= 9; $i++)
        {
            $pos = $endpos + 2;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
            $pos = $pos + 2;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
            $pos = $pos + 2;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
            
            $pos = $pos + 4;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        }
        #############################################################################
        #Endberechnung

        $start = $endpos;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Netto:');
        $field = 'J' . $start;
        $t = $start + 4;
        $fieldx = 'J' . $t;
        $xa = $endpos - 1;
        $formel = '=SUM(J' . $start_formel_pos . ':J' . $xa . ')';
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);	
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $fieldx = $field;
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'MwSt:');
        $field='J' . $start;
        $tdf = $start - 1;
        $field1 = 'J' . $tdf;
        $formel = '=' . $field1 . '/100 * ' . $taxrate;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Brutto inkl. ' . $taxrate . '% MwSt:');
            
        $field = 'J' . $start;
        $tdf = $start - 1;
        $field2 = 'J' . $tdf;
        $formel = '=' . $field1 . '+' . $field2;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            

        #############################################################################
        #Versand
        $start++;
        $start++;

        $ende = $start + 2;
        for ($i = $start; $i<= $ende; $i++) {
            $field = 'B' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
            $field = 'J' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        }
        $field='B'.$start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Versandart');
        $field = 'C' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $basketExtra['shipping.']['title']);
        $field = 'D' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, '');
        $field ='G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Netto:');


        $field = 'J' . $start;
        $versandpreis = $field;

        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $basketExtra['shipping.']['price']);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'MwSt:');
        $field = 'J' . $start;
        $tdf = $start - 1;
        $field1 = 'J' . $tdf;
        $formel = '=' . $field1 . '/100 * ' . $taxrate;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Brutto inkl. ' . $taxrate . '% MwSt:');
        $field = 'J' . $start;
        $tdf = $start - 1;
        $field2 = 'J' . $tdf;
        $formel = '=' . $field1 . '+' . $field2;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        ############################################################################



        #############################################################################
        #Rabatierung
        $start++;
        $start++;

        $ende = $start + 2;


        for ($i = 1; $i <= 9; $i++) {
            $pos = $start;
            $field = $cell_array[$i].$pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
        }

        for ($i = $start; $i <= $ende; $i++) {
            $field = 'B' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
            $field = 'J' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        }
        $field = 'B' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Rabattierung');
        $field = 'C' . $start;
        $fuerformel = $field;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, '0');
        $field='D' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, '%');
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Netto gesamt:');


        $field = 'J' . $start;
        $formel = '=(' . $fieldx . '+' . $versandpreis . ') / 100 * (100 - ' . $fuerformel . ')';
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'MwSt gesamt:');
        $field = 'J' . $start;
        $tdf = $start - 1;
        $field1 = 'J' . $tdf;
        $formel = '=' . $field1 . '/100 * ' . $taxrate;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        $start++;
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Brutto ges. inkl. ' . $taxrate . '% MwSt:');
        $field = 'J' . $start;
        $tdf = $start - 1;
        $field2 = 'J' . $tdf;
        $formel = '=' . $field1 . '+' . $field2;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, $formel);
        $objSpreadsheet->getActiveSheet()->GetStyle($field)->getNumberFormat()->setFormatCode('#,##0.00');
            
        ############################################################################
        $start = $start + 3;
        $ende = $start + 5;
        for ($i = $start - 1; $i <= $ende; $i++) {
            $field = 'B' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THICK);
            $field = 'J' . $i;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getRight()->setBorderStyle(Border::BORDER_THICK);
        }
        for ($i = 1; $i <= 9; $i++) {
            $pos = $start - 1;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getTop()->setBorderStyle(Border::BORDER_THICK);
            $pos = $pos + 6;
            $field = $cell_array[$i] . $pos;
            $objSpreadsheet->getActiveSheet()->getStyle($field)->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THICK);
        }

        $field = 'B' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Auszufüllen vom Auftraggeber:');
        $start = $start + 2;
        $field = 'B' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Ausführung in KW: ________________________________');
        $start = $start + 2;
        $field = 'B' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Unterschrift Kunde: ________________________________');
        $field = 'G' . $start;
        $objSpreadsheet->getActiveSheet()->SetCellValue($field, 'Auftrag erteilt am, Datum: _______________');


        ############################################################################

        $objDrawing = new Drawing();
        $objDrawing->setName('Logo');
        $objDrawing->setDescription('Logo');
        $objDrawing->setPath('./logo.jpg');
        $objDrawing->setCoordinates('B1');
        $objDrawing->setHeight(36);
        $objDrawing->setWorksheet($objSpreadsheet->getActiveSheet());

        ############################################################################

        $objWriter = new Xls($objSpreadsheet);

        $objWriter->save($filename);
        $file = $filename;
        $datei = basename($file);
        $size = filesize($file); 
        $ext = strtolower(substr($file, strlen($file) - 3, 3));
        ############################################################################

        $mailSubject = $config['mailSubject'] . ' - ' . $titelmitdatum;
        $mailFile = $config['mailFile'];
        $mailBody = '';

        if ($mailFile) {

            // mail file is fetched.
            $pathFilename = $GLOBALS['TSFE']->tmpl->getFileName($mailFile);
            $mailBody = file_get_contents($pathFilename);
        }

        $recipients = GeneralUtility::trimExplode(',', $config['mailTo']);
        foreach ($recipients as $recipient) {
            \JambageCom\Div2007\Utility\MailUtility::send(
                $recipient,
                $mailSubject,
                $tmp = '', // no plain text mail
                $mailBody,
                $config['mailFrom'],
                $config['mailName'],
                $filename, // attachment
                '', // cc
                '', // bcc
                '', // returnPath
                '', // replyTo
                QUOTATION_TT_PRODUCTS_EXT,
                'sendMail'
            );
        }

        #############################################################################
        header('Content-type: application/' . $ext); 
        header('Content-disposition: attachment; filename=' . $datei); 
        header('Content-Length: ' . $size); 
        header('Pragma: no-cache'); 
        header('Expires: 0'); 
        readfile($file);
    }
}

