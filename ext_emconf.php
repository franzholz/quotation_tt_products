<?php

########################################################################
# Extension Manager/Repository config file for ext "quotation_tt_products".
########################################################################

$EM_CONF[$_EXTKEY] = array(
    'title' => 'Quotation for Shop System',
    'description' => 'Export of the current shop basket into an Excel file in the form of a quotation.',
    'category' => 'misc',
    'version' => '0.2.0',
    'state' => 'stable',
    'uploadfolder' => 0,
	'createDirs' => 'fileadmin/data/quotation',
    'clearcacheonload' => 1,
    'author' => 'Franz Holzinger',
    'author_email' => 'franz@ttproducts.de',
    'author_company' => 'jambage.com',
    'constraints' => array(
        'depends' => array(
            'php' => '5.5.0-7.3.99',
            'typo3' => '6.2.0-9.5.99'
        ),
        'conflicts' => array(
        ),
        'suggests' => array(
            'base_excel' => '0.0.1-0.0.0'
        ),
    ),
);
