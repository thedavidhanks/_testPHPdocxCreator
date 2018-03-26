<?php
require_once 'vendor/autoload.php';
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Paper;


function reportDocxCreator($data) {
    
    $reportData = json_decode($data);
    $revisionNumber="Rev ".$reportData->docRev;
    
    // Creating the new document...
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $phpWord->getCompatibility()->setOoxmlVersion(15);

    $properties = $phpWord->getDocInfo();
    $properties->setCreator($reportData->author);
    $properties->setCompany($reportData->company);
    $properties->setTitle($reportData->reportType);
    $properties->setDescription($reportData->docDescription);
    $properties->setLastModifiedBy($reportData->lastModifier);
    $properties->setCreated(mktime(0, 0, 0, 3, 12, 2014));

    //
    //Styles
    //

    //Fonts
    $phpWord->addFontStyle('fsNormal', array(
            'name' => 'Arial',
            'size' => 10
    ));
    $phpWord->addFontStyle('fsTitle', array(
            'name' => 'Arial',
            'size' => 20,
            'bold' => true
    ));
	$phpWord->addFontStyle('fsSize8', array(
		'basedOn' => 'fsNormal',
		'size' => 8,
	));
	$phpWord->addFontStyle('fsSize9', array(
		'basedOn' => 'fsNormal',
		'size' => 9,
	));
    $phpWord->addFontStyle('fsSmallBold', array(
		'basedOn' => 'fsSize9',
		'bold' => true
	));
    $phpWord->addFontStyle('fsHeader', array(
        'name' => 'Arial',
        'size' => 8
    ));
    $phpWord->addFontStyle('fsSubTitle', array(
            'name' => 'Arial',
            'size' => 16,
            'bold' => true,
            'spaceAfter' => 0,
            'spaceBefore' => 0,
			'color' => '3A3B3D'
    ));
	$fsH1config = array(
        'name' => 'Arial',
        'size' => 14,
        'bold' => true,
        'color' => '8B2027'
    );
    $phpWord->addFontStyle('fsH1', $fsH1config);

    //Paragraphs
    $phpWord->addParagraphStyle('Normal', array(
            'alignment' => 'justified'
    ));
	$psCenteredConfig = array('alignment' => 'center');
	$phpWord->addParagraphStyle('psCentered',$psCenteredConfig );
	$phpWord->addParagraphStyle('psCenterTight', array(
            'alignment' => 'center',
			'spaceAfter' => 0,
            'spaceBefore' => 0
    ));
	$phpWord->addParagraphStyle('psLeftTight', array(
            'alignment' => 'left',
			'spaceAfter' => 0,
            'spaceBefore' => 0
    ));
	$phpWord->addParagraphStyle('psRightTight', array(
            'alignment' => 'right',
			'spaceAfter' => 0,
            'spaceBefore' => 0
    ));

    $styleParaTitle = array(
            'alignment' => 'right',
            'spaceAfter' => 2,
            'spaceBefore' => 2
    );

    $styleParaSubTitle = array(
            'alignment' => 'right',
            'spaceAfter' => 0,
            'spaceBefore' => 0
    );
    $styleParaRevs = array(
            'basedOn' => 'Normal',
            'alignment' => 'left'
    );

    //Sections
	$styleSectionDefault = array(
            'paperSize' => 'Letter',
            'marginTop' => Converter::inchToTwip(1.15),
            'marginLeft' => Converter::inchToTwip(.75),
            'marginRight' => Converter::inchToTwip(.75),
            'marginBottom' => Converter::inchToTwip(.5),
            'headerHeight' => Converter::inchToTwip(.3),
            'footerHeight' => Converter::inchToTwip(.3)
    );
    $styleSectionTitle = array(
            'paperSize' => 'Letter',
            'marginTop' => Converter::inchToTwip(.5),
            'marginLeft' => Converter::inchToTwip(.5),
            'marginRight' => Converter::inchToTwip(.5),
            'marginBottom' => Converter::inchToTwip(.3),
            'headerHeight' => Converter::inchToTwip(.3),
            'footerHeight' => Converter::inchToTwip(.4),
			'borderSize' => 40,
			'borderColor' => '9E3234'
    );
/*Merging styles
    $styleSectionTitle = array_merge($styleSectionDefault, [
        'borderSize' => 40,
        'borderColor' => '9E3234'
        ]
    );
 * 
 */
	
	//tables
	$phpWord->addTableStyle('tsPlain', array(
		'borderColor' => '000000',
		'borderSize' => 1,
		'cellMarginTop' => 0,
		'cellMarginBottom' => 0,
		'cellMarginLeft' => Converter::inchToTwip(0.08),
		'cellMarginRight' => Converter::inchToTwip(0.08)
	));
	$phpWord->addTableStyle('tsFooter', array(
		'cellMargin' => 0
	));
	
	//Titles
	$phpWord->addTitleStyle(1,$fsH1config, $psCenteredConfig);
	
	$vAlignCell = array('valign' => 'center');
    /*
     * BEGIN DOC CREATION
     */
    
    $sectionTitlePg = $phpWord->addSection($styleSectionTitle);

    //Add the coverpage image
    $width = 525;
    $height = 1302/1605*$width;
    $sectionTitlePg->addImage(
        'img/drivenByRealExp.jpg',
        array(
            'positioning'      => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posHorizontal'    => \PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_CENTER,
            'posHorizontalRel' => \PhpOffice\PhpWord\Style\Image::POSITION_RELATIVE_TO_PAGE,
            'marginTop'          => 0.75,
            'wrappingStyle'      => 'square',
            'width' => $width,//7.3" or 1605 pixels
            'height' => $height//5.92" or 1302 pixels
        )
    );

    //Add the coverpage Title, Rig, Client, Doc#, Rev, Release Date
    $sectionTitlePg->addText($reportData->reportType, 'fsTitle', $styleParaTitle);
    $sectionTitlePg->addText($reportData->reportSubType, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addTextBreak(1);
    $sectionTitlePg->addText('[Client Legal Name]', 'fsTitle', $styleParaTitle);
    $sectionTitlePg->addText('[Rig]', 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText('[Well Name]', 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addTextBreak(1);
    $sectionTitlePg->addText($reportData->docNumber, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText($revisionNumber, 'fsSubTitle', $styleParaSubTitle);
    $sectionTitlePg->addText('March 22, 2018', 'fsSubTitle', $styleParaSubTitle);

    //add the Coverpage OTCSolution image
    $sectionTitlePg->addImage(
        'img/OTCsolutions_cover.jpg',
        array(
            'height' => Converter::inchToPixel(.977),
            'wrappingStyle'      => 'square',
            'positioning' => \PhpOffice\PhpWord\Style\Image::POSITION_ABSOLUTE,
            'posVertical'      => \PhpOffice\PhpWord\Style\Image::POSITION_VERTICAL_BOTTOM,
            'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT,
            'marginLeft'       => .5,
            'marginBottom'        => 0
        )
    );

    //
    //REVISION PAGE
    $sectionRevPage=$phpWord->addSection($styleSectionDefault);
    //HEADER
	$header=$sectionRevPage->addHeader();
	$headerTable = $header->addTable('tsFooter');
	$headerTable->addRow();
	$headerTable->addCell(null)->addImage(
		'img/OTCsolutions_header.jpg',
        array(
            'height' => Converter::inchToPixel(.5625),
            'wrappingStyle'=> 'square',
			'positioning' => 'relative',
			'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT
        )
    );
	$headerRtCell = $headerTable->addCell(Converter::inchToTwip(4.34));
	$headerRtCell->addText($reportData->reportType, 'fsHeader', 'psRightTight');
	$headerRtCell->addText($reportData->docNumber, 'fsHeader', 'psRightTight');
	$headerRtCell->addText($revisionNumber, 'fsHeader', 'psRightTight');
	
	//FOOTER
	$footer=$sectionRevPage->addFooter();
	$footerTable = $footer->addTable('tsFooter');
	$footerTable->addRow();
	$footerTable->addCell(Converter::inchToTwip(3.03))->addText("www.otc-solutions.com", 'fsSize8', 'psLeftTight');
	$footerTable->addCell(Converter::inchToTwip(4.15))->addText("PROPRIETARY/CONFIDENTIAL", 'fsSize8', 'psLeftTight');
	//$footerTable->addCell(Converter::inchToTwip(.33))->addPreserveText("{PAGE}", "fsSize8", 'psRightTight');
    
	//REVISION HISTORY
    $sectionRevPage->addTextBreak(1, 'fsNormal');
    $sectionRevPage->addText('REVISION HISTORY', 'fsH1', $styleParaRevs);
	$tableRevHistory = $sectionRevPage->addTable('tsPlain');
	$tableRevHistory->addRow(Converter::inchToTwip(.3));
	$tableRevHistory->addCell(Converter::inchToTwip(.56), $vAlignCell)->addText("REV", 'fsSmallBold', 'psCenterTight');
	$tableRevHistory->addCell(Converter::inchToTwip(.94), $vAlignCell)->addText("Date of Issue", 'fsSmallBold', 'psCenterTight');
	$tableRevHistory->addCell(Converter::inchToTwip(2), $vAlignCell)->addText("Reason for Issue", 'fsSmallBold', 'psLeftTight');
	$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Prepared", 'fsSmallBold', 'psCenterTight');
	$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Checked", 'fsSmallBold', 'psCenterTight');
	$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Approved", 'fsSmallBold', 'psCenterTight');
	foreach($reportData->revisions as $revNokey => $value){
		$tableRevHistory->addRow(Converter::inchToTwip(.3));
		$tableRevHistory->addCell(null, $vAlignCell)->addText($revNokey, 'fsSize9', 'psCenterTight');
		$tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->date, 'fsSize9', 'psCenterTight');
		$tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->descShort, 'fsSize9', 'psLeftTight');
		$tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->preparedBy, 'fsSize9', 'psCenterTight');
		$tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->checkedBy, 'fsSize9', 'psCenterTight');
		$tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->approvedBy, 'fsSize9', 'psCenterTight');
	}
	$tableRevHistory->addRow();
	$tableRevHistory->addCell(Converter::inchToTwip(4.67),array('gridSpan' => 4))->addText("Initials:", 'fsSize9', 'psRightTight');
	$tableRevHistory->addCell();
	$tableRevHistory->addCell();
	
	$sectionRevPage->addTextBreak(2, 'fsNormal');
	
	//CHANGE DESCRIPTION
    $sectionRevPage->addText('CHANGE DESCRIPTION', 'fsH1', $styleParaRevs);
    $tableChangeDesc = $sectionRevPage->addTable('tsPlain');
	$tableChangeDesc->addRow(Converter::inchToTwip(.3));
	$tableChangeDesc->addCell(Converter::inchToTwip(.56), $vAlignCell)->addText("REV", 'fsSmallBold', 'psCenterTight');
	$tableChangeDesc->addCell(Converter::inchToTwip(6.44), $vAlignCell)->addText("Change Description", 'fsSmallBold', 'psLeftTight');
    foreach($reportData->revisions as $revNokey => $value){
		$tableChangeDesc->addRow(Converter::inchToTwip(.3));
		$tableChangeDesc->addCell(null,$vAlignCell)->addText($revNokey, 'fsSize9', 'psCenterTight');
		$tableChangeDesc->addCell(null,$vAlignCell)->addText($reportData->revisions->$revNokey->descLong, 'fsSize9', 'psLeftTight');
	}
	
	$sectionRevPage->addPageBreak();
	$sectionRevPage->addText('TABLE OF CONTENTS', 'fsH1', $styleParaRevs);
	
	//Main Section
	$sectionMainContent=$phpWord->addSection($styleSectionDefault);
	//HEADER
	$headerMain=$sectionMainContent->addHeader();
	$headerMainTable = $headerMain->addTable('tsFooter');
	$headerMainTable->addRow();
	$headerMainTable->addCell(null)->addImage(
		'img/OTCsolutions_header.jpg',
        array(
            'height' => Converter::inchToPixel(.5625),
            'wrappingStyle'=> 'square',
			'positioning' => 'relative',
			'posHorizontal' =>\PhpOffice\PhpWord\Style\Image::POSITION_HORIZONTAL_LEFT
        )
    );
	$headerMainRtCell = $headerMainTable->addCell(Converter::inchToTwip(4.34));
	$headerMainRtCell->addText($reportData->reportType, 'fsHeader', 'psRightTight');
	$headerMainRtCell->addText($reportData->docNumber, 'fsHeader', 'psRightTight');
	$headerMainRtCell->addText($revisionNumber, 'fsHeader', 'psRightTight');
	//FOOTER
	$footerMain=$sectionMainContent->addFooter();
	$footerMainTable = $footerMain->addTable('tsFooter');
	$footerMainTable->addRow();
	$footerMainTable->addCell(Converter::inchToTwip(3.03))->addText("www.otc-solutions.com", 'fsSize8', 'psLeftTight');
	$footerMainTable->addCell(Converter::inchToTwip(4.15))->addText("PROPRIETARY/CONFIDENTIAL", 'fsSize8', 'psLeftTight');
	$footerMainTable->addCell(Converter::inchToTwip(.33))->addPreserveText("{PAGE}", "fsSize8", 'psRightTight');
	
	//Shear Calculation Letter
	$sectionMainContent->addTitle('Shear Calculation Letter', 1);
	$sectionMainContent->addPageBreak();
	
	//Well Data & Equipment Information
	$sectionMainContent->addTitle('Well Data and Equipment Information', 1);
	$sectionMainContent->addPageBreak();
	
	//Pipe Data Summary
	$sectionMainContent->addTitle('Pipe Data Summary', 1);
	$sectionMainContent->addPageBreak();
	
	//Hydrostatic Effects
	$sectionMainContent->addTitle('Hydrostatic Effects', 1);
	$sectionMainContent->addPageBreak();
	
	//Pipe Detail Calculations
	$sectionMainContent->addTitle('Pipe No. X', 1);
	$sectionMainContent->addPageBreak();
	
	//References
	$sectionMainContent->addTitle('References', 1);
	$sectionMainContent->addPageBreak();
	
	//Appendix A
	$sectionMainContent->addTitle('Appendix A - Shear Test Report', 1);
	
	//Create the Table of Contents after all the pages have been created.
	$sectionRevPage->addTOC();
    // Saving the document as OOXML file...
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save('output/helloWorld.docx');
}

//The function will receive a JSON as an argument
$createDate = mktime(0, 0, 0, 3, 12, 2014);
$jsonArgument = <<<EOT
    {     
        "author": "David Hanks",
        "company" : "OTC Solutions",
        "docDescription" : "",
        "lastModifier" : "",
        "reportType" : "Shear Verification Package",
        "reportSubType" : "30 CFR §250.732 (a),(b)",
        "clientFullLegal" : "",
        "clientShortend" : "OES",
        "rig" : "",
        "wellName" : "",
        "docNumber": "SBS1_COX_180307_SVP",
        "docRev": 0,
        "revisions" : {
            "00" : {
                "date" : "2016-01-01",
                "descShort" : "Initial Release",
                "descLong" : "Initial Release",
                "preparedBy" : "David Hanks",
                "checkedBy" : "Steve Ronan",
                "approvedBy" : "David Hanks"
            },
            "01" : {
                    "date" : "2016-02-20",
                    "descShort" : "Fixed it",
                    "descLong" : "Something was really, really wrong.  It had to go.  It's all better now.",
                    "preparedBy" : "David Hanks",
                    "checkedBy" : "Steve Ronan",
                    "approvedBy" : "David Hanks"
            },
			"02": {
                    "date" : "2016-02-20",
                    "descShort" : "Fixed it",
                    "descLong" : "Something was kinda wrong.  We tweaked it a bit.  It's all better now.",
                    "preparedBy" : "This Guy",
                    "checkedBy" : "That Guy",
                    "approvedBy" : "Boss Man"
            }
        }
    }
EOT;
//echo $jsonArgument;
//$reportData = json_decode($jsonArgument);

reportDocxCreator($jsonArgument);
?>