<?php
require_once 'vendor/autoload.php';
require 'sections/SVPfunctions.php';
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Paper;


//The function will receive a JSON as an argument
$JSONdata = file_get_contents('testing/setup1.json');
  
$reportData = json_decode($JSONdata);
if($reportData->docRev < 10){$rNumber = "0".$reportData->docRev;}
else{$rNumber = $reportData->docRev;}
$revisionNumber="Rev ".$rNumber;

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

require_once 'sections/wordStyles.php';

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
$sectionTitlePg->addText($reportData->clientFullLegal, 'fsTitle', $styleParaTitle);
$sectionTitlePg->addText($reportData->Rig->name, 'fsSubTitle', $styleParaSubTitle);
$sectionTitlePg->addText($reportData->Well->name, 'fsSubTitle', $styleParaSubTitle);
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
generateHeader($header, $reportData, $revisionNumber);

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
$tableRevHistory->addRow($rsHeight);
$tableRevHistory->addCell(Converter::inchToTwip(.56), $vAlignCell)->addText("REV", 'fsSmallBold', 'psCenterTight');
$tableRevHistory->addCell(Converter::inchToTwip(.94), $vAlignCell)->addText("Date of Issue", 'fsSmallBold', 'psCenterTight');
$tableRevHistory->addCell(Converter::inchToTwip(2), $vAlignCell)->addText("Reason for Issue", 'fsSmallBold', 'psLeftTight');
$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Prepared", 'fsSmallBold', 'psCenterTight');
$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Checked", 'fsSmallBold', 'psCenterTight');
$tableRevHistory->addCell(Converter::inchToTwip(1.17), $vAlignCell)->addText("Approved", 'fsSmallBold', 'psCenterTight');
foreach($reportData->revisions as $revNokey => $value){
        $tableRevHistory->addRow($rsHeight);
        $tableRevHistory->addCell(null, $vAlignCell)->addText($revNokey, 'fsSize9', 'psCenterTight');
        $tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->date, 'fsSize9', 'psCenterTight');
        $tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->descShort, 'fsSize9', 'psLeftTight');
        $tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->preparedBy, 'fsSize9', 'psCenterTight');
        $tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->checkedBy, 'fsSize9', 'psCenterTight');
        $tableRevHistory->addCell(null, $vAlignCell)->addText($reportData->revisions->$revNokey->approvedBy, 'fsSize9', 'psCenterTight');
}
$tableRevHistory->addRow();
$tableRevHistory->addCell(Converter::inchToTwip(4.67),['gridSpan' => 4, 'valign' => 'center'])->addText("Initials:", 'fsSize9', 'psRightTight');
$tableRevHistory->addCell();
$tableRevHistory->addCell();

$sectionRevPage->addTextBreak(2, 'fsNormal');

//CHANGE DESCRIPTION
$sectionRevPage->addText('CHANGE DESCRIPTION', 'fsH1', $styleParaRevs);
$tableChangeDesc = $sectionRevPage->addTable('tsPlain');
$tableChangeDesc->addRow($rsHeight);
$tableChangeDesc->addCell(Converter::inchToTwip(.56), $vAlignCell)->addText("REV", 'fsSmallBold', 'psCenterTight');
$tableChangeDesc->addCell(Converter::inchToTwip(6.44), $vAlignCell)->addText("Change Description", 'fsSmallBold', 'psLeftTight');
foreach($reportData->revisions as $revNokey => $value){
        $tableChangeDesc->addRow($rsHeight);
        $tableChangeDesc->addCell(null,$vAlignCell)->addText($revNokey, 'fsSize9', 'psCenterTight');
        $tableChangeDesc->addCell(null,$vAlignCell)->addText($reportData->revisions->$revNokey->descLong, 'fsSize9', 'psLeftTight');
}

$sectionRevPage->addPageBreak();
$sectionRevPage->addText('TABLE OF CONTENTS', 'fsH1', $styleParaRevs);

//Main Section
$sectionMainContent=$phpWord->addSection($styleSectionDefault);
//HEADER
$headerMaincontent=$sectionMainContent->addHeader();
generateHeader($headerMaincontent, $reportData, $revisionNumber);


//FOOTER
$footerMain=$sectionMainContent->addFooter();
$footerMainTable = $footerMain->addTable('tsFooter');
$footerMainTable->addRow();
$footerMainTable->addCell(Converter::inchToTwip(3.03))->addText("www.otc-solutions.com", 'fsSize8', 'psLeftTight');
$footerMainTable->addCell(Converter::inchToTwip(4.15))->addText("PROPRIETARY/CONFIDENTIAL", 'fsSize8', 'psLeftTight');
$footerMainTable->addCell(Converter::inchToTwip(.33))->addPreserveText("{PAGE}", "fsSize8", 'psRightTight');

//Shear Calculation Letter
require_once 'sections/1-shearLetter.php';
$sectionMainContent->addPageBreak();

//Well Data & Equipment Information
$sectionMainContent->addTitle('Well Data and Equipment Information', 1);
if($reportData->Rig->location == "surface"){require_once 'sections/2-wdeiSurface.php';}
else{require_once 'sections/2-wdeiSubsea.php';}
$sectionMainContent->addPageBreak();

//Pipe Data Summary
require_once 'sections/3-pipeDataSummary.php';
$sectionMainContent->addPageBreak();

//Hydrostatic Effects
$sectionMainContent->addTitle('Hydrostatic Effects', 1);
require_once 'sections/4-hydrostaticEffects.php';
$sectionMainContent->addPageBreak();

//Pipe Detail Calculations
foreach($pipesArray as $pipeData){  //see 3-pipeDataSummary for $pipesArray
        $gradeName = $pipeData->strengthType === "grade" ? $pipeData->grade : "";
        $pipeDiameterStr = $pipeData->diameter;
        $pipeWallStr = $pipeData->wall;
	if($pipeData->pipeNo === 0){
		$sectionMainContent->addTitle('Test Pipe: ', 1);
        }else{$sectionMainContent->addTitle("Pipe No. {$pipeData->pipeNo}: {$gradeName}, {$pipeDiameterStr}\" OD, {$pipeWallStr}\" Wall", 1);}
	$tablePipeCalculations = $sectionMainContent->addTable('tsPlain');
	$tablePipeCalculations->addRow($rsHeight); 
	$tablePipeCalculations->addCell(Converter::inchtotwip(2.62),$csVbottomGrey)->addText("Definition", 'fsNormal', 'psCenterTable3');
	$tablePipeCalculations->addCell(Converter::inchtotwip(.62),$csVbottomGrey)->addText("Term", 'fsNormal', 'psCenterTable3');
	$tablePipeCalculations->addCell(Converter::inchtotwip(2.58),$csVbottomGrey)->addText("Equation", 'fsNormal', 'psCenterTable3');
	$tablePipeCalculations->addCell(Converter::inchtotwip(.71),$csVbottomGrey)->addText("Result", 'fsNormal', 'psCenterTable3');
	$tablePipeCalculations->addCell(Converter::inchtotwip(.46),$csVbottomGrey)->addText("Units", 'fsNormal', 'psCenterTable3');
	$tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Evaluation Yield Strength", 'fsNormal', 'psLeftTable3');
        addTermCell($tablePipeCalculations,"S","y");
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->evalStrength, 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText('psi', 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Outside Diameter of Pipe", 'fsNormal', 'psLeftTable3');
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("OD", 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeDiameterStr, 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText('in', 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addRow($rsHeight);
        if($pipeData->selectedMethod =="Cameron"){
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Pipe weight", 'fsNormal', 'psLeftTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("ppf", 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null,$csGrey);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->ppf, 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('ppf', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Shearing Constant", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"C","3");
            $tablePipeCalculations->addCell(null,$csGrey);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->C3, 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null,$csGrey);
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Calculated Force to Shear at Surface", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"F","CAM");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("C3 x Sy x ppf");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->CamForce, 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('kips', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Pressure Required at Sy", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"P","CAM");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Fcam / Ac x 1000 lb/kip");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('xxx', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('psi', 'fsNormal', 'psCenterTable3');
            
        }else if($pipeData->selectedMethod =="West"){
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Nominal Thickness of Pipe", 'fsNormal', 'psLeftTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Wall", 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null,$csGrey);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeWallStr, 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('in', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Pipe Cross-Sectional Area", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"A","g");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('OD2 - (OD - 2 x Wall)^2 x pi / 4', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->area, 'fsNormal', 'psCenterTable3');
            addSqInchCell($tablePipeCalculations);
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Calculated Force to Shear at Surface", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"F","YS");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->WestDef);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->WestForce, 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('kips', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addRow($rsHeight);
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Pressure Required at Sy", 'fsNormal', 'psLeftTable3');
            addTermCell($tablePipeCalculations,"P","WEST");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Fwest / Ac x 1000 lb/kip");
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('xxx', 'fsNormal', 'psCenterTable3');
            $tablePipeCalculations->addCell(null, $vAlignCell)->addText('psi', 'fsNormal', 'psCenterTable3');
        }
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Total Calculated Shear Pressure adjusted for operator effects", 'fsNormal', 'psLeftTable3');
        addTermCell($tablePipeCalculations,"P","T");
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("P + Padj");
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->OperatingPressure, 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText('psi', 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Operating Pressure to Shear and Seal", 'fsNormal', 'psLeftTable3');
        addTermCell($tablePipeCalculations,"P","OP");
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("MAX of Pys and Pseal");
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText($pipeData->OperatingPressure, 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText('psi', 'fsNormal', 'psCenterTable3');
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Is the calculated pressure less than or equal to 90% of the max operating pressure of the BOP?", 'fsNormal', 'psLeftTable3');
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csSpan2Good);
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Is the calculated pressure less than or equal to the supply pressure?", 'fsNormal', 'psLeftTable3');
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csSpan2Good);
        $tablePipeCalculations->addRow($rsHeight);
        $tablePipeCalculations->addCell(null, $vAlignCell)->addText("Is the calculated pressure less than or equal to the calculated pressure of the test pipe?", 'fsNormal', 'psLeftTable3');
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csGrey);
        $tablePipeCalculations->addCell(null,$csSpan2Good);
	$sectionMainContent->addPageBreak();
}

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
?>